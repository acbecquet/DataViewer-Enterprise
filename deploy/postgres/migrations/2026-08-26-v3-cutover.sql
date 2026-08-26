-- TPM template v3, Phase 3d - the standard-metric cutover.
--
-- Plan:  docs/superpowers/plans/2026-08-26-tpm-v3-phase3d-standard-cutover-plan.md
-- Index: docs/superpowers/plans/2026-07-30-tpm-v3-phase3-INDEX.md (D-3d-1/2/6/7)
--
-- WHAT THIS FILE DOES
--   1. bump_version(): drops updated_by from the no-op comparison (hazard H23)
--      and restores it in the no-op branch, so a suppressed no-op leaves the
--      row byte-identical. Latest CREATE OR REPLACE wins; init.sql and the
--      2026-07-15 / 2026-07-31 copies are superseded, not edited (precedent).
--   2. dve_commit_measurement(): the per-cell commit for the long format.
--      ONE signature - never overload it (the dve_commit_cell 5/6-arg
--      ambiguity is the cautionary tale, index D6).
--   3. dve_cutover_to_long_format(): idempotently migrates (re-running the
--      3b function), verifies parity AND the per-column sparse rule (the 3b
--      "parity alone is not sufficient" lesson, institutionalized), renames
--      data_rows -> data_rows_pre_v3 and samples -> samples_core, creates
--      the name-holder views, revokes, stamps schema_meta.
--   4. Calls the cutover at the tail. In the test container that runs against
--      an empty database (trivial). On the NAS this file is applied BY HAND
--      as a supervised v3.0.0 runbook step (D8/D-3d-2) - applying it IS the
--      cutover, so the runbook orders it LAST:
--        a. all clients closed; pg_dump backup taken (D7)
--        b. 2026-08-03-oil-units-to-grams.sql        (units first, D10)
--        c. 2026-07-31-v3-long-format.sql            (tables/views/function)
--        d. optional: SELECT * FROM dve_migrate_to_long_format(); + inspect
--        e. THIS FILE (re-migrates idempotently, verifies, cuts over)
--        f. install v3.0.0 clients
--   The app NEVER calls any of this (D-3d-2); it probes relkind at connect.
--
-- ROLLBACK: restore the step-a backup. data_rows_pre_v3 / samples_core keep
-- the full pre-cutover wide data as an on-database pre-image, but they are a
-- forensic/translation aid (v1 pending_edits, H9), NOT a rollback path -
-- writes made after cutover exist only in the long tables.

BEGIN;

-- ============================================================================
-- 1. bump_version - H23
-- ============================================================================
-- Copied from 2026-07-31-v3-long-format.sql with EXACTLY two changes, both in
-- the no-op branch: updated_by joins the ignore-list of the comparison, and
-- the branch restores OLD.updated_by. Without this, a whole-file save by user
-- B over rows last written by user A differs on updated_by alone and bumps
-- every otherwise-unchanged row - the exact churn D4 exists to prevent,
-- failing precisely in the multi-user case it exists for.
CREATE OR REPLACE FUNCTION bump_version() RETURNS TRIGGER AS $$
BEGIN
  -- No-op UPDATE on a scoped table (nothing material changed): keep the row
  -- inert so it doesn't churn version/updated_at or (via notify_row_change)
  -- storm live clients. Scoped to the whole-row session tables where the
  -- storm was observed plus the two v3 long tables (index decision D4); the
  -- TPM files/tests/samples hierarchy KEEPS its always-bump behavior.
  -- v3 Phase 3d / H23: updated_by no longer breaks the no-op detection - a
  -- whole-file save by another user over identical values must not bump ~13k
  -- measurement versions - and a suppressed no-op restores OLD.updated_by so
  -- the row stays byte-identical.
  IF TG_OP = 'UPDATE'
     AND TG_TABLE_NAME IN ('sensory_sessions', 'detailed_sensory_sessions',
                           'measurements', 'sample_headers')
     AND (to_jsonb(NEW) - 'version' - 'updated_at' - 'app_version' - 'updated_by')
         IS NOT DISTINCT FROM
         (to_jsonb(OLD) - 'version' - 'updated_at' - 'app_version' - 'updated_by')
  THEN
    NEW.version    := OLD.version;      -- undo any client-supplied version = version + 1
    NEW.updated_at := OLD.updated_at;
    NEW.updated_by := OLD.updated_by;   -- H23: a no-op leaves the row byte-identical
    RETURN NEW;
  END IF;
  -- Real change: refuse client-side version manipulation; we always +1.
  NEW.version    := OLD.version + 1;
  NEW.updated_at := now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- ============================================================================
-- 2. dve_commit_measurement - the long-format per-cell commit (H4 re-point)
-- ============================================================================
-- Identity is (sample_id, metric key, sort_order) - the same natural key the
-- whole-file save writes, so a per-cell commit and a whole-file save can never
-- disagree about which row they mean (the data_rows_v.id surrogate is banned
-- from write-back by its own column comment). Value routing comes from
-- metric_defs.value_type, the same single authority MetricDefCache uses.
-- Unknown key -> FALSE (mirrors dve_commit_cell's unknown-column contract;
-- the caller logs, the whole-file save still carries the edit).
CREATE OR REPLACE FUNCTION dve_commit_measurement(
    p_sample_id  BIGINT,
    p_key        TEXT,
    p_sort_order INTEGER,
    p_value      TEXT,
    p_uuid       TEXT
) RETURNS BOOLEAN AS $$
DECLARE
    v_def_id BIGINT;
    v_is_num BOOLEAN;
BEGIN
    SELECT md.id, md.value_type = 'number' INTO v_def_id, v_is_num
      FROM metric_defs md
     WHERE md.kind = 'metric' AND md.key = p_key;
    IF NOT FOUND THEN
        RETURN FALSE;
    END IF;
    PERFORM set_config('dve.live_column', p_key,   true);
    PERFORM set_config('dve.live_value',  p_value, true);
    IF v_is_num THEN
        INSERT INTO measurements (sample_id, metric_id, sort_order, value_num, updated_by)
        VALUES (p_sample_id, v_def_id, p_sort_order, p_value::DOUBLE PRECISION, p_uuid)
        ON CONFLICT (sample_id, metric_id, sort_order) DO UPDATE
           SET value_num = EXCLUDED.value_num, value_text = NULL,
               updated_by = EXCLUDED.updated_by;
    ELSE
        INSERT INTO measurements (sample_id, metric_id, sort_order, value_text, updated_by)
        VALUES (p_sample_id, v_def_id, p_sort_order, p_value, p_uuid)
        ON CONFLICT (sample_id, metric_id, sort_order) DO UPDATE
           SET value_text = EXCLUDED.value_text, value_num = NULL,
               updated_by = EXCLUDED.updated_by;
    END IF;
    RETURN TRUE;
END;
$$ LANGUAGE plpgsql;

COMMENT ON FUNCTION dve_commit_measurement(BIGINT, TEXT, INTEGER, TEXT, TEXT) IS
    'Per-cell commit for the v3 long format (Phase 3d, hazard H4). Keyed by the '
    'natural (sample_id, metric key, sort_order) identity - NEVER by data_rows_v.id, '
    'which is a synthetic surrogate. ONE signature by decision; do not overload.';

-- ============================================================================
-- 3. dve_cutover_to_long_format
-- ============================================================================
CREATE OR REPLACE FUNCTION dve_cutover_to_long_format(
    OUT already_cut             BOOLEAN,
    OUT measurements_inserted   BIGINT,
    OUT sample_headers_inserted BIGINT
) AS $$
DECLARE
    -- Pinned column lists, same values and same rationale as
    -- dve_migrate_to_long_format (a registry-driven list could silently start
    -- capturing a non-value column). If you change one, change both.
    c_metric_cols CONSTANT TEXT[] := ARRAY[
        'puffs', 'before_weight', 'after_weight', 'draw_pressure', 'resistance',
        'smell', 'clog', 'notes', 'tpm', 'tpm_power_density', 'variation_tpm',
        'oil_consumed', 'puffing_regime'];
    c_header_cols CONSTANT TEXT[] := ARRAY[
        'sample_name', 'sample_id', 'date', 'tester', 'media', 'viscosity',
        'resistance', 'voltage', 'power', 'heating_technology', 'puffing_regime',
        'initial_oil_mass', 'average_tpm', 'stddev_tpm', 'avg_power_density',
        'efficiency_percent', 'total_oil_consumed', 'total_puffs',
        'normalized_tpm', 'burn_status', 'clog_status', 'leak_status'];
    v_kind CHAR;
    v_col  TEXT;
    v_wide BIGINT;
    v_long BIGINT;
BEGIN
    already_cut := FALSE;
    SET LOCAL statement_timeout = 0;
    PERFORM pg_advisory_xact_lock(hashtext('dve_v3_cutover'));

    -- Idempotency: if data_rows already resolves to a view, the cutover has
    -- run. Supervised runbooks benefit from a safe re-run.
    SELECT c.relkind INTO v_kind FROM pg_class c
     WHERE c.oid = to_regclass('data_rows');
    IF v_kind = 'v' THEN
        RAISE NOTICE 'dve_cutover_to_long_format: data_rows is already a view - nothing to do.';
        already_cut := TRUE;
        measurements_inserted   := 0;
        sample_headers_inserted := 0;
        RETURN;
    END IF;
    IF v_kind IS NULL THEN
        RAISE EXCEPTION 'dve_cutover_to_long_format: no relation named data_rows. '
            'Apply init.sql / 2026-07-31-v3-long-format.sql first.';
    END IF;

    -- The 3b function must exist (H20 probe) and is idempotent, so re-running
    -- it here inserts only what a prior manual run left out.
    IF to_regprocedure('dve_migrate_to_long_format()') IS NULL THEN
        RAISE EXCEPTION 'dve_cutover_to_long_format: dve_migrate_to_long_format() is missing. '
            'Apply 2026-07-31-v3-long-format.sql first.';
    END IF;
    SELECT m.measurements_inserted, m.sample_headers_inserted
      INTO measurements_inserted, sample_headers_inserted
      FROM dve_migrate_to_long_format() m;

    -- ── Verification gate 1: row-count parity ───────────────────────────────
    SELECT count(*) INTO v_wide FROM data_rows;
    SELECT count(*) INTO v_long FROM data_rows_v;
    IF v_wide <> v_long THEN
        RAISE EXCEPTION 'cutover verify: data_rows has % rows but data_rows_v has % - refusing to cut over.',
            v_wide, v_long;
    END IF;
    SELECT count(*) INTO v_wide FROM samples;
    SELECT count(*) INTO v_long FROM samples_v;
    IF v_wide <> v_long THEN
        RAISE EXCEPTION 'cutover verify: samples has % rows but samples_v has % - refusing to cut over.',
            v_wide, v_long;
    END IF;

    -- ── Verification gate 2: the per-column sparse rule ─────────────────────
    -- The 3b rehearsal proved value-parity alone is blind to a dense
    -- migration (it pivots NULL back to NULL). The honest check is that each
    -- wide column's NOT NULL count equals its measurement/header count.
    FOREACH v_col IN ARRAY c_metric_cols LOOP
        EXECUTE format('SELECT count(*) FROM data_rows WHERE %I IS NOT NULL', v_col)
           INTO v_wide;
        SELECT count(*) INTO v_long
          FROM measurements m JOIN metric_defs md ON md.id = m.metric_id
         WHERE md.kind = 'metric' AND md.key = v_col;
        IF v_wide <> v_long THEN
            RAISE EXCEPTION 'cutover verify (sparse rule): data_rows.% has % non-NULL values but measurements holds % rows for that key.',
                v_col, v_wide, v_long;
        END IF;
    END LOOP;
    FOREACH v_col IN ARRAY c_header_cols LOOP
        EXECUTE format('SELECT count(*) FROM samples WHERE %I IS NOT NULL', v_col)
           INTO v_wide;
        SELECT count(*) INTO v_long
          FROM sample_headers sh JOIN metric_defs md ON md.id = sh.field_id
         WHERE md.kind = 'header' AND md.key = v_col;
        IF v_wide <> v_long THEN
            RAISE EXCEPTION 'cutover verify (sparse rule): samples.% has % non-NULL values but sample_headers holds % rows for that key.',
                v_col, v_wide, v_long;
        END IF;
    END LOOP;

    -- ── The cut ─────────────────────────────────────────────────────────────
    -- Views bind base relations by OID (verified live 2026-08-03): after the
    -- renames, data_rows_v and samples_v keep reading the renamed tables, and
    -- the old names are free for the name-holder views.
    ALTER TABLE data_rows RENAME TO data_rows_pre_v3;
    ALTER TABLE samples   RENAME TO samples_core;

    -- data_rows: pure replacement. SELECT * pins data_rows_v's 19 columns in
    -- canonical order at creation time.
    CREATE VIEW data_rows AS SELECT * FROM data_rows_v;

    -- samples: identity/audit from the surviving core table, the 22 values
    -- from the (structurally total, H21) samples_v. Column order matches the
    -- wide table's canonical 28-column projection (OfflineSnapshot regen).
    CREATE VIEW samples AS
    SELECT c.id, c.test_id, c.sort_order,
           v.sample_name, v.sample_id, v.date, v.tester, v.media, v.viscosity,
           v.resistance, v.voltage, v.power, v.heating_technology,
           v.puffing_regime, v.initial_oil_mass, v.average_tpm, v.stddev_tpm,
           v.avg_power_density, v.efficiency_percent, v.total_oil_consumed,
           v.total_puffs, v.normalized_tpm, v.burn_status, v.clog_status,
           v.leak_status,
           c.updated_at, c.updated_by, c.version
    FROM samples_core c
    LEFT JOIN samples_v v ON v.id = c.id;

    COMMENT ON VIEW data_rows IS
        'v3 Phase 3d name-holder: the wide data_rows table is renamed to data_rows_pre_v3 '
        '(frozen pre-image; v1 pending_edits translation) and this view serves every '
        'remaining wide reader. Read BY NAME, never by position. id is SYNTHETIC '
        '(MIN measurement id) - never a write-back identity; writes go to measurements '
        'keyed (sample_id, key, sort_order).';
    COMMENT ON VIEW samples IS
        'v3 Phase 3d name-holder: identity/audit columns come from samples_core (the '
        'surviving real table; id IS real and writable-back), the 22 value columns from '
        'samples_v/sample_headers. Wide writes go to samples_core (narrow) + sample_headers.';

    -- H16: new relations, explicit privilege posture.
    REVOKE ALL ON data_rows, samples FROM PUBLIC;
    IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'sensory_web') THEN
        REVOKE ALL ON data_rows, samples FROM sensory_web;
    END IF;

    INSERT INTO schema_meta (key, value)
    VALUES ('v3_long_format_cutover', now()::text)
    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;

    RAISE NOTICE 'dve_cutover_to_long_format: cut over. measurements inserted this run: %, sample_headers: %.',
        measurements_inserted, sample_headers_inserted;
END;
$$ LANGUAGE plpgsql;

COMMENT ON FUNCTION dve_cutover_to_long_format() IS
    'v3 Phase 3d cutover: idempotently (re-)runs dve_migrate_to_long_format, verifies '
    'row-count parity AND the per-column sparse rule, renames data_rows/samples aside '
    '(data_rows_pre_v3 / samples_core) and creates the name-holder views. Supervised '
    'runbook step - the application never calls this (index D-3d-2).';

-- ============================================================================
-- 4. Execute (trivial on an empty test container; on the NAS, applying this
--    file by hand IS the supervised cutover step).
-- ============================================================================
SELECT * FROM dve_cutover_to_long_format();

COMMIT;
