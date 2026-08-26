# TPM v3 Phase 3d - Standard-Metric Cutover Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans (in-context, per the project's token-efficiency directive) to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Move the 13 standard data-row metrics and 22 standard sample headers from the wide `data_rows` / `samples` columns into `measurements` / `sample_headers`, rename the wide tables aside, put the compat views under the old names, and leave every reader working and every write going to exactly one place.

**Architecture:** The read path stays wide SQL and silently starts reading the compat views (D1 pays off: zero read-path churn for standard metrics). The write path replaces per-row wide UPDATE/INSERT with the same batched keyed upserts 3c built for extras. `samples` survives as a narrow identity table (`samples_core`); `data_rows` becomes pure view. The cutover itself is one supervised SQL function applied to the NAS by hand, never run by the app; the app probes for post-cutover shape at connect and refuses cleanly otherwise.

**Tech Stack:** Qt 6.10 / QPSQL, PostgreSQL 16, PL/pgSQL, Qt Test, PowerShell provisioning.

**Master index:** `docs/superpowers/plans/2026-07-30-tpm-v3-phase3-INDEX.md`. Read its locked decisions (D1-D9), the hazard ledger (H1-H25), and the SEQUENCING CORRECTION before starting.
Read the memory `v2-4-1-regression-fix-batch` before touching the save path.

---

## Decisions locked for 3d (extend the index's D-series; record them there at wrap)

**D-3d-1 (cutover choreography).**
One new migration file `deploy/postgres/migrations/2026-08-26-v3-cutover.sql` defines `dve_cutover_to_long_format()`, which idempotently: runs `dve_migrate_to_long_format()` (itself idempotent), verifies per-column sparse counts and row-count parity (the 3b lesson: parity alone is insufficient, so the function checks the sparse rule per column), renames `data_rows -> data_rows_pre_v3` and `samples -> samples_core`, creates view `data_rows` (chained over `data_rows_v`) and view `samples` (samples_core LEFT JOIN samples_v), revokes, and stamps `schema_meta`.
The renamed wide tables are kept intact as the on-database pre-image.
They are NOT the production rollback path (production rollback = restore the pre-cutover pg_dump backup, D7); `data_rows_pre_v3` additionally serves as the translation table for v1 pending_edits (H9).

**D-3d-2 (who runs it).**
The app NEVER calls `dve_migrate_to_long_format()` or `dve_cutover_to_long_format()`.
Both are supervised v3.0.0 runbook steps applied to the NAS by hand (extends D8; the earlier "app wraps it in the advisory-lock gate" note in the 3b banner predates D8 and is superseded - a client must not run a full-table restructure over office wifi on first launch).
The app probes `relkind` of `data_rows` at connect: `'v'` = proceed; `'r'` or absent = refuse with a specific error, falling into the existing offline-snapshot path.

**D-3d-3 (write path).**
Standard data-row metrics write through the exact `flushExtras` batched-upsert machinery 3c built, keyed `(sample_id, metric_id, sort_order = row ordinal)`.
`DataRow::id` / `DataRow::version` are no longer read or written by the save path (closes H2; H13 dissolves - keyed upserts need no id anchors).
`samples_core` keeps the narrow id+version OCC upsert (test_id, sort_order, updated_by) exactly as today; the 22 values ride the `sample_headers` batch.
Sparse rule at write time = "a measurement exists exactly where today's wide bind was non-NULL": null QStrings and the `!hasPerRowRegime` regime bind are skipped, everything else always writes.
Standard keys resolve through `MetricDefCache` WITHOUT auto-registration; a missing standard key aborts the save (a broken seed must be loud).

**D-3d-4 (prune).**
The phase-A pre-image drops the `md.key NOT IN (<wide cols>)` exclusion: with the writer reproducing every standard metric, ALL of the file's measurements/sample_headers enter the pre-image, and pre-minus-post prunes deleted rows' measurements exactly as D9 designed.
The `data_rows` entry leaves the prune table list (no table to delete from; its deletion IS its measurements' deletion).
`samples` prune re-targets `samples_core` (FK cascade takes the child long rows and the frozen `data_rows_pre_v3` rows).
`appendExtra`'s wide-column skip STAYS: it now guards against a standard key sneaking in via a stale `extra` map (recovery JSON from a pre-3d build) and double-writing.

**D-3d-5 (H25 read filter).**
The two extras SELECTs in `loadFile` gain `AND md.key NOT IN (SELECT attname FROM pg_attribute WHERE attrelid = to_regclass('<table>') AND attnum > 0 AND NOT attisdropped)`.
Same SQL is correct pre- and post-cutover (pre-cutover the filter excludes keys that never appear; on a DB with no such relation, `to_regclass` is NULL, the subquery is empty, and nothing is filtered), so it lands before the flip.

**D-3d-6 (per-cell commit, H4).**
New stored function `dve_commit_measurement(p_sample_id, p_key, p_sort_order, p_value, p_uuid)` - ONE signature, no overloads (index decision: never repeat the dve_commit_cell overload ambiguity).
`LiveSync::commitMeasurement(sampleId, key, sortOrder, value)` replaces the `commitCell("data_rows", ...)` call in `onStoryCellEdited`; the dead `data_rows`/`samples`/`tests`/`files` entries leave the commitCell allowlist (audit-confirmed unreachable).
Offline queue: pending_edits gains `schema_version INTEGER` and `sort_order INTEGER` columns (ALTER-if-missing, the established pattern); measurement edits enqueue as v2 rows (`row_id` = sample id, `column_name` = metric key); the drain replays v2 via `dve_commit_measurement` and translates v1 `data_rows` rows through `data_rows_pre_v3` (id -> sample_id, sort_order); a v1 row whose id is absent there is dropped with a warning and one unsynced bump.

**D-3d-7 (H23).**
`bump_version()` removes `updated_by` from the no-op comparison AND restores `NEW.updated_by := OLD.updated_by` in the no-op branch, so a suppressed no-op leaves the row byte-identical.
Scope note: this also suppresses cross-client identical-content sensory saves, which is the DV-23-desired behavior (the DV-28 needs-save gates already stop clean sessions from issuing UPDATEs at all).
Delivered in the cutover .sql (latest CREATE OR REPLACE wins, same precedent as 2026-07-15 / 2026-07-31 superseding init.sql's copy without editing it).

**D-3d-8 (test architecture).**
`tests/start-test-postgres.ps1` provisions a SECOND database `dve_test_precut` with the same file list MINUS the cutover file; `dve_test` gets everything including the cutover (its .sql tail calls the function; trivial on an empty DB).
The migration-rehearsal harness (`tst_v3longformat`'s 3b half) re-homes to `dve_test_precut`; app-facing suites keep `dve_test`, now post-cutover.
The harness resets its private DB by UN-cutting at initTestCase (drop the two name-holder views, rename back, wipe); the un-cut lives ONLY in the harness so nobody mistakes it for a production rollback.
`tst_databasemanager` slots doing destructive long-table DDL re-home to the precut DB too.

**Conscious non-changes:**
- `notify_row_change` stays on `samples_core` (payload table name changes to `samples_core`; the audit proved zero consumers; broadcast volume matches today).
- `data_rows_pre_v3`'s triggers/index ride along inert.
- init.sql is NOT edited (migrations supersede, per precedent).
- The tests-table aggregates (`overall_avg_tpm` etc.) stay wide; `tests` does not migrate.
- Snapshot regen and the offline reader change ZERO lines: explicit name-based column lists hit the views (verified: samples view exposes exactly the 28 selected columns, data_rows view exactly the 19).
- `dve_uncut` is not shipped SQL; production rollback is the D7 backup.

**Known perf watch-item:** snapshot regen now pivots the whole `measurements` table through the view (hash agg).
Fine for close-time batch work at current data volume; eyeball at smoke, note in the index if slow.

---

## Task 1: H21 - make `samples_v` structurally total

**Files:**
- Modify: `deploy/postgres/migrations/2026-07-31-v3-long-format.sql` (the `samples_v` CREATE at ~line 374 and its comments at ~305-314, ~365-373, ~407-413)
- Test: `tests/tst_v3longformat/tst_v3longformat.cpp`

- [ ] **Step 1: Write the failing tripwire slot**

Add to the poka-yoke group (rolled-back transaction pattern, before the e2e gates in declaration order):

```cpp
// H21 (owner-approved fix 2026-08-26): samples_v must be structurally TOTAL -
// exactly one view row per samples row, even when every one of the 22 header
// columns is NULL and therefore (sparse rule, D2) zero sample_headers rows
// exist. Before the fix the view was driven FROM sample_headers, so such a
// sample had no group and vanished; every reader then had to remember to LEFT
// JOIN, which nothing enforced. After the fix the count parity below is an
// identity, not an assertion - this slot goes red only if someone edits the
// view back to an INNER form.
void TstV3LongFormat::samplesView_isTotal_allNullHeaderSampleSurvives()
{
    QVERIFY(db().transaction());
    QSqlQuery q(db());
    // A sample whose 22 value columns are ALL NULL. The columns carry
    // DEFAULT 0.0 so the NULLs must be explicit - which is exactly the
    // MigrationTool-verbatim-copy provenance H21 names.
    QVERIFY2(q.exec(
        "INSERT INTO samples (test_id, sort_order, sample_name, sample_id, date, "
        " tester, media, viscosity, resistance, voltage, power, heating_technology, "
        " puffing_regime, initial_oil_mass, average_tpm, stddev_tpm, avg_power_density, "
        " efficiency_percent, total_oil_consumed, total_puffs, normalized_tpm, "
        " burn_status, clog_status, leak_status, updated_by) "
        "SELECT id, 999, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, "
        " NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, "
        " 'h21-tripwire' FROM tests LIMIT 1 RETURNING id"),
        qPrintable(q.lastError().text()));
    QVERIFY(q.next());
    const qint64 ghostId = q.value(0).toLongLong();

    // No header rows may exist for it (nothing inserted any), yet it must
    // appear in samples_v with a NULL sample_name.
    QCOMPARE(scalar(QStringLiteral(
        "SELECT count(*) FROM sample_headers WHERE sample_id = %1").arg(ghostId)),
        static_cast<qint64>(0));
    QCOMPARE(scalar(QStringLiteral(
        "SELECT count(*) FROM samples_v WHERE id = %1").arg(ghostId)),
        static_cast<qint64>(1));
    QVERIFY(textScalar(QStringLiteral(
        "SELECT COALESCE(sample_name, '<null>') FROM samples_v WHERE id = %1")
            .arg(ghostId)) == QLatin1String("<null>"));

    // The global identity: one view row per samples row, always.
    QCOMPARE(scalar(QStringLiteral("SELECT count(*) FROM samples_v")),
             scalar(QStringLiteral("SELECT count(*) FROM samples")));
    db().rollback();
}
```

Declare the slot in the private-slots block beside the other poka-yokes.

- [ ] **Step 2: Run it to verify it fails against the current INNER-driven view**

```
tests\run-tests.ps1 (tst_v3longformat only, or build the suite and run the one slot)
```

Expected: FAIL on `count(samples_v) WHERE id = ghost` (0 vs 1).
If the container predates this session, re-provision first: `docker rm -f dve-test-pg; tests\start-test-postgres.ps1`.

- [ ] **Step 3: Replace the view definition**

In `2026-07-31-v3-long-format.sql`, replace the `CREATE OR REPLACE VIEW samples_v` statement with:

```sql
-- Driven FROM samples with a LEFT JOIN (H21 fix, owner-approved 2026-08-26):
-- exactly one row per samples row, always, whatever the header count. The
-- LEFT-ness lives INSIDE the view, enforced once, instead of being a
-- convention every reader must remember. A sample with zero header rows reads
-- as all-NULL - precisely what the wide row held - so no fourth pre-flight
-- abort is needed and an all-NULL legacy sample migrates faithfully instead
-- of blocking the migration. Still D2-sparse: nothing is materialized.
-- Safe under 3d's rename: Postgres binds views to base relations by OID
-- (verified live 2026-08-03), so ALTER TABLE samples RENAME leaves this view
-- reading the renamed table and a new `samples` view can take the old name.
CREATE OR REPLACE VIEW samples_v AS
SELECT
    s.id                                                                  AS id,
    MAX(sh.value_text) FILTER (WHERE md.key = 'sample_name')              AS sample_name,
    MAX(sh.value_text) FILTER (WHERE md.key = 'sample_id')                AS sample_id,
    MAX(sh.value_text) FILTER (WHERE md.key = 'date')                     AS date,
    MAX(sh.value_text) FILTER (WHERE md.key = 'tester')                   AS tester,
    MAX(sh.value_text) FILTER (WHERE md.key = 'media')                    AS media,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'viscosity')                AS viscosity,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'resistance')               AS resistance,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'voltage')                  AS voltage,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'power')                    AS power,
    MAX(sh.value_text) FILTER (WHERE md.key = 'heating_technology')       AS heating_technology,
    MAX(sh.value_text) FILTER (WHERE md.key = 'puffing_regime')           AS puffing_regime,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'initial_oil_mass')         AS initial_oil_mass,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'average_tpm')              AS average_tpm,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'stddev_tpm')               AS stddev_tpm,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'avg_power_density')        AS avg_power_density,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'efficiency_percent')       AS efficiency_percent,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'total_oil_consumed')       AS total_oil_consumed,
    (MAX(sh.value_num) FILTER (WHERE md.key = 'total_puffs'))::INTEGER    AS total_puffs,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'normalized_tpm')           AS normalized_tpm,
    MAX(sh.value_text) FILTER (WHERE md.key = 'burn_status')              AS burn_status,
    MAX(sh.value_text) FILTER (WHERE md.key = 'clog_status')              AS clog_status,
    MAX(sh.value_text) FILTER (WHERE md.key = 'leak_status')              AS leak_status
FROM samples s
LEFT JOIN sample_headers sh ON sh.sample_id = s.id
LEFT JOIN metric_defs    md ON md.id = sh.field_id AND md.kind = 'header'
GROUP BY s.id;
```

Also update: the `COMMENT ON VIEW samples_v` ("JOIN PARTNER" wording stays; add one sentence: "Structurally total since the H21 fix: exactly one row per samples row, including samples with zero header rows."), the "KNOWN LIMIT OF THE SPARSE RULE" block (~305: the samples exposure paragraph now reads "eliminated by the H21 total-view fix" instead of "milder"), and the "Correction (review...)" style notes if any conflict.

`CREATE OR REPLACE VIEW` cannot change an output column set, but this definition keeps the same 23 output columns in the same order, so a REPLACE over the 3b view succeeds.
Apply to the running container: re-provision (`docker rm -f dve-test-pg; tests\start-test-postgres.ps1`) - full re-provision, not a hand psql, so the file itself is proven.

- [ ] **Step 4: Run the harness - the new slot AND the whole rehearsal suite**

Expected: new slot PASSES; every existing parity/checksum/poka-yoke slot still green (the view change must not move any migrated value).

- [ ] **Step 5: Commit**

```bash
git add deploy/postgres/migrations/2026-07-31-v3-long-format.sql tests/tst_v3longformat/tst_v3longformat.cpp
git commit -m "fix(v3): samples_v is structurally total (H21) - LEFT JOIN from samples, tripwire slot"
```

---

## Task 2: the cutover migration file (SQL only, function not yet invoked anywhere)

**Files:**
- Create: `deploy/postgres/migrations/2026-08-26-v3-cutover.sql`
- Test: `tests/tst_v3longformat/tst_v3longformat.cpp` (function-level slots that are safe on the shared container)

- [ ] **Step 1: Write the file**

Complete content (use the Write tool; the file is new so MIP labeling is not a concern for SQL, but keep the repo's UTF-8/LF convention):

```sql
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
  IF TG_OP = 'UPDATE'
     AND TG_TABLE_NAME IN ('sensory_sessions', 'detailed_sensory_sessions',
                           'measurements', 'sample_headers')
     AND (to_jsonb(NEW) - 'version' - 'updated_at' - 'app_version' - 'updated_by')
         IS NOT DISTINCT FROM
         (to_jsonb(OLD) - 'version' - 'updated_at' - 'app_version' - 'updated_by')
  THEN
    NEW.version    := OLD.version;
    NEW.updated_at := OLD.updated_at;
    NEW.updated_by := OLD.updated_by;   -- H23: a no-op leaves the row byte-identical
    RETURN NEW;
  END IF;
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
    OUT already_cut            BOOLEAN,
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
    r      RECORD;
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
        measurements_inserted := 0;
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
```

- [ ] **Step 2: Function-level tests that are safe on the shared (still pre-cutover) container**

The cutover function itself is NOT invoked against the shared container in this task (it would flip a database five other suites share; that flip is Task 6).
But the other two functions are testable now.
Add to `tst_v3longformat.cpp`, rolled-back transaction pattern:

```cpp
// Phase 3d Task 2: dve_commit_measurement writes by natural key, routes by
// value_type, and returns FALSE for an unknown key. Runs inside a rolled-back
// transaction so nothing leaks into the shared container.
void TstV3LongFormat::commitMeasurement_upsertsByNaturalKeyAndType()
{
    QVERIFY(db().transaction());
    QSqlQuery q(db());
    // Any migrated sample will do as the FK target.
    const qint64 sid = scalar(QStringLiteral("SELECT min(id) FROM samples"));
    QVERIFY(sid > 0);

    auto commit = [&](const char* key, int so, const char* val) -> bool {
        QSqlQuery c(db());
        c.prepare(QStringLiteral("SELECT dve_commit_measurement(?, ?, ?, ?, ?)"));
        c.addBindValue(static_cast<qlonglong>(sid));
        c.addBindValue(QLatin1String(key));
        c.addBindValue(so);
        c.addBindValue(QLatin1String(val));
        c.addBindValue(QStringLiteral("tst-3d"));
        if (!c.exec() || !c.next()) return false;
        return c.value(0).toBool();
    };

    // number key -> value_num, text NULL
    QVERIFY(commit("tpm", 990, "7.25"));
    QCOMPARE(textScalar(QStringLiteral(
        "SELECT value_num::text || '|' || COALESCE(value_text, '<null>') "
        "FROM measurements m JOIN metric_defs md ON md.id = m.metric_id "
        "WHERE m.sample_id = %1 AND md.key = 'tpm' AND m.sort_order = 990")
            .arg(sid)),
        QStringLiteral("7.25|<null>"));
    // upsert same key: value replaced, still one row
    QVERIFY(commit("tpm", 990, "8.5"));
    QCOMPARE(scalar(QStringLiteral(
        "SELECT count(*) FROM measurements m JOIN metric_defs md ON md.id = m.metric_id "
        "WHERE m.sample_id = %1 AND md.key = 'tpm' AND m.sort_order = 990").arg(sid)),
        static_cast<qint64>(1));
    // text key -> value_text
    QVERIFY(commit("notes", 990, "burnt taste"));
    // unknown key -> FALSE, no row
    QVERIFY(!commit("no_such_metric_key_3d", 990, "x"));
    db().rollback();
}

// Phase 3d Task 2 / hazard H23: an UPDATE differing ONLY in updated_by is a
// no-op - version, updated_at AND updated_by all keep their old values.
void TstV3LongFormat::bumpVersion_updatedByAloneIsANoOp()
{
    QVERIFY(db().transaction());
    const qint64 sid = scalar(QStringLiteral("SELECT min(id) FROM samples"));
    QSqlQuery q(db());
    QVERIFY(q.exec(QStringLiteral(
        "INSERT INTO sample_headers (sample_id, field_id, value_text, updated_by) "
        "SELECT %1, id, 'userA', 'userA' FROM metric_defs "
        "WHERE kind = 'header' AND key = 'tester' RETURNING id").arg(sid)));
    QVERIFY(q.next());
    const qint64 rowId = q.value(0).toLongLong();

    QVERIFY(q.exec(QStringLiteral(
        "UPDATE sample_headers SET value_text = 'userA', updated_by = 'userB' "
        "WHERE id = %1").arg(rowId)));
    QCOMPARE(textScalar(QStringLiteral(
        "SELECT version::text || '|' || updated_by FROM sample_headers WHERE id = %1")
            .arg(rowId)),
        QStringLiteral("1|userA"));   // suppressed: no bump, updated_by restored

    QVERIFY(q.exec(QStringLiteral(
        "UPDATE sample_headers SET value_text = 'changed', updated_by = 'userB' "
        "WHERE id = %1").arg(rowId)));
    QCOMPARE(textScalar(QStringLiteral(
        "SELECT version::text || '|' || updated_by FROM sample_headers WHERE id = %1")
            .arg(rowId)),
        QStringLiteral("2|userB"));   // real change: bump + new author
    db().rollback();
}
```

- [ ] **Step 3: Run them to verify they FAIL** (functions absent / old bump_version).

Apply the two function definitions (sections 1-2 of the new file) to the running container by hand-psql for the red/green cycle, or re-provision after adjusting the provisioning exclusion below.
IMPORTANT: until Task 3 lands, the provisioning script would auto-apply the WHOLE cutover file (including the tail call) to the shared container and break every other suite.
So in THIS task, add the interim guard to `tests/start-test-postgres.ps1`: skip any file matching `*v3-cutover*` (one `-notlike` in the file loop, marked `# TEMP until Task 3 splits the databases`).
Then re-provision and apply sections 1+2 by hand psql for the slot run.

- [ ] **Step 4: Run the two new slots green; whole rehearsal suite still green.**

- [ ] **Step 5: Commit**

```bash
git add deploy/postgres/migrations/2026-08-26-v3-cutover.sql tests/tst_v3longformat/tst_v3longformat.cpp tests/start-test-postgres.ps1
git commit -m "feat(v3): cutover migration - dve_cutover_to_long_format, dve_commit_measurement, H23 bump_version fix"
```

---

## Task 3: two-database provisioning + rehearsal re-home + cutover rehearsal slots

**Files:**
- Modify: `tests/start-test-postgres.ps1`
- Modify: `tests/tst_v3longformat/tst_v3longformat.cpp`
- Modify: `tests/tst_databasemanager/tst_databasemanager.cpp` (destructive-DDL slots only)

- [ ] **Step 1: Provisioning split**

In `start-test-postgres.ps1`, after the existing `dve_test` provisioning loop:

1. Remove the Task-2 TEMP exclusion for `dve_test` (the cutover file now applies there - `dve_test` becomes post-cutover).
2. Create the second database and apply the same ordered file list EXCLUDING `*v3-cutover*`:

```powershell
# ── v3 Phase 3d (index D-3d-8): a second, PRE-cutover database for the
# migration-rehearsal harness. dve_test above is post-cutover (the cutover
# migration applies there), which is what every app-facing suite must test
# against - but tst_v3longformat's 3b half seeds wide rows and runs
# dve_migrate_to_long_format itself, which requires the wide tables to still
# be real tables. It gets its own database with everything EXCEPT the cutover.
& $psql -v ON_ERROR_STOP=1 -U test -d postgres -c "DROP DATABASE IF EXISTS dve_test_precut"
if ($LASTEXITCODE -ne 0) { throw "could not drop dve_test_precut" }
& $psql -v ON_ERROR_STOP=1 -U test -d postgres -c "CREATE DATABASE dve_test_precut OWNER test"
if ($LASTEXITCODE -ne 0) { throw "could not create dve_test_precut" }
foreach ($f in $sqlFiles) {                     # the SAME ordered list as above
    if ($f.Name -like '*v3-cutover*') { continue }   # THE difference
    Apply-SqlFile -Database 'dve_test_precut' -File $f   # same helper/machinery
}
```

Adapt the exact loop/helper names to the script's existing structure (it already has the per-file apply with ON_ERROR_STOP, the statement splitter, and the pg_cron strip - reuse those verbatim; do not fork the machinery).
`DVE_TEST_PG_CONN` continues to point at `dve_test`.

- [ ] **Step 2: Re-home the rehearsal harness**

In `tst_v3longformat.cpp`:

- Add `const char* const kPrecutConn = "tst_v3longformat_precut";` and a `dbPrecut()` accessor.
- In `initTestCase`, open a second QPSQL connection from `pgConfig()` with `c.database = "dve_test_precut"`.
- Re-point the 3b half (seed/wipe/migration/parity/checksum/poka-yoke/tripwire slots, including Task 1's H21 slot and Task 2's two function slots) at `dbPrecut()`.
  The MetricDefCache slots and the three e2e gate slots STAY on `kConn` (they drive production components against the post-cutover `dve_test`, which is exactly what 3d wants them to prove).
- Add the un-cut reset at the top of `initTestCase` (private-DB idempotency for re-runs):

```cpp
// D-3d-8: this suite CUTS ITS PRIVATE DB OVER in the cutover slots below, so a
// re-run must first put it back. This un-cut exists ONLY here - production
// rollback is the D7 backup, never a rename-back (post-cutover writes live
// only in the long tables and a rename-back would strand them).
{
    QSqlQuery q(dbPrecut());
    q.exec(QStringLiteral("SELECT c.relkind FROM pg_class c WHERE c.oid = to_regclass('data_rows')"));
    if (q.next() && q.value(0).toString() == QLatin1String("v")) {
        QVERIFY(q.exec(QStringLiteral("DROP VIEW data_rows")));
        QVERIFY(q.exec(QStringLiteral("DROP VIEW samples")));
        QVERIFY(q.exec(QStringLiteral("ALTER TABLE data_rows_pre_v3 RENAME TO data_rows")));
        QVERIFY(q.exec(QStringLiteral("ALTER TABLE samples_core RENAME TO samples")));
        QVERIFY(q.exec(QStringLiteral("DELETE FROM schema_meta WHERE key = 'v3_long_format_cutover'")));
    }
}
```

- [ ] **Step 3: Cutover rehearsal slots (declaration order: AFTER every pre-cutover slot, LAST in the precut group)**

```cpp
// ── Phase 3d: the cutover itself, rehearsed on the migrated fixture ─────────
// Declared after every slot that needs the wide tables to be real tables.
void TstV3LongFormat::cutover_migratesVerifiesRenamesAndCreatesViews()
{
    QSqlQuery q(dbPrecut());
    QVERIFY2(q.exec(QStringLiteral("SELECT * FROM dve_cutover_to_long_format()")),
             qPrintable(q.lastError().text()));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), false);              // already_cut = false, first run

    auto relkind = [&](const char* rel) {
        QSqlQuery k(dbPrecut());
        k.exec(QStringLiteral("SELECT c.relkind FROM pg_class c WHERE c.oid = to_regclass('%1')")
                   .arg(QLatin1String(rel)));
        return k.next() ? k.value(0).toString() : QString();
    };
    QCOMPARE(relkind("data_rows"),        QStringLiteral("v"));
    QCOMPARE(relkind("samples"),          QStringLiteral("v"));
    QCOMPARE(relkind("data_rows_pre_v3"), QStringLiteral("r"));
    QCOMPARE(relkind("samples_core"),     QStringLiteral("r"));

    // Every wide reader keeps working: the loadFile-shaped and regen-shaped
    // SELECTs run BY NAME against the views and return the fixture's counts.
    QCOMPARE(scalarPrecut(QStringLiteral("SELECT count(*) FROM data_rows")),
             static_cast<qint64>(m_dataRows));
    QCOMPARE(scalarPrecut(QStringLiteral("SELECT count(*) FROM samples")),
             static_cast<qint64>(m_samples));
    QCOMPARE(scalarPrecut(QStringLiteral(
        "SELECT count(*) FROM (SELECT id, test_id, sort_order, sample_name, sample_id, "
        "date, tester, media, viscosity, resistance, voltage, power, heating_technology, "
        "puffing_regime, initial_oil_mass, average_tpm, stddev_tpm, avg_power_density, "
        "efficiency_percent, total_oil_consumed, total_puffs, normalized_tpm, "
        "burn_status, clog_status, leak_status, updated_at, updated_by, version "
        "FROM samples) x")),
             static_cast<qint64>(m_samples));
    // Values survive: spot-check the tricky double at its known location.
    // (Reuse the kTrickyDouble constant and m_trickySampleId.)
    QCOMPARE(textScalarPrecut(QStringLiteral(
        "SELECT tpm::text FROM data_rows WHERE sample_id = %1 AND sort_order = 0")
            .arg(m_trickySampleId)),
        textScalarPrecut(QStringLiteral(
        "SELECT tpm::text FROM data_rows_pre_v3 WHERE sample_id = %1 AND sort_order = 0")
            .arg(m_trickySampleId)));
    QCOMPARE(textScalarPrecut(QStringLiteral(
        "SELECT value FROM schema_meta WHERE key = 'v3_long_format_cutover'")).isEmpty(),
        false);
}

void TstV3LongFormat::cutover_isIdempotent()
{
    QSqlQuery q(dbPrecut());
    QVERIFY2(q.exec(QStringLiteral("SELECT * FROM dve_cutover_to_long_format()")),
             qPrintable(q.lastError().text()));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toBool(), true);               // already_cut on re-run
}

void TstV3LongFormat::cutover_wideWritesAreRefusedByTheViews()
{
    // The single-source-of-truth guarantee: nothing can write the old shape
    // any more. Both statements must FAIL (grouped views are not updatable).
    QSqlQuery q(dbPrecut());
    QVERIFY(!q.exec(QStringLiteral(
        "UPDATE data_rows SET tpm = 1.0 WHERE sample_id = %1 AND sort_order = 0")
            .arg(m_trickySampleId)));
    QVERIFY(!q.exec(QStringLiteral(
        "INSERT INTO data_rows (sample_id, sort_order, puffs) VALUES (%1, 998, 1)")
            .arg(m_trickySampleId)));
}
```

Add `scalarPrecut` / `textScalarPrecut` helpers (same bodies as `scalar`/`textScalar` against `dbPrecut()`), or refactor the existing pair to take the connection.
Note: the existing `orphanTripwire_*` slots assert against `data_rows` as a table; confirm they are declared BEFORE the cutover slots and adjust their SQL only if they reference catalog shape.

- [ ] **Step 4: Re-home destructive tst_databasemanager slots**

`v3LongFormat_ensureSchemaDdlMatchesMigrationFile` drops the three long tables and lets ensureSchema rebuild them.
Post-cutover on `dve_test`, `DROP TABLE measurements` fails (the name-holder views depend on it) or a CASCADE would destroy the views for every later suite.
Re-home that slot (and any sibling doing destructive long-table DDL - audit the `v3LongFormat_*` group) to a second connection against `dve_test_precut`, taking care to run BEFORE tst_v3longformat's cutover slots would matter - they do not conflict because each suite re-provisions its own state: add the same un-cut guard at that slot's start if it can observe a cut-over precut DB (it can, when tst_v3longformat ran first in the same container lifetime).
Simplest robust form: the slot itself calls the un-cut sequence if `data_rows` is a view, then proceeds.
Extract the un-cut into a tiny shared header-only helper if duplicating it twice feels wrong: `tests/shared/v3precut.h` with `inline void dveUncutPrecut(QSqlDatabase&)`.

- [ ] **Step 5: Re-provision, run tst_v3longformat + tst_databasemanager**

```
docker rm -f dve-test-pg; tests\start-test-postgres.ps1
```

Expected: all green.
NOTE: from this step on, `dve_test` is POST-cutover and every OTHER DB suite (saveintegrity, livesync, twoclient, offlinesnapshot, storedfns, notificationlistener) is EXPECTED RED until Tasks 4-6 land.
Do not run the full suite between Tasks 3 and 6; run the targeted suites named in each task.
This is the one deliberate mid-phase red window; it is confined to this branch and ends inside this same working session.

- [ ] **Step 6: Commit**

```bash
git add tests/start-test-postgres.ps1 tests/tst_v3longformat/tst_v3longformat.cpp tests/tst_databasemanager/tst_databasemanager.cpp tests/shared/v3precut.h
git commit -m "feat(v3): two-database provisioning (dve_test post-cutover, dve_test_precut for rehearsal) + cutover rehearsal slots"
```

---

## Task 4: dual-mode wide-row seeding helper + fixture swaps

**Files:**
- Create: `tests/shared/seedrows.h`
- Modify: `tests/tst_livesync/tst_livesync.cpp` (~161-170)
- Modify: `tests/tst_twoclient_e2e/tst_twoclient_e2e.cpp` (~927-935)
- Modify: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp` (~218-280)
- Modify: `tests/tst_storedfns/tst_storedfns.cpp` (~138-146)
- Modify: `tests/tst_notificationlistener/tst_notificationlistener.cpp` (~142-150)
- Check-only: `tests/tst_migrationtool/tst_migrationtool.cpp` (its INSERTs at ~207-210 target the legacy SQLite side; verify and leave alone if so)

- [ ] **Step 1: The helper**

Header-only, no .pro surgery beyond INCLUDEPATH if `tests/shared` is not already on it (it is, if Task 3 created v3precut.h there; otherwise add `INCLUDEPATH += ../shared` to each touched .pro):

```cpp
#pragma once
// v3 Phase 3d test helper: seed a sample and its data rows in whichever shape
// the connected database speaks. Pre-cutover (data_rows is a real table) it
// INSERTs wide rows exactly as the fixtures always did; post-cutover it
// INSERTs the narrow samples_core row plus per-metric measurements. Fixtures
// switch to these helpers BEFORE the container flips, so the flip itself
// changes no test code (D-3d-8).
#include <QtSql>

namespace DVE { namespace TestSeed {

inline bool isPostCutover(QSqlDatabase& db)
{
    QSqlQuery q(db);
    if (!q.exec(QStringLiteral(
            "SELECT c.relkind FROM pg_class c WHERE c.oid = to_regclass('data_rows')")))
        return false;
    return q.next() && q.value(0).toString() == QLatin1String("v");
}

// Returns the new sample id, or -1. Wide pre-cutover; narrow core + headers
// post-cutover. Only sample_name is seeded as a header (the historical
// fixtures set nothing else).
inline qint64 seedSample(QSqlDatabase& db, qint64 testId, const QString& name,
                         int sortOrder = 0)
{
    QSqlQuery q(db);
    if (!isPostCutover(db)) {
        q.prepare(QStringLiteral(
            "INSERT INTO samples (test_id, sort_order, sample_name) "
            "VALUES (?, ?, ?) RETURNING id"));
        q.addBindValue(static_cast<qlonglong>(testId));
        q.addBindValue(sortOrder);
        q.addBindValue(name);
        if (!q.exec() || !q.next()) return -1;
        return q.value(0).toLongLong();
    }
    q.prepare(QStringLiteral(
        "INSERT INTO samples_core (test_id, sort_order) VALUES (?, ?) RETURNING id"));
    q.addBindValue(static_cast<qlonglong>(testId));
    q.addBindValue(sortOrder);
    if (!q.exec() || !q.next()) return -1;
    const qint64 sid = q.value(0).toLongLong();
    QSqlQuery h(db);
    h.prepare(QStringLiteral(
        "INSERT INTO sample_headers (sample_id, field_id, value_text, updated_by) "
        "SELECT ?, id, ?, 'test-seed' FROM metric_defs "
        "WHERE kind = 'header' AND key = 'sample_name'"));
    h.addBindValue(static_cast<qlonglong>(sid));
    h.addBindValue(name);
    if (!h.exec()) return -1;
    return sid;
}

// Seeds one data row carrying the given numeric metric (the fixtures only
// ever seed draw_pressure or nothing). Returns true on success. Post-cutover
// there is no data-row id; pre-cutover callers that need the id can query it
// back - the swapped fixtures below no longer need one.
inline bool seedDataRow(QSqlDatabase& db, qint64 sampleId, int sortOrder,
                        const char* numericKey = nullptr, double value = 0.0)
{
    QSqlQuery q(db);
    if (!isPostCutover(db)) {
        if (numericKey) {
            q.prepare(QStringLiteral(
                "INSERT INTO data_rows (sample_id, sort_order, %1) VALUES (?, ?, ?)")
                    .arg(QLatin1String(numericKey)));
            q.addBindValue(static_cast<qlonglong>(sampleId));
            q.addBindValue(sortOrder);
            q.addBindValue(value);
        } else {
            q.prepare(QStringLiteral(
                "INSERT INTO data_rows (sample_id, sort_order) VALUES (?, ?)"));
            q.addBindValue(static_cast<qlonglong>(sampleId));
            q.addBindValue(sortOrder);
        }
        return q.exec();
    }
    // Post-cutover: a data row IS its measurements (sparse rule). A row with
    // no metric at all cannot exist (H18); seed puffs = 0 as the anchor,
    // which is what a real parse always writes anyway.
    auto put = [&](const char* key, double v) {
        QSqlQuery m(db);
        m.prepare(QStringLiteral(
            "INSERT INTO measurements (sample_id, metric_id, sort_order, value_num, updated_by) "
            "SELECT ?, id, ?, ?, 'test-seed' FROM metric_defs "
            "WHERE kind = 'metric' AND key = ? "
            "ON CONFLICT (sample_id, metric_id, sort_order) DO UPDATE "
            "SET value_num = EXCLUDED.value_num"));
        m.addBindValue(static_cast<qlonglong>(sampleId));
        m.addBindValue(sortOrder);
        m.addBindValue(v);
        m.addBindValue(QLatin1String(key));
        return m.exec();
    };
    if (!put("puffs", 0.0)) return false;
    if (numericKey && !put(numericKey, value)) return false;
    return true;
}

}} // namespace DVE::TestSeed
```

- [ ] **Step 2: Swap the five suites' fixture INSERTs for helper calls**

Each is 1-2 statements; keep surrounding assertions.
Where a fixture captured the data-row id for later assertions (tst_livesync / tst_storedfns exercise `dve_commit_cell` on it), leave a `// Phase 3d Task 6 re-points this` marker; those assertions change in Task 6, not here.
`tst_offlinesnapshot`'s seed writes many columns - extend the helper call pattern rather than the helper (loop `seedDataRow` per row plus direct `sample_headers` seeding for the extra header values it sets, mirroring `seedSample`'s post-cutover branch; pre-cutover keep its existing wide INSERT under `!isPostCutover(db)`).
Keep the diff minimal: the goal is behavior-identical seeding pre-cutover and legal seeding post-cutover.

- [ ] **Step 3: Run the five suites against the CURRENT container state**

`dve_test` is already post-cutover (Task 3), so run them and expect: fixtures now SEED successfully; the suites may still be red in their ASSERT phases (production save/commit code is not cut over until Tasks 5-6).
Record which slots remain red and confirm each is an assert-phase failure, not a seed failure.

- [ ] **Step 4: Commit**

```bash
git add tests/shared/seedrows.h tests/tst_livesync tests/tst_twoclient_e2e tests/tst_offlinesnapshot tests/tst_storedfns tests/tst_notificationlistener
git commit -m "test(v3): dual-mode row seeding helper; fixtures legal on both sides of the cutover"
```

---

## Task 5: read-side and live-path changes that are correct in BOTH worlds

**Files:**
- Modify: `src/database/DatabaseManager.cpp` (loadFile extras SELECTs at ~1474 and ~1504; the connect probe near `open()`)
- Modify: `src/database/LiveSync.h` / `src/database/LiveSync.cpp` (commitMeasurement; allowlist prune; enqueue v2)
- Modify: `src/database/OfflineSnapshot.cpp` (pending_edits columns + drain translation)
- Modify: `src/MainWindow.cpp` (~3496-3520 onStoryCellEdited; the H6 dead-OCC block at ~254-270)
- Modify: `src/MainWindow.h` (`liveColumnForDataCol` -> `metricKeyForDataCol` if renaming; keep name if simpler)
- Tests: `tests/tst_livesync/tst_livesync.cpp`, `tests/tst_storedfns/tst_storedfns.cpp`, `tests/tst_databasemanager/tst_databasemanager.cpp`

- [ ] **Step 1: H25 filter (failing test first)**

Test (tst_databasemanager, against post-cutover `dve_test`): insert a `tpm` measurement AND a custom-key measurement for a loaded file's sample, call `loadFile`, assert `DataRow::extra` contains the custom key and does NOT contain `tpm`.
Run: FAILS (tpm appears in extra).

Then append to BOTH extras SELECTs in loadFile:

```sql
  AND md.key NOT IN (SELECT a.attname FROM pg_attribute a
                     WHERE a.attrelid = to_regclass('data_rows')
                       AND a.attnum > 0 AND NOT a.attisdropped)
```

(`'samples'` for the sample_headers SELECT.)
Update the big 3c comment block above them: the H25 paragraph flips from "3d must resolve" to "resolved - the filter below excludes every key that has a same-named wide column; the wide values arrive through the wide SELECTs above, which post-cutover read the name-holder views".
Re-run: PASSES.
Also re-run `tst_v3longformat`'s e2e gates (coil_temp must still round-trip - the filter must not eat open metrics).

- [ ] **Step 2: connect probe (failing test first)**

Test (tst_databasemanager): against the PRECUT database (pre-cutover shape), `DatabaseManager::open()` must fail with an error mentioning the runbook; against `dve_test` it succeeds.
Implementation in `DatabaseManager::open()` after `ensureSchema()`:

```cpp
// ── v3 Phase 3d (index D-3d-2): this build writes the long format ONLY. ──────
// Against a database whose data_rows is still a real table (the supervised
// v3.0.0 cutover has not been applied), saving would write measurements that
// no wide reader sees while the wide tables rot - silent split-brain. Refuse
// cleanly instead; the caller falls into the offline-snapshot path exactly as
// for an unreachable NAS.
{
    QSqlQuery probe(m_pg->queryDb());
    probe.exec(QStringLiteral(
        "SELECT c.relkind FROM pg_class c WHERE c.oid = to_regclass('data_rows')"));
    const QString kind = probe.next() ? probe.value(0).toString() : QString();
    if (kind != QLatin1String("v")) {
        m_lastError = QStringLiteral(
            "database is not cut over to the v3 long format (data_rows relkind='%1') - "
            "apply the v3.0.0 migration runbook before connecting this build")
                .arg(kind.isEmpty() ? QStringLiteral("absent") : kind);
        logDebug(m_lastError);
        close();
        return false;
    }
}
```

Match the surrounding open() error-path conventions (whatever close/cleanup the neighboring failure branches do - mirror them exactly).

- [ ] **Step 3: commitMeasurement + allowlist prune (failing tests first)**

tst_livesync: new test - `commitMeasurement` on a seeded sample upserts the measurement and the value is visible in `data_rows` (the view) at that (sample, sort_order); a second identical commit does not bump the measurement's version (H23 proof through the production path).
tst_storedfns: re-point its `dve_commit_cell`-on-data_rows tests to `dve_commit_measurement` (dve_commit_cell keeps its sensory-table tests untouched).

LiveSync.h/cpp:
- `bool commitMeasurement(qint64 sampleId, const QString& metricKey, int sortOrder, const QVariant& value);` - same throttle/queue plumbing as commitCell (pending map key grows a sortOrder; simplest: reuse the CellKey struct with table="measurements", rowId=sampleId, column=metricKey, and add an int sortOrder member with a comparator update).
- dispatch: `SELECT dve_commit_measurement(?, ?, ?, ?, ?)`.
- Allowlist: remove the `data_rows`, `samples`, `tests`, `files` entries from `isLiveSyncColumn`/`isLiveSyncTable` (audit-confirmed zero callers); keep both sensory tables.
- Offline branch: enqueue v2 (Step 4's shape).

MainWindow `onStoryCellEdited` (~3496):

```cpp
// Per-cell LiveSync, v3 long format (H4 re-point): keyed by the NATURAL
// identity (sample id, metric key, row ordinal). dr.id is the view's
// synthetic surrogate post-cutover and is deliberately not used.
if (m_liveSync && sample.id > 0) {
    const QString key = liveColumnForDataCol(col);   // same mapping, now a metric KEY
    if (!key.isEmpty() &&
        !m_liveSync->commitMeasurement(sample.id, key, dataRow, text)) {
        qWarning().noquote()
            << "[onStoryCellEdited] LiveSync rejected or could not queue the"
               " per-cell measurement commit; the edit is covered by the"
               " whole-file save. file=" << file->filePath
            << "sample=" << m_currentSampleIndex << "row=" << dataRow
            << "key=" << key << "sampleId=" << sample.id;
    }
}
```

(The `liveColumnForDataCol` mapping already returns exactly the metric keys - smell/clog/notes/puffing_regime/resistance; update its comment, keep the name.)
While here: read the H6 block at MainWindow.cpp:254-270 and LiveSync's `commitConflict`; if `commitConflict` still has zero consumers, delete the dead per-cell-OCC machinery (signal, m_versionLookup wiring, the 6-arg dve_commit_cell C++ call path if only data_rows used it - sensory uses the json variant); record the deletion in the index (H6: resolved by deletion).

- [ ] **Step 4: pending_edits v2 + drain translation (failing test first)**

tst_offlinesnapshot: new test - enqueue a v2 measurement edit (sample_id, key, sort_order), drain against the post-cutover DB, assert the measurement landed; enqueue a fabricated v1 data_rows edit whose row_id exists in `data_rows_pre_v3`, drain, assert translated; one whose row_id is absent, drain, assert dropped-with-warning and not retried.

OfflineSnapshot:
- Queue DDL block (~2298-2334): add `schema_version INTEGER` and `sort_order INTEGER` via the existing PRAGMA/ALTER-if-missing pattern.
- `enqueueCellEdit` gains an overload/params for v2 (schema_version=2, sort_order); v1 writers (sensory) keep enqueueing as today (schema_version NULL/1).
- `drainPendingEdits` (~2364): the replay callback grows the fields; in LiveSync's replay lambda (~322):

```cpp
// v3 Phase 3d (H9): three queue generations drain differently.
//  - v2 measurement rows: replay through dve_commit_measurement.
//  - v1 rows for data_rows: enqueued by a pre-cutover build against real
//    data_rows ids. Translate id -> (sample_id, sort_order) through the
//    frozen pre-image data_rows_pre_v3, then commit as a measurement. A row
//    that is not there (deleted before cutover) is dropped with a warning -
//    it has no honest target.
//  - v1 sensory rows: unchanged, dve_commit_cell path.
```

Translation SQL: `SELECT sample_id, sort_order FROM data_rows_pre_v3 WHERE id = ?` on the live connection.

- [ ] **Step 5: Run tst_livesync, tst_storedfns, tst_offlinesnapshot, tst_databasemanager**

Expected: the new tests green; remaining reds in livesync/twoclient confined to whole-file-save assert phases (Task 6).

- [ ] **Step 6: Commit**

```bash
git add src/database/DatabaseManager.cpp src/database/LiveSync.h src/database/LiveSync.cpp src/database/OfflineSnapshot.cpp src/MainWindow.cpp src/MainWindow.h tests/
git commit -m "feat(v3): H25 read filter, post-cutover connect probe, dve_commit_measurement live path, pending_edits v2 drain-or-translate"
```

---

## Task 6: persistFileCore - the standard-metric write cutover

**Files:**
- Modify: `src/database/DatabaseOps.cpp` (persistFileCore: statements ~571-609, phase A ~421-552, phase B row loop ~895-961 and sample upsert ~807-893, phase C ~1105-1165; freshChildVersion ~107)
- Modify: `src/database/MetricDefCache.h/.cpp` (lookup-only resolution for standard keys, if not already expressible)
- Tests: `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp`, `tests/tst_twoclient_e2e/tst_twoclient_e2e.cpp`

This is the H1 lift.
Re-read the two banner comments (phase A pre-image, phase C prune) before editing; every clause below is the "lifted DELIBERATELY, together with a write path that reproduces every one of them" they demand.

- [ ] **Step 1: The gate scenario, failing first**

`scenario24_standardMetricsSurviveLongFormatRoundTrip` in tst_saveintegrity_e2e: build a FileResult (the suite's existing builder), save, assert per-metric measurement counts (13 x rows, honoring the regime gate) and ZERO wide writes (data_rows_pre_v3 count for those samples unchanged); reload via loadFile, assert every DataRow member and all 22 sample fields byte-match memory; edit one cell + delete one row, re-save, assert the deleted row's measurements are GONE (prune) and the edit visible; save a third time UNCHANGED and assert `max(version)` over the file's measurements/sample_headers is unchanged (no-op suppression end-to-end).
Run: fails at the first save (old code UPDATEs the data_rows view).

- [ ] **Step 2: persistFileCore changes, in order**

a. **Samples statements go narrow, targeting samples_core:**

```cpp
if (!updateSample.prepare(
        "UPDATE samples_core SET test_id = ?, sort_order = ?, updated_by = ? "
        "WHERE id = ? AND version = ? RETURNING version") ||
    !insertSample.prepare(
        "INSERT INTO samples_core (test_id, sort_order, updated_by) "
        "VALUES (?, ?, ?) RETURNING id, version")) {
```

Bind loop shrinks accordingly (id/version OCC + classifyMissingUpdate("samples_core") + freshChildVersion("samples_core") - the FOR UPDATE in freshChildVersion must lock the real table, not the grouped view).
`classifyMissingUpdate`'s samples call sites pass "samples_core" too.

b. **The 22 header values join the header batch.** Before the existing `guardHeaders` extras loop, build the standard header entries from the members through the same `ExtraRow`/`flushExtras` machinery:

```cpp
// v3 Phase 3d: the 22 standard headers write through the SAME batch as the
// open ones. Resolution is lookup-only - a standard key missing from
// metric_defs means the seed/ensureSchema contract is broken and the save
// must abort loudly, never auto-register with an inferred type.
struct StdHeader { const char* key; QVariant value; };
const StdHeader stdHeaders[] = {
    { "sample_name",        sr.sampleName },
    { "sample_id",          sr.sampleID },
    { "date",               sr.date },
    { "tester",             sr.tester },
    { "media",              sr.media },
    { "viscosity",          sr.viscosity },
    { "resistance",         sr.resistance },
    { "voltage",            sr.voltage },
    { "power",              sr.power },
    { "heating_technology", sr.heatingTechnology },
    { "puffing_regime",     sr.puffingRegime },
    { "initial_oil_mass",   sr.initialOilMass },
    { "average_tpm",        sr.averageTPM },
    { "stddev_tpm",         sr.stdDevTPM },
    { "avg_power_density",  sr.averagePowerDensity },
    { "efficiency_percent", sr.efficiencyPercent },
    { "total_oil_consumed", sr.totalOilConsumed },
    { "total_puffs",        sr.totalPuffs },
    { "normalized_tpm",     sr.normalizedTPM },
    { "burn_status",        sr.burnStatus },
    { "clog_status",        sr.clogStatus },
    { "leak_status",        sr.leakStatus },
};
```

Append each through a new `appendStandard` twin of `appendExtra`: same null-skip (a null QString was a NULL wide bind - sparse), same encode-by-value_type, but resolution via lookup-only and NO wide-column skip (these ARE the wide columns), abort the save on a missing key.

c. **The 13 row metrics replace the updateRow/insertRow loop.** The whole `dr.id != -1` branch pair goes; in its place, per row, batch:

```cpp
const StdMetric stdMetrics[] = {
    { "puffs",             QVariant(dr.puffs) },
    { "before_weight",     QVariant(dr.beforeWeight) },
    { "after_weight",      QVariant(dr.afterWeight) },
    { "draw_pressure",     QVariant(dr.drawPressure) },
    { "resistance",        QVariant(dr.resistance) },
    { "smell",             QVariant(dr.smell) },
    { "clog",              QVariant(dr.clog) },
    { "notes",             QVariant(dr.notes) },
    { "tpm",               QVariant(dr.tpm) },
    { "tpm_power_density", QVariant(dr.tpmPowerDensity) },
    { "variation_tpm",     QVariant(dr.variationTPM) },
    { "oil_consumed",      QVariant(dr.oilConsumed) },
    { "puffing_regime",    sheet.hasPerRowRegime ? QVariant(dr.puffingRegime)
                                                 : QVariant() },  // sparse: skipped when absent
};
```

into the measurements batch (sortOrder = ri), flushed through the existing `flushExtras` (rename it `flushLongRows` if the name now grates; keep one function).
`dr.id`/`dr.version` are left untouched in the struct (loadFile still fills them; nothing writes through them - note it in a comment).
The extras pass stays exactly as-is after it (same batch vectors, so standard + extra share flushes and the 500-row cap).

d. **The guard becomes a hard requirement.** `longFormat`/`guardRows`/`guardHeaders` were best-effort in 3c (log and save wide anyway).
There is no wide fallback any more: if `metricDefs.load` fails or either catalog set is empty, the save must FAIL (rollback, OtherError, message naming the probe) - a half-save with no standard metrics is data loss.

e. **Pre-image lifts the exclusion (H1/H25 same breath, D-3d-4):** the two `preImageExcluding` calls lose their `AND md.key NOT IN (%1)` clause and bound sets (keep the helper for its error plumbing or inline it); rewrite the big H1 banner to its post-state: "the writer reproduces EVERY metric of every surviving row, so the full pre-image is prunable by pre-minus-post; the wide-column skip survives only in appendExtra, where it guards a stale recovery-JSON extra from double-writing a standard key".
The `preDataRowIds` SELECT and its `pruneOrphans("data_rows", ...)` call are DELETED (no data_rows table); `pruneOrphans("samples", ...)` becomes `pruneOrphans("samples_core", ...)`.

f. **updateRow/insertRow/updateImage prepares:** delete the row pair; images unchanged.

- [ ] **Step 3: Run scenario24 - green. Then the whole of tst_saveintegrity_e2e.**

Scenarios that INVERT and need edits (do them now):
- scenario18 (`noOpenMetrics_writesNoLongRows`): post-cutover every save writes long rows; re-scope it to "no open metrics writes no NON-standard rows" (count measurements with keys outside the wide set = 0).
- scenario23 (`wideColumnCollisionIsNotWritten`): the skip now means "a standard key arriving via extra is not written TWICE"; assert exactly one measurement for the colliding key with the MEMBER's value, not the extra's.
- scenario21 (deleted row takes its measurements) must now pass THROUGH the lifted prune - strengthen its assert to include standard keys.
- scenario17 (dbBrowserLoadKeepsWriteback) and scenario11 (removeSample) should pass unmodified - they assert through loadFile.

- [ ] **Step 4: tst_twoclient_e2e + tst_livesync full runs**

Two-client expectations to verify (and adjust asserts, not behavior): concurrent whole-file saves land row-level last-writer-wins as before (files FOR UPDATE serializes); a second client's identical save bumps NOTHING (H23+D4 through two real clients - add this assert if the suite lacks it).

- [ ] **Step 5: Full DB-suite sweep**

Run: tst_saveintegrity_e2e, tst_twoclient_e2e, tst_livesync, tst_offlinesnapshot, tst_storedfns, tst_notificationlistener, tst_databasemanager, tst_v3longformat.
tst_notificationlistener: fix its data_rows/samples NOTIFY assertions - post-cutover the payload table name for sample updates is `samples_core` and data-row traffic appears as NO notify (D3: measurements carry none); assert the new truth, do not re-add triggers.
Expected: ALL green.
This closes the red window opened in Task 3.

- [ ] **Step 6: Commit**

```bash
git add src/database/DatabaseOps.cpp src/database/MetricDefCache.h src/database/MetricDefCache.cpp tests/
git commit -m "feat(v3): standard-metric write cutover - batched keyed upserts, narrow samples_core, prune guard lifted with the writer (H1/H2/H8/H22)"
```

---

## Task 7: full-app verification + snapshot E2E + gates

**Files:** none new; fixes only where a gate fails.

- [ ] **Step 1: Clean `-Werror` app build** (decrypt first: `python tools/decrypt_via_copy.py --apply`).
- [ ] **Step 2: Full test suite** via `tests\run-tests.ps1` (NOT concurrently with an app build - documented CPU-starvation flake).
  Expected: all green; re-run any python-subprocess suite that fails standalone before treating it as real.
- [ ] **Step 3: Corpus gates** - corpus shadow harness and corpus round-trip harness at their 2c/3a baselines (26/0/6 and 38/0/0).
  These are parse-side and must be UNMOVED; any diff is a regression in something this phase had no business touching.
- [ ] **Step 4: Offline snapshot E2E** - run the tst_offlinesnapshot regen tests against post-cutover PG and eyeball the regen wall-clock in the log output vs a prior run (the view-pivot watch-item); note the number in the index.
- [ ] **Step 5: The 3c e2e gates** (manifest demo coil_temp through Postgres AND snapshot) - green, proving open metrics were not disturbed by the standard cutover.
- [ ] **Step 6: Commit any fixes**

```bash
git add -A
git commit -m "test(v3): 3d full-suite and corpus gates green"
```

---

## Task 8: the v3.0.0 migration runbook

**Files:**
- Create: `docs/superpowers/plans/2026-08-26-v3-migration-runbook.md`

- [ ] **Step 1: Write it.** Contents (each a numbered supervised step with its exact command and its verification query):
  1. Preconditions: every client closed (presence table empty), Synology freeze confirmed still in force, owner at the NAS.
  2. `pg_dump` backup + verify size/restorability (D7). THE rollback path.
  3. Apply `2026-08-03-oil-units-to-grams.sql` (D10: units BEFORE shape; it must not run while any pre-grams client can still write).
  4. Apply `2026-07-31-v3-long-format.sql`.
  5. Optional inspection: `SELECT * FROM dve_migrate_to_long_format();` + the parity/sparse spot-checks (copy the exact SQL from the cutover function's gates into the doc).
  6. Apply `2026-08-26-v3-cutover.sql` (idempotently re-migrates, verifies, cuts; paste the expected NOTICE output).
  7. Post-checks: relkind probes, `schema_meta` keys (`v3_long_format_cutover`, the oil key), a loadFile-shaped SELECT.
  8. Install v3.0.0 on one machine, open a known file, verify; then release to the rest.
  9. The window caveat: any v2.10.0 client that connects between 6 and its own upgrade will fail its saves with SQL errors (transactional, no corruption) - keep the window short; the runbook is a to-do list for the owner, who handles all NAS/Synology steps personally per standing rule.
  10. D7 note: the prod-dump REHEARSAL (restore last NAS backup locally, run steps 3-7 against it) still needs owner go-ahead and should happen before the real thing.
- [ ] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-08-26-v3-migration-runbook.md
git commit -m "docs(v3): supervised v3.0.0 migration runbook (oil -> long-format -> cutover)"
```

---

## Task 9: index + tracker + memory wrap

**Files:**
- Modify: `docs/superpowers/plans/2026-07-30-tpm-v3-phase3-INDEX.md`
- Modify: `docs/sprint-tracker.html` (per the sprint-tracker skill)

- [ ] **Step 1: Index updates:** record D-3d-1..8 in the locked-decisions section; mark 3d's sub-phase entry DONE with its gate results; hazard ledger sweep - H1, H2, H3 (verified, no change needed), H4, H5 (name the dedicated test added or explicitly accept), H6 (deleted or wired - whichever Step 5.3 found), H8, H9, H13 (dissolved by keyed upserts), H20 (probe live), H21 (fixed Task 1), H22 (re-verified through the lifted prune), H23 (fixed), H25 (fixed) - each gets its closing note; add the snapshot-regen perf number.
- [ ] **Step 2: Tracker refresh** (sprint-tracker skill; new commits, 3d done, next = internal smoke build).
- [ ] **Step 3: Commit**

```bash
git add docs/superpowers/plans/2026-07-30-tpm-v3-phase3-INDEX.md docs/sprint-tracker.html
git commit -m "docs(v3): Phase 3d wrap - decisions D-3d-1..8, hazard ledger closed out"
```

---

## Out of scope for 3d (explicitly)

- Running ANY migration against the NAS (supervised runbook, owner-driven, after v3.0.0 is otherwise ready).
- The D7 prod-dump rehearsal (needs owner go-ahead; the runbook documents it).
- Dropping `data_rows_pre_v3` / `samples_core`'s relic columns (a post-v3.0.0 cleanup once stability is proven).
- Phase 4 items: deleting the legacy parser/shadow harness, displaying extras, activating the did_burn/clog/leak header cells, per-measurement live UI.
- H5's dedicated sample-header-edit regression test IF it turns out one already exists in the 23 scenarios (audit in Task 6 Step 3; add only if genuinely missing).

## Self-review notes

- Spec coverage: 3d index bullet-list -> Task 6 (persistFileCore, prune, rename consumption), Task 2+3 (real migration machinery + rehearsal), Task 5 (dve_commit_measurement, pending_edits policy), Task 8 (runbook). H-ledger items each named in a task.
- The mid-phase red window (Tasks 3-6) is deliberate, bounded, and named in Task 3 Step 5.
- Type consistency: `ExtraRow`/`flushExtras` reused for standard batches; `commitMeasurement(qint64, QString, int, QVariant)` consistent across LiveSync.h, MainWindow call, and tests; `samples_core` spelled identically in SQL, C++, and helper.
