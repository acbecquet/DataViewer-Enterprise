-- TPM template v3, Phase 3b - long-format schema, compatibility views, and the
-- one-shot wide -> long migration function.
--
-- Plan: docs/superpowers/plans/2026-07-31-tpm-v3-phase3b-schema-migration-plan.md
-- Locked decisions: docs/superpowers/plans/2026-07-30-tpm-v3-phase3-INDEX.md
--
-- WHAT THIS FILE DOES
--   Creates `metric_defs`, `measurements`, `sample_headers` (EMPTY except for
--   the metric_defs seed), the read-only `data_rows_v` / `samples_v` bridges,
--   and `dve_migrate_to_long_format()`.
--
-- WHAT THIS FILE DOES NOT DO
--   It does not migrate a single row of data. `dve_migrate_to_long_format()` is
--   authored here and left UNCALLED; see the banner on the function itself.
--   No app read or write path changes in 3b - `data_rows` and `samples` stay
--   authoritative and untouched.
--
-- Idempotent throughout (IF NOT EXISTS / CREATE OR REPLACE / ON CONFLICT DO
-- NOTHING) because the test-container provisioning script and a future
-- ensureSchema() may both reach these objects. One BEGIN/COMMIT block. No
-- scheduled-job statements live here - stock postgres:16 in the test container
-- cannot run them and the provisioning script would strip them.

BEGIN;

-- ============================================================================
-- 1. Tables
-- ============================================================================

-- ── metric_defs: the persisted projection of the compiled MetricRegistry ─────
-- UNIQUE is (kind, key), NEVER (key). `resistance` and `puffing_regime` each
-- exist as BOTH a data_rows column (kind 'metric') and a samples column (kind
-- 'header'), and the ratified registry keeps metrics and header fields in
-- SEPARATE namespaces (spec section 5, rule 1 / D11). A single-column unique
-- constraint would make this schema structurally unable to represent the data
-- that is already in the database.
--
-- No CHECK on value_type on purpose: the ValueType enum grew by three members
-- in the registry ratification (spec 8.4) and will grow again. ensureSchema()
-- is best-effort and never throws, so a CHECK that rejected a new type would
-- turn a registry addition into a SILENT upsert failure.
CREATE TABLE IF NOT EXISTS metric_defs (
    id           BIGSERIAL PRIMARY KEY,
    kind         TEXT NOT NULL CHECK (kind IN ('metric', 'header')),
    key          TEXT NOT NULL,
    display_name TEXT NOT NULL,
    value_type   TEXT NOT NULL,   -- Manifest.cpp kTypeNames spelling: number/text/bool/mixed/numberlist/image
    unit         TEXT,            -- NULL = dimensionless; see the seed note on empty-string convergence
    role         TEXT,            -- Manifest.cpp kRoleNames spelling; NULL for kind='header' (HeaderFieldDef has no role)
    tags         JSONB,
    updated_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by   TEXT        NOT NULL DEFAULT 'migration',
    version      INTEGER     NOT NULL DEFAULT 1,
    UNIQUE (kind, key)
);

-- ── measurements: one row per (sample, metric, sort_order) ───────────────────
-- sort_order is NOT NULL here even though data_rows.sort_order is nullable.
-- That is deliberate - it is part of the identity of a measurement. The
-- migration function refuses to run against data that would violate it rather
-- than inventing a value (see its pre-flight).
--
-- No separate index on (sample_id): the UNIQUE constraint's btree is
-- (sample_id, metric_id, sort_order) and Postgres uses a leftmost prefix, so it
-- already serves the per-sample read. A redundant single-column index would
-- only add write cost to what will be the largest table in the database
-- (~13 rows per data_rows row). DEVIATION from plan Task 1 step 2, which asks
-- for the extra index; the intent (fast per-sample read) is satisfied.
CREATE TABLE IF NOT EXISTS measurements (
    id         BIGSERIAL PRIMARY KEY,
    sample_id  BIGINT  NOT NULL REFERENCES samples(id) ON DELETE CASCADE,
    metric_id  BIGINT  NOT NULL REFERENCES metric_defs(id),
    sort_order INTEGER NOT NULL,
    value_num  DOUBLE PRECISION,
    value_text TEXT,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by TEXT        NOT NULL DEFAULT 'migration',
    version    INTEGER     NOT NULL DEFAULT 1,
    UNIQUE (sample_id, metric_id, sort_order)
);

-- ── sample_headers: one row per (sample, header field) ───────────────────────
-- Same index reasoning as measurements: UNIQUE (sample_id, field_id) already
-- indexes sample_id as its leftmost column.
CREATE TABLE IF NOT EXISTS sample_headers (
    id         BIGSERIAL PRIMARY KEY,
    sample_id  BIGINT NOT NULL REFERENCES samples(id) ON DELETE CASCADE,
    field_id   BIGINT NOT NULL REFERENCES metric_defs(id),
    value_num  DOUBLE PRECISION,
    value_text TEXT,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by TEXT        NOT NULL DEFAULT 'migration',
    version    INTEGER     NOT NULL DEFAULT 1,
    UNIQUE (sample_id, field_id)
);

-- ============================================================================
-- 2. Triggers
-- ============================================================================

-- bump_version(), copied VERBATIM from
-- deploy/postgres/migrations/2026-07-15-notify-noop-suppression.sql, with
-- 'measurements' and 'sample_headers' added to the no-op suppression list
-- (index decision D4) and NOTHING else changed. A whole-file save re-writes
-- every measurement of every row; without this, unchanged measurements would
-- churn their version on every save and defeat the per-measurement optimistic
-- locking that is the point of the long format.
--
-- metric_defs is deliberately NOT in the suppression list: it is a small
-- vocabulary table that ensureSchema() upserts, not a save-path table, and D4
-- names only the two data tables.
CREATE OR REPLACE FUNCTION bump_version() RETURNS TRIGGER AS $$
BEGIN
  -- No-op UPDATE on a SESSION table (nothing material changed): keep the row
  -- inert so it doesn't churn version/updated_at or (via notify_row_change)
  -- storm live clients. Scoped to the whole-row session tables where the storm
  -- was observed; the TPM files/tests/samples/data_rows hierarchy KEEPS its
  -- always-bump behavior (a whole-file save's parent-row touch is part of that
  -- path's established multi-client OCC semantics and is out of scope here).
  -- v3 Phase 3b: measurements / sample_headers join the scoped list per index
  -- decision D4 - a whole-file save rewrites every child measurement, and an
  -- unchanged one must not consume a version.
  IF TG_OP = 'UPDATE'
     AND TG_TABLE_NAME IN ('sensory_sessions', 'detailed_sensory_sessions',
                           'measurements', 'sample_headers')
     AND (to_jsonb(NEW) - 'version' - 'updated_at' - 'app_version')
         IS NOT DISTINCT FROM
         (to_jsonb(OLD) - 'version' - 'updated_at' - 'app_version')
  THEN
    NEW.version    := OLD.version;      -- undo any client-supplied version = version + 1
    NEW.updated_at := OLD.updated_at;
    RETURN NEW;
  END IF;
  -- Real change: refuse client-side version manipulation; we always +1.
  NEW.version    := OLD.version + 1;
  NEW.updated_at := now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DO $$
DECLARE
  t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY['metric_defs', 'measurements', 'sample_headers'] LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_bump_version ON %I;
       CREATE TRIGGER trg_%I_bump_version
       BEFORE UPDATE ON %I
       FOR EACH ROW EXECUTE FUNCTION bump_version();',
       t, t, t, t
    );
  END LOOP;
END$$;

-- ── DELIBERATELY ABSENT: notify_row_change() on these three tables ───────────
-- Index decision D3. This is not an oversight and must not be "fixed".
--
-- The 2026-07-30 audit proved TPM row NOTIFYs drive ZERO UI today:
-- MainWindow::handleRemoteRowChange (src/MainWindow.cpp:3787-3822) handles only
-- op == "DELETE" for the three top-level resource tables and returns for
-- everything else, and LiveSync::cellChanged (src/database/LiveSync.cpp:530)
-- has exactly two consumers, both sensory JSON-path handlers - there is no TPM
-- subscriber at all.
--
-- Meanwhile one whole-file save fans ~1,500 data_rows UPDATEs into roughly
-- 13,000 measurement UPDATEs. Attaching the notify trigger would turn a single
-- save into a 13,000-NOTIFY broadcast to every live client: the DV-23 storm
-- class, for zero UI benefit. Resource-level NOTIFY on `files` is what
-- RowDeletedBanner actually consumes and it is untouched.
--
-- If a future phase needs per-measurement live updates, add a purpose-built
-- notifier with coalescing - do not attach notify_row_change() as-is.

-- ============================================================================
-- 3. metric_defs seed
-- ============================================================================
--
-- Covers EXACTLY the 13 data_rows value columns (kind 'metric') and the 22
-- samples value columns (kind 'header'), so dve_migrate_to_long_format() can
-- map every wide column. It is a floor, not a ceiling: ensureSchema() upserts
-- the full compiled MetricRegistry on top of this seed (plan Task 2), which
-- adds the metrics and header fields that have no wide column yet
-- (tpm_puff_density, puff_volume, resistance_final, did_burn, ...). Rows are
-- never deleted - registry naming policy rule 1, "keys are forever", so a key
-- that leaves the registry keeps its row and historical measurements still
-- resolve.
--
-- Display names / units / roles for the keys that EXIST in the compiled
-- registry are copied character for character from src/model/MetricRegistry.cpp
-- so that ensureSchema()'s ON CONFLICT DO UPDATE is a no-op rather than a
-- silent rewrite. Contract for that upsert: the registry stores "no unit" and
-- "no tags" as an EMPTY QString / empty QMap, and this seed stores them as SQL
-- NULL - ensureSchema must bind NULLIF(unit, '') (and a NULL jsonb for an empty
-- tag map) or the two representations will fight on every connect.
--
-- 11 of the 22 header keys have NO compiled-registry entry: sample_name is an
-- app-side label, average_tpm / stddev_tpm / avg_power_density /
-- efficiency_percent / total_oil_consumed / total_puffs / normalized_tpm are
-- the derived AGGREGATES of registry section 4, and burn_status / clog_status /
-- leak_status are the status_scan_v1 outputs (distinct from the registry's
-- did_burn / did_clog / did_leak header questions, which are read off the
-- sheet). They are seeded as kind 'header' because sample_headers is where the
-- migration must land them; ensureSchema will never touch these rows. Their
-- units are left NULL where the registry has not ratified one.

INSERT INTO metric_defs (kind, key, display_name, value_type, unit, role, tags) VALUES
  -- The 13 data_rows value columns, in CANONICAL projection order (the order at
  -- src/database/OfflineSnapshot.cpp:940-945). That is not always data_rows'
  -- PHYSICAL catalog order - where puffing_regime sits depends on the database's
  -- vintage; see the data_rows_v comment below. Seed order is cosmetic either
  -- way: the migration maps by key, never by position.
  ('metric', 'puffs',             'puffs',              'number', 'count',    'measured',    NULL),
  ('metric', 'before_weight',     'Before Weight (g)',  'number', 'g',        'measured',    NULL),
  ('metric', 'after_weight',      'After Weight (g)',   'number', 'g',        'measured',    NULL),
  ('metric', 'draw_pressure',     'Draw Pressure',      'number', 'kPa',      'measured',    NULL),
  ('metric', 'resistance',        'Resistance',         'number', 'ohm',      'measured',    NULL),
  ('metric', 'smell',             'Smell',              'text',   NULL,       'qualitative',
      '{"scale": "0-4; 1-4-era sheets leave blank for no event (D9)"}'),
  ('metric', 'clog',              'Clog',               'text',   NULL,       'qualitative', NULL),
  ('metric', 'notes',             'Notes',              'text',   NULL,       'qualitative', NULL),
  ('metric', 'tpm',               'TPM',                'number', 'mg/puff',  'derived',     NULL),
  ('metric', 'tpm_power_density', 'TPM Power Density',  'number', 'mg/(W*s)', 'derived',
      '{"approximation": "no transient power; TCR shifts instantaneous power (9.1)"}'),
  ('metric', 'variation_tpm',     'Variation in TPM',   'number', '%',        'derived',
      '{"definition": "template rolling CV in percent is canonical (D1)"}'),
  ('metric', 'oil_consumed',      'Oil Consumed',       'number', 'g',        'derived',     NULL),
  ('metric', 'puffing_regime',    'Puffing Regime',     'text',   NULL,       'qualitative',
      '{"encoding": "legacy composite string; canonical form is the 4 split metrics (9.2)"}'),

  -- The 22 samples value columns, in samples column order. role stays NULL:
  -- HeaderFieldDef carries no role, so ensureSchema writes NULL for every
  -- header and anything else here would be rewritten on first connect.
  ('header', 'sample_name',        'Sample Name',            'text',   NULL,      NULL, NULL),
  ('header', 'sample_id',          'Sample ID',              'text',   NULL,      NULL, NULL),
  ('header', 'date',               'Date',                   'text',   NULL,      NULL, NULL),
  ('header', 'tester',             'Tester',                 'text',   NULL,      NULL, NULL),
  ('header', 'media',              'Media',                  'text',   NULL,      NULL, NULL),
  ('header', 'viscosity',          'Viscosity',              'number', 'cP',      NULL, NULL),
  ('header', 'resistance',         'Resistance',             'number', 'ohm',     NULL, NULL),
  ('header', 'voltage',            'Voltage',                'number', 'V',       NULL, NULL),
  ('header', 'power',              'Power',                  'number', 'W',       NULL, NULL),
  ('header', 'heating_technology', 'Heating Technology',     'text',   NULL,      NULL, NULL),
  ('header', 'puffing_regime',     'Puffing Regime',         'text',   NULL,      NULL, NULL),
  ('header', 'initial_oil_mass',   'Initial Oil Mass',       'number', 'g',       NULL, NULL),
  -- Aggregates (registry section 4) - no compiled-registry entry.
  ('header', 'average_tpm',        'Average TPM',            'number', 'mg/puff', NULL, NULL),
  ('header', 'stddev_tpm',         'TPM Standard Deviation', 'number', 'mg/puff', NULL, NULL),
  -- Unit deliberately NULL: the mean_over_power aggregates are era-dependent
  -- (spec 2.1 / 9.1 - mg/(W*s) vs mg/(puff*W)) and no ruling picks one.
  ('header', 'avg_power_density',  'Average Power Density',  'number', NULL,      NULL, NULL),
  ('header', 'efficiency_percent', 'Usage Efficiency',       'number', '%',       NULL, NULL),
  -- Unit deliberately NULL: registry section 4 derives this from oil_consumed
  -- (grams, D4) but SheetProcessors.cpp:307-310 divides it by an mg quantity.
  -- Unresolved; do not guess in a seed.
  ('header', 'total_oil_consumed', 'Total Oil Consumed',     'number', NULL,      NULL, NULL),
  ('header', 'total_puffs',        'Total Puffs',            'number', 'count',   NULL, NULL),
  ('header', 'normalized_tpm',     'Normalized TPM',         'number', NULL,      NULL, NULL),
  ('header', 'burn_status',        'Burn Status',            'text',   NULL,      NULL, NULL),
  ('header', 'clog_status',        'Clog Status',            'text',   NULL,      NULL, NULL),
  ('header', 'leak_status',        'Leak Status',            'text',   NULL,      NULL, NULL)
ON CONFLICT (kind, key) DO NOTHING;

-- ============================================================================
-- 4. Privileges (hazard H16)
-- ============================================================================
--
-- The sensory_web role (deploy/postgres/migrations/2026-06-25-dv11-sensory-web-role.sql)
-- is the phone form's least-privilege login. Its REVOKE loop ran before these
-- tables existed, so every new table needs its own explicit revoke or the role
-- silently gains reach over lab data. The negative test that guards this class
-- lives at deploy/sensory-collect/tests/test_append.py:405-408.
REVOKE ALL ON metric_defs, measurements, sample_headers FROM PUBLIC;

DO $$
BEGIN
  IF EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'sensory_web') THEN
    REVOKE ALL ON metric_defs, measurements, sample_headers FROM sensory_web;
  END IF;
END$$;

-- ============================================================================
-- 5. Compatibility views (read-only bridges - index decision D1)
-- ============================================================================
--
-- These pivot the long tables back to the wide shape so the offline-snapshot
-- regen, MigrationTool, and any external SQL keep working while the C++ read
-- and write paths are cut over one at a time. They are READ-ONLY: no INSTEAD OF
-- triggers, by decision. Postgres does not fire a base table's row triggers
-- through a view, so a writable view would have to reimplement bump_version,
-- and optimistic locking would land on a synthetic view row - defeating the
-- per-measurement versioning that is the whole point.
--
-- In 3b they exist only so the rehearsal harness can prove parity. They become
-- load-bearing in 3d.
--
-- KNOWN LIMIT OF THE SPARSE RULE: a source row whose value columns are ALL NULL
-- produces zero long rows, so it has no group and DISAPPEARS from the view.
-- data_rows cannot express "this row exists but holds nothing" in the long
-- format. Today every data_rows numeric column carries DEFAULT 0.0 and the save
-- path always binds a value, so an all-NULL data row should not exist - but
-- MigrationTool copied legacy rows verbatim and nothing enforces it, so the
-- rehearsal must assert row counts on both sides (plan Task 5) and the
-- production-dump rehearsal (index D7) must re-check before the 3d cutover.
-- For samples the exposure is milder: samples_v is a JOIN partner, so an
-- all-NULL-header sample still appears as long as readers LEFT JOIN it.

CREATE OR REPLACE VIEW data_rows_v AS
SELECT
    -- SYNTHETIC. See COMMENT ON COLUMN below - never a write-back identity.
    MIN(m.id)                                                          AS id,
    m.sample_id                                                        AS sample_id,
    m.sort_order                                                       AS sort_order,
    MAX(m.value_num)  FILTER (WHERE md.key = 'puffs')                  AS puffs,
    MAX(m.value_num)  FILTER (WHERE md.key = 'before_weight')          AS before_weight,
    MAX(m.value_num)  FILTER (WHERE md.key = 'after_weight')           AS after_weight,
    MAX(m.value_num)  FILTER (WHERE md.key = 'draw_pressure')          AS draw_pressure,
    MAX(m.value_num)  FILTER (WHERE md.key = 'resistance')             AS resistance,
    MAX(m.value_text) FILTER (WHERE md.key = 'smell')                  AS smell,
    MAX(m.value_text) FILTER (WHERE md.key = 'clog')                   AS clog,
    MAX(m.value_text) FILTER (WHERE md.key = 'notes')                  AS notes,
    MAX(m.value_num)  FILTER (WHERE md.key = 'tpm')                    AS tpm,
    MAX(m.value_num)  FILTER (WHERE md.key = 'tpm_power_density')      AS tpm_power_density,
    MAX(m.value_num)  FILTER (WHERE md.key = 'variation_tpm')          AS variation_tpm,
    MAX(m.value_num)  FILTER (WHERE md.key = 'oil_consumed')           AS oil_consumed,
    MAX(m.value_text) FILTER (WHERE md.key = 'puffing_regime')         AS puffing_regime,
    MAX(m.updated_at)                                                  AS updated_at,
    (ARRAY_AGG(m.updated_by ORDER BY m.updated_at DESC, m.id DESC))[1] AS updated_by,
    MAX(m.version)                                                     AS version
FROM measurements m
JOIN metric_defs md ON md.id = m.metric_id AND md.kind = 'metric'
GROUP BY m.sample_id, m.sort_order;

COMMENT ON VIEW data_rows_v IS
    'Read-only pivot of measurements back to the data_rows wide shape (v3 Phase 3b, index D1). '
    'The column NAME SET matches data_rows exactly, 19 for 19. The column ORDER matches the '
    'canonical projection every in-repo reader uses (src/database/OfflineSnapshot.cpp:940-945). '
    'It does NOT necessarily match the PHYSICAL catalog order of data_rows, which is '
    'database-VINTAGE-dependent and therefore something no view can promise: a database built '
    'from a current init.sql has puffing_regime inline in canonical position, but one that '
    'predates that column received it by ALTER (the migration file, or ensureSchema''s '
    'kAdditiveColumns) and carries it physically LAST. Both shapes are live - a freshly '
    'provisioned dve-test-pg is the former, production is the latter - so an order-sensitive '
    'reader would pass its tests here and break on the NAS. Read this view BY NAME, never by '
    'position; that matters most in 3d, where this view takes the data_rows name. '
    'A NULL value column means NO measurement row exists '
    'for that (sample, metric, sort_order) - which under sparse materialization (index D2) is '
    'exactly what a NULL in the source data_rows column meant. A numeric 0 IS a real measurement '
    'and appears as 0, not NULL. Not updatable and deliberately has no INSTEAD OF trigger.';

COMMENT ON COLUMN data_rows_v.id IS
    'SYNTHETIC SURROGATE - MIN(measurements.id) over the group. This is NOT a data_rows.id and it '
    'is NOT stable: it changes if the lowest-id measurement of the row is deleted or re-inserted. '
    'Adequate for read-only consumers that just need a per-row key. NOTHING may use it as an '
    'identity to write back through; the writable identity is (sample_id, sort_order).';

-- samples_v exposes ONLY the 22 header columns. samples keeps its own identity
-- and audit columns (test_id, sort_order, updated_at, updated_by, version), so
-- this is a JOIN partner - `samples s JOIN samples_v v ON v.id = s.id` - and
-- NOT a drop-in replacement for samples the way data_rows_v is for data_rows.
--
-- Note the naming: the group key is called `id` (it IS samples.id), because
-- `sample_id` is already taken by the TEXT business identifier that is one of
-- the 22 header columns. Getting this backwards would collide two different
-- columns onto one name.
CREATE OR REPLACE VIEW samples_v AS
SELECT
    -- REAL, stable samples.id - unlike data_rows_v.id, this is the group key,
    -- not a synthetic MIN.
    sh.sample_id                                                          AS id,
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
    -- samples.total_puffs is INTEGER; value_num is DOUBLE PRECISION. Cast back
    -- so the view's type matches the wide column it stands in for.
    (MAX(sh.value_num) FILTER (WHERE md.key = 'total_puffs'))::INTEGER    AS total_puffs,
    MAX(sh.value_num)  FILTER (WHERE md.key = 'normalized_tpm')           AS normalized_tpm,
    MAX(sh.value_text) FILTER (WHERE md.key = 'burn_status')              AS burn_status,
    MAX(sh.value_text) FILTER (WHERE md.key = 'clog_status')              AS clog_status,
    MAX(sh.value_text) FILTER (WHERE md.key = 'leak_status')              AS leak_status
FROM sample_headers sh
JOIN metric_defs md ON md.id = sh.field_id AND md.kind = 'header'
GROUP BY sh.sample_id;

COMMENT ON VIEW samples_v IS
    'Read-only pivot of sample_headers back to the 22 samples header columns (v3 Phase 3b, index D1). '
    'JOIN PARTNER, NOT A REPLACEMENT: samples keeps test_id, sort_order and its own audit trio, so '
    'readers do `samples s JOIN samples_v v ON v.id = s.id`. `id` is the real samples.id; `sample_id` '
    'is the TEXT business identifier header field. A NULL value column means NO sample_headers row '
    'exists, which under sparse materialization (index D2) is exactly what a NULL in the source '
    'samples column meant. Not updatable and deliberately has no INSTEAD OF trigger.';

-- ============================================================================
-- 6. The one-shot migration function
-- ============================================================================

-- ┌──────────────────────────────────────────────────────────────────────────┐
-- │ THIS FUNCTION IS AUTHORED HERE AND IS NEVER CALLED IN PHASE 3b.          │
-- │                                                                          │
-- │ Nothing in ensureSchema() or any other application code path may invoke  │
-- │ it. Its caller is PHASE 3d, the standard-metric cutover, which wraps it  │
-- │ in the advisory-lock + schema_meta gate (the v242_legacy_score_normalize │
-- │ pattern at src/database/DatabaseManager.cpp:578-609). In 3b it is called │
-- │ only by the rehearsal harness, against a synthetic database.             │
-- │                                                                          │
-- │ WHY: the Phase 3 index SEQUENCING CORRECTION. Backfilling measurements   │
-- │ now would create standard-metric rows that NOTHING maintains until 3d,   │
-- │ so the migrated copy would rot for two phases while the wide tables kept │
-- │ taking writes. The migration runs once, at the moment the app is finally │
-- │ able to maintain those rows.                                             │
-- └──────────────────────────────────────────────────────────────────────────┘
--
-- Sparse rule (index D2), which is a CORRECTNESS requirement and not a size
-- optimization: a measurement row is inserted if and only if the source wide
-- column IS NOT NULL. Numeric zeros ARE inserted - they are real measurements.
-- SheetResult::hasPerRowRegime is derived at load time from
-- `puffing_regime IS NOT NULL` (src/database/DatabaseManager.cpp:1117-1118,
-- src/database/OfflineSnapshot.cpp:1695-1696), so materializing an empty regime
-- row for every data row would flip every old-template sheet into
-- per-row-regime mode. The rule is expressed literally below as the
-- `WHERE <col> IS NOT NULL` on each per-column INSERT.
--
-- Idempotent: ON CONFLICT DO NOTHING on both unique keys, so a re-run after a
-- partial failure inserts only what is missing.
CREATE OR REPLACE FUNCTION dve_migrate_to_long_format(
    OUT measurements_inserted   BIGINT,
    OUT sample_headers_inserted BIGINT,
    OUT data_rows_seen          BIGINT,
    OUT samples_seen            BIGINT
) AS $$
DECLARE
    -- The wide columns that migrate, pinned explicitly rather than derived from
    -- metric_defs at run time. metric_defs grows (ensureSchema upserts the whole
    -- compiled registry, and Phase 5's template builder adds more), and a
    -- key-driven join would silently start capturing a non-value column the day
    -- a registry key happened to collide with one.
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

    v_col         TEXT;
    v_def_id      BIGINT;
    v_is_num      BOOLEAN;
    v_target_col  TEXT;
    v_target_type TEXT;
    v_n           BIGINT;

    v_bad_id      BIGINT;
    v_dup_sample  BIGINT;
    v_dup_order   INTEGER;
    v_dup_count   BIGINT;
BEGIN
    -- A full-table transform over every data row; the app's connections carry a
    -- 10s statement_timeout, which this would blow through on a real database.
    -- Same precedent as dve_normalize_legacy_json.
    SET LOCAL statement_timeout = 0;

    measurements_inserted   := 0;
    sample_headers_inserted := 0;

    -- ── Pre-flight 1: NULL sort_order ───────────────────────────────────────
    -- data_rows.sort_order is `INTEGER DEFAULT 0` with NO NOT NULL, and
    -- MigrationTool copied legacy sort_order values verbatim. measurements.
    -- sort_order is NOT NULL and part of the row's identity, so a NULL here has
    -- no honest translation. COALESCE-ing it to 0 would fabricate an ordering
    -- and could collide with a real row 0. Abort and name the offender.
    SELECT d.id, d.sample_id INTO v_bad_id, v_dup_sample
      FROM data_rows d
     WHERE d.sort_order IS NULL
     LIMIT 1;
    IF FOUND THEN
        RAISE EXCEPTION
            'dve_migrate_to_long_format: data_rows.id=% (sample_id=%) has a NULL sort_order. '
            'measurements.sort_order is NOT NULL and identity-bearing; assign a real sort_order '
            'before migrating.', v_bad_id, v_dup_sample;
    END IF;

    -- ── Pre-flight 2: duplicate (sample_id, sort_order) ─────────────────────
    -- data_rows has NO unique constraint on (sample_id, sort_order) - only the
    -- PK on id - and MigrationTool copied legacy values verbatim, so duplicates
    -- are genuinely possible in real data. measurements.UNIQUE(sample_id,
    -- metric_id, sort_order) makes them illegal, and the ON CONFLICT DO NOTHING
    -- below would SILENTLY COLLAPSE the duplicates and lose a whole data row.
    -- This abort is the poka-yoke that makes running this against production
    -- safe.
    SELECT d.sample_id, d.sort_order, count(*)
      INTO v_dup_sample, v_dup_order, v_dup_count
      FROM data_rows d
     GROUP BY d.sample_id, d.sort_order
    HAVING count(*) > 1
     LIMIT 1;
    IF FOUND THEN
        RAISE EXCEPTION
            'dve_migrate_to_long_format: data_rows has duplicate (sample_id, sort_order) - '
            'sample_id=%, sort_order=% appears % times. measurements.UNIQUE(sample_id, metric_id, '
            'sort_order) cannot represent that, and migrating would silently collapse the '
            'duplicates. Resolve them before migrating.', v_dup_sample, v_dup_order, v_dup_count;
    END IF;

    -- ── Pre-flight 3: a data row whose 13 value columns are ALL NULL ────────
    -- Under the sparse rule such a row produces zero measurements, so it has no
    -- group in data_rows_v and VANISHES. The long format cannot express "this
    -- row exists but holds nothing" - there is no row without a value to hang
    -- it on. Refuse, exactly as for a duplicate or a NULL sort_order, rather
    -- than silently dropping a data row. Today every numeric column carries
    -- DEFAULT 0.0 and the save path always binds, so this should never fire -
    -- but MigrationTool copied legacy rows verbatim and nothing enforces it.
    --
    -- The predicate is built from c_metric_cols rather than spelled out, so it
    -- can never drift from the column list the migration actually reads.
    EXECUTE format(
        'SELECT id, sample_id, sort_order FROM data_rows WHERE %s LIMIT 1',
        (SELECT string_agg(format('%I IS NULL', c), ' AND ') FROM unnest(c_metric_cols) AS c))
      INTO v_bad_id, v_dup_sample, v_dup_order;
    -- NOT `IF FOUND`: EXECUTE deliberately does not touch FOUND (only plain
    -- SELECT INTO and the DML statements do), so an `IF FOUND` here reads
    -- whatever the previous pre-flight left behind and the abort never fires.
    -- EXECUTE does update GET DIAGNOSTICS, so ROW_COUNT is the honest test.
    GET DIAGNOSTICS v_n = ROW_COUNT;
    IF v_n > 0 THEN
        RAISE EXCEPTION
            'dve_migrate_to_long_format: data_rows.id=% (sample_id=%, sort_order=%) has ALL 13 '
            'value columns NULL. Sparse materialization would produce no measurement rows for it '
            'and the row would disappear from data_rows_v. Give it a value or delete it before '
            'migrating.', v_bad_id, v_dup_sample, v_dup_order;
    END IF;

    SELECT count(*) INTO data_rows_seen FROM data_rows;
    SELECT count(*) INTO samples_seen   FROM samples;

    -- ── data_rows -> measurements ───────────────────────────────────────────
    -- One set-based INSERT per column. The column name reaches the statement
    -- through format(%I) from the pinned array above, and the value_num /
    -- value_text routing comes from metric_defs.value_type - so the long
    -- format's typing is decided by the registry projection, not by this
    -- function's opinion. %s interpolation of the cast target is safe: it is
    -- one of exactly two literals chosen right here.
    FOREACH v_col IN ARRAY c_metric_cols LOOP
        SELECT md.id, md.value_type = 'number'
          INTO v_def_id, v_is_num
          FROM metric_defs md
         WHERE md.kind = 'metric' AND md.key = v_col;
        IF NOT FOUND THEN
            RAISE EXCEPTION
                'dve_migrate_to_long_format: metric_defs has no kind=''metric'' row for key %. '
                'The seed in this migration must cover every data_rows value column.', v_col;
        END IF;
        v_target_col  := CASE WHEN v_is_num THEN 'value_num' ELSE 'value_text' END;
        v_target_type := CASE WHEN v_is_num THEN 'DOUBLE PRECISION' ELSE 'TEXT' END;

        EXECUTE format($f$
            INSERT INTO measurements (sample_id, metric_id, sort_order, %2$I, updated_by)
            SELECT d.sample_id, $1, d.sort_order, d.%1$I::%3$s, d.updated_by
              FROM data_rows d
             WHERE d.%1$I IS NOT NULL
            ON CONFLICT (sample_id, metric_id, sort_order) DO NOTHING
        $f$, v_col, v_target_col, v_target_type)
        USING v_def_id;

        GET DIAGNOSTICS v_n = ROW_COUNT;
        measurements_inserted := measurements_inserted + v_n;
    END LOOP;

    -- ── samples -> sample_headers ───────────────────────────────────────────
    -- No sort_order: a header field is per-sample, so the key is
    -- (sample_id, field_id).
    FOREACH v_col IN ARRAY c_header_cols LOOP
        SELECT md.id, md.value_type = 'number'
          INTO v_def_id, v_is_num
          FROM metric_defs md
         WHERE md.kind = 'header' AND md.key = v_col;
        IF NOT FOUND THEN
            RAISE EXCEPTION
                'dve_migrate_to_long_format: metric_defs has no kind=''header'' row for key %. '
                'The seed in this migration must cover every samples value column.', v_col;
        END IF;
        v_target_col  := CASE WHEN v_is_num THEN 'value_num' ELSE 'value_text' END;
        v_target_type := CASE WHEN v_is_num THEN 'DOUBLE PRECISION' ELSE 'TEXT' END;

        EXECUTE format($f$
            INSERT INTO sample_headers (sample_id, field_id, %2$I, updated_by)
            SELECT s.id, $1, s.%1$I::%3$s, s.updated_by
              FROM samples s
             WHERE s.%1$I IS NOT NULL
            ON CONFLICT (sample_id, field_id) DO NOTHING
        $f$, v_col, v_target_col, v_target_type)
        USING v_def_id;

        GET DIAGNOSTICS v_n = ROW_COUNT;
        sample_headers_inserted := sample_headers_inserted + v_n;
    END LOOP;
END;
$$ LANGUAGE plpgsql;

COMMENT ON FUNCTION dve_migrate_to_long_format() IS
    'One-shot wide -> long data migration (v3 Phase 3b). AUTHORED, NOT EXECUTED: no application '
    'code path calls this in 3b. Phase 3d owns the gated invocation. Aborts rather than migrating '
    'if data_rows has a NULL or duplicate (sample_id, sort_order). Sparse per index D2: a row is '
    'written only where the wide column IS NOT NULL; numeric zeros are written. Idempotent.';

COMMIT;
