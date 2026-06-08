-- Migration: DATAVIEWER-2 — test-scoped sample-name presets.
--
-- v2.4.0 bug-fix batch. The sample-name dropdown in the Sensory and Detailed
-- Sensory panels becomes scoped to the *current test title* instead of a
-- single global pool, so the same sample name (e.g. "Sample 1") can live once
-- per test. This is delivered at runtime by DatabaseManager::ensureSchema()
-- (the self-heal that runs on every connect); the test harness does NOT apply
-- the files under deploy/postgres/migrations/ automatically. This file is the
-- durable CANONICAL RECORD of that runtime delivery — the source to reconcile
-- into deploy/postgres/init.sql when the baseline schema is next regenerated.
--
-- It faithfully mirrors what ensureSchema() does:
--   (1) create sensory_header_presets in the NEW test-scoped shape if absent,
--   (2) add test_name to a legacy table,
--   (3) drop the old (kind, value) uniqueness and swap to the test-scoped
--       expression index (kind, value, COALESCE(test_name, '')),
--   (4) register the shared version-bump + NOTIFY triggers (per 2026-05-20),
--   (5) one-time backfill of (test_title -> sample_name) presets from existing
--       session history, gated by a schema_meta marker.
--
-- Fully idempotent and re-runnable: every step uses IF EXISTS / IF NOT EXISTS /
-- ON CONFLICT DO NOTHING. The backfill keys on json_data->>'test_title' (the
-- test title), NOT session_name — this matches the live save/load path
-- (SensoryPanel / DetailedSensoryPanel pass the Test Title to
-- saveSensoryHeaderPresets; loadSampleNamesForTest matches the test_name column
-- on that title) and the corrected runtime backfill in ensureSchema().
--
-- NOTE: sensory_header_presets is intentionally NOT mirrored in OfflineSnapshot
-- (it never was). In offline read-only mode the sample-name dropdown simply
-- stays empty — out of scope here, no data loss.

BEGIN;

-- (1) Table in the test-scoped shape (incl. test_name). Columns/CHECKs copied
--     verbatim from ensureSchema()'s CREATE; ON a fresh DB this is the whole
--     table, on an existing DB it is a no-op and steps (2)/(3) do the healing.
CREATE TABLE IF NOT EXISTS sensory_header_presets (
    id          BIGSERIAL PRIMARY KEY,
    kind        TEXT NOT NULL CHECK (kind IN ('test_name', 'media', 'sample_name')),
    value       TEXT NOT NULL CHECK (length(trim(value)) > 0),
    test_name   TEXT,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    created_by  TEXT,
    version     INTEGER NOT NULL DEFAULT 1,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT
);

-- (2) Heal a legacy table that predates the test_name column.
ALTER TABLE sensory_header_presets ADD COLUMN IF NOT EXISTS test_name TEXT;

-- (3) Swap uniqueness from the legacy (kind, value) to the test-scoped
--     expression index. Drop the old constraint + companion lookup index, then
--     build the new unique + lookup indexes. The expression index lets the same
--     value exist once per test (and once globally, where test_name IS NULL).
ALTER TABLE sensory_header_presets
    DROP CONSTRAINT IF EXISTS sensory_header_presets_kind_value_key;
DROP INDEX IF EXISTS idx_sensory_header_presets_kind;

CREATE UNIQUE INDEX IF NOT EXISTS uq_shp_kind_value_test
    ON sensory_header_presets (kind, value, COALESCE(test_name, ''));
CREATE INDEX IF NOT EXISTS idx_shp_kind_test
    ON sensory_header_presets (kind, test_name);

-- (4) Register the table with the shared version-bump + NOTIFY trigger pools,
--     uniform with files/tests/samples/etc. Trigger names copied from the
--     2026-05-20 migration. The trigger functions (bump_version,
--     notify_row_change) ship in init.sql. Idempotent via DROP TRIGGER IF
--     EXISTS.
DO $$
BEGIN
    EXECUTE 'DROP TRIGGER IF EXISTS trg_sensory_header_presets_bump_version ON sensory_header_presets;
             CREATE TRIGGER trg_sensory_header_presets_bump_version
             BEFORE UPDATE ON sensory_header_presets
             FOR EACH ROW EXECUTE FUNCTION bump_version();';
    EXECUTE 'DROP TRIGGER IF EXISTS trg_sensory_header_presets_notify ON sensory_header_presets;
             CREATE TRIGGER trg_sensory_header_presets_notify
             AFTER INSERT OR UPDATE OR DELETE ON sensory_header_presets
             FOR EACH ROW EXECUTE FUNCTION notify_row_change();';
END$$;

-- (5) One-time backfill of (test_title -> sample_name) presets from existing
--     session history, for both sensory and detailed-sensory sessions. Every
--     distinct sample name ever saved under a non-blank test title becomes a
--     sample_name preset scoped to that title. Sessions with a blank/absent
--     test_title are skipped (the live path sends those names to the global
--     NULL pool, which a scoped lookup never reads). Keyed on
--     json_data->>'test_title' (NOT session_name). Idempotent via ON CONFLICT
--     against the test-scoped unique index.
INSERT INTO sensory_header_presets
    (kind, value, test_name, created_by, updated_by)
SELECT DISTINCT 'sample_name',
       trim(smp->>'name'),
       trim(ss.json_data->>'test_title'),
       'backfill', 'backfill'
FROM sensory_sessions ss
CROSS JOIN LATERAL
    jsonb_array_elements(ss.json_data->'samples') AS smp
WHERE ss.json_data->>'test_title' IS NOT NULL
  AND length(trim(ss.json_data->>'test_title')) > 0
  AND smp->>'name' IS NOT NULL
  AND length(trim(smp->>'name')) > 0
ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING;

INSERT INTO sensory_header_presets
    (kind, value, test_name, created_by, updated_by)
SELECT DISTINCT 'sample_name',
       trim(smp->>'name'),
       trim(ss.json_data->>'test_title'),
       'backfill', 'backfill'
FROM detailed_sensory_sessions ss
CROSS JOIN LATERAL
    jsonb_array_elements(ss.json_data->'samples') AS smp
WHERE ss.json_data->>'test_title' IS NOT NULL
  AND length(trim(ss.json_data->>'test_title')) > 0
  AND smp->>'name' IS NOT NULL
  AND length(trim(smp->>'name')) > 0
ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING;

-- Mark the one-time backfill as done so a future re-run is skipped (mirrors the
-- schema_meta gate in ensureSchema()).
INSERT INTO schema_meta (key, value)
VALUES ('dv2_sample_name_backfill', now()::text)
ON CONFLICT (key) DO NOTHING;

COMMIT;
