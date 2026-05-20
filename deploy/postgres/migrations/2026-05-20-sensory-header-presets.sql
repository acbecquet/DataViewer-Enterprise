-- Migration: sensory header presets (test name / media / sample name dropdowns).
--
-- v2.0.4 QoL feature. Sensory and Detailed Sensory panels gain a "Save
-- test headers" button that records the current Test Title, Media, and
-- per-sample names into a shared pool. Other users see the new values
-- in dropdowns the next time they open a session, avoiding the typo-
-- prone retype-the-exact-same-string flow.
--
-- Storage is global (one pool across all users); NOTIFY-driven so a save
-- on one workstation refreshes dropdowns on every other live install.
-- No OCC; values are immutable once inserted (uniqueness handles dedup,
-- and there's no reason to UPDATE an existing preset).

BEGIN;

CREATE TABLE IF NOT EXISTS sensory_header_presets (
    id          BIGSERIAL PRIMARY KEY,
    kind        TEXT NOT NULL CHECK (kind IN ('test_name', 'media', 'sample_name')),
    value       TEXT NOT NULL CHECK (length(trim(value)) > 0),
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    created_by  TEXT,
    -- version is unused (presets are immutable) but keeps the table in
    -- the same shape as the rest of the schema for trigger uniformity.
    version     INTEGER NOT NULL DEFAULT 1,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT,
    UNIQUE (kind, value)
);

CREATE INDEX IF NOT EXISTS idx_sensory_header_presets_kind
    ON sensory_header_presets (kind, value);

-- Register the table with both the version-bump trigger pool and the
-- NOTIFY trigger pool — uniform with files/tests/samples/etc.
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

COMMIT;
