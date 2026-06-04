-- DataViewer Enterprise — initial schema (Plan A)
-- Idempotent: re-running drops nothing; uses CREATE ... IF NOT EXISTS.
-- Convention: ON DELETE CASCADE on every child FK — child rows (images, data_rows, etc.) have no standalone meaning without their parent.

BEGIN;

-- ── files ────────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS files (
    id               BIGSERIAL PRIMARY KEY,
    file_path        TEXT NOT NULL,
    file_name        TEXT NOT NULL,
    loaded_at        TEXT NOT NULL,
    template_version TEXT,
    sheet_count      INTEGER DEFAULT 0,
    sample_count     INTEGER DEFAULT 0,
    updated_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by       TEXT        NOT NULL DEFAULT 'migration',
    version          INTEGER     NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_files_path ON files(file_path);

-- ── tests (sheets) ───────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS tests (
    id                 BIGSERIAL PRIMARY KEY,
    file_id            BIGINT NOT NULL REFERENCES files(id) ON DELETE CASCADE,
    sheet_name         TEXT NOT NULL,
    template_version   TEXT,
    overall_avg_tpm    DOUBLE PRECISION DEFAULT 0.0,
    overall_stddev_tpm DOUBLE PRECISION DEFAULT 0.0,
    is_raw_table       INTEGER DEFAULT 0,
    raw_grid           JSONB,
    sort_order         INTEGER DEFAULT 0,
    updated_at         TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by         TEXT        NOT NULL DEFAULT 'migration',
    version            INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_tests_file ON tests(file_id);

-- ── samples ──────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS samples (
    id                  BIGSERIAL PRIMARY KEY,
    test_id             BIGINT NOT NULL REFERENCES tests(id) ON DELETE CASCADE,
    sort_order          INTEGER DEFAULT 0,
    sample_name         TEXT,
    sample_id           TEXT,
    date                TEXT,
    tester              TEXT,
    media               TEXT,
    viscosity           DOUBLE PRECISION DEFAULT 0.0,
    resistance          DOUBLE PRECISION DEFAULT 0.0,
    voltage             DOUBLE PRECISION DEFAULT 0.0,
    power               DOUBLE PRECISION DEFAULT 0.0,
    heating_technology  TEXT,
    puffing_regime      TEXT,
    initial_oil_mass    DOUBLE PRECISION DEFAULT 0.0,
    average_tpm         DOUBLE PRECISION DEFAULT 0.0,
    stddev_tpm          DOUBLE PRECISION DEFAULT 0.0,
    avg_power_density   DOUBLE PRECISION DEFAULT 0.0,
    efficiency_percent  DOUBLE PRECISION DEFAULT 0.0,
    total_oil_consumed  DOUBLE PRECISION DEFAULT 0.0,
    total_puffs         INTEGER DEFAULT 0,
    normalized_tpm      DOUBLE PRECISION DEFAULT 0.0,
    burn_status         TEXT,
    clog_status         TEXT,
    leak_status         TEXT,
    updated_at          TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by          TEXT        NOT NULL DEFAULT 'migration',
    version             INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_samples_test ON samples(test_id);

-- ── data_rows ────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS data_rows (
    id                BIGSERIAL PRIMARY KEY,
    sample_id         BIGINT NOT NULL REFERENCES samples(id) ON DELETE CASCADE,
    sort_order        INTEGER DEFAULT 0,
    puffs             DOUBLE PRECISION DEFAULT 0.0,
    before_weight     DOUBLE PRECISION DEFAULT 0.0,
    after_weight      DOUBLE PRECISION DEFAULT 0.0,
    draw_pressure     DOUBLE PRECISION DEFAULT 0.0,
    resistance        DOUBLE PRECISION DEFAULT 0.0,
    smell             TEXT,
    clog              TEXT,
    notes             TEXT,
    tpm               DOUBLE PRECISION DEFAULT 0.0,
    tpm_power_density DOUBLE PRECISION DEFAULT 0.0,
    variation_tpm     DOUBLE PRECISION DEFAULT 0.0,
    oil_consumed      DOUBLE PRECISION DEFAULT 0.0,
    puffing_regime    TEXT,
    updated_at        TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by        TEXT        NOT NULL DEFAULT 'migration',
    version           INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_data_rows_sample ON data_rows(sample_id);

-- ── images (per-sample) ──────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS images (
    id          BIGSERIAL PRIMARY KEY,
    sample_id   BIGINT NOT NULL REFERENCES samples(id) ON DELETE CASCADE,
    sort_order  INTEGER DEFAULT 0,
    file_name   TEXT,
    image_data  BYTEA,
    layout_x    DOUBLE PRECISION,
    layout_y    DOUBLE PRECISION,
    layout_w    DOUBLE PRECISION,
    layout_h    DOUBLE PRECISION,
    crop_x      DOUBLE PRECISION,
    crop_y      DOUBLE PRECISION,
    crop_w      DOUBLE PRECISION,
    crop_h      DOUBLE PRECISION,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT        NOT NULL DEFAULT 'migration',
    version     INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_images_sample ON images(sample_id);

-- ── sensory_sessions ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS sensory_sessions (
    id            BIGSERIAL PRIMARY KEY,
    session_name  TEXT,
    tester_name   TEXT,
    assessor_name TEXT,
    media         TEXT,
    puff_length   TEXT,
    date          TEXT,
    timestamp     TEXT,
    json_data     JSONB,
    layout_json   JSONB,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by    TEXT        NOT NULL DEFAULT 'migration',
    version       INTEGER     NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_sensory_sessions_key
    ON sensory_sessions(session_name, tester_name, date);

-- ── sensory_images ───────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS sensory_images (
    id          BIGSERIAL PRIMARY KEY,
    session_id  BIGINT NOT NULL REFERENCES sensory_sessions(id) ON DELETE CASCADE,
    sort_order  INTEGER DEFAULT 0,
    file_name   TEXT,
    image_data  BYTEA,
    layout_x    DOUBLE PRECISION,
    layout_y    DOUBLE PRECISION,
    layout_w    DOUBLE PRECISION,
    layout_h    DOUBLE PRECISION,
    crop_x      DOUBLE PRECISION,
    crop_y      DOUBLE PRECISION,
    crop_w      DOUBLE PRECISION,
    crop_h      DOUBLE PRECISION,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT        NOT NULL DEFAULT 'migration',
    version     INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_sensory_images_session ON sensory_images(session_id);

-- ── cell_focus (live per-cell editing presence) ──────────────────────────────
-- One row per (user, table, row, column) currently being edited. Distinct
-- from `presence` (which is per-resource); cell_focus is the fine-grained
-- "user X is in column Y of row Z" view that drives the colored border +
-- name flag overlay in collaborative tables.
CREATE TABLE IF NOT EXISTS cell_focus (
    user_uuid    UUID        NOT NULL,
    table_name   TEXT        NOT NULL,
    row_id       BIGINT      NOT NULL,
    column_name  TEXT        NOT NULL,
    user_name    TEXT        NOT NULL,
    user_color   TEXT        NOT NULL,
    started_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (user_uuid, table_name, row_id, column_name)
);
CREATE INDEX IF NOT EXISTS idx_cell_focus_target
    ON cell_focus(table_name, row_id);

-- ── detailed_sensory_sessions ────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS detailed_sensory_sessions (
    id            BIGSERIAL PRIMARY KEY,
    session_name  TEXT,
    tester_name   TEXT,
    assessor_name TEXT,
    media         TEXT,
    date          TEXT,
    timestamp     TEXT,
    json_data     JSONB,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by    TEXT        NOT NULL DEFAULT 'migration',
    version       INTEGER     NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_detailed_sensory_sessions_key
    ON detailed_sensory_sessions(session_name, tester_name, date);

-- ── detailed_sensory_images ──────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS detailed_sensory_images (
    id          BIGSERIAL PRIMARY KEY,
    session_id  BIGINT NOT NULL REFERENCES detailed_sensory_sessions(id) ON DELETE CASCADE,
    sort_order  INTEGER DEFAULT 0,
    file_name   TEXT,
    image_data  BYTEA,
    layout_x    DOUBLE PRECISION,
    layout_y    DOUBLE PRECISION,
    layout_w    DOUBLE PRECISION,
    layout_h    DOUBLE PRECISION,
    crop_x      DOUBLE PRECISION,
    crop_y      DOUBLE PRECISION,
    crop_w      DOUBLE PRECISION,
    crop_h      DOUBLE PRECISION,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT        NOT NULL DEFAULT 'migration',
    version     INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_detailed_sensory_images_session
    ON detailed_sensory_images(session_id);

-- ── settings (key/value) ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS settings (
    key        TEXT PRIMARY KEY,
    value      TEXT,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by TEXT        NOT NULL DEFAULT 'migration',
    version    INTEGER     NOT NULL DEFAULT 1
);

-- ── presence (live "who has this open" view) ─────────────────────────────────
CREATE TABLE IF NOT EXISTS presence (
    user_uuid      UUID        NOT NULL,
    user_name      TEXT        NOT NULL,
    user_color     TEXT        NOT NULL CHECK (user_color ~ '^#[0-9A-Fa-f]{6}$'),
    resource_type  TEXT        NOT NULL CHECK (resource_type IN ('file', 'sensory_session', 'detailed_sensory_session')),
    resource_id    BIGINT      NOT NULL,
    intent         TEXT        NOT NULL DEFAULT 'viewing' CHECK (intent IN ('viewing', 'editing')),
    last_heartbeat TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (user_uuid, resource_type, resource_id)
);
CREATE INDEX IF NOT EXISTS idx_presence_resource ON presence(resource_type, resource_id);
COMMENT ON COLUMN presence.resource_id IS
    'Polymorphic reference: points to files.id, sensory_sessions.id, or detailed_sensory_sessions.id depending on resource_type. Cannot be enforced as a foreign key by Postgres; cleanup of orphaned presence rows is the application''s responsibility (handled by pg_cron stale-row deletion in Task 5).';

-- ── schema_meta (migration provenance, schema versioning) ────────────────────
CREATE TABLE IF NOT EXISTS schema_meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- ── bump_version: auto-increments version + stamps updated_at on every UPDATE
CREATE OR REPLACE FUNCTION bump_version() RETURNS TRIGGER AS $$
BEGIN
  -- Refuse client-side version manipulation: if NEW.version differs from
  -- OLD.version, that's a bug in the app, not an upgrade. We always +1.
  NEW.version    := OLD.version + 1;
  NEW.updated_at := now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DO $$
DECLARE
  t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'files', 'tests', 'samples', 'data_rows', 'images',
    'sensory_sessions', 'sensory_images',
    'detailed_sensory_sessions', 'detailed_sensory_images',
    'settings'
  ] LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_bump_version ON %I;
       CREATE TRIGGER trg_%I_bump_version
       BEFORE UPDATE ON %I
       FOR EACH ROW EXECUTE FUNCTION bump_version();',
       t, t, t, t
    );
  END LOOP;
END$$;

-- ── notify_row_change: fires NOTIFY on every INSERT/UPDATE/DELETE for the
--    10 editable data tables. Channel: 'dataviewer_changes'. Payload JSON:
--    {table, op, id, updated_by} plus optional {column, new_value} when the
--    caller has set the dve.live_column / dve.live_value session vars
--    (single-column UPDATEs only — multi-column UPDATEs leave them unset
--    and the client falls back to re-SELECT). Trigger is AFTER ROW so it
--    sees the result of any BEFORE trigger (e.g., bump_version) — so the
--    payload's updated_by reflects the final post-trigger state.
CREATE OR REPLACE FUNCTION notify_row_change() RETURNS TRIGGER AS $$
DECLARE
  payload JSONB;
BEGIN
  payload := jsonb_build_object(
    'table',      TG_TABLE_NAME,
    'op',         TG_OP,
    'id',         COALESCE(NEW.id, OLD.id),
    'updated_by', COALESCE(NEW.updated_by, OLD.updated_by)
  );
  IF current_setting('dve.live_column', true) <> '' THEN
    payload := payload || jsonb_build_object(
      'column',    current_setting('dve.live_column', true),
      'new_value', current_setting('dve.live_value',  true)
    );
  END IF;
  PERFORM pg_notify('dataviewer_changes', payload::text);
  RETURN NULL;
END;
$$ LANGUAGE plpgsql;

-- settings is intentionally NOT in this array: its PK is `key TEXT`, not `id`,
-- so the trigger's NEW.id reference would fail. settings is app config and
-- doesn't need live NOTIFY-driven UI refresh.
DO $$
DECLARE
  t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'files', 'tests', 'samples', 'data_rows', 'images',
    'sensory_sessions', 'sensory_images',
    'detailed_sensory_sessions', 'detailed_sensory_images'
  ] LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_notify ON %I;
       CREATE TRIGGER trg_%I_notify
       AFTER INSERT OR UPDATE OR DELETE ON %I
       FOR EACH ROW EXECUTE FUNCTION notify_row_change();',
       t, t, t, t
    );
  END LOOP;
END$$;

-- ── notify_presence_change: separate channel 'dataviewer_presence' for
--    high-frequency presence heartbeats. Distinct from data changes so the
--    UI can throttle/batch presence events independently.
CREATE OR REPLACE FUNCTION notify_presence_change() RETURNS TRIGGER AS $$
BEGIN
  PERFORM pg_notify(
    'dataviewer_presence',
    json_build_object(
      'op',            TG_OP,
      'user_uuid',     COALESCE(NEW.user_uuid::text, OLD.user_uuid::text),
      'resource_type', COALESCE(NEW.resource_type, OLD.resource_type),
      'resource_id',   COALESCE(NEW.resource_id, OLD.resource_id),
      'intent',        COALESCE(NEW.intent, OLD.intent)
    )::text
  );
  RETURN NULL;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_presence_notify ON presence;
CREATE TRIGGER trg_presence_notify
AFTER INSERT OR UPDATE OR DELETE ON presence
FOR EACH ROW EXECUTE FUNCTION notify_presence_change();

-- ── notify_cell_focus: separate channel 'dataviewer_cell_focus' for the
--    per-cell editing presence broadcasts. Distinct from data and resource
--    presence so the UI can subscribe selectively.
CREATE OR REPLACE FUNCTION notify_cell_focus() RETURNS TRIGGER AS $$
DECLARE
  payload JSONB;
BEGIN
  payload := jsonb_build_object(
    'op',         TG_OP,
    'user_uuid',  COALESCE(NEW.user_uuid::text, OLD.user_uuid::text),
    'user_name',  COALESCE(NEW.user_name, OLD.user_name),
    'user_color', COALESCE(NEW.user_color, OLD.user_color),
    'table',      COALESCE(NEW.table_name, OLD.table_name),
    'row_id',     COALESCE(NEW.row_id, OLD.row_id),
    'column',     COALESCE(NEW.column_name, OLD.column_name)
  );
  PERFORM pg_notify('dataviewer_cell_focus', payload::text);
  RETURN NULL;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_cell_focus_notify ON cell_focus;
CREATE TRIGGER trg_cell_focus_notify
AFTER INSERT OR UPDATE OR DELETE ON cell_focus
FOR EACH ROW EXECUTE FUNCTION notify_cell_focus();

-- ── Server-side single-round-trip commit helpers (v2.0.1 perf fix) ──────────
-- Originally the client did BEGIN / set_config / UPDATE / COMMIT as four
-- separate round trips. On a LAN with even modest latency that's hundreds
-- of milliseconds per cell edit, which blocks the UI thread. These wrap
-- the whole sequence so the client only needs one call.
--
-- Client still owns table+column allowlisting (LiveSync::isLiveSyncTable /
-- isLiveSyncColumn) so the format()/EXECUTE pair below stays safe.
CREATE OR REPLACE FUNCTION dve_commit_cell(
    p_table   TEXT,
    p_row_id  BIGINT,
    p_column  TEXT,
    p_value   TEXT,
    p_uuid    TEXT
) RETURNS BOOLEAN AS $$
DECLARE
    affected INT;
BEGIN
    PERFORM set_config('dve.live_column', p_column, true);
    PERFORM set_config('dve.live_value',  p_value,  true);
    EXECUTE format(
        'UPDATE %I SET %I = $1, version = version + 1, '
        'updated_at = now(), updated_by = $2 WHERE id = $3',
        p_table, p_column
    ) USING p_value, p_uuid, p_row_id;
    GET DIAGNOSTICS affected = ROW_COUNT;
    RETURN affected > 0;
END;
$$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION dve_commit_cell_json(
    p_table     TEXT,
    p_row_id    BIGINT,
    p_path_text TEXT,         -- e.g. 'samples[2].name' (without the json_path: prefix)
    p_path_arr  TEXT[],       -- parsed array form: {samples,2,name}
    p_value     TEXT,
    p_uuid      TEXT
) RETURNS BOOLEAN AS $$
DECLARE
    affected INT;
BEGIN
    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true);
    PERFORM set_config('dve.live_value',  p_value,                     true);
    EXECUTE format(
        'UPDATE %I SET json_data = jsonb_set(json_data, $1, '
        'to_jsonb($2::text)::jsonb, true), '
        'version = version + 1, updated_at = now(), updated_by = $3 '
        'WHERE id = $4',
        p_table
    ) USING p_path_arr, p_value, p_uuid, p_row_id;
    GET DIAGNOSTICS affected = ROW_COUNT;
    RETURN affected > 0;
END;
$$ LANGUAGE plpgsql;

CREATE OR REPLACE FUNCTION dve_focus_cell(
    p_uuid       TEXT,
    p_table      TEXT,
    p_row_id     BIGINT,
    p_column     TEXT,
    p_user_name  TEXT,
    p_user_color TEXT
) RETURNS VOID AS $$
BEGIN
    DELETE FROM cell_focus WHERE user_uuid = p_uuid::uuid;
    INSERT INTO cell_focus(user_uuid, table_name, row_id, column_name,
                            user_name, user_color)
    VALUES (p_uuid::uuid, p_table, p_row_id, p_column, p_user_name, p_user_color);
END;
$$ LANGUAGE plpgsql;

COMMIT;

-- ── pg_cron: schedule presence stale-row cleanup every minute ───────────────
-- Must run as superuser; the postgres:16 image creates the extension when
-- shared_preload_libraries=pg_cron is set (see docker-compose command:).
-- Runs OUTSIDE the main BEGIN/COMMIT because CREATE EXTENSION is a non-
-- transactional, superuser-only operation.
CREATE EXTENSION IF NOT EXISTS pg_cron;

-- cron.schedule(jobname, ...) is idempotent: re-running updates an existing
-- job with the same name rather than creating a duplicate. No explicit
-- unschedule needed.
SELECT cron.schedule(
  'dve_presence_cleanup',
  '* * * * *',  -- every minute
  $$ DELETE FROM presence
     WHERE last_heartbeat < now() - INTERVAL '30 seconds' $$
);

SELECT cron.schedule(
  'dve_cell_focus_cleanup',
  '*/30 * * * * *',
  $$DELETE FROM cell_focus WHERE started_at < now() - INTERVAL '30 seconds'$$
);
