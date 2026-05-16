-- v2.0 -> v2.0.1: live collaborative editing
-- Adds the cell_focus table and extends the row-changed NOTIFY payload
-- with column + new_value. Idempotent.

BEGIN;

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

-- Extend the existing row-changed trigger function so the NOTIFY payload
-- carries the column that changed and its new value when the operation
-- is a single-column UPDATE. (Multi-column UPDATEs send a payload
-- without column/new_value -- the client falls back to re-SELECT.)
-- Existing function is named notify_row_change and is AFTER ROW -> returns NULL.
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
    -- live-sync extension: include the column + new value when the
    -- caller set a session var indicating a single-column UPDATE.
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

-- New channel: cell_focus broadcasts. Trigger fires on every INSERT,
-- UPDATE, DELETE on cell_focus and pushes the row.
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

COMMIT;
