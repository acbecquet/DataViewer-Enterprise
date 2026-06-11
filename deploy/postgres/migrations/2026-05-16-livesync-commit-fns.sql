-- Migration: server-side commit functions for v2.0.1 perf fix.
-- Reduces every per-cell edit from 4 round trips (BEGIN/set_config/UPDATE/COMMIT)
-- to 1 by wrapping the whole sequence in a function call. Idempotent.

BEGIN;

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
    p_path_text TEXT,
    p_path_arr  TEXT[],
    p_value     TEXT,
    p_uuid      TEXT
) RETURNS BOOLEAN AS $$
DECLARE
    affected INT;
BEGIN
    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true);
    PERFORM set_config('dve.live_value',  p_value,                     true);
    -- DATAVIEWER-4: numeric-looking values store as JSON NUMBERS (not strings),
    -- so live-streamed scores don't revert on read. See the 2026-06-10 migration.
    -- (This 6-arg form is superseded at runtime by the 7-arg OCC version from
    -- 2026-05-17; kept consistent for replay-in-order correctness.)
    EXECUTE format(
        'UPDATE %I SET json_data = jsonb_set(json_data, $1, '
        '(CASE WHEN $2 ~ ''^-?[0-9]+(\.[0-9]+)?$'' '
        '      THEN to_jsonb($2::numeric) ELSE to_jsonb($2::text) END)::jsonb, '
        'true), '
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
