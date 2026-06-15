-- Migration: v2.4.2 A2 — 6-arg dve_commit_cell_json numeric JSONB storage.
--
-- v2.4.1 healed the 7-arg OCC overload (the one LiveSyncWorker dispatches) to
-- store numeric-looking scores as JSON NUMBERS instead of strings (DATAVIEWER-4
-- root). This applies the same fix to the legacy 6-arg overload so any old/alt
-- caller can't re-introduce string-typed scores. Delivered at runtime by
-- DatabaseManager::ensureSchema() (catalog-guarded on prosrc, pronargs=6);
-- canonical record here. Idempotent (CREATE OR REPLACE, unchanged signature).

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
    EXECUTE format(
        'UPDATE %I SET json_data = jsonb_set(json_data, $1, '
        '(CASE WHEN $2 ~ ''^-?[0-9]+(\.[0-9]+)?$'' '
        '      THEN to_jsonb($2::numeric) ELSE to_jsonb($2::text) END)::jsonb, '
        'true), updated_by = $3 WHERE id = $4',
        p_table
    ) USING p_path_arr, p_value, p_uuid, p_row_id;
    GET DIAGNOSTICS affected = ROW_COUNT;
    RETURN affected > 0;
END;
$$ LANGUAGE plpgsql;
