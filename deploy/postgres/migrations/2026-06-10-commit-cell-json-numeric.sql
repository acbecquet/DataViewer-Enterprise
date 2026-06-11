-- Migration: DATAVIEWER-4 root cause — numeric JSONB for live-streamed scores.
--
-- v2.4.0 save/sync fix batch. THE ORIGINAL "scores reset to 5" REVERT (live
-- since v2.0.1). The LiveSync per-cell stored function dve_commit_cell_json
-- stored EVERY value via to_jsonb($2::text), so a numeric score streamed live
-- landed in the session JSONB as a STRING ("7.0"). The C++ readers
-- (sensorySessionFromJson / detailedSensorySessionFromJson) read scores with
-- QJsonValue::toDouble(default), which returns the DEFAULT for a string-typed
-- value — so every score a user changed via the live per-cell stream reverted
-- to its default (5.0 sensory / 1.0 detailed) on the next fresh DB load.
--
-- Fix (writer half — stops NEW corruption): store a numeric-looking value as a
-- JSON NUMBER, falling back to text for genuine strings (names, free text):
--
--     CASE WHEN $2 ~ '^-?[0-9]+(\.[0-9]+)?$'
--          THEN to_jsonb($2::numeric)
--          ELSE to_jsonb($2::text) END
--
-- The reader half (C++ jsonToDouble/jsonToInt tolerant coercion) repairs the
-- EXISTING string-typed rows already in the production DB; it ships in the same
-- build and is the critical half for historical data.
--
-- CANONICAL SIGNATURE. The function that LiveSyncWorker actually calls is the
-- 7-argument OCC form introduced by 2026-05-17-v2.0.2-fixes.sql
-- (p_expected_version INT DEFAULT NULL) — NOT the 6-arg v2.0.1 form in
-- init.sql / 2026-05-16-livesync-commit-fns.sql (which lacks OCC and was
-- superseded). CREATE OR REPLACE with a different argument list creates an
-- OVERLOAD, not a replacement, so the live DB carries both; only the 7-arg one
-- is dispatched (the call binds 7 positional args). This migration REPLACES the
-- 7-arg form in place — same name, same argument list, same OCC behavior — and
-- changes ONLY the to_jsonb expression. The scalar dve_commit_cell is left
-- untouched (out of scope).
--
-- Delivered at runtime by DatabaseManager::ensureSchema() (the self-heal that
-- runs on every connect); the live NAS DB and the ephemeral test container both
-- get the new body when DataViewer next opens a connection. This file is the
-- durable CANONICAL RECORD of that runtime delivery — the source to reconcile
-- into deploy/postgres/init.sql when the baseline schema is next regenerated.
--
-- Fully idempotent and re-runnable: CREATE OR REPLACE FUNCTION with an
-- unchanged signature is a cheap in-place body swap.

BEGIN;

CREATE OR REPLACE FUNCTION dve_commit_cell_json(
    p_table             TEXT,
    p_row_id            BIGINT,
    p_path_text         TEXT,
    p_path_arr          TEXT[],
    p_value             TEXT,
    p_uuid              TEXT,
    p_expected_version  INT DEFAULT NULL
) RETURNS BOOLEAN AS $$
DECLARE
    affected INT;
BEGIN
    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true);
    PERFORM set_config('dve.live_value',  p_value,                     true);
    -- Store numeric-looking values as JSON NUMBERS; everything else as text.
    IF p_expected_version IS NULL THEN
        EXECUTE format(
            'UPDATE %I SET json_data = jsonb_set(json_data, $1, '
            '(CASE WHEN $2 ~ ''^-?[0-9]+(\.[0-9]+)?$'' '
            '      THEN to_jsonb($2::numeric) ELSE to_jsonb($2::text) END)::jsonb, '
            'true), updated_by = $3 '
            'WHERE id = $4',
            p_table
        ) USING p_path_arr, p_value, p_uuid, p_row_id;
    ELSE
        EXECUTE format(
            'UPDATE %I SET json_data = jsonb_set(json_data, $1, '
            '(CASE WHEN $2 ~ ''^-?[0-9]+(\.[0-9]+)?$'' '
            '      THEN to_jsonb($2::numeric) ELSE to_jsonb($2::text) END)::jsonb, '
            'true), updated_by = $3 '
            'WHERE id = $4 AND version = $5',
            p_table
        ) USING p_path_arr, p_value, p_uuid, p_row_id, p_expected_version;
    END IF;
    GET DIAGNOSTICS affected = ROW_COUNT;
    RETURN affected > 0;
END;
$$ LANGUAGE plpgsql;

COMMIT;
