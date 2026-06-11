-- Migration: v2.0.2 save & DB correctness fixes.
--
-- Restores optimistic-concurrency control in the per-cell commit path
-- (silently disabled by the v2.0.1 perf-fix migration), scopes the cell-
-- focus DELETE to a single cell, replaces the invalid 6-field pg_cron
-- expression with a valid 5-field one, and adds a new stored function for
-- OCC-checked sensory layout saves.
--
-- All function CREATE OR REPLACE calls add a `p_expected_version INT
-- DEFAULT NULL` parameter on the existing scalar/json commit functions.
-- When NULL, behavior matches v2.0.1 (no OCC) so older clients still work
-- against this migration; when non-NULL, the UPDATE adds `AND version =
-- $N` and returns FALSE on miss. Apply this migration to the NAS Postgres
-- BEFORE rolling out the v2.0.2 installer.
--
-- Findings addressed:
--   C1: dve_commit_cell / dve_commit_cell_json gain OCC
--   M6: explicit version = version + 1 removed; bump_version trigger owns it
--   M4: dve_focus_cell DELETE scoped to "other cells", not all of user's focus
--   M5: dve_cell_focus_cleanup cron expression changed from invalid 6-field
--       '*/30 * * * * *' to valid 5-field '* * * * *'
--   H4: new dve_commit_session_layout for OCC-checked layout_json saves

BEGIN;

-- ── C1 + M6 ─ scalar cell commit with optional OCC, version managed by trigger
-- p_value is TEXT but the target column may be DOUBLE PRECISION (data_rows.tpm,
-- samples.viscosity, etc.). Pg doesn't implicitly cast text → numeric in
-- dynamic SQL, so we look up the column's declared type via pg_attribute and
-- insert an explicit ::<type> cast into the EXECUTE'd statement. This also
-- fixes a latent pre-v2.0.2 bug where the v2.0.1 stored function never
-- actually wrote to DOUBLE PRECISION columns (the open-coded write path in
-- DatabaseManager bound numeric QVariants directly and dodged the issue).
CREATE OR REPLACE FUNCTION dve_commit_cell(
    p_table             TEXT,
    p_row_id            BIGINT,
    p_column            TEXT,
    p_value             TEXT,
    p_uuid              TEXT,
    p_expected_version  INT DEFAULT NULL
) RETURNS BOOLEAN AS $$
DECLARE
    affected INT;
    col_type TEXT;
BEGIN
    SELECT format_type(atttypid, atttypmod)
      INTO col_type
      FROM pg_attribute
     WHERE attrelid = p_table::regclass
       AND attname  = p_column
       AND NOT attisdropped;
    IF col_type IS NULL THEN
        RAISE EXCEPTION 'dve_commit_cell: unknown column %.%', p_table, p_column;
    END IF;

    PERFORM set_config('dve.live_column', p_column, true);
    PERFORM set_config('dve.live_value',  p_value,  true);

    IF p_expected_version IS NULL THEN
        EXECUTE format(
            'UPDATE %I SET %I = $1::%s, updated_by = $2 WHERE id = $3',
            p_table, p_column, col_type
        ) USING p_value, p_uuid, p_row_id;
    ELSE
        EXECUTE format(
            'UPDATE %I SET %I = $1::%s, updated_by = $2 '
            'WHERE id = $3 AND version = $4',
            p_table, p_column, col_type
        ) USING p_value, p_uuid, p_row_id, p_expected_version;
    END IF;
    GET DIAGNOSTICS affected = ROW_COUNT;
    RETURN affected > 0;
END;
$$ LANGUAGE plpgsql;

-- ── C1 + M6 ─ JSON-path commit with optional OCC
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
    -- DATAVIEWER-4: store numeric-looking values as JSON NUMBERS, not strings.
    -- to_jsonb(text) wrote scores as "7.0"; the C++ reader's toDouble(default)
    -- returned the DEFAULT for a string-typed value, so every live-streamed
    -- score reverted on the next fresh DB load. The regex matches a plain
    -- integer/decimal (optional leading minus); names / free text still store
    -- as text. See deploy/postgres/migrations/2026-06-10-commit-cell-json-numeric.sql.
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

-- ── M4 ─ scope dve_focus_cell to not blow away the user's other focus rows.
-- cell_focus PK is (user_uuid, table_name, row_id, column_name), so a user
-- can legitimately have focus in multiple cells across panels. ON CONFLICT
-- refreshes the started_at timestamp on a re-focus of the same cell.
CREATE OR REPLACE FUNCTION dve_focus_cell(
    p_uuid       TEXT,
    p_table      TEXT,
    p_row_id     BIGINT,
    p_column     TEXT,
    p_user_name  TEXT,
    p_user_color TEXT
) RETURNS VOID AS $$
BEGIN
    DELETE FROM cell_focus
     WHERE user_uuid = p_uuid::uuid
       AND NOT (table_name = p_table
                AND row_id = p_row_id
                AND column_name = p_column);
    INSERT INTO cell_focus(user_uuid, table_name, row_id, column_name,
                           user_name, user_color, started_at)
    VALUES (p_uuid::uuid, p_table, p_row_id, p_column,
            p_user_name, p_user_color, now())
    ON CONFLICT (user_uuid, table_name, row_id, column_name)
    DO UPDATE SET user_name  = EXCLUDED.user_name,
                  user_color = EXCLUDED.user_color,
                  started_at = now();
END;
$$ LANGUAGE plpgsql;

-- ── H4 ─ OCC-checked sensory layout commit. Returns the new version on
-- success, NULL on stale-version conflict or unknown row. The caller
-- (DatabaseManager::saveSensoryLayout) decides whether NULL means "show
-- a stale-data dialog" or "fetch the latest and retry".
CREATE OR REPLACE FUNCTION dve_commit_session_layout(
    p_table             TEXT,
    p_session_id        BIGINT,
    p_layout            JSONB,
    p_uuid              TEXT,
    p_expected_version  INT
) RETURNS INT AS $$
DECLARE
    new_version INT;
BEGIN
    IF p_table NOT IN ('sensory_sessions') THEN
        RAISE EXCEPTION 'dve_commit_session_layout: invalid table %', p_table;
    END IF;
    PERFORM set_config('dve.live_column', 'layout_json',    true);
    PERFORM set_config('dve.live_value',  p_layout::text,   true);
    EXECUTE format(
        'UPDATE %I SET layout_json = $1, updated_by = $2 '
        'WHERE id = $3 AND version = $4 '
        'RETURNING version',
        p_table
    )
    USING p_layout, p_uuid, p_session_id, p_expected_version
    INTO new_version;
    RETURN new_version;
END;
$$ LANGUAGE plpgsql;

COMMIT;

-- ── M5 ─ pg_cron: the v2.0.1 cell_focus cleanup line used '*/30 * * * * *'
-- which is a 6-field expression. pg_cron uses standard 5-field cron and
-- rejects (or silently ignores) the 6-field form, so the cleanup job has
-- never run and dead cell_focus rows accumulate forever. Replace with a
-- valid 5-field expression. cron.schedule is idempotent on jobname.
-- Lives outside the BEGIN/COMMIT because cron operations are not always
-- transactional.
SELECT cron.schedule(
  'dve_cell_focus_cleanup',
  '* * * * *',
  $$DELETE FROM cell_focus WHERE started_at < now() - INTERVAL '30 seconds'$$
);
