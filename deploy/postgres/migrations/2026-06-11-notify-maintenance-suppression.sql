-- 2026-06-11-notify-maintenance-suppression.sql
-- v2.4.2 R4 (Sub-plan 2, Task 1): suppress NOTIFY for maintenance/bulk writes.
--
-- The nightly legacy-score normalizer (dve_normalize_legacy_json) rewrites
-- json_data across many rows in one statement. Without this guard, the AFTER
-- ROW trigger notify_row_change would fan out one NOTIFY per touched row and
-- storm every live client with a redundant refresh burst (the conditions that
-- resurrect the DATAVIEWER-4 reset-to-5 bug under a flapping network). A bulk
-- writer now sets the dve.maintenance GUC to '1' for its transaction and the
-- trigger short-circuits before building/sending any payload.
--
-- Body is byte-faithful to init.sql's notify_row_change PLUS the early-return;
-- the trigger wiring is unchanged, so CREATE OR REPLACE FUNCTION swaps the body
-- in place with no DROP/CREATE TRIGGER needed. Idempotent: re-running just
-- re-replaces the same definition.

CREATE OR REPLACE FUNCTION notify_row_change() RETURNS TRIGGER AS $$
DECLARE
  payload JSONB;
BEGIN
  -- v2.4.2 R4: maintenance/bulk writes (the legacy-score normalizer) set this
  -- GUC so they don't fan out a NOTIFY per row and storm every live client.
  IF current_setting('dve.maintenance', true) = '1' THEN RETURN NULL; END IF;
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
