-- DV-23: suppress no-op UPDATEs so they neither version-bump nor NOTIFY.
--
-- Symptom: a client that re-saved an unchanged sensory_sessions row ~1/sec was
-- firing a continuous NOTIFY stream to every live client. The row-change
-- trigger fires on EVERY UPDATE - even one that writes identical values - and
-- bump_version bumped version + updated_at on every UPDATE too. So a client
-- stuck in a re-save loop stormed everyone regardless of whether anything
-- actually changed.
--
-- Fix at the trigger layer so the storm dies no matter which caller loops,
-- SCOPED to the whole-row session tables where it was observed:
--   * bump_version: on sensory_sessions / detailed_sensory_sessions, if no
--     NON-AUDIT column actually changed, leave version and updated_at untouched
--     (RETURN NEW, not NULL, so the row still "updates" to identical values and
--     RETURNING still yields it - callers' OCC checks and the *_commit_* stored
--     functions are unaffected). The TPM files/tests/samples/data_rows hierarchy
--     is deliberately NOT scoped in: a whole-file save touches the parent row as
--     part of that path's multi-client OCC semantics (out of scope for DV-23).
--   * notify_row_change: skip pg_notify on an UPDATE whose version was NOT
--     bumped (i.e. a suppressed no-op). INSERT/DELETE always notify. This is
--     coupled to bump_version's decision, so it follows the same scoping.
--
-- "Material change" = any column other than version / updated_at / app_version
-- differs. Compared via jsonb so json_data equality is key-order-insensitive.
-- updated_by IS compared: a write by a DIFFERENT author is a real event that
-- should version + notify (the reported storm is the SAME client re-writing the
-- same row, so updated_by is identical between its writes and it stays a no-op).
-- app_version is subtracted defensively (bump_version runs before the
-- app_version stamp trigger; the key is simply absent on tables without that
-- column, which is a safe no-op).
--
-- Idempotent: pure CREATE OR REPLACE. The triggers reference these functions by
-- name and pick up the new bodies automatically, so no trigger DDL is needed.
-- Safe to re-run.

CREATE OR REPLACE FUNCTION bump_version() RETURNS TRIGGER AS $$
BEGIN
  -- No-op UPDATE on a SESSION table (nothing material changed): keep the row
  -- inert so it doesn't churn version/updated_at or (via notify_row_change)
  -- storm live clients. Scoped to the whole-row session tables where the storm
  -- was observed; the TPM files/tests/samples/data_rows hierarchy KEEPS its
  -- always-bump behavior (a whole-file save's parent-row touch is part of that
  -- path's established multi-client OCC semantics and is out of scope here).
  IF TG_OP = 'UPDATE'
     AND TG_TABLE_NAME IN ('sensory_sessions', 'detailed_sensory_sessions')
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

CREATE OR REPLACE FUNCTION notify_row_change() RETURNS TRIGGER AS $$
DECLARE
  payload JSONB;
BEGIN
  -- Maintenance/bulk writes (the legacy-score normalizer) set this GUC so they
  -- don't fan out a NOTIFY per row and storm every live client.
  IF current_setting('dve.maintenance', true) = '1' THEN RETURN NULL; END IF;
  -- DV-23: bump_version left the version unchanged on a no-op UPDATE, so nothing
  -- really changed -> don't NOTIFY. INSERT/DELETE always notify.
  IF TG_OP = 'UPDATE' AND NEW.version = OLD.version THEN RETURN NULL; END IF;
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
