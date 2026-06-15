-- Migration: v2.4.2 A1 — server-side app_version era stamping.
--
-- Adds a nullable app_version TEXT column to files / sensory_sessions /
-- detailed_sensory_sessions and a BEFORE INSERT OR UPDATE trigger that fills it
-- from current_setting('application_name') (the app sets application_name=
-- DataViewer/<ver> in its connect string). Fail-safe: nullable, no CHECK, never
-- RAISE, COALESCE so a NULL-name (old-client) session never blanks a good stamp.
--
-- Delivered at runtime by DatabaseManager::ensureSchema(); this file is the
-- durable CANONICAL RECORD to reconcile into init.sql at the next baseline.
-- Fully idempotent / re-runnable.

BEGIN;

ALTER TABLE files                     ADD COLUMN IF NOT EXISTS app_version TEXT;
ALTER TABLE sensory_sessions          ADD COLUMN IF NOT EXISTS app_version TEXT;
ALTER TABLE detailed_sensory_sessions ADD COLUMN IF NOT EXISTS app_version TEXT;

CREATE OR REPLACE FUNCTION dve_stamp_app_version() RETURNS TRIGGER AS $$
BEGIN
  NEW.app_version := COALESCE(
      NEW.app_version,
      NULLIF(current_setting('application_name', true), ''));
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DO $$
DECLARE t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY['files','sensory_sessions','detailed_sensory_sessions']
  LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_stamp_app_version ON %I;
       CREATE TRIGGER trg_%I_stamp_app_version
       BEFORE INSERT OR UPDATE ON %I
       FOR EACH ROW EXECUTE FUNCTION dve_stamp_app_version();',
      t, t, t, t);
  END LOOP;
END$$;

COMMIT;
