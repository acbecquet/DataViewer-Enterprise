-- Plan B: persist SOP/raw-table grid so DB-only loads render it.
-- Additive & NULL-default: old binaries (which SELECT explicit columns) are unaffected.
-- Re-runnable.
BEGIN;
ALTER TABLE tests ADD COLUMN IF NOT EXISTS raw_grid JSONB;
COMMENT ON COLUMN tests.raw_grid IS
  'SOP/raw-table sheet content {"headers":[...],"rows":[[...],...]}; NULL for normal sheets.';
COMMIT;
