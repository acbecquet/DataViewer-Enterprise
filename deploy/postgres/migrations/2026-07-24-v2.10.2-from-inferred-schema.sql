-- v2.10.2 - persist the header-driven schema-inference flag on tests.
-- Additive & non-destructive: existing rows default to 0 (standard 12-wide
-- layout), so every already-stored sheet keeps today's read/write behavior.
-- A non-zero value marks a sheet parsed by SchemaInference (non-standard
-- 13/8-column block layout, e.g. the S26 Cart-era / UserSim sheets) whose
-- 12-wide Excel write-back must stay disabled on reopen - without this column a
-- DB round-trip would silently re-enable write-back against a misaligned file.
-- Re-runnable.
BEGIN;
ALTER TABLE tests ADD COLUMN IF NOT EXISTS from_inferred_schema INTEGER NOT NULL DEFAULT 0;
COMMIT;
