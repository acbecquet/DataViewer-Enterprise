-- Migration: F6 (v2.5.0) — versioned TPM file re-adds via (file_path, added_at).
--
-- v2.4.0 save/sync fix batch. USER REQUIREMENT: "for TPM files, each file
-- should get a date and time extension to the end for the database, in case
-- new versions of the same test are added." Until now files.file_path was
-- UNIQUE (idx_files_path), so re-adding the same .xlsx LATER silently
-- UPDATEd/overwrote the existing DB record — losing the prior version.
--
-- Desired semantics (decided):
--   * every fresh load-from-DISK mints a NEW files row stamped with the add
--     time (added_at);
--   * in-session saves keep UPDATEing that same row by id (added_at untouched);
--   * loads FROM the DB adopt the loaded row's id (continue that version);
--   * old rows remain as historical versions.
--
-- Delivered at runtime by DatabaseManager::ensureSchema() (the self-heal that
-- runs on every connect); the test harness does NOT apply the files under
-- deploy/postgres/migrations/ automatically. This file is the durable
-- CANONICAL RECORD of that runtime delivery — the source to reconcile into
-- deploy/postgres/init.sql when the baseline schema is next regenerated.
--
-- It faithfully mirrors what ensureSchema() does:
--   (1) add the additive files.added_at column if a legacy DB lacks it;
--   (2) swap uniqueness from the legacy single-column UNIQUE(file_path) to the
--       composite UNIQUE(file_path, added_at) — letting the same path live once
--       per add-time (i.e. once per version).
--
-- Fully idempotent and re-runnable: every step uses IF EXISTS / IF NOT EXISTS.
-- added_at is NOT NULL DEFAULT now(); on a legacy table with existing rows the
-- ADD COLUMN backfills every row with the migration time, which is distinct
-- per path (one row per path pre-migration), so the new composite unique index
-- builds cleanly.

BEGIN;

-- (1) Heal a legacy files table that predates the added_at column. NOT NULL
--     DEFAULT now() stamps existing rows at migration time; the new index can
--     then build because pre-migration each path had at most one row.
ALTER TABLE files
    ADD COLUMN IF NOT EXISTS added_at TIMESTAMPTZ NOT NULL DEFAULT now();

-- (2) Swap uniqueness from legacy (file_path) to the composite
--     (file_path, added_at). Drop the old single-column unique index/constraint
--     (it lived as idx_files_path in init.sql; a constraint form may exist on
--     very old DBs), then build the composite unique index. IF EXISTS / IF NOT
--     EXISTS keep it idempotent and re-runnable.
DROP INDEX IF EXISTS idx_files_path;
ALTER TABLE files DROP CONSTRAINT IF EXISTS files_file_path_key;

CREATE UNIQUE INDEX IF NOT EXISTS uq_files_path_added
    ON files(file_path, added_at);

COMMIT;
