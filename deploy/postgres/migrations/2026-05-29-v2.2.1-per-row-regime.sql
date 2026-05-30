-- v2.2.1 â per-row Puffing Regime column on data_rows.
-- Additive & non-destructive: old rows keep puffing_regime = NULL (read as
-- old template); new-template rows store a (possibly empty) string. Re-runnable.
BEGIN;
ALTER TABLE data_rows ADD COLUMN IF NOT EXISTS puffing_regime TEXT;
INSERT INTO schema_meta(key, value) VALUES ('schema_version', '3')
    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
COMMIT;
