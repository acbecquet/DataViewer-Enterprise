# v3.0.0 NAS migration runbook (supervised)

Owner-executed, step by step, on the production Postgres at the office NAS (dataviewer-db stack, port 5433).
The application NEVER runs any of this (index D-3d-2): the v3.0.0 build only PROBES for the post-cutover shape at connect and refuses a database that is not there yet, naming this runbook.
Every step is idempotent, so a failed step can be re-run after fixing its cause.

Do not start until v3.0.0 is otherwise fully ready to install (Phase 4 complete, internal smokes passed): between step 6 and each user's upgrade, v2.10.0 clients fail their saves with SQL errors (transactional - no corruption, but no saves), so the window must be as short as you can make it.

## 0. Preconditions

- [ ] Every client closed. Verify: `SELECT count(*) FROM presence;` returns 0 (or only stale rows older than the session timeout).
- [ ] The Synology release freeze is still in force (users on v2.10.0; nothing new published).
- [ ] You have shell/psql access to the NAS container and a copy of this repo's `deploy/postgres/migrations/` at the commit being shipped.

## 1. Backup (D7 - THE rollback path)

```bash
docker exec dataviewer-db pg_dump -U <user> -Fc -d <dbname> > dve_pre_v3_$(date +%Y%m%d).dump
```

- [ ] The dump file is non-trivially sized and `pg_restore --list` on it prints a table of contents.
- [ ] Copy it somewhere OFF the NAS as well.

Rollback at ANY later step = restore this dump.
`data_rows_pre_v3` / `samples_core` keep the pre-cutover wide data on the database as a forensic aid and the v1 pending-edits translation source, but they are NOT a rollback path: writes made after the cutover exist only in the long tables.

## 2. Oil units (D10 - must come FIRST)

Apply `2026-08-03-oil-units-to-grams.sql` by hand (psql `\i` or stdin).
It recomputes per-row and total oil from the before/after weights (idempotent, mixed-unit safe) and must not run while any pre-grams client can still write - which is why it lives in this runbook and not earlier.

- [ ] Spot-check: `SELECT oil_consumed FROM data_rows WHERE before_weight > 0 AND after_weight > 0 LIMIT 5;` shows gram-scale values (fractions of a gram), not 30-50-range milligrams.
- [ ] `SELECT value FROM schema_meta WHERE key = 'oil_units_grams';` is stamped.

## 3. Long-format schema

Apply `2026-07-31-v3-long-format.sql`.
Creates `metric_defs` (seeded), `measurements`, `sample_headers`, the compat views `data_rows_v` / `samples_v`, and `dve_migrate_to_long_format()`.
Touches no existing data.

- [ ] `SELECT count(*) FROM metric_defs;` >= 35.

## 4. Cutover machinery

Apply `2026-08-26-v3-cutover-1-functions.sql`.
Delivers the H23 `bump_version()` fix, `dve_commit_measurement()`, and `dve_cutover_to_long_format()`.
Still touches no data.

## 5. Optional inspection pass

```sql
SELECT * FROM dve_migrate_to_long_format();
```

Runs the data migration WITHOUT cutting over (wide tables stay authoritative; the function is idempotent so step 6 re-runs it for free).
Inspect at leisure:

```sql
-- row-count parity
SELECT (SELECT count(*) FROM data_rows)  AS wide_rows,
       (SELECT count(*) FROM data_rows_v) AS long_rows;
SELECT (SELECT count(*) FROM samples)   AS wide_samples,
       (SELECT count(*) FROM samples_v)  AS long_samples;
-- spot value parity
SELECT d.id, d.tpm, v.tpm FROM data_rows d
JOIN data_rows_v v USING (sample_id, sort_order)
WHERE d.tpm IS DISTINCT FROM v.tpm LIMIT 5;   -- expect 0 rows
```

If the migration RAISEs, it names the offending row (NULL sort_order, duplicate (sample_id, sort_order), or an all-NULL data row).
Fix that row by hand, re-run.

## 6. THE CUTOVER

Apply `2026-08-26-v3-cutover-2-execute.sql` (one statement: `SELECT * FROM dve_cutover_to_long_format();`).
The function re-runs the migration idempotently, verifies row-count parity AND the per-column sparse rule (refusing loudly on any mismatch), renames `data_rows -> data_rows_pre_v3` and `samples -> samples_core`, creates the name-holder views, revokes, and stamps `schema_meta`.

Expected output: one NOTICE with the inserted counts and a result row with `already_cut = f`.
A re-run prints `already_cut = t` and changes nothing.

- [ ] `SELECT c.relkind FROM pg_class c WHERE c.oid = to_regclass('data_rows');` returns `v`.
- [ ] `SELECT c.relkind FROM pg_class c WHERE c.oid = to_regclass('samples');` returns `v`.
- [ ] `SELECT value FROM schema_meta WHERE key = 'v3_long_format_cutover';` is stamped.
- [ ] A loadFile-shaped read works: `SELECT sample_name, tester, average_tpm FROM samples LIMIT 3;`

## 7. First client

Install v3.0.0 on ONE machine (yours).
Open a known file from the DB browser; verify values, edit a cell, save, reload.
The connect must succeed silently (the probe passes); a pre-cutover complaint here means step 6 did not take.

## 8. Fleet

Release v3.0.0 through the normal channel (owner-performed Synology transfer - ends the freeze).
Users upgrade via the hourly auto-update check; until each does, that user's saves fail with SQL errors and their edits stay in Excel/recovery - keep the window short and tell the lab.

## Rehearsal note (D7)

Before the real thing, rehearse steps 2-6 against a RESTORED COPY of a production dump on a local container.
The test-container harness already rehearses them against synthetic prod-shaped data on every run (tst_v3longformat), and the cutover function itself re-verifies parity + the sparse rule against the real data in step 6 - but a prod-dump rehearsal additionally smokes real-data oddities (the MigrationTool-era rows H17/H18/H21 anticipate).
Owner go-ahead required for pulling the dump.
