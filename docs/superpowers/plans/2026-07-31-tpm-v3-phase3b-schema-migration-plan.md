# TPM v3 Phase 3b (long-format schema + migration, authored and rehearsed) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stand up the long-format schema (`metric_defs`, `sample_headers`, `measurements`), the read-only compatibility views, and the one-shot data migration - and prove the migration is value-exact on prod-shaped data - **without changing one line of app behavior and without migrating any live data.**

**Architecture:** The tables are created empty and nothing reads or writes them yet.
The migration is authored as a server-side function that is never called automatically; 3d calls it at the real cutover, when the app is finally able to maintain those rows.
This ordering is deliberate - see the SEQUENCING CORRECTION section of the Phase 3 index. Backfilling `measurements` now would create standard-metric rows that nothing maintains for two phases, so the migrated copy would rot.
The compat views exist in 3b purely so the rehearsal harness can prove parity; they become load-bearing in 3d when `data_rows` / `samples` are renamed aside and the views take their names.

**Tech Stack:** PostgreSQL 16 (container `dve-test-pg`, port 5433), C++17 / Qt 6.10 (qmake + MinGW), Qt Test.

---

## Machine + repo rules (read first)

- Create all new source files with the Write tool ONLY - never python file writes (MIP labeling), never heredoc/echo.
- **NEVER use `git stash`** - it rewrites the working tree, re-triggers the MIP labeler and re-encrypts source mid-session, and the stash stack is shared across worktrees. Use a temporary commit instead.
- Ciphertext (`%TSD-Header-###%`): read via `git show HEAD:<path>`; run `python tools/decrypt_via_copy.py --apply` from repo root before any build.
- Public repo: `tests/corpus/` gitignored; never commit real workbooks or `results.txt` artifacts.
- Branch `worktree-tpm-template-v3-research`. Commit per task; plain dashes; NO Co-Authored-By.
- Qt Test stdout is INVISIBLE - always `-o results.txt,txt`, and `/c/Qt/6.10.1/mingw_64/bin` MUST be on PATH when running test exes (silent death otherwise).
- Suite inner loop (from suite dir): `export PATH="/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH"`, `export DVE_TEST_PG_CONN="host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"`, prepend `vendor/libpq-16` to PATH, qmake, mingw32-make, `release/<suite>.exe -o results.txt,txt`.
- Run PG-dependent suites ONE AT A TIME - they wipe shared tables.
- Full-suite gate: `powershell -ExecutionPolicy Bypass -File tests/run-tests.ps1` from repo root. NEVER concurrently with an app build.
- Never touch the production/NAS database. NEVER kill Excel processes.
- -Werror -Wall -Wextra -Wpedantic.

## Reference: verified seams (2026-07-30 audit)

- **Provisioning:** `tests/start-test-postgres.ps1` applies `init.sql`'s pre-pg_cron block then every `deploy/postgres/migrations/*.sql` in filename order, now with `ON_ERROR_STOP=1` and a per-file exit check. A new migration file is picked up automatically. Production has NEVER auto-applied a migration file - `docker-compose.yml` mounts only `init.sql`, and only on an empty data dir - so `ensureSchema()` is the real delivery mechanism for the NAS.
- **`ensureSchema()`** (`src/database/DatabaseManager.cpp:129-880`) runs on every `open()`/`reopen()`, is best-effort and never throws. It reconciles COLUMNS only via `kAdditiveColumns[]` (`:149-166`) - **it cannot create tables** (H12). The catalog-guarded `CREATE` precedent to copy is the `dve_stamp_app_version` trigger block at `:218-260` (probe `pg_trigger` / `pg_proc`, then create).
- **One-shot gating precedent** (`:578-609`): `db.transaction()` -> `SELECT pg_advisory_xact_lock(<id>)` -> probe `schema_meta` for a key -> run -> `INSERT INTO schema_meta ... ON CONFLICT DO NOTHING`.
- **`bump_version()`** (`init.sql:264-289`, current body `2026-07-15-notify-noop-suppression.sql:36-59`) sets `version := OLD.version + 1` and refuses client-supplied values; its no-op suppression branch is scoped to `('sensory_sessions','detailed_sensory_sessions')`.
- **`notify_row_change()`** (`init.sql:345-392`) dereferences `NEW.id`/`OLD.id`; `settings` is excluded from the notify pool precisely because its PK is `key`.
- **The wide surface to mirror:** `data_rows` 13 value columns (`puffs, before_weight, after_weight, draw_pressure, resistance, smell, clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed, puffing_regime`) + `sort_order`. `samples` 22 value columns (5 identity/header, 7 device, 7 derived aggregates, 3 status) + `sort_order`. Full lists in `init.sql:52-106`.
- **Namespace collision is real:** `resistance` and `puffing_regime` each exist on BOTH `samples` and `data_rows`, and metrics vs header fields are separate namespaces in the registry. Uniqueness must therefore be `(kind, key)`, never `(key)`.
- **`data_rows` has NO unique constraint on `(sample_id, sort_order)`** - only PK on `id`. `MigrationTool` copied legacy `sort_order` verbatim, so duplicates are possible in real data.
- **`MetricRegistry`** (`src/model/MetricRegistry.h`) exposes `allMetrics()` / `allHeaderFields()` returning `QVector<MetricDef>` / `QVector<HeaderFieldDef>`; keys are already the snake_case column names. `MetricDef` carries key, display name, `ValueType`, unit, role, tags.
- **`sensory_web` role** (`2026-06-25-dv11-sensory-web-role.sql`) is granted narrowly and has an explicit `REVOKE EXECUTE` loop over every `dve_commit_cell%` overload, plus a negative test at `deploy/sensory-collect/tests/test_append.py:405-408`.

## Non-goals

- **NO migration of live data.** The migration function is authored and rehearsed only; nothing calls it outside the harness. 3d runs it for real.
- NO app read or write path changes - `DatabaseOps.cpp` and `DatabaseManager`'s read path SQL stay byte-identical. `extra` persistence is 3c.
- NO renaming of `data_rows` / `samples`, and NO `INSTEAD OF` triggers on the views. The views are read-only bridges (index D1).
- NO notify trigger on the new tables (index D3).
- NO `dve_commit_measurement` yet - that lands with the LiveSync cutover in 3d.
- NO changes to `MigrationTool` beyond its wipe/order lists if a test demands it.
- NO production database access of any kind.

---

### Task 1: The long-format schema migration file

**Files:**
- Create: `deploy/postgres/migrations/2026-07-31-v3-long-format.sql`

Write it as one `BEGIN; ... COMMIT;` block. It must be idempotent (`IF NOT EXISTS` / `CREATE OR REPLACE`) because the provisioning script and `ensureSchema` may both see it. Do NOT put any `pg_cron` statement in it.

**Steps:**

- [ ] `metric_defs`: `id BIGSERIAL PK`, `kind TEXT NOT NULL CHECK (kind IN ('metric','header'))`, `key TEXT NOT NULL`, `display_name TEXT NOT NULL`, `value_type TEXT NOT NULL`, `unit TEXT`, `role TEXT`, `tags JSONB`, plus the standard audit trio (`updated_at TIMESTAMPTZ NOT NULL DEFAULT now()`, `updated_by TEXT NOT NULL DEFAULT 'migration'`, `version INTEGER NOT NULL DEFAULT 1`). **`UNIQUE(kind, key)`** - not `UNIQUE(key)`, because `resistance` and `puffing_regime` legitimately exist in both namespaces.
- [ ] `measurements`: `id BIGSERIAL PK`, `sample_id BIGINT NOT NULL REFERENCES samples(id) ON DELETE CASCADE`, `metric_id BIGINT NOT NULL REFERENCES metric_defs(id)`, `sort_order INTEGER NOT NULL`, `value_num DOUBLE PRECISION`, `value_text TEXT`, audit trio, **`UNIQUE(sample_id, metric_id, sort_order)`**. Index on `(sample_id)` for the per-sample read.
- [ ] `sample_headers`: `id BIGSERIAL PK`, `sample_id BIGINT NOT NULL REFERENCES samples(id) ON DELETE CASCADE`, `field_id BIGINT NOT NULL REFERENCES metric_defs(id)`, `value_num DOUBLE PRECISION`, `value_text TEXT`, audit trio, **`UNIQUE(sample_id, field_id)`**. Index on `(sample_id)`.
- [ ] Attach `bump_version()` BEFORE UPDATE triggers to all three tables. **Do NOT attach `notify_row_change()`** - index D3; leave a comment in the file saying why (TPM row notifications drive zero UI, and one whole-file save would emit ~13k of them).
- [ ] Extend `bump_version()`'s no-op suppression list to include `measurements` and `sample_headers` (index D4). Use `CREATE OR REPLACE FUNCTION` and keep the rest of the body byte-identical to the current version in `2026-07-15-notify-noop-suppression.sql` - copy it, add the two names, change nothing else.
- [ ] Seed `metric_defs` with every standard metric and header key, `INSERT ... ON CONFLICT (kind, key) DO NOTHING`. These must cover exactly the 13 `data_rows` value columns (kind `metric`) and the 22 `samples` value columns (kind `header`) so the migration can map every column. Source the display names/units/roles from `docs/superpowers/specs/2026-07-27-tpm-v3-vocabulary-registry-draft.md`, and note in a comment that `ensureSchema` upserts the full compiled registry on top of this seed (Task 2).
- [ ] `REVOKE ALL ON metric_defs, measurements, sample_headers FROM PUBLIC, sensory_web;` (H16), with a comment pointing at the negative test.

**Verification:**

- [ ] Recreate the container (`docker rm -f dve-test-pg` then `tests/start-test-postgres.ps1`) and confirm the script applies the new file with no error and reports the expected object counts.
- [ ] Query the catalog to confirm: three tables present; `UNIQUE(kind,key)` on `metric_defs`; `UNIQUE(sample_id, metric_id, sort_order)` on `measurements`; `bump_version` triggers present on all three; **`notify_row_change` triggers ABSENT on all three**; seed row counts are 13 metrics + 22 headers.

---

### Task 2: `ensureSchema` creates the tables and upserts the compiled registry

**Files:**
- Modify: `src/database/DatabaseManager.cpp` (`ensureSchema`)
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp`

**Why:** production never auto-applies migration files, so without this the NAS would never get the tables (H12). And the registry will keep growing (Phase 5's template builder adds metrics), so seeding must be driven by the compiled `MetricRegistry`, not a frozen SQL literal.

**Steps:**

- [ ] Add a catalog-guarded table-creation block modelled on the `dve_stamp_app_version` precedent at `:218-260`: probe `pg_class` for each of the three tables, and create any that are missing using the SAME DDL as Task 1's migration. Keep the DDL in one place if you can do so cleanly; if you must duplicate it, add a test that asserts the two definitions agree (compare `information_schema.columns` after each path).
- [ ] Upsert `metric_defs` from `MetricRegistry::allMetrics()` (kind `metric`) and `allHeaderFields()` (kind `header`): `INSERT ... ON CONFLICT (kind, key) DO UPDATE` on display name / value_type / unit / role / tags. Never delete rows - a key that leaves the registry must keep its row so historical measurements still resolve (registry naming policy: keys are forever).
- [ ] Follow the existing best-effort contract: log via `logDebug` and never throw. A DB that refuses the DDL must not stop the app from opening.
- [ ] Red first: a test that drops one of the three tables out of band, calls `ensureSchema`, and asserts it is recreated; and a test asserting every `MetricRegistry` metric and header key has a matching `metric_defs` row after `ensureSchema`.

---

### Task 3: Compatibility views

**Files:**
- Modify: `deploy/postgres/migrations/2026-07-31-v3-long-format.sql` (same file, same transaction)

**Steps:**

- [ ] `data_rows_v`: pivot `measurements` back to the wide `data_rows` shape - one row per `(sample_id, sort_order)`, one output column per standard metric key, using `MAX(value_num) FILTER (WHERE md.key = '...')` / `MAX(value_text) FILTER (...)`. Join `metric_defs` on `metric_id`.
- [ ] Synthesize the non-value columns a reader needs: `version` as `MAX(m.version)` over the group, `updated_at` as `MAX(m.updated_at)`, `updated_by` as the `updated_by` of the most recently updated member. For `id`, emit `MIN(m.id)` and **document in a comment that this is NOT a stable `data_rows.id`** - it is a synthetic surrogate adequate for read-only consumers and nothing may treat it as an identity to write back through.
- [ ] `samples_v`: same treatment over `sample_headers`, one row per `sample_id`, one column per standard header key. Note it exposes only the header columns - `samples`' own identity/audit columns still come from `samples` itself, so `samples_v` is a JOIN partner, not a replacement. Document that.
- [ ] Both views must be `CREATE OR REPLACE VIEW`. Grant nothing to `sensory_web`.
- [ ] Add a comment in each view stating that a NULL output column means "no measurement row exists", which under sparse materialization (index D2) is exactly what the source NULL meant.

---

### Task 4: The migration function (authored, NOT executed)

**Files:**
- Modify: `deploy/postgres/migrations/2026-07-31-v3-long-format.sql`

Write `dve_migrate_to_long_format()` returning a summary (row counts). It is **never called by application code in this phase** - only by the rehearsal harness.

**Steps:**

- [ ] **Pre-flight abort:** first statement checks `SELECT 1 FROM data_rows GROUP BY sample_id, sort_order HAVING count(*) > 1 LIMIT 1`. If any row exists, `RAISE EXCEPTION` naming the offending `sample_id`/`sort_order` and abort - the new `UNIQUE(sample_id, metric_id, sort_order)` would change what is legal, and silently collapsing duplicates would lose data. This is the poka-yoke that makes running it against real data safe.
- [ ] **Sparse rule (index D2), stated as the implementation rule:** insert a `measurements` row for a `(data_row, metric)` pair **if and only if the corresponding wide column IS NOT NULL**. Numeric zeros are real measurements and ARE inserted. This is not an optimization - `hasPerRowRegime` is derived from `puffing_regime IS NOT NULL`, so materializing an empty regime for every row would flip every old-template sheet into per-row-regime mode.
- [ ] Route each column to `value_num` or `value_text` by the registry's `value_type`. Carry `updated_by` from the source row; let the trigger own `updated_at`/`version`.
- [ ] Same treatment for `samples` -> `sample_headers`, keyed on `(sample_id, field_id)`.
- [ ] Make it idempotent: `ON CONFLICT DO NOTHING` on both unique keys, so a re-run after a partial failure is safe.
- [ ] Return counts (`measurements_inserted`, `sample_headers_inserted`, `data_rows_seen`, `samples_seen`) so the harness can assert on them.
- [ ] Do NOT gate it with `schema_meta` here and do NOT call it from `ensureSchema` - 3d owns the gated invocation. Put a comment at the top saying exactly that.

---

### Task 5: Rehearsal harness

**Files:**
- Create: `tests/tst_v3longformat/tst_v3longformat.pro`, `tests/tst_v3longformat/tst_v3longformat.cpp`
- Modify: `tests/tests.pro` (SUBDIRS)

This is the phase's real gate. It must be able to fail.

**Steps:**

- [ ] Seed a prod-shaped dataset directly into the wide tables: several files, multiple sheets each, several samples per sheet, tens of data rows per sample, with a realistic mix of NULL and non-NULL text columns (`smell`/`clog`/`notes` mostly NULL, `puffing_regime` NULL on some sheets and set on others) and numeric zeros present. Use the shapes from the audit: 1-6 sheets/file, ~4-9 samples/sheet, 12-33 rows/sample.
- [ ] Run `dve_migrate_to_long_format()`.
- [ ] **Value parity:** assert `data_rows_v` is row-for-row identical to `data_rows` on every one of the 13 value columns, joined on `(sample_id, sort_order)`, with matching row counts on both sides. Do the same for `samples_v` vs `samples` on all 22 header columns. Use `IS NOT DISTINCT FROM` so NULLs compare equal.
- [ ] **Per-metric checksums:** for each numeric metric, compare `sum`, `count`, `min`, `max` between source column and migrated rows. For each text metric compare `count` and an aggregate hash of the ordered values.
- [ ] **Sparse rule:** assert that a source NULL produced NO measurement row (not a row with NULL value), and that a source `0.0` DID produce a row with `value_num = 0`.
- [ ] **`hasPerRowRegime` preservation:** assert that a sheet whose `puffing_regime` was entirely NULL produces zero `puffing_regime` measurement rows, so the existing derivation still yields false.
- [ ] **Pre-flight abort:** insert a deliberate duplicate `(sample_id, sort_order)` and assert the function RAISEs and inserts nothing.
- [ ] **Idempotence:** run the migration twice and assert the second run inserts zero additional rows and parity still holds.
- [ ] Add the three new tables to this suite's wipe list, and to the other wipe lists only where a test actually breaks without it (H15 - do not churn all seven speculatively).

---

### Task 6: Gates and wrap

- [ ] `python tools/decrypt_via_copy.py --apply`, then a clean app build - `-Werror` must stay clean. The app is unchanged, so this is a regression check on Task 2 only.
- [ ] Full suite, not concurrent with a build. Record counts.
- [ ] Corpus shadow and corpus round-trip harnesses (`DVE_TEST_CORPUS_DIR=tests/corpus`) - both must match the Phase 2c/3a baseline exactly (26/0/6 and 38/0/0). Nothing in 3b should be able to move them; if either moves, stop and investigate.
- [ ] Confirm the app still opens against a migrated database and behaves identically - the wide tables are still authoritative and untouched.
- [ ] Independent review of the whole 3b diff against this plan.
- [ ] Update the tracker and the v3 memory topic file.
