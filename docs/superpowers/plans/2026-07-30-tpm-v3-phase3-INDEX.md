# TPM v3 Phase 3 (DB long-format) - master index

**Status:** 3a in progress (2026-07-30). 3b/3c/3d planned, not started.

**Goal:** Move the Postgres schema from fixed metric columns to a long/tidy format - one row per `(sample, metric, sort_order)` - so that custom and reordered metrics survive a database round-trip, and so `DataRow::extra` (the open metric map that Phase 2 fills for inference and manifest sheets) finally persists.

**Predecessor:** Phase 2c, merged to `main` at `8826a8e`.
Workbooks are already self-describing and parse by name; the database is the last layer that still hardcodes the 12-column template.

---

## Pre-flight audit (2026-07-30)

Three read-only audits were run against `8826a8e` before planning: the save/persist path, the full DB surface, and the live-sync/snapshot/serialization layers.
The plan's own coexistence rule required this after the v2.10 merge, and it paid for itself - three findings below invert what the original research plan assumed.

### Finding 1 - the NOTIFY scaling risk is avoidable entirely, not mitigable

The research plan assumed live updates would need re-pointing at measurement granularity, and that a whole-file save fanning ~1,500 `data_rows` UPDATEs into ~13,000 measurement UPDATEs would be a live-sync scaling problem in the DV-23 storm class.

The audit proved that TPM row NOTIFYs drive **zero** UI today.
`MainWindow::handleRemoteRowChange` (`src/MainWindow.cpp:3787-3822`) handles only `op == "DELETE"` for the three top-level resource tables and returns for everything else, with a comment saying so explicitly.
`LiveSync::cellChanged` (`src/database/LiveSync.cpp:530`) has exactly two consumers, both sensory JSON-path handlers; there is no TPM subscriber at all.

So the decision is not "how do we throttle 13,000 NOTIFYs" but "don't emit them."
See D3.

### Finding 2 - the orphan prune is the top data-loss risk, and it already has a live instance

`persistFileCore` phase C (`src/database/DatabaseOps.cpp:759-784`) captures a pre-image of every child id belonging to the file and DELETEs anything the in-memory model did not reproduce.
Today that is safe because the wide columns are exactly the in-memory fields.

It is already unsafe for `DataRow::extra`: `DatabaseManager::loadFile` cannot read extras (they are not persisted), so load-from-database then save silently drops them.
The gap is documented in-code at `src/MainWindow.cpp:1181-1186` and names Phase 3 as the fix.
Under long format the pre-image becomes every metric of every row, so any metric the model fails to reproduce is deleted on the next save.

This makes "persist `extra`" a prerequisite for the write path, not a nice-to-have, and it makes the prune rework the highest-scrutiny task in 3d.

### Finding 3 - the tool we would verify the migration with is silently broken

`tests/start-test-postgres.ps1:85` applies every migration with `ON_ERROR_STOP=0` and pipes psql output to `Out-Null`, so a failing migration produces a clean-looking provisioning run and a silently wrong schema.
The same script (`:60-65`) slices only the first `BEGIN;...COMMIT;` block out of `init.sql`, so any DDL after that first `COMMIT;` is invisible to the entire test suite.

Fixing this is task 1 of 3a. Every later gate reads through this script.

---

## Locked decisions

**D1 - Compatibility views are read-only bridges; the C++ write path goes direct to the long tables.**
Making writes work through views needs `INSTEAD OF` triggers, and Postgres does not fire the base tables' `bump_version` / `notify_row_change` row triggers through a view, so they would have to be reimplemented in the trigger body.
Worse, optimistic-locking would land on a synthetic view row, defeating the per-measurement versioning that is the point of the exercise.
Views exist so the offline-snapshot regen, `MigrationTool`, and any external SQL keep working unchanged while the C++ read and write paths are cut over one at a time.

**D2 - Sparse materialization: a measurement row exists exactly where the wide column is non-NULL today.**
Numeric zeros are real measurements and are written; SQL NULLs are absent.
This is not a size optimization, it is a correctness requirement: `SheetResult::hasPerRowRegime` is derived from `puffing_regime IS NOT NULL` (`src/database/DatabaseManager.cpp:1117-1118`, `src/database/OfflineSnapshot.cpp:1695-1696`), so a dense migration that materialized an empty regime for every row would flip every old-template sheet into per-row-regime mode.
The rule also makes migration parity mechanically checkable: after migrating, `data_rows_v` must be row-for-row identical to the pre-migration `data_rows`.

**D2a - `tests.has_per_row_regime` becomes an explicit column.**
D2 keeps the existing derivation correct, but the derivation stays a landmine: any future change to what gets materialized silently re-labels historical sheets.
One additive column with a backfill from the current derivation removes the class.
Cheap poka-yoke; cut it only if 3b runs long.

**D3 - `measurements` and `sample_headers` do NOT get the `notify_row_change` trigger.**
Justified by Finding 1: no client consumes TPM row NOTIFYs.
Resource-level NOTIFY on `files` is what `RowDeletedBanner` actually uses and it is untouched.
This removes the largest scaling risk in the phase at the cost of zero code.

**D4 - `bump_version` no-op suppression extends to the new tables.**
Currently scoped to the two session tables (`deploy/postgres/migrations/2026-07-15-notify-noop-suppression.sql:44-48`).
Prevents version churn on unchanged measurements during whole-file saves.

**D5 - A new `dve_commit_measurement(...)` stored function for per-cell live commits.**
`dve_commit_cell` is `format('UPDATE %I SET %I = $1::%s ... WHERE id = $3')` against a `pg_attribute` type lookup - it structurally requires the metric to be a column.
It stays, unchanged, for the sensory JSONB path.
Note the latent overload hazard: a 5-positional-arg call to `dve_commit_cell` is already ambiguous (verified live against the test container) because the legacy 5-arg form coexists with the 6-arg-with-DEFAULT form.
Do not add overloads to the existing functions; add a new name.

**D6 - Table creation rides `ensureSchema`; the one-shot data migration does not.**
`ensureSchema` runs on every connect from every client, is best-effort, and never throws - correct for additive DDL (catalog-guarded `CREATE` precedent at `src/database/DatabaseManager.cpp:218-260`), wrong for a one-shot data transform.
The data migration is advisory-locked and `schema_meta`-gated, modelled on the existing `v242_legacy_score_normalize` pattern (`src/database/DatabaseManager.cpp:578-609`).
Note that production has never auto-applied a file from `deploy/postgres/migrations/` - `docker-compose.yml` mounts only `init.sql`, and only on an empty data directory.

**D7 - Rehearsal happens twice: synthetic now, production dump before the NAS step.**
3b rehearses on a prod-shaped synthetic database seeded from the corpus workbooks, which gets the parity and checksum gates without touching production.
The real production-dump rehearsal is a hard gate before the NAS migration, which happens at v3.0.0 ship time.
**That step needs the owner's go-ahead** - taking a dump reads the live NAS database, and the standing guardrail is that live queries are approved case by case.

---

## Sub-phases and gates

### 3a - Pre-flight hardening (no schema change)

Fix the defects that would silently mask a Phase 3 regression, plus the two that compound directly under long format.
Nothing here changes behavior a user would notice; everything here is verifiable today.
Detailed plan: `2026-07-30-tpm-v3-phase3a-preflight-hardening-plan.md`.

**Gate:** full suite green, and the test-container provisioning script demonstrably fails loudly on a deliberately broken migration.

### SEQUENCING CORRECTION (2026-07-31) - read the rationale before touching 3c

The original split (3c = read path, 3d = write path) is **unshippable at any midpoint**.
If reads move to the long tables while writes still go to the wide ones, the app reads from tables that writes never update: saving appears to work and reloading shows stale data.
The reverse order fails the same way.
Read and write cannot be split across a release boundary for the *same* metric.

The fix is to split by **which metrics move**, not by read-vs-write.

Open metrics - `DataRow::extra` / `SampleResult::extra`, the custom columns Phase 2 already parses - have **no wide equivalent at all**, so they cannot go stale by construction.
They move first, read and write together, in a phase that leaves the standard 13-column path completely untouched.
The standard metrics move later, read and write together, in a single coordinated cutover.

This also fixes a second flaw: 3b as written would have backfilled `measurements` with standard-metric rows that nothing maintained until 3d, so the migrated copy would rot for two phases.
**3b therefore authors and rehearses the migration but does not run it against live data.**
It executes for real at the 3d cutover, when the app is finally ready to maintain those rows.

Net effect on the end state: none. Net effect on risk: each phase is independently correct and shippable, and the midpoint smoke has a real user-visible win instead of a broken build.

### 3b - Schema + migration authored and rehearsed (no app behavior change at all)

`metric_defs` / `sample_headers` / `measurements` with `UNIQUE(sample_id, metric_id, sort_order)`, the compat views, registry seeding projected from `MetricRegistry`, and the one-shot migration **authored and proven on a rehearsal copy, NOT executed against live data**.
Tables are created (empty); nothing else about the running app changes.
Pre-flight production checks belong here - in particular `SELECT sample_id, sort_order, count(*) FROM data_rows GROUP BY 1,2 HAVING count(*) > 1`, because `data_rows` has no uniqueness on that pair today and the new constraint would change what is legal.

**Gate:** on a prod-shaped synthetic database, running the migration produces `data_rows_v` / `samples_v` row-for-row identical to the source tables with matching per-metric checksums; and the app, rebuilt against the new schema, behaves exactly as before (full suite green, corpus harnesses unmoved).

### 3c - Open metrics persist (read and write together)

`DataRow::extra` and `SampleResult::extra` are written to `measurements` / `sample_headers` inside the existing save transaction and read back on load, in Postgres and in the offline snapshot.
The standard 13 columns are **not touched** - they keep flowing through the wide path exactly as today, so the regression surface is near zero.
Snapshot schema version 3 to 4, with all five coupled sites updated atomically.
Closes H1's live instance: extras are no longer silently dropped by a load-then-save.

**Gate:** the manifest demo workbook's custom `coil_temp` column survives save, close, and reload-from-database - the first time a custom metric survives the DB. Plus the 16 save-integrity scenarios and the full suite.

### 3d - Standard-metric cutover (read and write together)

Run the real migration; `persistFileCore` writes the standard metrics as measurements; the orphan prune is reworked against the long shape; `data_rows` / `samples` are renamed aside and replaced by the compat views so every remaining reader keeps working; LiveSync commits per measurement through `dve_commit_measurement`; the offline `pending_edits` queue gains a schema version plus a drain-or-translate policy for entries captured before the cutover.

**Gate:** the 16 save-integrity E2E scenarios, the two-client E2E suite, and a new scenario proving a standard metric survives edit, save, reload, and re-save without being pruned.

---

## Hazard ledger (carried from the audit; each must be closed or consciously accepted)

| # | Hazard | Where | Closes in |
|---|---|---|---|
| H1 | Orphan prune deletes any metric the in-memory model fails to reproduce | `DatabaseOps.cpp:759-784` | 3d (prereq: `extra` persisted in 3c) |
| H2 | Child rows have no natural key; identity is a carried `qint64` only | `DatabaseOps.cpp:453/517/604/694` | 3d |
| H3 | Background persist writes a value copy; only file id+version come back | `PersistWorker.h:31`, `MainWindow.cpp:2861-2862` | 3d |
| H4 | `commitCell` rejects unknown columns and the caller ignores the return | `LiveSync.cpp:147-151`, `MainWindow.cpp:3496-3500` | 3a (surface it), 3d (re-point it) |
| H5 | Sample-header edits have exactly one persistence route (whole-file save) | `MainWindow.cpp:2145-2297` | 3d (needs a dedicated regression test) |
| H6 | Per-cell OCC is already disabled and `commitConflict` has no consumer | `MainWindow.cpp:257-270`, `LiveSync.cpp:75-79` | 3d (decide: wire it or delete it) |
| H7 | `saveFile()` copies the struct and discards the entire writeback | `DatabaseManager.cpp:921-927` | 3a |
| H8 | Transaction size: statement count is already O(children), long format multiplies by metrics | `DatabaseOps.cpp:182-786` | 3d (batch the child writes) |
| H9 | `pending_edits.sqlite` persists wide column names on user machines | `OfflineSnapshot.cpp:1996-2005` | 3d |
| H10 | Snapshot freshness fingerprint hardcodes 9 tables in 5 coupled sites | `OfflineSnapshot.cpp:377-385/422-443/55` | 3c |
| H11 | Dirty-set path-key inconsistency: recovery restore uses raw `==` | `MainWindow.cpp:5654` vs `:3274/:5110` | 3a |
| H12 | `ensureSchema` reconciles columns only; it cannot create tables | `DatabaseManager.cpp:129-205` | 3b |
| H13 | `DbRepair` re-derives child identity by name and skips data rows entirely | `DbRepair.cpp:61-86` | 3d |
| H14 | `assertColumnArity` is `Q_ASSERT_X` (debug-only) and is not called for the three migrating tables | `OfflineSnapshot.cpp:304-313/847/888/924` | 3a |
| H15 | Hardcoded table lists in 7 test wipe helpers, `MigrationTool`, and the deployment self-test | see audit | 3b/3c |
| H16 | `sensory_web` role needs an explicit REVOKE on any new table plus a negative test | `2026-06-25-dv11-sensory-web-role.sql` | 3b |
| H17 | `data_rows.sort_order` is nullable (`INTEGER DEFAULT 0`) but `measurements.sort_order` is NOT NULL and identity-bearing; `COALESCE(...,0)` would fabricate an ordering and could collide with a real row 0 | `init.sql:85-106` | 3b (second pre-flight abort) |
| H18 | An all-NULL `data_rows` row produces zero measurements under the sparse rule, so it vanishes from `data_rows_v` - the long format cannot express "this row exists but holds nothing" | sparse rule, D2 | 3b (pre-flight abort + row-count parity), re-check at D7 |
| H19 | `metric_defs` carries `bump_version` but is not in the no-op suppression list, so an unguarded `ON CONFLICT DO UPDATE` in `ensureSchema` would bump ~39 rows on every connect from every client | 3b Task 2 | 3b DONE (guarded with a 5-column `IS DISTINCT FROM` WHERE; verified 0 rows written across 100 of 102 connects) |
| **H20** | **`ensureSchema` delivers only the TABLES to production. The migration file also carries the two compat views, the `bump_version()` body update that adds the new tables to the no-op suppression list (D4), the 11 seed-only `metric_defs` rows, and `dve_migrate_to_long_format()` itself - none of which reach the NAS.** On production `kind='header'` is 28, not 39 and not 22, so the migration would `RAISE EXCEPTION` on the first missing key | 3b Task 2 | **3d - hard gate, see the delivery decision below** |

| H21 | H18's pre-flight covers `data_rows` only. A **sample** whose 22 header columns are all NULL produces zero `sample_headers` rows and silently vanishes from `samples_v` - no abort fires. The SQL calls this "milder" because readers `LEFT JOIN`, but that makes correctness depend on every future reader choosing LEFT, which nothing enforces | migration pre-flights | 3d - needs an explicit ruling before the cutover |

### Parity alone is NOT a sufficient migration gate (proven 2026-07-31)

The 3b harness was deliberately made to fail three ways. The important result: with the sparse guard removed from the migration - i.e. a **dense** migration, the exact D2 violation that would flip every old-template sheet into per-row-regime mode - `dataRowsView_matchesDataRows` **stayed GREEN**.
A dense migration pivots a NULL-valued measurement back to NULL, so a value-parity check cannot see it.
What caught it was the sparse-rule assertion and the `hasPerRowRegime` assertion (638 phantom regime rows, 15,951 measurements instead of 11,639).

**Consequence for 3d:** the production-dump rehearsal must run the sparse-rule and regime assertions, not just parity. A parity-only gate would sign off on a migration that silently re-labels historical sheets.

### D8 - NAS delivery of the long-format schema (decided 2026-07-31, from H20)

`ensureSchema` keeps creating the three tables (plus their `bump_version` triggers and the H16 revokes) so that 3c works against a NAS in any state - open metrics only need the tables and the registry-sourced `metric_defs` rows, both of which `ensureSchema` provides.

Everything else in the migration file - the views, the `bump_version` body update, the 11 seed-only rows, and the migration function - is delivered by **applying `deploy/postgres/migrations/2026-07-31-v3-long-format.sql` to the NAS by hand**, as a step in the supervised v3.0.0 migration runbook.
That is the right call rather than extending `ensureSchema`: duplicating a 25-line plpgsql function body in C++ is a live drift risk, the v3.0.0 deploy is already a supervised step requiring a backup and owner go-ahead (D7), and every NAS schema change since v2.0 has reached production either through `ensureSchema` or by hand anyway.
**3d must not run the migration until that manual step is confirmed applied**, and should probe for the function's existence and abort with a clear message if it is missing.

## Registry gaps found while authoring 3b (2026-07-31)

- **11 of the 22 `samples` value columns have no `HeaderFieldDef`**: `sample_name`, the 7 derived aggregates (`average_tpm`, `stddev_tpm`, `avg_power_density`, `efficiency_percent`, `total_oil_consumed`, `total_puffs`, `normalized_tpm`), and the 3 status strings (`burn_status`, `clog_status`, `leak_status`).
  They are seeded as `kind='header'` with `role='derived'` where appropriate rather than inventing a third `kind` - `sample_headers` is where the migration must land them either way, and `role` already carries the distinction.
  Consequence for tests: `kind='header'` is **22 on a migration-only database and 39 after `ensureSchema` upserts the compiled registry**. Assert against the right one.
- `did_burn` / `did_clog` / `did_leak` in the registry are the sheet's Y/N questions and are **NOT** the same fields as `burn_status` / `clog_status` / `leak_status`. Keep the keys separate.
- **Unit conflict needing owner ratification:** the spec derives `total_oil_consumed` from `oil_consumed` in grams (D4), but `SheetProcessors.cpp:307-310` computes `eff = totalOilConsumed / (initialOilMass * 1000)`, i.e. treats it as milligrams. Seeded as NULL unit rather than guessing. Same for `avg_power_density` and `normalized_tpm`, which are era-dependent per spec 2.1/9.1.

## Accepted implementation deviations from the written plan (2026-07-31)

- `samples_v`'s group key is named `id` (it genuinely is `samples.id`), leaving `sample_id` free for the TEXT business identifier that is one of the 22 header columns. The plan's "one row per sample_id" phrasing would have collided two different columns onto one name. Join as `samples s JOIN samples_v v ON v.id = s.id`.
- The plan's separate single-column indexes on `measurements(sample_id)` and `sample_headers(sample_id)` are omitted: the UNIQUE btrees lead with `sample_id`, so Postgres serves those lookups from the leftmost prefix. A redundant index would only add write cost to what will be the largest table in the database.
- The migration uses one set-based `INSERT` per column over a pinned column array via `format(%I)`, not a `to_jsonb` key-join: the sparse rule then appears literally as `WHERE d.<col> IS NOT NULL`, float8 values never round-trip through jsonb, and a future registry key colliding with `id`/`sort_order` cannot silently start capturing a non-value column.

## Dead code the audit found; resolve rather than port

- The entire `cell_focus` loop is inert: the outbound timer is constructed and connected (`MainWindow.cpp:515-527`) but never started, and the inbound `cellFocused` / `cellBlurred` signals have no consumers.
- `UniqueViolationDialog` is compiled but never instantiated; unique violations are resolved by silent auto-suffix retry loops.
- The `samples`, `tests`, and `files` entries in `LiveSync::isLiveSyncColumn` are unreachable - nothing calls `commitCell` for those tables.

## Assets

`MetricRegistry` already uses exactly the snake_case keys that are the current column names (`before_weight`, `draw_pressure`, `tpm_power_density`, `heating_technology`, ...).
`metric_defs` is a persisted projection of the compiled registry, not a new vocabulary, so seeding is mechanical and the migration's metric mapping is an identity map for every standard column.
