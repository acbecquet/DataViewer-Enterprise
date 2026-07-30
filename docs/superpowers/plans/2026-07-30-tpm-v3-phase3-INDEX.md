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

### 3b - Schema, migration, rehearsal (SQL + harness; no app behavior change)

`metric_defs` / `sample_headers` / `measurements` with `UNIQUE(sample_id, metric_id, sort_order)`, the two compat views, registry seeding projected from `MetricRegistry`, and the gated one-shot migration.
Pre-flight production checks belong here too - in particular `SELECT sample_id, sort_order, count(*) FROM data_rows GROUP BY 1,2 HAVING count(*) > 1`, because `data_rows` has no uniqueness on that pair today and the new constraint would change what is legal.

**Gate:** on a prod-shaped synthetic database, `data_rows_v` and `samples_v` are row-for-row identical to the pre-migration tables, per-metric checksums match, and the app still runs entirely unchanged against the migrated schema (it is still reading the wide tables, which are now views).

### 3c - Read path

`DatabaseManager` and `OfflineSnapshot` read from the long tables; `DataRow::extra` and `SampleResult::extra` round-trip for the first time; snapshot schema version 3 to 4 with its five coupled sites updated atomically.

**Gate:** dual-run diff - the same file loaded through the wide read and the long read on migrated data produces identical `FileResult`s - plus a new test that a custom metric survives a database round-trip.

### 3d - Write path

`persistFileCore` writes measurements and sample headers, the orphan prune is reworked against the long shape, LiveSync commits per measurement through `dve_commit_measurement`, and the offline `pending_edits` queue gains a schema version plus a drain-or-translate policy for entries captured before the cutover.

**Gate:** the 16 save-integrity E2E scenarios, the two-client E2E suite, and a new scenario proving a custom metric survives edit, save, reload, and re-save without being pruned.

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

## Dead code the audit found; resolve rather than port

- The entire `cell_focus` loop is inert: the outbound timer is constructed and connected (`MainWindow.cpp:515-527`) but never started, and the inbound `cellFocused` / `cellBlurred` signals have no consumers.
- `UniqueViolationDialog` is compiled but never instantiated; unique violations are resolved by silent auto-suffix retry loops.
- The `samples`, `tests`, and `files` entries in `LiveSync::isLiveSyncColumn` are unreachable - nothing calls `commitCell` for those tables.

## Assets

`MetricRegistry` already uses exactly the snake_case keys that are the current column names (`before_weight`, `draw_pressure`, `tpm_power_density`, `heating_technology`, ...).
`metric_defs` is a persisted projection of the compiled registry, not a new vocabulary, so seeding is mechanical and the migration's metric mapping is an identity map for every standard column.
