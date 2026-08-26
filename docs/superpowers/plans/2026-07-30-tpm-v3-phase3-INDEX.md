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

**D10 - Oil is measured in GRAMS everywhere (owner ruling, 2026-08-03).**
"Let's have everything in grams for oil. Please make sure units are consistent throughout."
This resolves the unit conflict flagged under "Registry gaps" below, and it resolves it in favour of what almost every layer already believed.
The template's own column header is `OilCum(g)` and its formula `=(TPM*puffs)/1000` divides milligrams down to grams; `MetricRegistry` declares `oil_consumed` and `initial_oil_mass` as `"g"`; `samples.initial_oil_mass` has always been grams.
The single dissenter was `SheetProcessor::calculateMetrics`, which scaled the gram-valued weight difference **up** by 1000, after which five display sites divided by 1000 again to undo it.
The app therefore looked correct while `data_rows.oil_consumed` and `samples.total_oil_consumed` held milligrams in a column whose registry unit said grams - a latent 1000x error for anything that trusts `metric_defs.unit` (Phase 4 UI, Phase 5 template builder, exports).
The compute path now produces grams and the five compensating conversions are deleted.

Stored data is normalized by `deploy/postgres/migrations/2026-08-03-oil-units-to-grams.sql`, which **recomputes** from `before_weight` / `after_weight` rather than dividing by 1000.
That matters: between the code change and the migration, a client can save a row that is already grams, and a blind `/1000` would divide it a second time with no way to tell the two apart afterwards.
Recomputing is idempotent, immune to a half-converted column, and reproduces exactly what the app computes on its next load.
It is a **separate file from the long-format migration on purpose** - `dve_migrate_to_long_format()` is a pure shape transform whose gate asserts `data_rows_v` is value-identical to `data_rows`, and folding a unit change into it would destroy that referee.
Ordering in the v3.0.0 runbook: oil migration first, long-format migration second.
It must not be applied early: v2.10.x clients still compute milligrams and would write them straight back.

**D11 - Every `samples` value column has a registry definition; sample-wide and per-row failure fields stay distinct (owner ruling, 2026-08-03).**
"We should cover all the columns. Again, extra columns are not a problem. But there is a difference between burn? leak? clog? and burn. leak. clog. One is sample-wide one is for one set of puffs."
Two changes follow.
First, the 11 header keys that existed only in the migration seed (`sample_name`, the 7 derived aggregates, the 3 status columns) are now real `HeaderFieldDef`s, so `metric_defs` carries a registry-backed display name, type and unit for all 22 columns instead of 11 anonymous rows.
The seed and the registry were reconciled character for character in the same change, because `ensureSchema`'s upsert would otherwise rewrite those rows on every connect (the convergence test catches exactly this).
Second, per-row `burn` and `leak` metrics join the existing per-row `clog`, so a template with those data columns lands on curated keys instead of colliding with the sample-wide questions.
`did_burn` / `did_clog` / `did_leak` remain the sample-wide fields and keep their `burn_status` / `clog_status` / `leak_status` storage columns, so the migration's column-to-key mapping stays an identity map for all 22.

**Do not use punctuation as the discriminator.** The question mark is not what separates the two concepts - the *band* is.
Real inferred-layout sheets spell a per-row data column `Clog?` (see `tests/tst_v3inference`), so `Clog?` is legitimately an alias in both namespaces: `metricByAlias()` resolves data-column headers, `headerByLabel()` resolves header-band labels, and nothing looks up both.

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

### 3d - Standard-metric cutover (read and write together) - DONE 2026-08-26

Run the real migration; `persistFileCore` writes the standard metrics as measurements; the orphan prune is reworked against the long shape; `data_rows` / `samples` are renamed aside and replaced by the compat views so every remaining reader keeps working; LiveSync commits per measurement through `dve_commit_measurement`; the offline `pending_edits` queue gains a schema version plus a drain-or-translate policy for entries captured before the cutover.

**Gate:** the 16 save-integrity E2E scenarios, the two-client E2E suite, and a new scenario proving a standard metric survives edit, save, reload, and re-save without being pruned.

**DONE (plan `2026-08-26-tpm-v3-phase3d-standard-cutover-plan.md`, commits `1087ae6`..the 3d wrap).**
Gates achieved: full suite green after the tst_migrationtool re-scope, the whole DB surface green in SEQUENCE on a fresh container (28/33/93/20/12/15/35/9 = 255 passed 0 failed), scenario25 = the edit/row-delete/no-churn-re-save gate, corpus shadow + round-trip harnesses re-run at their baselines.
Decisions D-3d-1..8 below; execution deviations and the production defects the flip smoked out are in the "3d execution findings" section.

#### D-3d decisions (locked at plan time, 2026-08-26)

- **D-3d-1 (cutover choreography):** one migration pair - `2026-08-26-v3-cutover-1-functions.sql` (H23 bump_version fix + `dve_commit_measurement` + `dve_cutover_to_long_format`; definitions only, safe anywhere) and `2026-08-26-v3-cutover-2-execute.sql` (the one-statement invocation; applying it IS the cutover). The cutover function idempotently re-runs the data migration, verifies row-count parity AND the per-column sparse rule (the 3b lesson institutionalized server-side), renames `data_rows -> data_rows_pre_v3` / `samples -> samples_core`, creates the name-holder views, revokes, stamps `schema_meta`. The renamed wide tables stay as the on-database pre-image and the v1 pending-edits translation source - NOT a rollback path (rollback = the D7 backup).
- **D-3d-2 (who runs it):** the app NEVER runs the migration or the cutover; both are supervised runbook steps (`2026-08-26-v3-migration-runbook.md`). This supersedes the 3b banner's "caller is 3d wrapping it in the advisory-lock gate" - a client must not run a full-table restructure over office wifi. The app's side is `DatabaseManager::verifyCutoverSchema()`: open()/reopen() refuse a pre-cutover database, naming the runbook, and fall into the offline path.
- **D-3d-3 (write path):** standard metrics ride the 3c batched keyed upserts, `(sample_id, metric_id/field_id, sort_order = row ordinal)`; `samples_core` keeps the narrow id+version OCC; `DataRow::id/version` are neither read nor written (H2 closed; H13 dissolves - keyed upserts need no id anchors); standard-key resolution is LOOKUP-ONLY (`MetricDefCache::lookup`), aborting the save on a broken seed. Sparse at write = "a measurement exists exactly where the old wide bind was non-NULL".
- **D-3d-4 (prune):** the phase-A `md.key NOT IN (<wide>)` exclusion lifted together with the writer that reproduces every standard metric (the 3c H1 banner's stated condition); `data_rows` left the prune list (a row IS its measurements); `samples` prune targets `samples_core`. `appendExtra`'s wide-column skip SURVIVES as the double-write guard for stale pre-3d recovery-JSON extras.
- **D-3d-5 (H25):** both loadFile extras SELECTs exclude keys with a same-named wide column via a `to_regclass`/pg_attribute subquery - same SQL correct in both worlds; the OFFLINE reader got the SQLite twin via `pragma_table_info` (found live by the 3c e2e gate, see findings).
- **D-3d-6 (per-cell, H4):** `dve_commit_measurement(sample_id, key, sort_order, value, uuid)` - ONE signature, no OCC arm (LWW by design, per-measurement no-op suppression on top); `LiveSync::commitMeasurement` + worker slot with the same reconnect-retry idiom; `data_rows`/`samples`/`tests`/`files` left the commitCell allowlist (zero callers). Unknown key = permanent, logged, never enqueued.
- **D-3d-7 (H23):** `bump_version()` drops `updated_by` from the no-op comparison AND restores it in the suppressed branch (byte-identical no-ops). Applies to the sensory pair too - identical-content cross-user saves stop bumping, which is the DV-23-desired behavior.
- **D-3d-8 (test architecture):** `start-test-postgres.ps1` provisions TWO databases with the same splitter machinery: `dve_test` post-cutover (every app-facing suite) and `dve_test_precut` pre-cutover minus only the execute file (the migration rehearsal + `tst_migrationtool`). The un-cut reset lives in `tests/common/PrecutReset.h`, shared by every suite that needs the pre-cutover shape, so suite order cannot poison it. `tests/common/SeedRows.h` seeds fixtures in whichever shape the connected database speaks.

#### 3d execution findings (2026-08-26)

- **Two production defects only the flip could smoke out, both caught by existing tests the moment the container went post-cutover:**
  1. `ensureSchema`'s long-table rebuild DDL said `REFERENCES samples(id)` - an FK cannot target the post-cutover VIEW, so a post-cutover NAS could never self-heal a dropped `measurements` table. The referent now resolves live (`samples_core` when it exists as a table). The DDL-parity digest still matches because a migration-built FK follows the rename by OID and prints `samples_core` too.
  2. **Qt 6 `QVariant::isNull` no longer sees a CONTAINED null QString**, so the sparse-skip never fired for untouched text members and every one minted a measurement row with BOTH value columns NULL - a D2 violation that was also latent for 3c extras. `isAbsentValue()` in persistFileCore now guards both appenders. Found by scenario23's both-null probe plus a statement-log capture showing 13 tuples where 10 belonged.
- The offline snapshot reader's missing H25 filter leaked all 20 standard headers into `SampleResult::extra` on the first post-cutover offline read - caught by the 3c manifest-demo offline gate, fixed via `pragma_table_info`.
- `MigrationTool` (the pre-v2.0 SQLite importer, CLI-only) now refuses a post-cutover target with the honest reason; its suite re-homed to the precut database with a refusal slot against the main one. Resolves its H15 entry without porting a retired tool.
- Scenario/count inversions: "measurements holds only extras" assertions across 4 suites became non-standard-key counts (`md.key NOT IN (<wide columns>)`), since standard metrics legitimately live there now; out-of-band "delete/update a data row" fixtures became measurement-group mutations.
- H6 note: per-cell OCC stays DORMANT (VersionLookup never registered in production, LWW by design); the occGuard tests re-target sensory_sessions, the only per-cell dve_commit_cell surface left. Deleting the machinery is deferred to the Phase 4 scaffolding sweep with the rest of the dead code.
- Corpus growth findings (the DB Data corpus grew to 44 files during the pause; parse code byte-identical to main, so both are PRE-EXISTING): `Gembox HTHH.xlsx` - the LEGACY referee's naive Project-join emits " HTHH-1" when the Project: cell is empty while production trims to the raw cell value (documented skip in tst_v3shadow); `T58G 510 Standardized Test old.xlsx` - the inferred provenance on its free-form 'Test Plan' TRACKING sheet maps the custom `progress` column one row off (documented skip in tst_v3roundtrip's mappedDomainIdentity; **Phase 4 inference-hardening item**: tracking-sheet layouts with sparse/merged header rows need a provenance sanity pass).
- Hazard sweep: H1/H2/H4/H8/H9/H13/H20/H21/H22/H23/H25 CLOSED by this sub-phase; H3 verified needing no change (child ids are no longer save-path inputs); H5 covered by scenario25 through the header batch; H6 consciously dormant (above); H24 remains ongoing discipline; D7's prod-dump rehearsal remains owner-gated (runbook step 0/rehearsal note).

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

| H21 | H18's pre-flight covers `data_rows` only. A **sample** whose 22 header columns are all NULL produces zero `sample_headers` rows and silently vanishes from `samples_v` - no abort fires. The SQL calls this "milder" because readers `LEFT JOIN`, but that makes correctness depend on every future reader choosing LEFT, which nothing enforces | migration pre-flights | **investigated 2026-08-03, fix recommended below - awaiting owner go-ahead** |

### H21 investigation (2026-08-03) - recommended fix: make `samples_v` structurally total

**Reachability.** Low but nonzero, and with the same provenance as H18.
The `samples` DDL gives all 12 numeric header columns `DEFAULT 0.0` / `DEFAULT 0`, and the save path always binds, so the application cannot produce an all-NULL sample.
`MigrationTool` copies legacy rows verbatim and nothing enforces the invariant, which is exactly how H18's data-row case can arise too.

**The structural asymmetry that decides the fix.** H18 and H21 look like the same bug and are not.
`data_rows_v` **replaces** `data_rows` in 3d - it takes the name and the wide table goes away.
A data row *is* its values, so there is no surviving identity table to anchor an empty row to; the long format genuinely cannot express "this row exists but holds nothing", and a hard pre-flight abort is the only honest answer.
`samples` is different: the identity table **survives 3d by design**, keeping `test_id`, `sort_order` and its audit trio, which is why `samples_v` is documented as a JOIN partner rather than a replacement.
So the sample case has an anchor and the data-row case does not.
That is the crisp version of what the SQL comment was gesturing at with "milder" - and it is also why "milder" was the wrong conclusion to stop at: it turned a structural property into a convention every future reader has to remember.

**Recommended fix.** Drive the view from `samples` with a `LEFT JOIN` instead of from `sample_headers`:

```sql
CREATE OR REPLACE VIEW samples_v AS
SELECT s.id                                                      AS id,
       MAX(sh.value_text) FILTER (WHERE md.key = 'sample_name')  AS sample_name,
       ...                                                       -- the other 21 unchanged
FROM samples s
LEFT JOIN sample_headers sh ON sh.sample_id = s.id
LEFT JOIN metric_defs    md ON md.id = sh.field_id AND md.kind = 'header'
GROUP BY s.id;
```

Why this is the best available answer:

- **It removes the hazard instead of detecting it.** Exactly one view row per `samples` row, always, whatever the header count. Row-count parity stops being an assertion and becomes an identity.
- **It stays D2-compliant.** Nothing is materialized. A sample with no header rows reads as all-NULL, which is precisely what the wide row held.
- **The LEFT-ness moves inside the view**, enforced once, instead of being a convention every caller must honor. That is the specific complaint in the hazard text.
- **No fourth pre-flight abort is needed.** Once the view is total, an all-NULL sample is no longer lossy, so aborting an entire production migration over one junk legacy row would be strictly worse than migrating it faithfully.

**Safety under 3d's rename - verified live, not assumed.** Postgres binds views to base relations by OID, not by name.
Confirmed in the test container: after `ALTER TABLE _h21_base RENAME TO _h21_base_renamed`, `pg_get_viewdef` reports `FROM _h21_base_renamed`, and a *new* view can then be created under the old name on top of the old view without a dependency cycle.
So `samples_v` reading `FROM samples` survives 3d renaming `samples` aside, and does not block a later `samples` view.

**Keep the tripwire.** The harness assertion `count(samples_v) = count(samples)` becomes a tautology under this fix - which is the point. It goes red only if someone edits the view back to an INNER JOIN.

**Rejected alternative:** materializing a sentinel header row so every sample has at least one. It violates D2, and 3b already proved that value-parity gates are blind to dense materialization, so it would be a change no existing gate could police.

| **H22** | **`measurements` has no referential link to `data_rows`.** Its only lifecycle FK is `sample_id -> samples(id)`, so the orphan prune's `DELETE FROM data_rows` leaves the row's measurements behind. Delete row 5 of 10 and save: survivors renumber, the stale measurements re-bind to the wrong rows on next load, and `sort_order` 9 becomes a `data_rows_v` group with no `data_rows` partner | `.sql:69-80`, prune at `DatabaseOps.cpp:759-784` | **3c - hard gate, see D9** |
| **H25** | **The 3c extras READ has no `kind`/`key` filter** - correct now, because `measurements` holds exclusively extras. The moment 3d puts standard metrics there, this read starts pulling `tpm` / `before_weight` / ... into `DataRow::extra` **alongside** the wide columns still being read above: two live representations of one value, and whichever the next save writes last wins silently. Exact mirror of the prune's H1 guard and must be resolved in the same breath | banner at `DatabaseManager.cpp:1520` | 3d |
| H24 | **Cross-suite landmine:** `metric_defs` is deliberately never wiped, and `tst_databasemanager::v3LongFormat_ensureSchemaMatchesRegistry` asserts an EXACT row count against the compiled registry. Any auto-registered custom key that survives a suite turns a DIFFERENT suite red for an unrelated reason. `tst_saveintegrity_e2e::wipe()` now deletes exactly the keys its scenarios invent, via one `openMetricKeys()` helper. **Every future test that saves a custom column must do the same** - this will bite the 3c manifest-demo gate (`coil_temp`) and the snapshot task | shared vocabulary table | ongoing discipline |
| H23 | **CONFIRMED LIVE 2026-07-31** (a 3c scenario exercises it and only passes because `updated_by` is unchanged within one client). Deferred to 3d deliberately, not from doubt: the fix is one token (`- 'updated_by'`), but `bump_version` is a SHARED trigger governing 11 tables including the two sensory tables whose DV-23 storm fix this is, and 3c's entire design premise is a near-zero regression surface. Practical impact today is small - `measurements` is empty in production until 3d. D4's no-op suppression compares `to_jsonb(NEW) - 'version' - 'updated_at' - 'app_version'`, which still contains `updated_by`. The save path binds `updated_by` on every child UPDATE, so in 3d a whole-file save by user B over rows last written by user A differs on every otherwise-unchanged row and bumps all ~13k versions - the exact churn D4 exists to prevent, failing precisely in the multi-user case it exists for | `.sql:123-128` | 3d |

### D9 - `measurements` lifecycle vs `data_rows` (decided 2026-07-31, from H22)

The tempting fix - `measurements.data_row_id BIGINT REFERENCES data_rows(id) ON DELETE CASCADE` - **cannot survive 3d**, where `data_rows` is renamed aside and its name is taken by a view. A foreign key onto a view is impossible, so the FK would have to be dropped at exactly the moment the data matters most.

Therefore: **the prune owns it.** When 3c's save path removes a `data_rows` row it must also delete that row's measurements by `(sample_id, sort_order)`, in the same transaction, driven by the same surviving-row set the prune already computes.

A tripwire lands in the 3b harness now, before any code depends on it: assert that no measurement exists at a `(sample_id, sort_order)` absent from `data_rows`.
It passes today - nothing writes measurements yet - and turns red the moment 3c gets this wrong.

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
  They are seeded as `kind='header'` rather than inventing a third `kind` - `sample_headers` is where the migration must land them either way.
  Correction (review, 2026-07-31): the SQL seeds `role` as **NULL for all 22 headers**, not `role='derived'` for the aggregates as this section originally claimed. That is fine and self-consistent - `ensureSchema` also writes NULL role for every header field, so the two paths converge - but the distinction between a measured header and a derived aggregate is currently **not** recorded anywhere in the database. If Phase 4 or 5 needs it, that is a registry change, not a schema change.
  Consequence for tests: `kind='header'` is **22 on a migration-only database and 39 after `ensureSchema` upserts the compiled registry**. Assert against the right one.
- **CLOSED 2026-08-03 by D11.** All 11 are now `HeaderFieldDef`s, so nothing is seed-only any more and the expected header count is simply `allHeaderFields().size()`. `tst_databasemanager`'s `seedOnlyHeaderKeys()` became `alwaysPresentHeaderKeys()` - it no longer adds to the count, and survives as a tripwire against dropping a key that historical `sample_headers` rows resolve through.
- `did_burn` / `did_clog` / `did_leak` in the registry are the sheet's Y/N questions and are **NOT** the same fields as the per-row `burn` / `clog` / `leak` metrics. Keep the keys separate.
  `burn_status` / `clog_status` / `leak_status` are the *storage columns* for the `did_*` questions, not a third concept - `SheetProcessor` writes the same answer to both. See D11.
- **CLOSED 2026-08-03 by D10** (was: unit conflict needing owner ratification). `total_oil_consumed` is grams; the seed now says `'g'` and `SheetProcessors` no longer divides by an mg quantity.
  `avg_power_density` and `normalized_tpm` keep a NULL/empty unit deliberately - both are era-dependent per spec 2.1/9.1 and the oil ruling did not cover them. Seeding a guess there would recreate the exact failure mode D10 just corrected.

## Accepted implementation deviations from the written plan (2026-07-31)

- `samples_v`'s group key is named `id` (it genuinely is `samples.id`), leaving `sample_id` free for the TEXT business identifier that is one of the 22 header columns. The plan's "one row per sample_id" phrasing would have collided two different columns onto one name. Join as `samples s JOIN samples_v v ON v.id = s.id`.
- The plan's separate single-column indexes on `measurements(sample_id)` and `sample_headers(sample_id)` are omitted: the UNIQUE btrees lead with `sample_id`, so Postgres serves those lookups from the leftmost prefix. A redundant index would only add write cost to what will be the largest table in the database.
- The migration uses one set-based `INSERT` per column over a pinned column array via `format(%I)`, not a `to_jsonb` key-join: the sparse rule then appears literally as `WHERE d.<col> IS NOT NULL`, float8 values never round-trip through jsonb, and a future registry key colliding with `id`/`sort_order` cannot silently start capturing a non-value column.

## Open finding from D11: the sample-wide failure answers are never actually read

Registering the vocabulary (D11) exposed that the standard path does not read the cells it names.

The template puts the three sample-wide questions in the **header band** - `Burn?` at row 1, `Clog?` at row 2, `Leak?` at row 3, label at block col +10 and answer at +11 (`docs/superpowers/specs/template-cell-map.md`).
Nothing reads those answer cells.
`ExcelReader::SampleMetadata` has no burn/clog/leak member and the extractor never touches col +10/+11; `StandardSchema` declares no header cell for `did_burn` / `did_clog` / `did_leak`.

What populates `burnStatus` / `clogStatus` / `leakStatus` instead is a **keyword scan of the Notes column** - `LegacyAdapter.cpp:258-269`, whose own comment concedes that the other strategy's "raw cols 8-10 Y/N flags don't exist on these layouts".
So a note reading "burned out around puff 300" yields `burnStatus = "Y"`, a note reading "no burn" yields `"N"`, and the tester's actual answer in the header band is ignored entirely.
`SheetProcessors.cpp:180-201` has the same shape, scanning data-band columns 8-10, which on the current template are the TPM / TPM-PD / Var% formula columns.

**Why this was not fixed in the same change.** Reading those cells changes parse output, and `SampleResult::extra` is serialized unconditionally (`ReportDataJson.cpp:247`), so a newly-read `did_burn` would appear as a diff in `tst_v3shadow` - the byte-identity gate that has held since Phase 1.
This is the same situation as the old-era col 10/11/12 re-keying that 2a deliberately left dormant: register the vocabulary now, activate at the coordinated Phase 3/4 flip.

**Phase 4 activation item.** Give `did_burn` / `did_clog` / `did_leak` their header cells in `StandardSchema`, retire the notes keyword scan, and decide what a disagreement between the header answer and the notes text should mean.
The keyword scan should probably survive as a fallback for legacy files that genuinely have no header answer - but it should stop being the primary source.

## Dead code the audit found; resolve rather than port

- The entire `cell_focus` loop is inert: the outbound timer is constructed and connected (`MainWindow.cpp:515-527`) but never started, and the inbound `cellFocused` / `cellBlurred` signals have no consumers.
- `UniqueViolationDialog` is compiled but never instantiated; unique violations are resolved by silent auto-suffix retry loops.
- The `samples`, `tests`, and `files` entries in `LiveSync::isLiveSyncColumn` are unreachable - nothing calls `commitCell` for those tables.

## Assets

`MetricRegistry` already uses exactly the snake_case keys that are the current column names (`before_weight`, `draw_pressure`, `tpm_power_density`, `heating_technology`, ...).
`metric_defs` is a persisted projection of the compiled registry, not a new vocabulary, so seeding is mechanical and the migration's metric mapping is an identity map for every standard column.
