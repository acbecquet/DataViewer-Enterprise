# Plan B — DB-Load Render Fix — Design

**Date:** 2026-06-04
**Status:** Approved (design); ready for implementation planning
**Branch:** `feature/db-load-fix`
**Part of:** the v2.2.x critical-bug initiative. Sibling plans: **A** = save paths + Settings (SHIPPED v2.2.4); **C** = auto-recovery + crash snapshot (the biggest; see `memory/plan-c-auto-recovery.md`). Build one at a time; this is second.

---

## Goal

Make loading a TPM file from the database render **identically to loading it from the source `.xlsx`** — for every user, without needing the source file present — and repair the existing database so colleagues on the *current (broken) app version* see correct data immediately.

## Root cause (verified three ways: static trace, real-Postgres round-trip, save-path trace)

The author is unaffected because `MainWindow::onLoadFromDatabase` re-parses the source `.xlsx` when `QFile::exists(dbResult.filePath)` is true (only on the author's machine). Everyone else gets the DB-only reconstruction from `DatabaseManager::loadFile`, which is incomplete in two ways:

- **Part 1 — SOP/raw sheets (schema gap, ongoing):** The `tests` table has **no column** for the raw grid. `SheetResult::rawHeaders`/`rawRows` are never written and never read; `isRawTable` survives but the grid is empty → blank "Content" table. Proven by a save→reload round-trip against real Postgres (grid came back empty). Same gap in `OfflineSnapshot`.
- **Part 2 — normal sheets (missing data, not a code bug):** The `data_rows` write/read round-trip is **provably correct** (a normal sheet with non-zero weights round-trips perfectly through real Postgres). Affected files simply have **no `data_rows`** in the DB — they have `samples` rows (so `samples.average_tpm` renders the bar) but the per-row grid was never persisted, almost certainly because the v2.0 Postgres migration imported aggregates without per-row data. The display/plot filter `beforeWeight==0 || afterWeight==0` then drops every row, leaving only the averageTPM bar. The per-row data exists **only in the source `.xlsx`** and can be recovered only by re-parsing it.

User-confirmed: affected = **essentially all files** (systematic, consistent with a one-time migration), and **~95% of source `.xlsx` are accessible** (the rest the user can retrieve).

## Key insight — repairing the *data* fixes colleagues without an app update

The old `loadFile` reconstructs `data_rows` correctly **when they exist** (proven). So **backfilling `data_rows` into the DB makes the current/broken app render normal sheets correctly with no client update**. The new `raw_grid` column is additive and ignored by old binaries, so SOP/raw sheets are fixed only for updated clients (acceptable per the user; coworkers may not update in time, so the priority is the normal per-row data, which the old binary can already read).

## Tech Stack

C++17 / Qt 6.10 (QPSQL + libpq, `QJsonDocument` for JSONB, `QProcess`/bundled Python via the existing `ExcelReader` pipeline), Postgres 16 (JSONB), namespace `DVE`, qmake + MinGW, `-Werror -Wall -Wextra`. Migrations are hand-applied numbered `.sql` files (no auto-runner). Tests via `tests/tests.pro` + the ephemeral test Postgres (`tests/start-test-postgres.ps1`).

---

## Locked decisions (from brainstorming)

1. **Raw-grid storage = JSONB column `raw_grid` on `tests`** (not a child table). SOP grids are bounded; mirrors existing JSONB usage; additive + NULL-default so old binaries are unaffected.
2. **Repair = a `--repair-db` flag on `DataViewer.exe`** (headless; reuses the full C++ pipeline + bundled openpyxl) — a "run-it-once standalone" command, not a separate executable.
3. **Keep the `QFile::exists` re-parse branch** as belt-and-suspenders (redundant after repair, but harmless and serves as the per-load self-heal).

## Components

### 1. Schema migration — `deploy/postgres/migrations/2026-06-04-v2.x-tests-raw-grid.sql`
```sql
ALTER TABLE tests ADD COLUMN IF NOT EXISTS raw_grid JSONB;
COMMENT ON COLUMN tests.raw_grid IS
  'SOP/raw-table sheet content {"headers":[...],"rows":[[...],...]}; NULL for normal sheets.';
```
Also add the column to `deploy/postgres/init.sql`'s `tests` definition (for fresh installs). Additive, NULL-default — safe in the mixed-version window (old binaries SELECT explicit columns, never `raw_grid`; old INSERTs leave it NULL).

**Ops:** the user applies this `.sql` to the NAS Postgres manually (their DB-admin step, like the Synology drop). The plan hands them the file; we never touch the NAS.

### 2. Write path — persist the raw grid
- `DatabaseManager::tryWriteFile` (tests INSERT ~DatabaseManager.cpp:378-386 + the UPDATE counterpart): when `sheet.isRawTable`, serialize `{headers: rawHeaders, rows: rawRows}` to a JSON string and bind it into a new `raw_grid` column; else bind NULL.
- `OfflineSnapshot::regenerate` (tests copy ~OfflineSnapshot.cpp:429-443): add `raw_grid` to the SELECT-from-PG + INSERT-into-SQLite column lists (store the JSON text; SQLite has no JSONB but TEXT round-trips).
- One shared serializer helper (e.g. `rawGridToJson(headers, rows)` / `rawGridFromJson(str, headers&, rows&)`) so write/read/repair use one format.

### 3. Read path — reconstruct the raw grid
- `DatabaseManager::loadFile` (tests SELECT ~DatabaseManager.cpp:871-884): add `raw_grid` to the SELECT; deserialize into `sheet.rawHeaders`/`rawRows` when non-NULL.
- `OfflineSnapshot::loadFile` (tests SELECT ~OfflineSnapshot.cpp:966): same addition.

### 4. Standalone repair — `DataViewer.exe --repair-db [--source-dir <dir>] [--dry-run] [--report <path>]`
Headless (no GUI), parsed in `src/main.cpp` alongside the existing `--self-test`. Logic in a new `src/database/DbRepair.{h,cpp}`:
- Connect via the existing `db.conf` path.
- Enumerate every file row in the DB.
- For each, locate its `.xlsx`: the recorded `files.file_path` if it exists, else search `--source-dir` (recursively) by `file_name`.
- If found: `DataProcessor::processFile` (re-parse, populates `data_rows` + `rawHeaders/rawRows`), inherit the DB `file id/version`, then `DatabaseManager::saveFile` → backfills complete `data_rows` **and** `raw_grid`.
- Classify each file: **healed** / **already-complete** (skip) / **skipped-no-xlsx**.
- Print a summary table; write a JSON report (`--report`, default `%TEMP%\dataviewer_repair.json`). `--dry-run` reports what *would* be healed without writing. Exit code: 0 if no errors (even with skips), non-zero on connection/write failure.

This is the lever that un-breaks colleagues on the old version: once `data_rows` exist, their existing `loadFile` renders normal sheets.

### 5. Automatic on author load (forward self-heal)
- `onLoadFromDatabase`'s `QFile::exists` branch already re-parses + `saveFile`; with Component 2 it now also writes `raw_grid`. Verify it reliably persists `data_rows` (the round-trip confirms `saveFile` does). 
- Confirm `onFileLoadFinished` (open-from-`.xlsx`) likewise persists complete data. **Implementation Task 1 will pin whether any active save-trigger gap exists** (vs. the migration-legacy hypothesis) and close it if so — the goal is that any file the author opens heals its own DB row.

### 6. Detection / safeguard (no more silent blanks)
- Add `bool dbDataIncomplete = false;` to `SheetResult` (`src/pipeline/ReportData.h`).
- In `DatabaseManager::loadFile` (and `OfflineSnapshot::loadFile`), set it when a sheet reloads incomplete: `(!isRawTable && sampleHasAggregatesButNoRows) || (isRawTable && rawHeaders.isEmpty())`.
- `MainWindow` shows a non-blocking banner (reuse the `OfflineBanner`/`RowDeletedBanner` pattern) when a loaded file has any incomplete sheet: *"This file's per-row data isn't in the database. Run Tools → Repair, or open the source .xlsx."* Offer a Repair affordance when the `.xlsx` is present.

### 7. Tests (extend `tests/tst_databasemanager`, which already has the two probes added during investigation)
- **Raw-grid round-trip** (now PASSES after Components 2-3): save a sheet with `rawHeaders`/`rawRows` → reload → grid intact.
- **Repair backfill:** seed the test DB with a file whose samples exist but `data_rows` is empty (mimicking the migration state); run the repair logic against a test `.xlsx`; assert `data_rows` + `raw_grid` are now populated and a subsequent `loadFile` renders rows.
- **Detection flag:** load an incomplete file → assert `dbDataIncomplete` is set; a complete file → not set.
- **OfflineSnapshot parity:** raw_grid round-trips through `regenerate` + `loadFile`.

## Mixed-version compatibility (explicit)

| Scenario | Normal sheets | SOP/raw sheets |
|---|---|---|
| **Old binary** after DB repair | ✅ renders (reads `data_rows`) | ❌ still blank (can't read `raw_grid`) — until they update |
| **New binary** after DB repair | ✅ | ✅ |
| Old binary *saves* a file | writes `data_rows` (fine); leaves `raw_grid` NULL | a churned re-INSERT could drop a previously-set `raw_grid` during the transition — re-established by repair / next new-binary save |

The repair + the additive column are safe to apply before all clients update.

## Files

**Create:**
- `deploy/postgres/migrations/2026-06-04-v2.x-tests-raw-grid.sql`
- `src/database/DbRepair.h`, `src/database/DbRepair.cpp`

**Modify:**
- `deploy/postgres/init.sql` (add `raw_grid` to `tests`)
- `src/database/DatabaseManager.{h,cpp}` (write/read `raw_grid`; `dbDataIncomplete` detection; shared `rawGridToJson`/`FromJson`)
- `src/database/OfflineSnapshot.cpp` (regenerate + loadFile `raw_grid`; detection parity)
- `src/pipeline/ReportData.h` (`SheetResult::dbDataIncomplete`)
- `src/main.cpp` (`--repair-db` arg parsing → `DbRepair`)
- `src/MainWindow.{h,cpp}` (incomplete-data banner; verify/ensure load-save persists complete data; optional Tools "Repair" entry point)
- a banner widget (reuse/extend `src/widgets/OfflineBanner` pattern) or a small new one
- `DataViewerEnterprise.pro` (add `DbRepair.{cpp,h}`)
- `tests/tst_databasemanager/tst_databasemanager.cpp` (+ `.pro` if new deps)

## Non-goals (explicitly out of scope for Plan B)

- The broader optimistic-concurrency / child-id-backfill churn in `tryWriteFile` (that's the Postgres-multiuser initiative's territory — see its index; do not refactor it here beyond what's needed).
- Making the OLD binary render SOP/raw sheets (impossible without the new column read — they update).
- Auto-applying the migration to the NAS (the user's manual ops step).
- Plan C (auto-recovery) and any Plan A follow-ups.

## Open items to confirm during implementation
- **Active vs. legacy save gap:** Task 1 pins whether the current app *ever* saves a normal file without `data_rows` (vs. it being purely migration-legacy). If active, fix it so the backfill isn't re-broken.
- **`raw_grid` size:** confirm the largest real SOP sheets serialize to a reasonable JSONB size (well under the existing 100 MB / libpq guards). If any are huge, revisit child-table storage (currently ruled out).

## Build / version

No version bump for the spec/plan. The VERSION bump (→ next patch, e.g. `2.2.5`) + installer build happens at the end of implementation, followed by the user's installer eyeball-test, the user applying the migration `.sql` to the NAS, then merge `feature/db-load-fix` → `main` and the user's Synology drop.
