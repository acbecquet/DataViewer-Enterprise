# Plan B — DB-Load Render Fix — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Loading a TPM file from the database renders identically to loading the source `.xlsx`, for every user, without needing the file present — and a `--repair-db` backfill makes colleagues on the *current (broken)* app version render correctly immediately.

**Architecture:** Add a `raw_grid` JSONB column to `tests` (additive, old-binary-safe) so SOP/raw sheets persist; write/read it in `DatabaseManager` + `OfflineSnapshot` via a shared JSON helper. Backfill missing `data_rows` (the v2.0 migration never wrote them) via a headless `DataViewer.exe --repair-db` that re-parses each file's `.xlsx`. Add per-load self-heal + an incomplete-data banner so blanks stop being silent.

**Tech Stack:** C++17 / Qt 6.10 (QPSQL/libpq, `QJsonDocument`), Postgres 16 (JSONB), namespace `DVE`, qmake + MinGW, `-Werror -Wall -Wextra`. DB tests use the ephemeral test Postgres.

**Design spec:** `docs/superpowers/specs/2026-06-04-plan-b-db-load-fix-design.md`

---

## Conventions for EVERY build/test step (read once)

- **MIP decrypt before any C++ build/test:** `python tools/decrypt_via_copy.py --apply` from the repo root (idempotent). Re-run if a tool shows `%TSD-Header-###%` ciphertext.
- **Create new source files via Python delete-and-rewrite** (avoid MIP labels). Edit existing files with the Edit tool after decrypting.
- **Commit immediately after writing a new file**; verify plaintext with `git show HEAD:<path> | head -3`.
- **DB-dependent tests** need the throwaway Postgres: `powershell -ExecutionPolicy Bypass -File tests/start-test-postgres.ps1` (sets `$env:DVE_TEST_PG_CONN`, prepends `vendor\libpq-16` to PATH). Tear down with `docker rm -f dve-test-pg`. Use FORWARD slashes for `-File tests/...ps1` (bash eats backslashes).
- **Run the unit suite:** `powershell -ExecutionPolicy Bypass -File tests/run-tests.ps1` (auto-detects `DVE_TEST_PG_CONN`; DB suites skip cleanly without it).
- **Compile the app** (incremental release — fast, strictest warnings): `cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake -spec win32-g++ DataViewerEnterprise.pro CONFIG+=release && mingw32-make -j8"`.
- Work on branch `feature/db-load-fix`. No VERSION bump until the Release section.
- The two investigation probes (`testNormalSheetRoundTrip_nonZeroWeights`, `testRawSheetRoundTrip_rawDataNotPersisted`) already exist in `tests/tst_databasemanager` — Task 5 flips the raw-sheet one from "asserts empty" to "asserts survives".

---

## Task 1: Pin (and if needed, fix) the save-completeness gap

**Goal:** Determine whether the *current* app ever persists a normal TPM file to the DB **without** its `data_rows` (active bug), or whether empty-`data_rows` files are purely v2.0-migration legacy. The repair + auto-heal rely on `saveFile`-on-load being complete, so close any active gap.

**Files:** Investigate `src/MainWindow.cpp` (`onFileLoadFinished`, `onUpdateDatabase`, `m_dbSaveTimer` setup, Ctrl+U handler, every `m_db->saveFile`/`tryWriteFile` call) and `src/database/DatabaseManager.cpp` (`tryWriteFile`, any sample/aggregate write path). Possibly `MigrationTool.cpp`.

- [ ] **Step 1: Trace every DB write trigger** and record, for each: what FileResult is passed and whether `samples[].rows` is populated at that moment. Identify any path that writes `tests`/`samples` (aggregates) without entering the `INSERT INTO data_rows` loop, or any condition/early-return that skips rows.
- [ ] **Step 2: Decide active vs. legacy.** If a current path can persist a normal file with empty `data_rows`, it's ACTIVE → go to Step 3. If rows are always written when present (and the empty-rows files can only come from the v2.0 migration), it's LEGACY → document in the task report and skip to Task 2.
- [ ] **Step 3 (only if ACTIVE): write a failing test** in `tests/tst_databasemanager` reproducing the gap (e.g. the specific save sequence that lands a normal sheet with empty `data_rows`), against the test DB. Run → confirm it fails (rows missing).
- [ ] **Step 4 (only if ACTIVE): fix** the save path so the gap closes (e.g. back-fill child ids after the first `saveFile`, or force a complete write on the offending trigger) — the minimal change that makes the test pass without disturbing the OCC/child-id design (that broader churn is a Plan-B non-goal).
- [ ] **Step 5: run the suite** (`run-tests.ps1` with the test DB) → green.
- [ ] **Step 6: commit** (only if a fix landed):
```bash
git add src/MainWindow.cpp src/database/DatabaseManager.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "fix(db): persist complete data_rows on <trigger> (close save-completeness gap)"
```
Report the active-vs-legacy determination either way.

---

## Task 2: Shared raw-grid JSON serializer (TDD)

**Files:**
- Create: `src/database/RawGridJson.h`, `src/database/RawGridJson.cpp`
- Create: `tests/tst_rawgridjson/tst_rawgridjson.pro`, `tests/tst_rawgridjson/tst_rawgridjson.cpp`
- Modify: `DataViewerEnterprise.pro`, `tests/tests.pro`

- [ ] **Step 1: Write the failing test** — `tests/tst_rawgridjson/tst_rawgridjson.cpp`:
```cpp
#include <QtTest>
#include "database/RawGridJson.h"
using namespace DVE;

class TestRawGridJson : public QObject {
    Q_OBJECT
private slots:
    void roundTrip_headersAndRows();
    void empty_serializesToEmptyString();
    void emptyString_parsesToEmpty();
    void specialChars_survive();
    void ragged_rows_survive();
};

void TestRawGridJson::roundTrip_headersAndRows() {
    const QStringList h{"A","B","C"};
    const QVector<QStringList> r{{"1","2","3"},{"x","y","z"}};
    const QString j = rawGridToJson(h, r);
    QVERIFY(!j.isEmpty());
    QStringList h2; QVector<QStringList> r2;
    rawGridFromJson(j, h2, r2);
    QCOMPARE(h2, h);
    QCOMPARE(r2, r);
}
void TestRawGridJson::empty_serializesToEmptyString() {
    QCOMPARE(rawGridToJson({}, {}), QString());
}
void TestRawGridJson::emptyString_parsesToEmpty() {
    QStringList h{"stale"}; QVector<QStringList> r{{"stale"}};
    rawGridFromJson(QString(), h, r);
    QVERIFY(h.isEmpty());
    QVERIFY(r.isEmpty());
}
void TestRawGridJson::specialChars_survive() {
    const QStringList h{"He said \"hi\""};
    const QVector<QStringList> r{{"a,b","ünïcode","line\nbreak"}};
    QStringList h2; QVector<QStringList> r2;
    rawGridFromJson(rawGridToJson(h, r), h2, r2);
    QCOMPARE(h2, h);
    QCOMPARE(r2, r);
}
void TestRawGridJson::ragged_rows_survive() {
    const QStringList h{"A","B"};
    const QVector<QStringList> r{{"1"},{"1","2","3"}};
    QStringList h2; QVector<QStringList> r2;
    rawGridFromJson(rawGridToJson(h, r), h2, r2);
    QCOMPARE(r2, r);
}
QTEST_MAIN(TestRawGridJson)
#include "tst_rawgridjson.moc"
```

- [ ] **Step 2: Create the test `.pro`** — `tests/tst_rawgridjson/tst_rawgridjson.pro` (mirror a sibling like `tst_outputpaths`):
```pro
QT += core testlib
QT -= gui
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_rawgridjson
INCLUDEPATH += ../../src ../../src/database
SOURCES += tst_rawgridjson.cpp \
           ../../src/database/RawGridJson.cpp
HEADERS += ../../src/database/RawGridJson.h
```

- [ ] **Step 3: Register** `tst_rawgridjson` in `tests/tests.pro` SUBDIRS.

- [ ] **Step 4: Run → FAILS to compile** (no `RawGridJson`). Decrypt first, then `run-tests.ps1`.

- [ ] **Step 5: Create the header** — `src/database/RawGridJson.h`:
```cpp
#pragma once
#include <QString>
#include <QStringList>
#include <QVector>
namespace DVE {
// {"headers":["h1",...],"rows":[["a","b",...],...]}.  Empty headers AND rows -> "".
QString rawGridToJson(const QStringList& headers, const QVector<QStringList>& rows);
// Parse back; empty/invalid input clears both outputs.
void rawGridFromJson(const QString& json, QStringList& headers, QVector<QStringList>& rows);
} // namespace DVE
```

- [ ] **Step 6: Create the implementation** — `src/database/RawGridJson.cpp`:
```cpp
#include "database/RawGridJson.h"
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
namespace DVE {
QString rawGridToJson(const QStringList& headers, const QVector<QStringList>& rows)
{
    if (headers.isEmpty() && rows.isEmpty())
        return QString();
    QJsonArray hdr;
    for (const QString& h : headers) hdr.append(h);
    QJsonArray rws;
    for (const QStringList& r : rows) {
        QJsonArray cells;
        for (const QString& c : r) cells.append(c);
        rws.append(cells);
    }
    QJsonObject obj;
    obj.insert(QStringLiteral("headers"), hdr);
    obj.insert(QStringLiteral("rows"), rws);
    return QString::fromUtf8(QJsonDocument(obj).toJson(QJsonDocument::Compact));
}
void rawGridFromJson(const QString& json, QStringList& headers, QVector<QStringList>& rows)
{
    headers.clear();
    rows.clear();
    if (json.isEmpty()) return;
    const QJsonDocument doc = QJsonDocument::fromJson(json.toUtf8());
    if (!doc.isObject()) return;
    const QJsonObject obj = doc.object();
    const QJsonArray hdr = obj.value(QStringLiteral("headers")).toArray();
    for (const QJsonValue& h : hdr) headers.append(h.toString());
    const QJsonArray rws = obj.value(QStringLiteral("rows")).toArray();
    for (const QJsonValue& rv : rws) {
        QStringList cells;
        const QJsonArray r = rv.toArray();
        for (const QJsonValue& c : r) cells.append(c.toString());
        rows.append(cells);
    }
}
} // namespace DVE
```

- [ ] **Step 7: Register in app `.pro`** — add `src/database/RawGridJson.cpp` to SOURCES and `.h` to HEADERS in `DataViewerEnterprise.pro`.

- [ ] **Step 8: Run → PASSES.** `run-tests.ps1`.

- [ ] **Step 9: commit**
```bash
git add src/database/RawGridJson.h src/database/RawGridJson.cpp tests/tst_rawgridjson/ DataViewerEnterprise.pro tests/tests.pro
git commit -m "feat(db): RawGridJson serializer for SOP raw-table storage"
```

---

## Task 3: Schema migration — `raw_grid` column

**Files:**
- Create: `deploy/postgres/migrations/2026-06-04-v2.x-tests-raw-grid.sql`
- Modify: `deploy/postgres/init.sql`

- [ ] **Step 1: Create the migration** (Python delete-and-rewrite or Write) — `deploy/postgres/migrations/2026-06-04-v2.x-tests-raw-grid.sql`:
```sql
-- Plan B: persist SOP/raw-table grid so DB-only loads render it.
-- Additive + NULL-default: old binaries (which SELECT explicit columns) are unaffected.
ALTER TABLE tests ADD COLUMN IF NOT EXISTS raw_grid JSONB;
COMMENT ON COLUMN tests.raw_grid IS
  'SOP/raw-table sheet content {"headers":[...],"rows":[[...],...]}; NULL for normal sheets.';
```

- [ ] **Step 2: Add the column to `init.sql`** — in the `CREATE TABLE tests (...)` definition (read the file; add `raw_grid JSONB,` near `is_raw_table`), so fresh installs and the test DB have it. Match existing formatting.

- [ ] **Step 3: Verify the test DB applies it** — tear down + restart: `docker rm -f dve-test-pg; powershell -ExecutionPolicy Bypass -File tests/start-test-postgres.ps1`, then confirm the column exists (the test DB applies `init.sql`). A quick check: the Task 4 round-trip will fail to bind `raw_grid` if it's missing.

- [ ] **Step 4: commit**
```bash
git add deploy/postgres/migrations/2026-06-04-v2.x-tests-raw-grid.sql deploy/postgres/init.sql
git commit -m "feat(db): add tests.raw_grid JSONB column (migration + init.sql)"
```

---

## Task 4: Write + read `raw_grid` in DatabaseManager (TDD against the test DB)

**Files:** Modify `src/database/DatabaseManager.cpp` (+`.h` if needed), `tests/tst_databasemanager/tst_databasemanager.cpp`. `#include "database/RawGridJson.h"`.

- [ ] **Step 1: Flip the existing probe to expect survival.** In `tests/tst_databasemanager`, update `testRawSheetRoundTrip_rawDataNotPersisted` (rename to `testRawSheetRoundTrip_gridSurvives`) so it now asserts the reloaded sheet's `rawHeaders`/`rawRows` **equal the originals** (and `isRawTable` true). Run with the test DB → it FAILS (grid still dropped).
- [ ] **Step 2: Write path** — in `DatabaseManager::tryWriteFile`, the `tests` INSERT (~:378-386) and its UPDATE counterpart: add `raw_grid` to the column list and bind `sheet.isRawTable ? rawGridToJson(sheet.rawHeaders, sheet.rawRows) : QString()`. Bind an empty/none as SQL NULL (use `QVariant(QMetaType(QMetaType::QString))` or the codebase's existing null-bind idiom — read how other nullable text columns are bound here). Keep all other bindings unchanged.
- [ ] **Step 3: Read path** — in `DatabaseManager::loadFile`, the `tests` SELECT (~:871-884): add `raw_grid` to the SELECT; after building each `SheetResult`, `rawGridFromJson(q.value(<idx>).toString(), sheet.rawHeaders, sheet.rawRows)` (no-op when NULL/empty).
- [ ] **Step 4: Build + run** the test (decrypt; test DB) → the round-trip now PASSES (grid survives). Confirm the normal-sheet probe still passes.
- [ ] **Step 5: commit**
```bash
git add src/database/DatabaseManager.cpp src/database/DatabaseManager.h tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "feat(db): persist + reconstruct SOP raw_grid in DatabaseManager"
```

---

## Task 5: OfflineSnapshot `raw_grid` parity

**Files:** Modify `src/database/OfflineSnapshot.cpp`. `#include "database/RawGridJson.h"`.

- [ ] **Step 1: `regenerate`** (~:429-443): add `raw_grid` to the SELECT-from-Postgres and the INSERT-into-SQLite column lists for `tests` (store the JSON as TEXT; SQLite has no JSONB but text round-trips).
- [ ] **Step 2: `loadFile`** (~:966): add `raw_grid` to the `tests` SELECT and `rawGridFromJson(...)` into `sheet.rawHeaders`/`rawRows`.
- [ ] **Step 3: Test** — add a snapshot round-trip case (or extend the offline-snapshot suite): save a raw sheet to PG, `regenerate` the SQLite snapshot, `OfflineSnapshot::loadFile` → assert grid survives. Build + run (test DB).
- [ ] **Step 4: commit**
```bash
git add src/database/OfflineSnapshot.cpp tests/<offline-snapshot-test>.cpp
git commit -m "feat(db): OfflineSnapshot raw_grid parity (regenerate + loadFile)"
```

---

## Task 6: Incomplete-data detection flag

**Files:** Modify `src/pipeline/ReportData.h`, `src/database/DatabaseManager.cpp`, `src/database/OfflineSnapshot.cpp`, `tests/tst_databasemanager`.

- [ ] **Step 1: Add the flag** — in `SheetResult` (`src/pipeline/ReportData.h`): `bool dbDataIncomplete = false;` (with a comment: set by DB loaders when a sheet reloads without its per-row/raw content). It defaults false on the Excel path.
- [ ] **Step 2: Write the failing test** — in `tst_databasemanager`: seed a normal sheet whose `samples` exist (with `averageTPM>0`) but `data_rows` is empty (insert sample, skip rows); `loadFile` → assert `sheet.dbDataIncomplete == true`. Also assert a complete sheet → `false`. Run → fails.
- [ ] **Step 3: Set the flag** — in `DatabaseManager::loadFile` (and `OfflineSnapshot::loadFile`), after assembling each sheet: `sheet.dbDataIncomplete = (!sheet.isRawTable && <any sample has averageTPM>0 (or hasSamples) but rows empty>) || (sheet.isRawTable && sheet.rawHeaders.isEmpty());`.
- [ ] **Step 4: Build + run** → passes.
- [ ] **Step 5: commit**
```bash
git add src/pipeline/ReportData.h src/database/DatabaseManager.cpp src/database/OfflineSnapshot.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "feat(db): flag sheets that reload with incomplete DB data"
```

---

## Task 7: `--repair-db` backfill (CLI + module, TDD against the test DB)

**Files:**
- Create: `src/database/DbRepair.h`, `src/database/DbRepair.cpp`
- Modify: `src/main.cpp` (arg parsing), `DataViewerEnterprise.pro`, `tests/tst_databasemanager` (or a new `tst_dbrepair`).

- [ ] **Step 1: Design the module** — `DbRepair` exposes a headless entry, e.g.:
```cpp
namespace DVE {
struct RepairOptions { QString sourceDir; bool dryRun = false; QString reportPath; };
struct RepairSummary { int healed = 0; int alreadyComplete = 0; int skippedNoXlsx = 0; int failed = 0;
                       QStringList details; };
// Connects via the existing db.conf, enumerates files, locates each .xlsx
// (files.file_path if it exists, else recursive search of opts.sourceDir by file_name),
// re-parses via DataProcessor::processFile, and DatabaseManager::saveFile to backfill
// data_rows + raw_grid. Returns a summary; writes JSON to opts.reportPath.
RepairSummary runDbRepair(DatabaseManager& db, DataProcessor& proc, const RepairOptions& opts);
}
```
Read `DatabaseManager`'s public API for "list all files" (or add a small `QVector<FileResult> listFilesShallow()` that returns id + filePath + fileName without children) and how `saveFile` inherits ids.

- [ ] **Step 2: Write the failing integration test** (test DB): seed a file with samples but empty `data_rows` (the migration state) whose `files.file_path` points at a small fixture `.xlsx` you add under `tests/tst_databasemanager/data/`; run `runDbRepair` (dryRun=false); assert the summary marks it `healed` and a subsequent `loadFile` returns non-empty `rows` (+ raw_grid for an SOP fixture). Run → fails (no `DbRepair`).
- [ ] **Step 3: Implement `DbRepair`** to make the test pass: enumerate → locate `.xlsx` → `processFile` → inherit DB `id`/`version` → `saveFile` → classify. `dryRun` skips the `saveFile`. Write the JSON report.
- [ ] **Step 4: Wire `--repair-db` into `src/main.cpp`** — alongside `--self-test`: parse `--repair-db [--source-dir <dir>] [--dry-run] [--report <path>]`, construct `DatabaseManager`/`DataProcessor`, call `runDbRepair`, print the summary table, set exit code (0 unless a connection/write failure). No GUI is created on this path.
- [ ] **Step 5: Register** `DbRepair.{cpp,h}` in `DataViewerEnterprise.pro`. Build the app (release) → clean.
- [ ] **Step 6: Run the suite** (test DB) → green. Optionally smoke-test the CLI against the test DB: `release/DataViewer.exe --repair-db --dry-run --source-dir tests/tst_databasemanager/data` (with `DVE_TEST_PG_CONN`/db.conf pointed at the test DB) and eyeball the summary.
- [ ] **Step 7: commit**
```bash
git add src/database/DbRepair.h src/database/DbRepair.cpp src/main.cpp DataViewerEnterprise.pro tests/tst_databasemanager/
git commit -m "feat(db): --repair-db backfill of data_rows + raw_grid from source .xlsx"
```

---

## Task 8: Incomplete-data banner + per-load self-heal

**Files:** Modify `src/MainWindow.{h,cpp}`; reuse/extend the `src/widgets/OfflineBanner` pattern (or add a small `IncompleteDataBanner`). No unit test (UI) — gate is a clean compile + manual.

- [ ] **Step 1: Banner** — when a loaded FileResult has any sheet with `dbDataIncomplete`, show a non-blocking banner: *"This file's per-row data isn't in the database. Run Tools → Repair, or open the source .xlsx."* Wire it into the same place `OfflineBanner`/`RowDeletedBanner` are shown/hidden. Hide it for complete files.
- [ ] **Step 2: Self-heal confirmation** — verify `onLoadFromDatabase`'s `QFile::exists` re-parse branch now also persists `raw_grid` (it goes through `saveFile`, which Task 4 covers) and that it writes `data_rows` (round-trip-proven). If Task 1 found an active gap, ensure this path is on the fixed code. No behavior change beyond confirming completeness.
- [ ] **Step 3: Compile** (release) → clean `-Werror`.
- [ ] **Step 4: Manual check (deferred to eyeball-test):** load a known-incomplete file → banner appears; load a complete file → no banner. (Subagent reports DONE on clean compile + correct wiring; runtime check is the user's.)
- [ ] **Step 5: commit**
```bash
git add src/MainWindow.h src/MainWindow.cpp src/widgets/
git commit -m "feat(ui): banner when a DB-loaded file has incomplete per-row/raw data"
```

---

## Task 9: Full verification

- [ ] **Step 1:** Decrypt + clean `-Werror` release build (`mingw32-make clean && mingw32-make -j8`). Expected: clean, `release\DataViewer.exe` produced.
- [ ] **Step 2:** Start the test DB; run the full suite (`run-tests.ps1`) → all pass incl. `tst_rawgridjson`, the DatabaseManager raw-grid + detection + repair tests, and OfflineSnapshot parity. Tear down the container.
- [ ] **Step 3: Manual end-to-end** (the user, at eyeball-test): apply the migration to a test/NAS DB; load an SOP-bearing file from the DB on a machine without the `.xlsx` → SOP grid renders; run `DataViewer.exe --repair-db --dry-run` → sane summary; run it for real → affected files heal; an *old-version* client then renders normal sheets.

---

## Release (after all tasks pass + user approval)

1. Bump `VERSION` (→ `2.2.5`) and write `release_overview/release_overview_v_2_2_5.txt` (commit atomically for plaintext).
2. Build the installer via **rebuild-dataviewer** (clean rebuild after the VERSION bump → `build_installer.bat`); confirm `dist\DataViewer-setup.exe` = `2.2.5`.
3. **User eyeball-tests** the installed build; **user applies the migration `.sql` to the NAS Postgres** and runs `--repair-db`. Only after approval: merge `feature/db-load-fix` → `main`, push; **user does the Synology drop**.
4. Then start **Plan C** (auto-recovery) — see `memory/plan-c-auto-recovery.md`.
