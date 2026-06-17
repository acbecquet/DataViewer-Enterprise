# v2.4.2 Sub-plan 3 — Durability (MIP / crash / off-thread / clock) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax. Each task: implementer → spec review → quality review → (for R5/R6) an adversarial durability/threading lens.

**Goal:** Harden the three local durability stores against MIP encryption, make the offline-snapshot promotion crash-safe and clock-honest, stop the up-to-30 s UI freeze on Excel write-back, and stop silently swallowing offline-edit failures — so one crash-while-offline, one MIP-locked file, or one large workbook can't quietly destroy recovery/offline/sync.

**Architecture:** Tier-3 of the v2.4.2 batch (spec `docs/superpowers/specs/2026-06-11-v242-backcompat-resilience-design.md`, items R5, R6-remainder, R7, R7b). Builds on SP1 (transport + `app_version` columns) and SP2 (convergence). Wraps to **v2.4.4** with an installer for smoke-testing. Ground truth for current code is the durability-surface map (file:line refs below are from it). The MIP-read fallback reuses the proven ExcelReader bundled-python pattern (`ExcelReader.cpp:604-702`); off-thread Excel mirrors RecoveryManager's `QtConcurrent::run` + `QFutureWatcher` re-entrancy guard (`RecoveryManager.cpp:380-396`).

**Tech Stack:** C++17 / Qt 6.10 (QtSql, QtConcurrent, QProcess), SQLite (snapshot + pending_edits), PostgreSQL 16, bundled Python 3.11, Windows file APIs (`MoveFileExW`), qmake + MinGW.

**Key learnings carried from SP1/SP2 (apply to every task):**
- MIP: before edits/builds run `python tools/decrypt_via_copy.py --apply` (idempotent; ignore the known `tests/excel/test_atomic_delete.py` re-label). Create NEW files via Python round-trip + immediate `git add`; verify committed blobs `git show HEAD:<path> | head`. Commit trailer `Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>`.
- The run-PATH gotcha is real: a test exe needs Qt + MinGW + libpq DLLs **next to the exe** or Qt reports `can not load requested driver QPSQL` and `initTestCase` fails with a misleading connection error. Always `cp vendor/libpq-16/*.dll` into the build output dir and run with the full PATH.
- A green unit test can mask broken production wiring (SP2 keystone). For R5/R6, adversarially verify the real failure mode (crafted ciphertext file; an actually-blocked write), not just the happy path.
- Branch `feature/v2.4.0-bugfix-batch` (current HEAD `9c9e96d`); do NOT push/merge; never touch Synology. Build is `-Werror -Wall -Wextra -Wpedantic`.

---

## Test environment

Container `dve-test-pg` on :5433 is up (`SELECT now()` works). Env: `DVE_TEST_PG_CONN='host=127.0.0.1 port=5433 dbname=dve_test user=test password=test'`.

**Build & run ONE suite** (substitute `<suite>` = `tst_offlinesnapshot`, `tst_recoverymanager`, `tst_saveintegrity_e2e`):
```
cd "tests/<suite>" && cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe <suite>.pro && mingw32-make.exe -j8"
cp ../../vendor/libpq-16/*.dll ./release/ 2>/dev/null
DVE_TEST_PG_CONN='host=127.0.0.1 port=5433 dbname=dve_test user=test password=test' PATH="<repo>/vendor/libpq-16:/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH" ./release/<suite>.exe -o res.txt,txt ; grep -E "PASS|FAIL|Totals" res.txt ; rm -f res.txt
```
**Incremental app compile** (proves MainWindow/threading code builds clean; not a clean rebuild — that's Task 5):
```
cd build && cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro CONFIG+=release && mingw32-make.exe -j8 release"
```

---

## File Structure

- `src/database/OfflineSnapshot.{h,cpp}` — R7b (3× CREATE/SELECT/INSERT + count assertion), R7 (promotion via `MoveFileExW`, `synchronous=FULL`, `source_schema_version` validation in `openReadOnly`, server-clock stamp, max-staleness), R5 (MIP fallback for snapshot.sqlite + pending_edits.sqlite open; surface `enqueueCellEdit` failures).
- `src/utils/RecoveryManager.{h,cpp}` — R5 (MIP fallback in `readAll`; surface read failure instead of silent empty).
- `src/utils/MipFallback.{h,cpp}` *(new)* — shared helper: detect `%TSD-Header-###%`, decrypt-to-temp via bundled python, return a readable temp path (or loud failure). One source of truth reused by RecoveryManager + OfflineSnapshot.
- `src/database/LiveSync.{h,cpp}` — R5 (propagate/emit `enqueueCellEdit` failure: unsynced counter + distinct signal).
- `src/MainWindow.{h,cpp}` — R6 (off-thread Excel write-back + tiered timeout + close-path synchronous finish); R5 wiring (loud warning on undecodable store; unsynced-edit indicator).
- Tests: `tests/tst_offlinesnapshot/` (R7b round-trip + count assertion; R7 schema-version reject + clock; R5 MIP fallback + enqueue surface), `tests/tst_recoverymanager/` (R5 MIP fallback), a focused `tests/tst_mipfallback/` *(new, optional)* for the helper, plus build + smoke for R6.
- `DataViewerEnterprise.pro` — VERSION 2.4.3 → 2.4.4; add the new source files.

---

## Task 1 — R7b: snapshot `app_version` columns + column-count assertion

**Files:** `src/database/OfflineSnapshot.cpp` (CREATE/SELECT/INSERT for the 3 session-ish tables + a debug assertion), `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`.

Why first: smallest, lowest-risk, same file as R7, and it feeds SP4's offline era display. Bumps the snapshot schema version (consumed by Task 2's validation).

- [ ] **Step 1: Write the failing test** `snapshot_carriesAppVersionColumn()` in `tst_offlinesnapshot.cpp` (follow the suite's existing regenerate-from-PG fixture). Seed a `files` row (and one sensory + one detailed) in PG with a non-NULL `app_version` (e.g. `'DataViewer/2.4.1'`); regenerate the snapshot; open read-only; assert the snapshot row's `app_version` matches. RED because the column isn't copied yet (the read returns empty/zero column).
- [ ] **Step 2: Run → FAIL.** Build+run `tst_offlinesnapshot`.
- [ ] **Step 3: Add the column to the 3 SQLite `CREATE TABLE`s** — `files` (`OfflineSnapshot.cpp:39-51`), `sensory_sessions` (`:149-163`), `detailed_sensory_sessions` (`:185-197`): add `app_version TEXT`.
- [ ] **Step 4: Extend the 3 SELECT/INSERT pairs** to include `app_version` (append it last to keep diffs minimal), and bump each bind-loop bound by 1:
  - files SELECT (`:407-409`) + INSERT (`:417-420`) + loop `c < 11` → `c < 12` (`:422`).
  - sensory SELECT (`:573-577`) + INSERT (`:585-588`) + loop `c < 13` → `c < 14` (`:590`).
  - detailed SELECT (`:639-642`) + INSERT (`:650-653`) + loop `c < 11` → `c < 12` (`:655`).
- [ ] **Step 5: Add the count-assertion footgun guard.** Introduce a small helper used at each copy site that asserts the SELECT column count equals the INSERT placeholder count equals the loop bound, in debug builds:
```cpp
// OfflineSnapshot.cpp (file-local). Catches the hand-maintained SELECT/INSERT/loop drift.
static inline void assertColumnArity(QSqlQuery& sel, int insertPlaceholders, int loopBound, const char* table) {
    Q_ASSERT_X(sel.record().count() == insertPlaceholders && insertPlaceholders == loopBound,
               "OfflineSnapshot::regenerate", table);
    Q_UNUSED(sel); Q_UNUSED(insertPlaceholders); Q_UNUSED(loopBound); Q_UNUSED(table); // release: no-op
}
```
  Call it once per table right after the SELECT `exec` (before the bind loop) with the literal placeholder count + loop bound. (Q_ASSERT compiles out in release, so `-Werror` is satisfied via the Q_UNUSED line; verify no unused-variable warning.)
- [ ] **Step 6: Bump the snapshot schema version constant** from `"2"` to `"3"` at the `_snapshot_meta` write (`OfflineSnapshot.cpp:752-753`). (Task 2 wires the read-side validation; bumping now records that the schema shape changed.) If a named constant doesn't exist, introduce `static constexpr int kSnapshotSchemaVersion = 3;` and use it both at the write and (Task 2) the read.
- [ ] **Step 7: Run → PASS** (app_version round-trips; assertion green). Build+run `tst_offlinesnapshot`.
- [ ] **Step 8: Commit** (OfflineSnapshot.cpp + test).

---

## Task 2 — R7: crash-safe promotion + clock discipline + schema-version validation

**Files:** `src/database/OfflineSnapshot.cpp`, `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`.

- [ ] **Step 1: Write failing tests** in `tst_offlinesnapshot.cpp`:
  - `snapshot_rejectsStaleSchemaVersion()`: regenerate a snapshot, then directly `UPDATE _snapshot_meta SET value='2' WHERE key='source_schema_version'` in the SQLite file; call `openReadOnly()`; assert it returns **false** (or sets an "incompatible" flag) so the app treats it as no usable snapshot. RED: today `openReadOnly` ignores the field and opens anyway.
  - `snapshot_stampIsServerClock()`: regenerate against PG; read `snapshotTakenAt()`; assert it is within a few seconds of the PG `SELECT now()` captured around regeneration (NOT the client `QDateTime::currentDateTimeUtc()` — to make the test meaningful, the implementer may need a seam to compare against server time; at minimum assert the stamp is read from the regenerate query, see Step 4). RED: today it stamps client UTC.
- [ ] **Step 2: Run → FAIL.**
- [ ] **Step 3: Atomic promotion via `MoveFileExW`.** Replace the delete-then-`QFile::rename` block (`OfflineSnapshot.cpp:787-802`) — which has a crash-unsafe gap between `QFile::remove(prodPath)` and `QFile::rename` — with an atomic Windows replace (no delete-first):
```cpp
#include <windows.h>   // top of file, guarded by #ifdef Q_OS_WIN if cross-compile matters (this is Windows-only)
...
// keep the explicit WAL/SHM sidecar removal of the OLD prod file's siblings only AFTER success,
// or remove the tmp's siblings; the tmp was built with its own WAL checkpointed (see Step 5).
const std::wstring tmpW  = QDir::toNativeSeparators(tmpPath).toStdWString();
const std::wstring prodW = QDir::toNativeSeparators(prodPath).toStdWString();
if (!MoveFileExW(tmpW.c_str(), prodW.c_str(),
                 MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
    m_lastError = QStringLiteral("regenerate: atomic replace failed (err %1) for %2")
                      .arg(GetLastError()).arg(prodPath);
    cleanup();
    return false;
}
```
  `MOVEFILE_REPLACE_EXISTING` removes the delete-first gap; `MOVEFILE_WRITE_THROUGH` flushes the rename to disk so a crash can't leave a dangling directory entry. Keep the `QSqlDatabase::removeDatabase` (`:778`) BEFORE the move (releases the file handle). Drop the now-unneeded `QFile::remove(prodPath)` and `QFile::rename`.
- [ ] **Step 4: Durable tmp + server-clock stamp.** In `regenerate()`: change the tmp connection pragma (`:358-362`) `synchronous=NORMAL` → `synchronous=FULL` (so the tmp's content is on disk before the move); run a `PRAGMA wal_checkpoint(TRUNCATE)` + close before the move so there are no tmp WAL/SHM siblings to carry. For the stamp (`:738-761`): source the time from the **PG server** — capture `SELECT now()` (UTC) on the source connection during regeneration and bind that ISO string as `snapshot_taken_at` instead of `QDateTime::currentDateTimeUtc()`. (The source PG connection is in scope in `regenerate`.)
- [ ] **Step 5: Validate `source_schema_version` in `openReadOnly`.** After the `SELECT 1 FROM _snapshot_meta` check (`:842`), read `source_schema_version`; if it != `kSnapshotSchemaVersion` (3), set `m_lastError` ("snapshot schema vN, expected v3 — ignoring stale snapshot"), close, and return false. Add a max-staleness check: if `snapshotTakenAt()` is older than a threshold (e.g. 30 days — a `static constexpr` or config), log a loud warning (the offline banner can surface it) but still open (stale-but-usable is better than nothing). Don't reject on staleness alone — only on schema mismatch.
- [ ] **Step 6: Run → PASS** (stale-schema rejected; server-clock stamp). Build+run.
- [ ] **Step 7: Commit** (OfflineSnapshot.cpp + tests).

---

## Task 3 — R5: MIP-resilient durability stores + loud enqueue failures

**Files:** `src/utils/MipFallback.{h,cpp}` (new), `src/utils/RecoveryManager.cpp`, `src/database/OfflineSnapshot.cpp`, `src/database/LiveSync.{h,cpp}`, `src/MainWindow.cpp` (wiring), tests.

- [ ] **Step 1: Create the shared MIP-fallback helper** (new file via Python round-trip + `git add`). `src/utils/MipFallback.h`:
```cpp
#pragma once
#include <QString>
namespace DVE {
// Detect the MIP/AIP magic and, if present, decrypt-to-temp via bundled python
// (which is on the MIP allowlist, exactly like ExcelReader). Returns a path to a
// readable PLAINTEXT temp copy on success, or QString() on failure (caller warns loudly).
// `looksEncrypted(path)` peeks the first bytes for "%TSD-Header-###%".
bool   looksEncrypted(const QString& path);
QString decryptToTempViaPython(const QString& path, const QString& pythonExe, QString& errOut);
}
```
  `MipFallback.cpp`: `looksEncrypted` reads the first ~32 bytes and checks for `%TSD-Header-###%`. `decryptToTempViaPython` writes a tiny python script (`open(src,'rb').read()` → `open(dst,'wb').write(...)`) and runs bundled python (reuse `MainWindow::findPython` resolution logic, or pass the exe in) to copy the decrypted bytes to a `QTemporaryFile` path the caller then opens. Mirror `ExcelReader::runPythonReader` (`ExcelReader.cpp:604-702`) for the QProcess + timeout shape. **Add to `DataViewerEnterprise.pro` SOURCES/HEADERS.**
- [ ] **Step 2: Write failing tests** (`tests/tst_offlinesnapshot` + `tests/tst_recoverymanager`, or a focused `tst_mipfallback`): craft a file whose bytes start with `%TSD-Header-###%` followed by junk; assert `looksEncrypted()` is true and that the store's open path takes the fallback branch / surfaces a loud error rather than returning silent-empty. (True MIP decryption is only verifiable on the work machine; the unit test pins the detection + fallback-branch wiring with a crafted marker. State this limitation in the test comment.)
- [ ] **Step 3: RecoveryManager fallback.** In `readAll()` (`RecoveryManager.cpp:418-458`): when `index.json` open fails OR `QJsonDocument::fromJson` errors, check `looksEncrypted()`; if so, `decryptToTempViaPython` and re-read the temp; if that ALSO fails (or python unavailable), set an error flag + log loudly (do NOT return silent-empty). Same for each blob. Add a `bool m_lastReadFailed`/`QString m_lastError` so `hasRecoverable()` can distinguish "nothing to recover" from "couldn't read what's there" — the latter must reach the user (MainWindow surfaces a warning at the `maybeOfferRecovery()` site, `MainWindow.cpp:5114`).
- [ ] **Step 4: OfflineSnapshot fallback.** In `openReadOnly()` (`:812-854`) and `ensureQueueOpen()` (`:1351-1413`): if `m_db.open()`/`m_queueDb.open()` fails AND `looksEncrypted(path)`, `decryptToTempViaPython` to a temp `.sqlite` and open THAT (read-only for the snapshot). If even that fails, set `m_lastError` + return false (caller now surfaces it — see Step 6). Make the two production `openReadOnly()` callers (`MainWindow.cpp:146`, `:5667`) check the return and, on a decode failure (not mere absence), show a loud one-time warning.
- [ ] **Step 5: Stop swallowing `enqueueCellEdit` failures.** `OfflineSnapshot::enqueueCellEdit` (`:1415-1434`) already returns bool; the swallow is at the call sites. In `LiveSync.cpp` (`:181`, `:228`, and propagate `:155`): when `enqueueCellEdit` returns false, increment an unsynced-failure counter and emit a NEW distinct signal (e.g. `LiveSync::offlineEnqueueFailed(int unsyncedCount)`); retain the edit in an in-memory fallback list so it isn't lost. Declare the signal + counter in `LiveSync.h`.
- [ ] **Step 6: Wire the loud indicators in MainWindow** — connect `LiveSync::offlineEnqueueFailed` to a status-bar/indicator update (reuse the existing sync-indicator path, `updateDbSyncIndicator`); show a one-time warning when a durability store is MIP-undecodable. Keep it non-blocking (never a modal that stops work).
- [ ] **Step 7: Run → PASS** (`tst_offlinesnapshot`, `tst_recoverymanager`); incremental app compile clean.
- [ ] **Step 8: Commit** (MipFallback + RecoveryManager + OfflineSnapshot + LiveSync + MainWindow + tests + .pro).

---

## Task 4 — R6: off-thread Excel write-back + tiered timeout (HIGHEST RISK — extra review)

**Files:** `src/MainWindow.{h,cpp}`. **This moves a freshly-stabilized synchronous save path off the UI thread — the riskiest change in the batch. It gets the strongest adversarial review (threading/regression lens) on top of spec+quality.**

Current blocking path (from the map): `writeCellsToExcel` (`:6480-6534`) / `deleteRowFromExcel` (`:6442-6472`) → `runPython` (`:6154-6200`, `QProcess::waitForFinished(30000)` on the UI thread). Debounce timer `flushExcelWrites` (`:6560-6587`) + 5 inline-flush close paths (`:450, 2260, 4758, 5036, 6009`). Must preserve `markFileModified`/`m_modifiedFilePaths`/`updateDbSyncIndicator` semantics + the rate-limited failure warning (`m_excelWriteFailureShown`).

- [ ] **Step 1: Design note (no test yet) — the seam.** Keep `runPython` synchronous (it's reused elsewhere). Add an off-thread wrapper for the Excel-write case only: `QtConcurrent::run` returning a small result struct `{bool ok; QString error;}` watched by a `QFutureWatcher` member, mirroring RecoveryManager's `m_flushInFlight`/`m_flushPending`/`m_flushFuture` re-entrancy guard (`RecoveryManager.cpp:380-396`, `.h:130-132`). The worker lambda calls the existing python-invocation logic; it must NOT touch any QWidget or shared mutable state — it takes a COPY of the cells + paths and returns the result.
- [ ] **Step 2: Write what's testable.** Threading is largely build+smoke-verified, but pin the decomposition: extract the pure "build the python script + args for these cells" into a small testable function (no QProcess), and unit-test that the script/args for a given `QVector<CellWrite>` match what the synchronous path produced (so the off-thread move can't silently change the payload). Add to `tst_saveintegrity_e2e` or a focused MainWindow-adjacent test if a seam exists; otherwise document that R6 is build + smoke verified and the payload-equivalence is the automated guard.
- [ ] **Step 3: Implement the async flush.** `flushExcelWrites()` (`:6560-6587`): instead of calling `writeCellsToExcel` synchronously, snapshot `m_pendingWrites`+paths, clear the pending set optimistically into an "in-flight" copy, and `QtConcurrent::run` the worker; on the `QFutureWatcher::finished` slot (UI thread) read the result — on success reset `m_excelWriteFailureShown`; on failure re-queue the in-flight cells into `m_pendingWrites` and show the rate-limited warning. Re-entrancy: if a flush is already in flight, set a "pending" flag and re-flush on completion (don't spawn concurrent python writes to the same file). All `m_modifiedFilePaths`/`updateDbSyncIndicator` mutations stay on the UI thread (in the finished slot or before dispatch).
- [ ] **Step 4: Tiered timeout.** The worker's `waitForFinished` uses an interactive budget (5–8 s) for debounced single-cell flushes and a batch budget (30 s) for bulk operations (add/remove row, multi-cell). Pass the budget into the worker; on timeout treat as failure (re-queue + warn).
- [ ] **Step 5: Close-path synchronous finish.** The 5 inline-flush sites + `~MainWindow` (`:450`) must NOT leave a write in flight on a background thread as the app exits. Add a `finishExcelWritesBlocking()` that, if a flush is in flight, waits it out (mirror RecoveryManager's `flushNow(true)` wait-out, `RecoveryManager.cpp:363-364`) and runs any still-pending writes synchronously. Call it from the close paths in place of the inline `flushExcelWrites()`.
- [ ] **Step 6: Verify.** Incremental app compile clean under `-Werror`. Run `tst_saveintegrity_e2e` (must stay green — the DB save path is unchanged). Manual smoke (user, at acceptance): edit a large workbook's cell → UI stays responsive; close mid-write → no corruption, no hang.
- [ ] **Step 7: Commit** (MainWindow.cpp/.h + any test).

---

## Task 5 — v2.4.4 bump + clean rebuild + verify + installer

**Files:** `DataViewerEnterprise.pro` (VERSION 2.4.3 → 2.4.4; ensure new SOURCES/HEADERS for MipFallback are listed), `release_overview/release_overview_v_2_4_4.txt` (new).

- [ ] **Step 1: Bump** VERSION → 2.4.4. Confirm `src/utils/MipFallback.{cpp,h}` are in the `.pro`.
- [ ] **Step 2: MIP decrypt + clean rebuild** (`make clean` required on VERSION bump). Must be clean under `-Werror`.
- [ ] **Step 3: Run all affected suites fresh** — `tst_offlinesnapshot`, `tst_recoverymanager`, `tst_saveintegrity_e2e`, plus `tst_livesync`/`tst_databasemanager` (regression). All green.
- [ ] **Step 4: Build the installer** via the `rebuild-dataviewer` flow (clean rebuild + `tools\prepare_python_embed.bat` + `build_installer.bat`); verify `release\DataViewer.exe` + `dist\DataViewer-setup.exe` both report 2.4.4. **Do NOT touch Synology.**
- [ ] **Step 5: Write + commit `release_overview_v_2_4_4.txt`** (customer-readable: crash-safe offline cache; recovers gracefully from locked/encrypted local files instead of losing data; no more "Not Responding" freeze when saving large workbooks; clearer warnings when something can't sync).
- [ ] **Step 6: Commit** bump + overview. Hand off: v2.4.4 installer for a real-connection + crash/large-file smoke test; capture learnings to feed SP4.

---

## Self-Review

**Spec coverage:** R5 (MIP fallback for all 3 stores + loud enqueue failures) → Task 3. R6 (off-thread Excel + tiered timeout) → Task 4. R7 (crash-safe promotion + clock + schema-version) → Task 2. R7b (snapshot app_version + count assertion) → Task 1. ✓

**SP1/SP2 learnings applied:** libpq-DLL-next-to-exe in every run recipe; MIP Python round-trip for new files; adversarial verification of the real failure mode for R5 (crafted ciphertext) and R6 (blocked write / mid-write close); the keystone lesson (test the production path, not a parallel reimplementation) → Task 4 Step 2 pins payload-equivalence.

**Risk callouts:**
- **Task 4 (R6) is the highest-risk** — it relocates the stabilized save path off-thread. It gets an extra adversarial threading/regression lens, keeps `runPython` synchronous for reuse + close paths, and guards re-entrancy. If risk appetite is low, R6 can be reduced to *tiered-timeout-only* (still synchronous) as a fallback — but that does not remove the freeze, only bounds it; flag this to the owner before executing Task 4.
- **R5 MIP decryption efficacy** depends on the bundled python being MIP-allowlisted (same assumption ExcelReader already relies on). The unit tests pin detection + fallback wiring; true decryption is verified only on the work machine.
- **Placeholders:** none — every step names exact files/lines (from the durability map) and gives the non-obvious code; mechanical edits reference the map's locations.

**Type/name consistency:** `kSnapshotSchemaVersion` (Task 1 write ↔ Task 2 read), `MipFallback::looksEncrypted/decryptToTempViaPython`, `LiveSync::offlineEnqueueFailed`, `finishExcelWritesBlocking` — used consistently across tasks.
