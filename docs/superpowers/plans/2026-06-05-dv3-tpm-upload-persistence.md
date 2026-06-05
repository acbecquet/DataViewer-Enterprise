# DATAVIEWER-3: Harden TPM load-time DB persistence (v2.3.2) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make freshly-loaded TPM files persist reliably to Postgres by checking the save result, keeping a failed save marked dirty for retry, and surfacing the failure to the user — instead of silently discarding a failed `saveFile()`.

**Architecture:** A pure, unit-tested helper `classifyLoadSaveResult(WriteResult) -> LoadSavePolicy` encodes the policy; a new `MainWindow::persistLoadedFile(int)` method uses it to drive the dirty-set + status indicator after each load-time write, replacing the two fire-and-forget `saveFile()` call sites in `onFileLoadFinished()`. No schema change.

**Tech Stack:** C++17, Qt 6.10 (Core/Widgets/Sql/Test), qmake + MinGW 13.1, PostgreSQL 16 (QPSQL).

---

## Machine conventions (read first)

- **MIP encryption.** Source files in the working tree may be encrypted at rest (`%TSD-Header-###%`). Two rules:
  - **Create new source files via Python delete-and-rewrite** (not the editor's create path), e.g. `python -c "open(r'src/database/WriteOutcome.h','w',encoding='utf-8',newline='\n').write(CONTENT)"`, so they don't inherit a MIP label.
  - **Before every C++ build, run** `python tools\decrypt_via_copy.py --apply` from the repo root.
- **Windows / PowerShell.** Use `;` not `&&` in PowerShell; backtick continuation. Commands below are shown for `cmd.exe`/batch unless noted.
- **Branch:** all work happens on `feature/v2.4.0-bugfix-batch` (already checked out).
- **Warnings are errors** (`-Werror -Wall -Wextra -Wpedantic`). New code must be warning-clean.

## File structure

| Action | Path | Responsibility |
|---|---|---|
| Create | `src/database/WriteOutcome.h` | `LoadSavePolicy` enum + `classifyLoadSaveResult()` declaration |
| Create | `src/database/WriteOutcome.cpp` | Pure mapping `WriteResult -> LoadSavePolicy` |
| Create | `tests/tst_writeoutcome/tst_writeoutcome.cpp` | Unit test for the mapping (no DB) |
| Create | `tests/tst_writeoutcome/tst_writeoutcome.pro` | qmake project for the unit test |
| Modify | `tests/tests.pro` | Register `tst_writeoutcome` in `SUBDIRS` |
| Modify | `DataViewerEnterprise.pro` | Add `WriteOutcome.{h,cpp}`; bump `VERSION` 2.3.1 → 2.3.2 |
| Modify | `src/MainWindow.h` | Declare `void persistLoadedFile(int fileIndex);` |
| Modify | `src/MainWindow.cpp` | `#include` + implement `persistLoadedFile`; replace the two save sites |
| Modify | `CLAUDE.md` | Record the semver release policy |

Reference (already verified in the codebase):
- `enum class WriteResult { Success, VersionMismatch, RowDeleted, UniqueViolation, OfflineReadOnly, OtherError };` — `src/database/DatabaseManager.h:32`
- `enum DbStatus { DbStatusOk, DbStatusModified, DbStatusDisconnected };` — `src/MainWindow.h:78`
- `WriteResult tryWriteFile(FileResult&);` (mutable, back-fills id/version) — `DatabaseManager.h:109`
- `FileResult loadFileByPath(const QString&) const;` — `DatabaseManager.h:121`
- `QString lastError() const;` — `DatabaseManager.h:254`
- `QSet<QString> m_modifiedFilePaths;` — `src/MainWindow.h:290`
- The two save sites: `src/MainWindow.cpp:2361` (reload-in-place) and `:2385` (new file), each followed by an **unconditional** `m_modifiedFilePaths.remove(...)` + `updateDbSyncIndicator()`.

---

## Task 1: Pure `classifyLoadSaveResult` helper + unit test (TDD)

**Files:**
- Create: `src/database/WriteOutcome.h`
- Create: `src/database/WriteOutcome.cpp`
- Create: `tests/tst_writeoutcome/tst_writeoutcome.cpp`
- Create: `tests/tst_writeoutcome/tst_writeoutcome.pro`
- Modify: `tests/tests.pro`

- [ ] **Step 1: Create the header `src/database/WriteOutcome.h`**

```cpp
#pragma once

#include "DatabaseManager.h"   // for DVE::WriteResult

namespace DVE {

// How MainWindow should react to the result of an automatic, load-time
// saveFile()/tryWriteFile(). Keeps the policy (clear vs. keep the dirty flag,
// and which message to surface) in one pure, unit-tested place. DATAVIEWER-3.
enum class LoadSavePolicy {
    Saved,         // WriteResult::Success         -> clear dirty, no message
    RetryOffline,  // WriteResult::OfflineReadOnly -> keep dirty, info "offline"
    RetryConflict, // Version/RowDeleted           -> keep dirty, warn "changed by another user"
    RetryError     // UniqueViolation/OtherError   -> keep dirty, warn generic failure
};

LoadSavePolicy classifyLoadSaveResult(WriteResult r);

} // namespace DVE
```

- [ ] **Step 2: Create a STUB `src/database/WriteOutcome.cpp`** (intentionally wrong, so the test fails at runtime)

```cpp
#include "WriteOutcome.h"

namespace DVE {

LoadSavePolicy classifyLoadSaveResult(WriteResult /*r*/)
{
    return LoadSavePolicy::Saved;   // STUB: replaced in Step 7
}

} // namespace DVE
```

- [ ] **Step 3: Create the test `tests/tst_writeoutcome/tst_writeoutcome.cpp`**

```cpp
#include <QtTest>
#include "WriteOutcome.h"

using namespace DVE;

class tst_WriteOutcome : public QObject
{
    Q_OBJECT
private slots:
    void success_isSaved();
    void offline_isRetryOffline();
    void versionMismatch_isRetryConflict();
    void rowDeleted_isRetryConflict();
    void uniqueViolation_isRetryError();
    void otherError_isRetryError();
};

void tst_WriteOutcome::success_isSaved()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::Success)),         int(LoadSavePolicy::Saved)); }

void tst_WriteOutcome::offline_isRetryOffline()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::OfflineReadOnly)), int(LoadSavePolicy::RetryOffline)); }

void tst_WriteOutcome::versionMismatch_isRetryConflict()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::VersionMismatch)), int(LoadSavePolicy::RetryConflict)); }

void tst_WriteOutcome::rowDeleted_isRetryConflict()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::RowDeleted)),      int(LoadSavePolicy::RetryConflict)); }

void tst_WriteOutcome::uniqueViolation_isRetryError()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::UniqueViolation)), int(LoadSavePolicy::RetryError)); }

void tst_WriteOutcome::otherError_isRetryError()
{ QCOMPARE(int(classifyLoadSaveResult(WriteResult::OtherError)),      int(LoadSavePolicy::RetryError)); }

QTEST_APPLESS_MAIN(tst_WriteOutcome)
#include "tst_writeoutcome.moc"
```

- [ ] **Step 4: Create `tests/tst_writeoutcome/tst_writeoutcome.pro`**

```pro
QT += core sql testlib network
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils \
               ../../src/reporting ../../src/plotting ../../src/database ../common

TARGET = tst_writeoutcome

SOURCES += tst_writeoutcome.cpp \
           ../../src/database/WriteOutcome.cpp

HEADERS += ../../src/database/WriteOutcome.h
```

- [ ] **Step 5: Register the target in `tests/tests.pro`**

Insert `tst_writeoutcome` right after `tst_databasemanager`. Replace:

```pro
    tst_databasemanager \
    tst_offlinesnapshot \
```
with:
```pro
    tst_databasemanager \
    tst_writeoutcome \
    tst_offlinesnapshot \
```

- [ ] **Step 6: Build & run the test — expect FAIL**

```bat
python tools\decrypt_via_copy.py --apply
set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH%
cd tests\tst_writeoutcome
qmake CONFIG+=release
mingw32-make -j8
release\tst_writeoutcome.exe
cd ..\..
```
Expected: builds clean; test run prints `FAIL!` for the offline/versionMismatch/rowDeleted/uniqueViolation/otherError slots and exits non-zero (the stub returns `Saved` for everything).

- [ ] **Step 7: Implement the real `src/database/WriteOutcome.cpp`** (replace the stub body)

```cpp
#include "WriteOutcome.h"

namespace DVE {

LoadSavePolicy classifyLoadSaveResult(WriteResult r)
{
    switch (r) {
    case WriteResult::Success:         return LoadSavePolicy::Saved;
    case WriteResult::OfflineReadOnly: return LoadSavePolicy::RetryOffline;
    case WriteResult::VersionMismatch: return LoadSavePolicy::RetryConflict;
    case WriteResult::RowDeleted:      return LoadSavePolicy::RetryConflict;
    case WriteResult::UniqueViolation: return LoadSavePolicy::RetryError;
    case WriteResult::OtherError:      return LoadSavePolicy::RetryError;
    }
    return LoadSavePolicy::RetryError;  // defensive; unreachable for a valid enum
}

} // namespace DVE
```

- [ ] **Step 8: Rebuild & run the test — expect PASS**

```bat
cd tests\tst_writeoutcome
mingw32-make -j8
release\tst_writeoutcome.exe
cd ..\..
```
Expected: `Totals: 6 passed, 0 failed, 0 skipped`, exit code 0.

- [ ] **Step 9: Commit**

```bat
git add src/database/WriteOutcome.h src/database/WriteOutcome.cpp tests/tst_writeoutcome/ tests/tests.pro
git commit -m "feat(db): classifyLoadSaveResult helper + unit test (DATAVIEWER-3)" -m "Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: Wire `persistLoadedFile` into MainWindow

**Files:**
- Modify: `DataViewerEnterprise.pro` (add the new source/header)
- Modify: `src/MainWindow.h:~460` (declare the method)
- Modify: `src/MainWindow.cpp` (include, implement, replace 2 call sites)

- [ ] **Step 1: Add `WriteOutcome` to `DataViewerEnterprise.pro`**

In `SOURCES`, replace:
```pro
    src/database/RawGridJson.cpp \
    src/ui/NewFileDialog.cpp \
```
with:
```pro
    src/database/RawGridJson.cpp \
    src/database/WriteOutcome.cpp \
    src/ui/NewFileDialog.cpp \
```
In `HEADERS`, replace:
```pro
    src/database/RawGridJson.h \
    src/ui/NewFileDialog.h \
```
with:
```pro
    src/database/RawGridJson.h \
    src/database/WriteOutcome.h \
    src/ui/NewFileDialog.h \
```

- [ ] **Step 2: Declare the method in `src/MainWindow.h`**

After the line `    void updateDbSyncIndicator();` (~line 460), add:
```cpp
    // DATAVIEWER-3: persist a freshly loaded/refreshed TPM file and reflect the
    // WriteResult (keep dirty + surface on failure, clear on success).
    void persistLoadedFile(int fileIndex);
```

- [ ] **Step 3: Add the include to `src/MainWindow.cpp`**

After the existing `#include "database/LiveSync.h"` (~line 16), add:
```cpp
#include "database/WriteOutcome.h"
```

- [ ] **Step 4: Implement `persistLoadedFile`** — add this definition immediately **before** `void MainWindow::onFileLoadFinished()` (~line 2300)

```cpp
// DATAVIEWER-3: persist a freshly loaded/refreshed TPM file to the database and
// reflect the outcome. Unlike the old fire-and-forget saveFile(), this checks
// the WriteResult: a failed save keeps the file marked dirty (so the Ctrl+U
// batch / close-flush retries it) and surfaces the reason, instead of silently
// dropping it. Uses the mutable tryWriteFile overload so id/version are stamped
// back into m_loadedFiles for presence dots + subsequent saves.
void MainWindow::persistLoadedFile(int fileIndex)
{
    if (fileIndex < 0 || fileIndex >= m_loadedFiles.size())
        return;
    FileResult& fr = m_loadedFiles[fileIndex];

    // No database configured: nothing to persist; mirror the pre-DATAVIEWER-3
    // behavior of treating the file as not-dirty for the (no-op) indicator.
    if (!m_db) {
        m_modifiedFilePaths.remove(fr.filePath);
        updateDbSyncIndicator();
        return;
    }

    DVE::WriteResult r = m_db->tryWriteFile(fr);

    // One-shot optimistic-concurrency recovery: the DB row changed or was
    // deleted since we inherited id/version. Re-inherit from the current row
    // and retry exactly once (no loop).
    if (r == DVE::WriteResult::VersionMismatch || r == DVE::WriteResult::RowDeleted) {
        const FileResult dbRow = m_db->loadFileByPath(fr.filePath);
        if (dbRow.id > 0) {
            fr.id      = dbRow.id;
            fr.version = dbRow.version;
            r = m_db->tryWriteFile(fr);
        }
    }

    const QString name = fr.fileName.isEmpty()
                             ? QFileInfo(fr.filePath).fileName()
                             : fr.fileName;

    switch (DVE::classifyLoadSaveResult(r)) {
    case DVE::LoadSavePolicy::Saved:
        m_modifiedFilePaths.remove(fr.filePath);
        break;
    case DVE::LoadSavePolicy::RetryOffline:
        m_modifiedFilePaths.insert(fr.filePath);
        updateStatusBar(tr("Offline: '%1' was not saved to the database; "
                           "it will be retried when the connection returns.").arg(name));
        qWarning().noquote() << "[persistLoadedFile] offline, not saved:"
                             << name << "-" << m_db->lastError();
        break;
    case DVE::LoadSavePolicy::RetryConflict:
        m_modifiedFilePaths.insert(fr.filePath);
        updateStatusBar(tr("'%1' was changed by another user and was not saved; "
                           "press Ctrl+U to retry.").arg(name));
        qWarning().noquote() << "[persistLoadedFile] OCC conflict, not saved:"
                             << name << "-" << m_db->lastError();
        break;
    case DVE::LoadSavePolicy::RetryError:
        m_modifiedFilePaths.insert(fr.filePath);
        updateStatusBar(tr("Failed to save '%1' to the database: %2")
                            .arg(name, m_db->lastError()));
        qWarning().noquote() << "[persistLoadedFile] save failed:"
                             << name << "-" << m_db->lastError();
        break;
    }

    updateDbSyncIndicator();
}
```

- [ ] **Step 5: Replace the reload-in-place save site** in `onFileLoadFinished()` (~line 2361). Replace:
```cpp
            if (m_db) m_db->saveFile(m_loadedFiles[i]);
            m_modifiedFilePaths.remove(result.filePath);
            updateDbSyncIndicator();
```
with:
```cpp
            persistLoadedFile(i);
```

- [ ] **Step 6: Replace the new-file save site** in `onFileLoadFinished()` (~line 2385). Replace:
```cpp
    if (m_db) m_db->saveFile(m_loadedFiles[m_currentFileIndex]);
    m_modifiedFilePaths.remove(result.filePath);
    updateDbSyncIndicator();
```
with:
```cpp
    persistLoadedFile(m_currentFileIndex);
```

- [ ] **Step 7: Build the whole app — expect clean compile** (incremental, no VERSION change yet)

```bat
python tools\decrypt_via_copy.py --apply
set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH%
qmake CONFIG+=release DataViewerEnterprise.pro
mingw32-make -j8
```
Expected: builds to `release\DataViewer.exe` with no warnings/errors. (`-Werror` will fail the build on any issue.)

- [ ] **Step 8: Commit**

```bat
git add DataViewerEnterprise.pro src/MainWindow.h src/MainWindow.cpp
git commit -m "fix(db): check load-time TPM save result; keep dirty + surface failures (DATAVIEWER-3)" -m "onFileLoadFinished now routes both save sites through persistLoadedFile, which keeps a file dirty for retry and shows offline/conflict/error status instead of silently dropping a failed saveFile()." -m "Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

Note: keeping a failed file dirty means the existing `onUpdateDatabase()` batch (Ctrl+U / close-flush) automatically retries it — no change needed there. Auto-flush on reconnect for whole TPM files is out of scope (per-cell reconnect is LiveSync's job).

---

## Task 3: Version bump to v2.3.2, record semver policy, full verification, installer

**Files:**
- Modify: `DataViewerEnterprise.pro:14` (VERSION)
- Modify: `CLAUDE.md` (Release Workflow)

- [ ] **Step 1: Bump VERSION** in `DataViewerEnterprise.pro`. Replace `VERSION = 2.3.1` with `VERSION = 2.3.2`.

- [ ] **Step 2: Record the semver policy in `CLAUDE.md`.** Read `CLAUDE.md`, find the `## Release Workflow` section, and insert this subsection at its end (immediately before `## Deployment Self-Test`):

```markdown
### Versioning scheme

Semantic `x.y.z`:
- **Patch (`z`)** — internal build, **not deployed**. Staged fixes verified locally.
- **Minor (`y`)** — deployable release (the version dropped on Synology).
- **Major (`x`)** — fundamental changes.

Batch fixes ship as consecutive internal patch builds, then wrap into one deployable
minor release (e.g. v2.3.2 / v2.3.3 / v2.3.4 internal → v2.4.0 deploy).
```

- [ ] **Step 3: Clean rebuild (VERSION bumps need a clean rebuild)**

```bat
python tools\decrypt_via_copy.py --apply
set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH%
qmake CONFIG+=release DataViewerEnterprise.pro
mingw32-make clean
mingw32-make -j8
```
Expected: `release\DataViewer.exe` rebuilt; its FileVersion is `2.3.2`.

- [ ] **Step 4: Run the full test suite** (ephemeral Postgres for the DB suites; `-Rebuild` re-runs qmake so the new `tst_writeoutcome` is picked up)

```powershell
.\tests\start-test-postgres.ps1
.\tests\run-tests.ps1 -Rebuild
```
Expected: final line `Results: N passed, 0 failed, M skipped` with **0 failed** and `tst_writeoutcome` showing `PASS  tst_writeoutcome  6 passed`. (DB suites that need `DVE_TEST_PG_CONN` run because `start-test-postgres.ps1` set it.)

- [ ] **Step 5: Build the internal installer (v2.3.2)**

```bat
tools\prepare_python_embed.bat
build_installer.bat
```
Expected: `build_installer.bat` confirms `release\DataViewer.exe` FileVersion == 2.3.2 and writes `dist\DataViewer-setup.exe`. **Do not** copy it anywhere — surface the path to Charlie. No Synology.

- [ ] **Step 6: Commit**

```bat
git add DataViewerEnterprise.pro CLAUDE.md
git commit -m "chore(release): v2.3.2 internal — TPM load-time persistence hardening (DATAVIEWER-3)" -m "Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 7: Mark the Plane issue Ready for Release** (via the `plane` MCP, not a shell command): set DATAVIEWER-3 (`a5ce8e44-ac39-4de6-8be5-cfd6b0bd575b`) state to **Ready for Release** (`ba78adb9-bdf5-4586-a675-ae5f0de2f8c7`). Then stop and hand the `dist\DataViewer-setup.exe` path to Charlie for eyeball-testing. **Do not deploy.**

---

## Self-review

**Spec coverage** (against `docs/superpowers/specs/2026-06-05-bugfix-batch-design.md` §3):
- "capture the WriteResult" → Task 2 Step 4 (`tryWriteFile` result captured).
- "clear dirty only on Success; keep dirty otherwise" → `persistLoadedFile` switch.
- "surface via DB status indicator + log; offline/conflict/error messages" → switch + `qWarning`.
- "re-inherit id/version + retry once on Version/RowDeleted" → one-shot OCC block.
- "no schema change; transactional" → confirmed (`tryWriteFile` is BEGIN..COMMIT).
- Verification "(a) healthy save clears dirty / (b) failure keeps dirty" → the pure helper is unit-tested (Task 1); the GUI wiring is build-verified + confirmed on the work machine via the now-visible status/log. (MainWindow's async load path is not unit-testable without a GUI harness; this is the honest seam.)

**Placeholder scan:** none — every step has concrete code/commands/expected output.

**Type consistency:** `LoadSavePolicy { Saved, RetryOffline, RetryConflict, RetryError }` and `WriteResult { Success, VersionMismatch, RowDeleted, UniqueViolation, OfflineReadOnly, OtherError }` are used identically across the helper, test, and `persistLoadedFile`. Method name `persistLoadedFile(int)` matches between `MainWindow.h` decl and the two call sites.
