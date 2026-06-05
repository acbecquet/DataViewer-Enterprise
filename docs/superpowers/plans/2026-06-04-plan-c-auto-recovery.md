# Plan C — Auto-Recovery + Crash Snapshot — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax. **Every C++ source file on this machine must be created via the Python delete-and-rewrite pattern** (MIP labels corrupt Edit/Write output); run `python tools/decrypt_via_copy.py --apply` before each build. All app code is in `namespace DVE`. Build is `-Werror -Wall -Wextra -Wpedantic`.

**Goal:** On any non-clean termination (crash or the updater's `std::_Exit`), restore the user's complete in-memory work across TPM, Sensory, and Detailed Sensory modes — including never-saved edits — via a constantly-flushed rolling snapshot; plus a consolidated clean-close save flow.

**Architecture:** A new `RecoveryManager` (`src/utils/`) owns a per-item JSON snapshot under `%LOCALAPPDATA%/SDR/DataViewer Enterprise/Recovery/`, flushed debounced + off-thread on every mutation and synchronously before the updater's hard exit. On startup the live dir is renamed to `Recovery_prev/`; if it's non-empty (prior instance died uncleanly — `SingleInstance` guarantees one instance), MainWindow offers recovery (auto-prompt + a Tools→Recover dialog). Clean close clears both dirs.

**Tech Stack:** C++17 / Qt 6.10 (`QJsonDocument`, `QtConcurrent`, `QSaveFile` for atomic writes), qmake + MinGW, `namespace DVE`. New JSON serializers for the `FileResult` tree; reuse `SensorySession` JSON; promote `DetailedSensorySession` JSON out of `DatabaseManager.cpp`.

---

## File Structure

**Create:**
- `src/pipeline/ReportDataJson.{h,cpp}` — `FileResult`↔JSON (+ nested sheet/sample/row).
- `src/pipeline/DetailedSensoryData.cpp` — new TU; promoted `DetailedSensorySession` JSON + `isPlaceholderSession`.
- `src/utils/RecoveryManager.{h,cpp}` — snapshot store, detection, cadence, flush.
- `src/ui/RecoverDialog.{h,cpp}` — selective-reload dialog for Tools→Recover.
- `tests/tst_reportdatajson/` (+ `.pro`), `tests/tst_recoverymanager/` (+ `.pro`).

**Modify:**
- `src/database/DatabaseManager.cpp` — call promoted detailed-session JSON instead of the anon-namespace copies.
- `src/pipeline/DetailedSensoryData.h` — declare promoted functions + `isPlaceholderSession`.
- `src/utils/UpdateChecker.{h,cpp}` — `preExitHook` invoked before `std::_Exit(0)`.
- `src/MainWindow.{h,cpp}` — own `RecoveryManager`; `noteDirty` hooks; startup recovery prompt; Tools→Recover; consolidated `closeEvent`; honor `m_detailedSensorySessionsDirty`.
- `DataViewerEnterprise.pro`, `tests/tests.pro` — new sources/tests; `VERSION` bump at the end.

---

## Phase 1 — Serialization foundation

### Task C1: `ReportDataJson` — FileResult tree ↔ JSON

**Files:** Create `src/pipeline/ReportDataJson.{h,cpp}`, `tests/tst_reportdatajson/tst_reportdatajson.{cpp,pro}`.

**Coverage decision:** serialize all data-bearing fields. **Omit** `SheetResult::tpmTrend`/`puffCounts` (recomputed on restore from rows), `SheetResult::images` (embedded Excel bytes, re-derivable), `SheetResult::dbDataIncomplete` (transient). Keep `columnHeaders`, `hasPerRowRegime`, `isRawTable`, `rawHeaders`, `rawRows`, `extra` (via `QJsonObject::fromVariantMap`), all image-ref vectors, and every `id`/`version`.

- [ ] **Step 1: Header.** Create `src/pipeline/ReportDataJson.h`:
```cpp
#pragma once
#include <QJsonObject>
#include "pipeline/ReportData.h"
namespace DVE {
QJsonObject fileResultToJson(const FileResult& f);
FileResult  fileResultFromJson(const QJsonObject& obj);
} // namespace DVE
```

- [ ] **Step 2: Failing test.** Create `tests/tst_reportdatajson/tst_reportdatajson.cpp` — build a fully-populated `FileResult` (2 sheets: one normal sheet with a sample carrying 2 `DataRow`s + image vectors + `extra`, one raw/SOP sheet with `rawHeaders`/`rawRows`), round-trip through `fileResultToJson`→`fileResultFromJson`, and `QCOMPARE` every persisted field. Include `id`/`version` on each level and a `QRectF` in `imageLayouts`. Assert omitted fields (`tpmTrend`) are empty after round-trip (documents the contract).

- [ ] **Step 3: `.pro`.** Create `tests/tst_reportdatajson/tst_reportdatajson.pro`:
```pro
QT += core testlib
QT -= gui
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_reportdatajson
INCLUDEPATH += ../../src ../../src/pipeline
SOURCES += tst_reportdatajson.cpp ../../src/pipeline/ReportDataJson.cpp
HEADERS += ../../src/pipeline/ReportDataJson.h ../../src/pipeline/ReportData.h
```
Add `tst_reportdatajson` to `tests/tests.pro` SUBDIRS.

- [ ] **Step 4: Implement** `src/pipeline/ReportDataJson.cpp` (write via Python). Use these helpers and field mappings (mirror the existing `sensorySessionToJson` pattern):
```cpp
#include "pipeline/ReportDataJson.h"
#include <QJsonArray>
namespace DVE {
namespace {
QJsonObject rectToJson(const QRectF& r) {
    return QJsonObject{{"x",r.x()},{"y",r.y()},{"w",r.width()},{"h",r.height()}};
}
QRectF rectFromJson(const QJsonObject& o) {
    return QRectF(o["x"].toDouble(), o["y"].toDouble(), o["w"].toDouble(), o["h"].toDouble());
}
QJsonObject rowToJson(const DataRow& d) {
    return QJsonObject{
        {"puffs",d.puffs},{"before",d.beforeWeight},{"after",d.afterWeight},
        {"draw_pressure",d.drawPressure},{"resistance",d.resistance},
        {"puffing_regime",d.puffingRegime},{"smell",d.smell},{"clog",d.clog},{"notes",d.notes},
        {"tpm",d.tpm},{"tpm_pd",d.tpmPowerDensity},{"variation",d.variationTPM},
        {"oil_consumed",d.oilConsumed},{"id",static_cast<double>(d.id)},{"version",d.version}};
}
DataRow rowFromJson(const QJsonObject& o) {
    DataRow d;
    d.puffs=o["puffs"].toDouble(); d.beforeWeight=o["before"].toDouble();
    d.afterWeight=o["after"].toDouble(); d.drawPressure=o["draw_pressure"].toDouble();
    d.resistance=o["resistance"].toDouble(); d.puffingRegime=o["puffing_regime"].toString();
    d.smell=o["smell"].toString(); d.clog=o["clog"].toString(); d.notes=o["notes"].toString();
    d.tpm=o["tpm"].toDouble(); d.tpmPowerDensity=o["tpm_pd"].toDouble();
    d.variationTPM=o["variation"].toDouble(); d.oilConsumed=o["oil_consumed"].toDouble();
    d.id=static_cast<qint64>(o["id"].toDouble(-1)); d.version=o["version"].toInt();
    return d;
}
// sampleToJson/sampleFromJson: every SampleResult scalar (see §1 field table),
//   "rows" -> array of rowToJson, "extra" -> QJsonObject::fromVariantMap(s.extra),
//   image vectors -> parallel arrays ("image_paths", "image_layouts"=rectToJson[],
//   "image_crops"=rectToJson[], "image_ids", "image_versions"), id/version.
// sheetToJson/sheetFromJson: sheetName, templateVersion, hasPerRowRegime,
//   columnHeaders, overallAvgTPM, overallStdDevTPM, isRawTable, rawHeaders,
//   rawRows (array of string arrays), "samples" -> array, id/version.
} // anon
QJsonObject fileResultToJson(const FileResult& f) { /* filePath, fileName,
    templateVersion, sheetNames, "sheets"->array(sheetToJson), id, version */ }
FileResult fileResultFromJson(const QJsonObject& obj) { /* inverse */ }
} // namespace DVE
```
Implement the elided helpers fully field-for-field per §1 of the facts report. For `qint64` ids use `static_cast<double>` on write and `static_cast<qint64>(...toDouble(-1))` on read (JSON has no 64-bit int). For `extra`: `QJsonObject::fromVariantMap(s.extra)` / `obj["extra"].toObject().toVariantMap()`.

- [ ] **Step 5: Run test → pass.** `release\tst_reportdatajson.exe -o res.txt,txt` then read `res.txt` (QtTest stdout doesn't survive the cmd→bash pipe on this machine; use the `-o file` form). Expect `0 failed`.

- [ ] **Step 6: Add to app `.pro`** (`SOURCES`+`HEADERS`) and **commit** (`feat(recovery): FileResult tree JSON serializer + round-trip test`).

### Task C2: Promote `DetailedSensorySession` JSON + add `isPlaceholderSession`

**Files:** Create `src/pipeline/DetailedSensoryData.cpp`; modify `src/pipeline/DetailedSensoryData.h`, `src/database/DatabaseManager.cpp`; create `tests/tst_detailedsensoryjson/`.

- [ ] **Step 1: Declarations** in `DetailedSensoryData.h` (bottom, inside `namespace DVE`, mirroring `SensoryData.h`):
```cpp
class QJsonObject;
namespace DVE {
QJsonObject detailedSensorySessionToJson(const DetailedSensorySession& s);
DetailedSensorySession detailedSensorySessionFromJson(const QJsonObject& obj);
bool isPlaceholderSession(const DetailedSensorySession& s); // untitled & no real data
} // namespace DVE
```

- [ ] **Step 2: Failing round-trip test** `tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp` — populate a `DetailedSensorySession` (header fields + 2 samples with scores), round-trip via the new `QJsonObject` pair, `QCOMPARE` fields. Plus assert `isPlaceholderSession` is true for a fresh `"New Session"` and false once a sample/title is set. `.pro` links `DetailedSensoryData.cpp` only (no DB deps). Add to `tests/tests.pro`.

- [ ] **Step 3: Implement** `src/pipeline/DetailedSensoryData.cpp` (write via Python): move the bodies of `serializeDetailedSensoryJson`/`deserializeDetailedSensoryJson` from `DatabaseManager.cpp` (facts §3) into `QJsonObject`-based `detailedSensorySessionToJson`/`FromJson` (drop the `QJsonDocument` string wrapping — return/accept the object). Add `isPlaceholderSession`: true when `sessionName` is empty or `"New Session"` AND `samples` is empty (or all samples empty) AND header fields blank. Include `<QJsonObject>`,`<QJsonArray>`; `kDetailedAllMetrics`/`kDetailedMetricMaxScore` come from the header.

- [ ] **Step 4: Rewire `DatabaseManager.cpp`** — delete the two anon-namespace copies; at the four call sites (facts §3: ~lines 2215, 2276, 2370, 2409) wrap with `QString::fromUtf8(QJsonDocument(detailedSensorySessionToJson(s)).toJson(Compact))` on write and `detailedSensorySessionFromJson(QJsonDocument::fromJson(bytes).object())` on read. (Leave `OfflineSnapshot.cpp`'s independent `…JsonLocal` copies untouched — out of scope.)

- [ ] **Step 5: Build the app** (`qmake CONFIG+=release && mingw32-make release -j8`) to confirm DatabaseManager still compiles + links; **run** `tst_detailedsensoryjson` (`-o` form) → pass.

- [ ] **Step 6: `.pro`** (+ `DetailedSensoryData.cpp` to SOURCES) and **commit** (`refactor(recovery): promote DetailedSensorySession JSON to DetailedSensoryData; add isPlaceholderSession`).

---

## Phase 2 — RecoveryManager

### Task C3: RecoveryManager — snapshot store (write/read)

**Files:** Create `src/utils/RecoveryManager.{h,cpp}`, `tests/tst_recoverymanager/`.

- [ ] **Step 1: Header** `src/utils/RecoveryManager.h` — types + store API:
```cpp
#pragma once
#include <QObject>
#include <QString>
#include <QVector>
#include <QJsonObject>
namespace DVE {
enum class RecoveryKind { Tpm, Sensory, Detailed };
struct RecoveryEntry {
    RecoveryKind kind; QString id; QString displayName; QString sourcePath;
    bool dirty = false; QString blobFile; QJsonObject payload; // payload filled by readAll()
};
class RecoveryManager : public QObject {
    Q_OBJECT
public:
    explicit RecoveryManager(QObject* parent=nullptr); // resolves dir = AppLocalDataLocation/Recovery
    QString liveDir() const;  QString prevDir() const; // <root>/Recovery , <root>/Recovery_prev
    // Store primitives (Task C3):
    bool writeItem(const RecoveryEntry& e, const QJsonObject& payload); // atomic blob + index update
    bool removeItem(RecoveryKind kind, const QString& id);
    QVector<RecoveryEntry> readAll(const QString& dir) const; // index + each blob payload
    void setDirOverride(const QString& dir); // tests
private:
    QString m_root; // override or AppLocalDataLocation
    QVector<RecoveryEntry> m_index; // current live index mirror
    bool writeIndex();
    static QString blobName(RecoveryKind, const QString& id);
};
} // namespace DVE
```

- [ ] **Step 2: Failing test** `tst_recoverymanager.cpp` (`testCase: storeRoundTrip`): `setDirOverride(QTemporaryDir)`, `writeItem` two entries (one Tpm with a `fileResultToJson` payload, one Sensory), then `readAll(liveDir())` → 2 entries, payloads equal, `index.json` exists. `.pro` links `RecoveryManager.cpp ReportDataJson.cpp` (+ ReportData). Add to `tests/tests.pro`.

- [ ] **Step 3: Implement store** in `RecoveryManager.cpp` (write via Python): dir = `setDirOverride` value else `QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation)`; `liveDir()=root+"/Recovery"`, `prevDir()=root+"/Recovery_prev"`. `writeItem`: `QDir().mkpath(liveDir())`; write blob with **`QSaveFile`** (atomic) named `blobName(kind,id)` (e.g. `tpm_<id>.json`) containing `QJsonDocument(payload).toJson()`; update `m_index` (replace by kind+id) and `writeIndex()` (also `QSaveFile`). `readAll(dir)`: parse `dir/index.json` → entries; for each, load its blob into `payload`. `blobName`: `<kind>_<sanitized id>.json`.

- [ ] **Step 4: Run test (`-o` form) → pass. Commit** (`feat(recovery): RecoveryManager snapshot store (atomic write/read)`).

### Task C4: RecoveryManager — detection + lifecycle

- [ ] **Step 1: Extend header** with:
```cpp
void adoptPreviousSession();             // startup: rename Recovery -> Recovery_prev, fresh Recovery
bool hasRecoverable() const;             // Recovery_prev/index.json exists & non-empty
QVector<RecoveryEntry> recoverableItems() const; // readAll(prevDir())
void clear();                            // remove Recovery and Recovery_prev (clean exit)
```

- [ ] **Step 2: Failing test** `detectionLifecycle`: with override dir, `writeItem` two entries into live; `adoptPreviousSession()` → `liveDir()` now empty/fresh, `hasRecoverable()` true, `recoverableItems().size()==2`; `clear()` → both dirs gone, `hasRecoverable()` false.

- [ ] **Step 3: Implement.** `adoptPreviousSession`: if `prevDir()` exists, `QDir(prevDir()).removeRecursively()`; if `liveDir()` exists, `QDir().rename(liveDir(), prevDir())`; reset `m_index`. `hasRecoverable`: `QFile::exists(prevDir()+"/index.json")` && `!readAll(prevDir()).isEmpty()`. `recoverableItems`: `readAll(prevDir())`. `clear`: `removeRecursively()` both; clear `m_index`.

- [ ] **Step 4: Run test (`-o`) → pass. Commit** (`feat(recovery): RecoveryManager detection + move-to-prev lifecycle`).

### Task C5: RecoveryManager — cadence + off-thread flush + state provider

- [ ] **Step 1: Extend header** with the capture interface + debounce:
```cpp
// MainWindow supplies the current full state on demand:
using StateProvider = std::function<QVector<RecoveryEntry>()>; // entries WITH payloads
void setStateProvider(StateProvider p);
void noteDirty();                 // arms ~2s debounce
void flushNow(bool synchronous);  // capture via provider, write all, prune removed
private: QTimer* m_debounce; QTimer* m_safety; StateProvider m_provider;
```

- [ ] **Step 2: Failing test** `flushViaProvider`: set a provider returning 1 Tpm entry; `flushNow(true)` → `readAll(liveDir())` has it. Change provider to return 0 entries; `flushNow(true)` → live blob pruned (removed items deleted).

- [ ] **Step 3: Implement.** ctor: `m_debounce` (single-shot 2000 ms) → `flushNow(false)`; `m_safety` (30000 ms repeating) → `flushNow(false)`. `noteDirty()`: `m_debounce->start()`. `flushNow(sync)`: snapshot `m_provider()`; the actual disk writes run on a worker via `QtConcurrent::run` when `!sync` (UI must never block — see the v2.0.6 freeze lesson), or inline when `sync`; diff against `m_index` and delete pruned blobs; rewrite index. Guard re-entrancy with a flag.

- [ ] **Step 4: Run test (`-o`) → pass. Commit** (`feat(recovery): debounced off-thread flush + state provider`).

---

## Phase 3 — Integration

### Task C6: MainWindow — own RecoveryManager + capture hooks

**Files:** `src/MainWindow.{h,cpp}`.

- [ ] **Step 1:** Add member `RecoveryManager* m_recovery = nullptr;` and `#include "utils/RecoveryManager.h"`. Construct in the ctor after panels are available; `m_recovery->adoptPreviousSession();` early (before any flush) so a crashed prior session is preserved.

- [ ] **Step 2: State provider.** Implement `QVector<RecoveryEntry> MainWindow::captureRecoveryState() const`: for each `FileResult` in `m_loadedFiles` → `RecoveryEntry{Tpm, filePath-or-synthetic-id, fileName, filePath, m_modifiedFilePaths.contains(filePath), payload=fileResultToJson(f)}`; for each `m_sensoryPanel->allSessions()` → Sensory entry, payload `sensorySessionToJson(s)`; for each `m_detailedSensoryPanel->allSessions()` → Detailed entry, payload `detailedSensorySessionToJson(s)`. Set via `m_recovery->setStateProvider([this]{ return captureRecoveryState(); });`.

- [ ] **Step 3: Dirty hooks.** Call `m_recovery->noteDirty()` in: `markFileModified()` (covers all 7 TPM sites), the `SensoryPanel::sessionsChanged` lambda (~3535), and the `DetailedSensoryPanel::sessionsChanged` lambda (~3600). Also `noteDirty()` on file open/close (so the index tracks the open set).

- [ ] **Step 4: Build app (`-Werror`) → clean. Commit** (`feat(recovery): MainWindow owns RecoveryManager; capture hooks on all three stores`). (No unit test — exercised via build + later manual; capture logic is thin glue over tested serializers.)

### Task C7: UpdateChecker — synchronous flush before `std::_Exit`

**Files:** `src/utils/UpdateChecker.{h,cpp}`, `src/MainWindow.cpp`.

- [ ] **Step 1:** Add to `UpdateChecker`: `void setPreExitHook(std::function<void()> hook);` + member `std::function<void()> m_preExitHook;` (`#include <functional>`).

- [ ] **Step 2:** In the Update Now lambda (facts §8, before `std::_Exit(0)` at line 278): `if (m_preExitHook) m_preExitHook();`.

- [ ] **Step 3:** In MainWindow, where the `UpdateChecker` is created/used, set the hook: `updateChecker->setPreExitHook([this]{ if (m_recovery) m_recovery->flushNow(true); });` (synchronous capture so the snapshot is complete before the hard exit).

- [ ] **Step 4: Build → clean. Commit** (`feat(recovery): flush recovery snapshot before updater std::_Exit`).

### Task C8: Startup recovery prompt

**Files:** `src/MainWindow.{h,cpp}`.

- [ ] **Step 1:** After panels are constructed and the window is ready (end of ctor or a `QTimer::singleShot(0,...)`), call `maybeOfferRecovery()`.

- [ ] **Step 2: Implement** `maybeOfferRecovery()`: if `!m_recovery->hasRecoverable()` return. Count items; `QMessageBox::question(this, "Recover Previous Session", QString("%1 file(s)/session(s) were open when DataViewer last closed unexpectedly.\nReload them?").arg(n), Yes|No)`. **Yes** → `restoreItems(m_recovery->recoverableItems())` then leave `Recovery_prev` in place until a clean close. **No** → keep `Recovery_prev` so Tools→Recover still works this session.

- [ ] **Step 3: Implement** `restoreItems(const QVector<RecoveryEntry>&)`: group by kind. Tpm → `fileResultFromJson(payload)`, recompute `tpmTrend`/`puffCounts` via the existing per-sheet recompute (call `recalculateSampleMetrics`/equivalent), append to `m_loadedFiles`, `m_modifiedFilePaths.insert(filePath)`; Sensory/Detailed → build session vectors, `m_sensoryPanel->loadSessions(...)` / `m_detailedSensoryPanel->loadSessions(...)`, set the matching dirty flag, then `inheritExistingIdsAndVersions()` so re-save uses the OCC path. Refresh trees/combos; switch to the mode of the first restored item.

- [ ] **Step 4: Build → clean. Commit** (`feat(recovery): startup prompt restores previous-session work`).

### Task C9: RecoverDialog + Tools→Recover button

**Files:** Create `src/ui/RecoverDialog.{h,cpp}`; modify `src/MainWindow.{h,cpp}`.

- [ ] **Step 1: RecoverDialog** — a `QDialog` taking `QVector<RecoveryEntry>`; a `QListWidget` with a checkable item per entry (`"[TPM] name — unsaved"`); `selected()` returns the checked subset; OK/Cancel. Write via Python.

- [ ] **Step 2: Tools button** (facts §7) — in `buildToolsTab`, add a "Recovery" group + `addLargeButton("Recover", AppTheme::icon("rotate-ccw"), "Restore unsaved work from the last session")`, `connect`→`onRecover`. Verify the icon name resolves (`AppTheme::icon`); fall back to an existing name if not.

- [ ] **Step 3:** `onRecover()`: items = `m_recovery->recoverableItems()`; if empty, `QMessageBox::information(...,"Nothing to recover")`; else open `RecoverDialog`; on accept `restoreItems(dlg.selected())`.

- [ ] **Step 4:** Add `RecoverDialog.{cpp,h}` to `.pro`. **Build → clean. Commit** (`feat(recovery): Tools -> Recover dialog for selective reload`).

### Task C10: Consolidated close flow + detailed-dirty fix

**Files:** `src/MainWindow.{h,cpp}`.

- [ ] **Step 1: Fix the dirty-flag gap (regression first).** In `promptSaveDatabase()` (facts §5): include `m_detailedSensorySessionsDirty` in the `hasSensory`/parts logic; in `updateDbSyncIndicator()` add the same branch; ensure `onUpdateDatabase()` saves detailed sessions and resets `m_detailedSensorySessionsDirty=false` after a successful save (it is currently never reset).

- [ ] **Step 2: Unsaved inventory + dialog.** Add `struct UnsavedItem { RecoveryKind kind; QString label; bool untitled; };` and `QVector<UnsavedItem> unsavedInventory() const` (TPM via `m_modifiedFilePaths`; Sensory via `m_sensorySessionsDirty` + `!isPlaceholderSession` per session; Detailed via `m_detailedSensorySessionsDirty` + `!isPlaceholderSession`). Replace the body of `promptSaveDatabase()`'s prompt (or wrap it) with one dialog: if inventory empty return true; else `QMessageBox` **Save All / Discard / Cancel**. Cancel→false (veto). Discard→true. Save All→ persist: named TPM files via existing save; for untitled sensory/detailed sessions, route through the existing per-mode save which already pops a Save-As (one per untitled), then `onUpdateDatabase()`; return true.

- [ ] **Step 3: Recovery teardown on clean close.** At the end of `closeEvent`, after the save path succeeds (and after the existing ImageCache wipe + snapshot regen), call `m_recovery->clear()` so no spurious prompt next launch. (The `std::_Exit` path in C7 deliberately does NOT clear — it flushes and leaves the dir.)

- [ ] **Step 4: Build → clean. Commit** (`feat(recovery): consolidated close prompt + honor detailed-sensory dirty flag; clear recovery on clean exit`).

---

## Phase 4 — Verify + release

### Task C11: Full build + test sweep

- [ ] **Step 1:** `python tools/decrypt_via_copy.py --apply`; full `qmake CONFIG+=release && mingw32-make release -j8` → clean `-Werror`.
- [ ] **Step 2:** Build + run the new tests (`tst_reportdatajson`, `tst_detailedsensoryjson`, `tst_recoverymanager`) via the `-o res.txt,txt` form → all pass.
- [ ] **Step 3:** Restart the test Postgres fresh (`docker rm -f dve-test-pg` + `tests\start-test-postgres.ps1`) and run the DB-touching suites (`tst_databasemanager` especially, since C2 rewired its detailed-session JSON) → green except known flakes (`tst_excelexporter::testRoundTripFormatE`, `tst_responsivelayout`).
- [ ] **Step 4: Commit** any test-wiring fixes.

### Task C12: Version bump + installer + release overview

- [ ] **Step 1:** Bump `VERSION = 2.3.0` in `DataViewerEnterprise.pro` (minor — auto-recovery is a feature; bundles the Settings-paths fix).
- [ ] **Step 2:** Clean rebuild (`mingw32-make clean` then build) + `tools\prepare_python_embed.bat` + `build_installer.bat`; verify `release\DataViewer.exe` and `dist\DataViewer-setup.exe` both report `2.3.0`.
- [ ] **Step 3:** Write `release_overview/release_overview_v_2_3_0.txt` (auto-recovery; consolidated close prompt; Settings paths apply to all save/load/new).
- [ ] **Step 4: Commit** (`chore(release): v2.3.0 — auto-recovery + crash snapshot`). Hand the installer to the user for eyeball-test (incl. a real crash/update-kill recovery), then merge to `main` + user's Synology drop.

---

## Manual verification (the part unit tests can't cover)
- Load files in all three modes, make edits, **kill the process** (Task Manager) → relaunch → recovery prompt → Reload all → edits intact.
- Same but click **No** → Tools→Recover → selective reload works.
- Trigger a **real update** (`std::_Exit` path) → relaunch → recovery offered (the original bug).
- Clean close with unsaved work → consolidated Save All / Discard / Cancel behaves; a dirty **detailed** session now prompts (regression fix).
- Clean close → next launch shows **no** spurious recovery prompt.

## Non-goals (carried from the spec)
Auto-update download/replace internals; cross-machine recovery; undo-stack recovery; per-keystroke snapshots; collapsing `OfflineSnapshot.cpp`'s duplicate `…JsonLocal` serializers.

**(I-1) Sensory/Detailed image associations are not recovered.** Sensory/Detailed recovery restores session **data** (scores, comments, fields) but **NOT** image associations — the `sensorySessionToJson` / `detailedSensorySessionToJson` serializers omit images (image bytes live in the DB `images` table), and the recovery snapshot reuses those serializers. TPM recovery (via `fileResultToJson`) **does** include image refs. We deliberately do **not** change the sensory/detailed serializers (that would alter the DB blob format); recovered sessions re-associate images from the DB on next save.

## Self-review notes
- Spec coverage: snapshot store (C3) ✓, detection/move-to-prev (C4) ✓, cadence/off-thread (C5) ✓, serialization incl. detailed promotion (C1/C2) ✓, updater flush (C7) ✓, consolidated close + detailed-dirty fix (C10) ✓, reopen prompt (C8) ✓, Recover dialog (C9) ✓, restore reconciliation (C8 step 3) ✓, images-as-path-refs (C1 coverage decision) ✓.
- Type consistency: `RecoveryEntry`, `RecoveryKind`, `StateProvider`, `fileResultToJson/FromJson`, `detailedSensorySessionToJson/FromJson` used identically across tasks.
- Test-output gotcha encoded in every run step (use `-o res.txt,txt`; QtTest stdout is swallowed by the cmd→bash pipe on this machine).
