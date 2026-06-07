# Zero Save-As — auto-filename + required-fields guardrail (sensory & detailed) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Eliminate every "Save As" dialog in Sensory and Detailed Sensory modes — the on-disk filename is auto-derived from the session label (test + tester + round) under the configured default directory, written silently alongside the DB. Add one guardrail: a non-empty **test name** and **tester** are required before any save/update; explicit saves block with an instruction, the background 5 s auto-save silently skips incomplete sessions.

**Architecture:** A shared pure path helper (`OutputPaths::autoSavePath`) replaces the three `QFileDialog::getSaveFileName` call sites. Two pure predicates (`isSensorySessionSavable` / `isDetailedSessionSavable`) define "has a usable natural key" (title + tester non-empty — the same columns as the DB unique index `(session_name, tester_name, date)`, so this is also a correctness fix against `('','',date)` key collisions). The predicates gate three behaviors: interactive saves block + instruct, background auto-save skips silently, per-item Close refuses to drop an unnamed-but-non-empty session. TPM mode is untouched.

**Tech Stack:** C++17 / Qt 6.10 (QtTest, QFileDialog, QMessageBox), qmake + MinGW, PostgreSQL 16 (QPSQL). MIP: run `python tools/decrypt_via_copy.py --apply` before any build; decrypt-before-edit. Build is `-Werror -Wall -Wextra -Wpedantic`. Test runner `tests\run-tests.ps1 [-Filter <suite>]`.

**Behavioral changes (intended — confirmed with Charlie):**
1. **Zero Save-As prompts** in sensory/detailed. Files auto-name to `test - tester - round` (sensory) / `test - tester` (detailed), sanitized, in the configured default dir; same values → same file → overwrite (intended dedup that kills filename mismatches).
2. **Test name + tester are required.** Explicit Save / Ctrl+S / Ctrl+U / explicit Close on a session with real data but a blank title or tester is blocked with an instruction. This **replaces** Sensory's old "auto-generate a default test name" soft-prompt (`resolveTestName`/`nextDefaultTestName`).
3. **Live-sync activates only after the header is complete.** A new session isn't persisted (no disk file, no DB row, no per-cell sync) until it has a name + tester; the close-guardrail prevents silent loss, and unsaved in-memory state still rides the existing recovery snapshot.

---

## File structure

- **`src/utils/OutputPaths.h` / `.cpp`** — add pure `static QString autoSavePath(ReportMode, const QString& sessionLabel, const QString& lastUsedDir, const QString& ext)`. Single source of the auto-derived path (`resolveDir` + `sanitize`).
- **`src/pipeline/SensoryData.h` / `.cpp`** — add pure `bool isSensorySessionSavable(const SensorySession&)`.
- **`src/pipeline/DetailedSensoryData.h` / `.cpp`** — add pure `bool isDetailedSessionSavable(const DetailedSensorySession&)`.
- **`src/ui/SensoryPanel.cpp`** — `save()`: hard require-fields guard at top; replace Save-As block with `autoSavePath`; remove the `resolveTestName` soft-prompt path.
- **`src/ui/DetailedSensoryPanel.cpp`** — `save()`: same.
- **`src/MainWindow.cpp`** — `onUpdateDatabase` loops (background skip / interactive summary), `saveSensorySessionsBeforeClose` / `saveDetailedSensorySessionsBeforeClose` (refuse unnamed non-empty), `promptSaveDatabase` (gate the disk-courtesy `save()` calls to savable sessions).
- **`DataViewerEnterprise.pro`** — `VERSION` 2.3.4 → 2.3.5.
- **Tests:** `tests/tst_outputpaths/`, `tests/tst_sensorydataplaceholder/`, `tests/tst_detailedsensoryjson/`.

---

## Task 1: `OutputPaths::autoSavePath` helper + unit test

**Files:**
- Modify: `src/utils/OutputPaths.h` (declare), `src/utils/OutputPaths.cpp` (define)
- Test: `tests/tst_outputpaths/tst_outputpaths.cpp`

> First read `OutputPaths.h/.cpp` to confirm: the `ReportMode` enum (values include `Sensory`, `DetailedSensory`), the signature of `resolveDir(ReportMode, const QString& lastUsedDir)`, and `sanitize(const QString&)`. Use the real signatures.

- [ ] **Step 1: Write the failing test.** In `tst_outputpaths.cpp` add a `private slot`:

```cpp
void autoSavePath_joinsResolvedDirSanitizedLabelAndExt()
{
    using namespace DVE;
    const QString dir = OutputPaths::resolveDir(ReportMode::Sensory, QString());
    // label with spaces + the " - " separator must be sanitized into the base
    const QString p = OutputPaths::autoSavePath(ReportMode::Sensory,
                                                "My Test - Alice - 1", QString(), ".xlsx");
    QVERIFY(p.startsWith(dir));
    QVERIFY(p.endsWith(".xlsx"));
    // base is the sanitized label (no raw forbidden chars); compare against sanitize()
    const QString expectedBase = OutputPaths::sanitize("My Test - Alice - 1");
    QVERIFY(p.contains(expectedBase));
    // ext normalization: passing "xlsx" (no dot) yields the same result
    QCOMPARE(OutputPaths::autoSavePath(ReportMode::Sensory, "X", QString(), "xlsx"),
             OutputPaths::autoSavePath(ReportMode::Sensory, "X", QString(), ".xlsx"));
    // empty label falls back to a non-empty base (no path ending in "/.xlsx")
    const QString empty = OutputPaths::autoSavePath(ReportMode::Sensory, "   ", QString(), ".xlsx");
    QVERIFY(!empty.endsWith("/.xlsx"));
}
```

- [ ] **Step 2: Run to verify it fails (undeclared).**
Run: `tests\run-tests.ps1 -Filter tst_outputpaths`
Expected: build FAILS — `autoSavePath` not a member.

- [ ] **Step 3: Declare** in `src/utils/OutputPaths.h` (public, beside `resolveDir`/`sanitize`):
```cpp
    // Auto-derived save path: resolveDir(mode,lastUsedDir) + "/" + sanitize(label) + ext.
    // Single source of truth for sensory/detailed silent auto-save filenames
    // (replaces the Save-As dialog). `ext` may be given with or without a leading dot.
    // An empty/whitespace label falls back to "untitled" so the path is always valid.
    static QString autoSavePath(ReportMode mode, const QString& sessionLabel,
                                const QString& lastUsedDir, const QString& ext);
```

- [ ] **Step 4: Implement** in `src/utils/OutputPaths.cpp`:
```cpp
QString OutputPaths::autoSavePath(ReportMode mode, const QString& sessionLabel,
                                  const QString& lastUsedDir, const QString& ext)
{
    const QString dir = resolveDir(mode, lastUsedDir);
    QString base = sanitize(sessionLabel);
    if (base.trimmed().isEmpty()) base = QStringLiteral("untitled");
    QString e = ext;
    if (!e.isEmpty() && !e.startsWith(QLatin1Char('.'))) e.prepend(QLatin1Char('.'));
    return dir + QLatin1Char('/') + base + e;
}
```
> If `sanitize` already trims/handles empties, keep the explicit fallback anyway (defensive). If `resolveDir` returns a path with a trailing slash, adjust the join to avoid `//` (read it; most resolvers return no trailing slash).

- [ ] **Step 5: Run to verify it passes.**
Run: `tests\run-tests.ps1 -Filter tst_outputpaths`
Expected: PASS.

- [ ] **Step 6: Commit.**
```bash
git add src/utils/OutputPaths.h src/utils/OutputPaths.cpp tests/tst_outputpaths/tst_outputpaths.cpp
git commit -m "feat(paths): autoSavePath helper for silent auto-named saves"
```

---

## Task 2: Pure `isSavable` predicates + unit tests

**Files:**
- Modify: `src/pipeline/SensoryData.h` / `.cpp`, `src/pipeline/DetailedSensoryData.h` / `.cpp`
- Test: `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp`, `tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp`

> Read how the tester+round are stored. Sensory folds tester+round into `SensorySession::testerName` via `combineTesterRound` (`src/ui/TesterRound.h`), so the predicate must strip the round with `splitTesterRound(...).tester` before checking emptiness. Detailed has no round — `testerName` is the raw tester. Confirm `splitTesterRound`'s return shape (a struct with a `.tester` field) and include `TesterRound.h` in `SensoryData.cpp` if needed.

- [ ] **Step 1: Write failing sensory test.** In `tst_sensorydataplaceholder.cpp`:
```cpp
void isSensorySavable_requiresTitleAndTester()
{
    DVE::SensorySession s;
    s.testTitle = "";  s.testerName = DVE::combineTesterRound("Alice", "1");
    QVERIFY(!DVE::isSensorySessionSavable(s));          // no title
    s.testTitle = "T"; s.testerName = DVE::combineTesterRound("", "1");
    QVERIFY(!DVE::isSensorySessionSavable(s));          // no tester (round only)
    s.testTitle = "T"; s.testerName = DVE::combineTesterRound("Alice", "1");
    QVERIFY(DVE::isSensorySessionSavable(s));           // both present
    s.testTitle = "   "; s.testerName = DVE::combineTesterRound("Alice", "1");
    QVERIFY(!DVE::isSensorySessionSavable(s));          // whitespace-only title
}
```
> Confirm `combineTesterRound`/`splitTesterRound` are accessible from the test (include `ui/TesterRound.h`). If `combineTesterRound("", "1")` yields a non-empty string like "1", the predicate's `splitTesterRound(...).tester` must still come back empty — verify and adjust the test's expectation to match real semantics while keeping the intent (empty tester ⇒ not savable).

- [ ] **Step 2: Write failing detailed test.** In `tst_detailedsensoryjson.cpp`:
```cpp
void isDetailedSavable_requiresTitleAndTester()
{
    DVE::DetailedSensorySession s;
    s.testTitle = "";  s.testerName = "Alice"; QVERIFY(!DVE::isDetailedSessionSavable(s));
    s.testTitle = "T"; s.testerName = "";      QVERIFY(!DVE::isDetailedSessionSavable(s));
    s.testTitle = "T"; s.testerName = "Alice"; QVERIFY(DVE::isDetailedSessionSavable(s));
}
```

- [ ] **Step 3: Run both — verify FAIL (undeclared).**
Run: `tests\run-tests.ps1 -Filter tst_sensorydataplaceholder` and `-Filter tst_detailedsensoryjson`.

- [ ] **Step 4: Declare + implement.** In `src/pipeline/SensoryData.h` (namespace DVE):
```cpp
// True iff the session has the non-empty test title AND tester it needs for a
// valid natural key (session_name, tester_name, date). Round is stripped before
// the tester check. Pure. Used to gate saves (block interactive / skip background).
bool isSensorySessionSavable(const SensorySession& s);
```
`src/pipeline/SensoryData.cpp` (add `#include "ui/TesterRound.h"` — verify the include path; it may be `"TesterRound.h"` via the .pro include paths):
```cpp
bool isSensorySessionSavable(const SensorySession& s)
{
    if (s.testTitle.trimmed().isEmpty()) return false;
    const QString tester = splitTesterRound(s.testerName).tester;
    return !tester.trimmed().isEmpty();
}
```
`src/pipeline/DetailedSensoryData.h`:
```cpp
// True iff the detailed session has a non-empty test title AND tester (its
// natural-key columns). Pure. Gates saves.
bool isDetailedSessionSavable(const DetailedSensorySession& s);
```
`src/pipeline/DetailedSensoryData.cpp`:
```cpp
bool isDetailedSessionSavable(const DetailedSensorySession& s)
{
    return !s.testTitle.trimmed().isEmpty() && !s.testerName.trimmed().isEmpty();
}
```

- [ ] **Step 5: Run both — verify PASS.**

- [ ] **Step 6: Commit.**
```bash
git add src/pipeline/SensoryData.h src/pipeline/SensoryData.cpp src/pipeline/DetailedSensoryData.h src/pipeline/DetailedSensoryData.cpp tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp
git commit -m "feat(sensory): pure isSavable predicates (title+tester required for natural key)"
```

---

## Task 3: Sensory `save()` — auto-path + hard require-fields guard

**Files:** Modify `src/ui/SensoryPanel.cpp` (and `.h` only if a small helper is added). Test: build + manual (GUI path).

> Read `SensoryPanel::save()` (~1474-1547), `resolveTestName()` (~1435-1468), `sessionLabel()` (~1335), `buildSession()` (~911), and how `m_savePath` feeds `saveToJson`/`saveToExcel` (~1546/1557). Identify the Save-As block (~1488-1505) and the `m_testTitleEdit`/`m_testerEdit`/`m_roundCombo` members.

- [ ] **Step 1: Add the hard guard at the top of `save()`.** Before any save work:
```cpp
    // Required fields for a valid natural key + clean filename. No Save-As fallback,
    // no auto-name: the user must name the test and tester.
    if (m_testTitleEdit->text().trimmed().isEmpty()
        || m_testerEdit->text().trimmed().isEmpty()) {
        QMessageBox::warning(this, tr("Test Name and Tester Required"),
            tr("Enter a test name and a tester before saving — both are needed "
               "to store and find this session reliably."));
        if (m_testTitleEdit->text().trimmed().isEmpty()) m_testTitleEdit->setFocus();
        else m_testerEdit->setFocus();
        return;
    }
```
Then DELETE the `resolveTestName()` call and its blank-title soft-prompt branch (the hard guard replaces it). If `resolveTestName()` is now unused, remove the method too (and its declaration) — confirm no other caller via grep.

- [ ] **Step 2: Replace the Save-As block with the auto-derived path.** Where `save()` currently builds `needsSaveAs` / calls `QFileDialog::getSaveFileName`, replace with:
```cpp
    // Auto-derive the on-disk path from the session label; no dialog, ever.
    const SensorySession cur = buildSession();
    m_savePath = OutputPaths::autoSavePath(ReportMode::Sensory, sessionLabel(cur),
                                           m_lastBrowseDir, QString());
```
> CRITICAL: match how `m_savePath` is consumed. If `saveToJson`/`saveToExcel` append their own extensions to `m_savePath` (base path), pass `QString()` ext (as above) so you don't double-append. If they expect a full path with extension, pass the right ext and call them accordingly. Read those two methods and wire it so the produced files are exactly what the old dialog default would have produced (test - tester - round). Add `#include "utils/OutputPaths.h"` if missing. Keep `setLastBrowseDir(...)`/`m_lastBrowseDir` update behavior consistent (the dir is now the resolved default; do not overwrite `m_lastBrowseDir` with the auto dir if that would drift the user's browse history — prefer leaving `m_lastBrowseDir` as-is and only using it as the resolver hint).

- [ ] **Step 3: Build the app** (incremental release in `build\`); confirm warning-free under `-Werror` and that `save()` no longer references `QFileDialog`.
Run: `python tools\decrypt_via_copy.py --apply` then `mingw32-make -j8` in `build` (qmake first if needed).

- [ ] **Step 4: Commit.**
```bash
git add src/ui/SensoryPanel.cpp src/ui/SensoryPanel.h
git commit -m "feat(sensory): silent auto-named save + required test-name/tester guard (no Save-As)"
```

---

## Task 4: Detailed `save()` — auto-path + hard require-fields guard

**Files:** Modify `src/ui/DetailedSensoryPanel.cpp` (and `.h` if needed).

> Read `DetailedSensoryPanel::save()` (~1128-1160), `sessionLabel()` (~1006), `buildSession()` (~1017), the Save-As block (~1135-1144), and `m_testTitleEdit`/`m_testerEdit` members. Detailed has no round and writes `.xlsx` only.

- [ ] **Step 1: Hard guard at top of `save()`** (mirror Task 3):
```cpp
    if (m_testTitleEdit->text().trimmed().isEmpty()
        || m_testerEdit->text().trimmed().isEmpty()) {
        QMessageBox::warning(this, tr("Test Name and Tester Required"),
            tr("Enter a test name and a tester before saving — both are needed "
               "to store and find this session reliably."));
        if (m_testTitleEdit->text().trimmed().isEmpty()) m_testTitleEdit->setFocus();
        else m_testerEdit->setFocus();
        return;
    }
```

- [ ] **Step 2: Replace the Save-As block** with:
```cpp
    const DetailedSensorySession cur = buildSession();
    m_savePath = OutputPaths::autoSavePath(ReportMode::DetailedSensory, sessionLabel(cur),
                                           m_lastBrowseDir, QStringLiteral(".xlsx"));
```
> Adapt to how `m_savePath`/`saveToExcel` consume the path (full path vs base). Detailed's filename base must be `sessionLabel` (test - tester), NOT the current `testTitle`-only fallback (~1136). Add `#include "utils/OutputPaths.h"` if missing.

- [ ] **Step 3: Build** (incremental, warning-free under -Werror; no `QFileDialog` left in `save()`).

- [ ] **Step 4: Commit.**
```bash
git add src/ui/DetailedSensoryPanel.cpp src/ui/DetailedSensoryPanel.h
git commit -m "feat(detailed-sensory): silent auto-named save + required test-name/tester guard (no Save-As)"
```

---

## Task 5: MainWindow — background skip / interactive summary / close guard

**Files:** Modify `src/MainWindow.cpp` (`onUpdateDatabase` ~4522; `saveSensorySessionsBeforeClose` ~2403; `saveDetailedSensorySessionsBeforeClose` ~2442; `promptSaveDatabase` ~5139).

> Read `onUpdateDatabase` (sensory loop ~4576-4654, detailed loop ~4656-4720), the two close helpers, and `promptSaveDatabase` (~5177-5185). The signature is `onUpdateDatabase(bool flushPending)` — `true` = interactive (Ctrl+U, program-close), `false` = background 5 s autosave.

- [ ] **Step 1: Gate the `onUpdateDatabase` per-session loops.** In BOTH loops, after the existing `isPlaceholderSession` skip, add a savability gate. Use `DVE::isSensorySessionSavable(sess)` / `DVE::isDetailedSessionSavable(sess)`:
```cpp
        if (!DVE::isSensorySessionSavable(sess)) {
            if (flushPending) incompleteNames << sessionDisplayName(sess);  // collect for one summary
            continue;                                                       // never persist an unkeyed session
        }
```
Declare `QStringList incompleteNames;` before the loops (shared across sensory+detailed) and a small local `sessionDisplayName(...)` (or inline `sess.testTitle.isEmpty() ? tr("(unnamed)") : sess.testTitle`). After BOTH loops, if `flushPending && !incompleteNames.isEmpty()`, show ONE message:
```cpp
    if (flushPending && !incompleteNames.isEmpty()) {
        QMessageBox::information(this, tr("Some Sessions Not Saved"),
            tr("These sessions need a test name and tester before they can be saved:\n\n  - %1")
                .arg(incompleteNames.join(QStringLiteral("\n  - "))));
    }
```
Place the message before the existing success/failure status update so counts stay accurate. The background path (`flushPending==false`) just `continue;`s silently — no nag, no UI block.

- [ ] **Step 2: Gate the per-item Close helpers.** In `saveSensorySessionsBeforeClose`, for each non-placeholder index, if `!isSensorySessionSavable(sess)` → it must NOT be closed (would lose unkeyed data): append to `failed`, and collect a name to instruct once after the loop:
```cpp
        if (!DVE::isSensorySessionSavable(sess)) { failed.append(idx); needName << idx; continue; }
```
After the loop, if `needName` non-empty, `QMessageBox::warning(this, tr("Name Required to Close"), tr("Add a test name and tester to the session(s) you're closing — they can't be saved (or safely closed) without them."));`. Mirror in `saveDetailedSensorySessionsBeforeClose`. (A placeholder/empty session is already skipped by the existing `isPlaceholderSession` check and closes freely — no nag.)

- [ ] **Step 3: Gate the program-close disk-courtesy `save()` calls.** In `promptSaveDatabase` (~5177-5183), only call the panel `save()` when the current session is savable, so app shutdown never hits the new hard-guard modal:
```cpp
    if (m_sensorySessionsDirty && m_sensoryPanel && !m_sensoryPanel->hasSavePath()
        && m_sensoryPanel->currentSessionSavable())   // new tiny panel accessor: title+tester non-empty
        m_sensoryPanel->save();
    if (m_detailedSensorySessionsDirty && m_detailedSensoryPanel && !m_detailedSensoryPanel->hasSavePath()
        && m_detailedSensoryPanel->currentSessionSavable())
        m_detailedSensoryPanel->save();
    onUpdateDatabase(/*flushPending=*/true);   // already skips+summarizes incomplete sessions
```
Add a trivial public `bool currentSessionSavable() const` to each panel: `return isSensorySessionSavable(buildSession());` (or read the two edits). Incomplete sessions are then neither disk-saved nor DB-saved on close, but survive via the existing recovery snapshot (`m_recovery->flushNow(true)` in `closeEvent`), so no data loss — they reappear on next launch to be named.

- [ ] **Step 4: Build** the app (incremental release, warning-free under -Werror). Run `tests\run-tests.ps1 -Filter tst_mainwindow_remotecell` to confirm the harness still builds/passes.

- [ ] **Step 5: Commit.**
```bash
git add src/MainWindow.cpp src/ui/SensoryPanel.h src/ui/SensoryPanel.cpp src/ui/DetailedSensoryPanel.h src/ui/DetailedSensoryPanel.cpp
git commit -m "feat(persist): require test-name+tester to save; background auto-save skips incomplete sessions"
```

---

## Task 6: Regression sweep, version bump v2.3.5, build, installer, Plane

**Files:** `DataViewerEnterprise.pro`.

- [ ] **Step 1: Regression sweep.** Full Qt suite: `tests\run-tests.ps1` (test PG `dve-test-pg` running). Confirm the new unit tests pass (`tst_outputpaths`, `tst_sensorydataplaceholder`, `tst_detailedsensoryjson`) and no NEW failures beyond the known pre-existing ones (`sensoryHeaderPresets_roundTrip`, `tst_storedfns` probe, `tst_responsivelayout` segfault, the stale `tst_excelexporter` artifact — all unrelated init.sql-drift / pre-existing).

- [ ] **Step 2: Bump VERSION** in `DataViewerEnterprise.pro` from `2.3.4` to `2.3.5`.

- [ ] **Step 3: Clean release build** (VERSION change needs clean rebuild):
```
python tools\decrypt_via_copy.py --apply
qmake CONFIG+=release   ::(canonical: in repo root, DESTDIR=release\)
mingw32-make clean && mingw32-make -j8
```
Expected: `release\DataViewer.exe` FileVersion = `2.3.5.0`, warning-free under -Werror.

- [ ] **Step 4: Build the installer** (do NOT touch Synology):
```
tools\prepare_python_embed.bat   :: only if release\python_bundle.zip absent
build_installer.bat              :: -> dist\DataViewer-setup.exe (version guard must pass at 2.3.5)
```

- [ ] **Step 5: Commit the version bump.**
```bash
git add DataViewerEnterprise.pro
git commit -m "chore(release): v2.3.5 internal -- zero Save-As + required test-name/tester guardrail"
```

- [ ] **Step 6: Plane.** Set the tracking issue to *Ready for Release* with a resolution comment (root cause: Save-As dialogs + no required-field validation caused filename mismatches and `('','',date)` key collisions; fix: `OutputPaths::autoSavePath` silent auto-naming in both sensory modes + `isSavable` predicates gating interactive-block / background-skip / close-refuse; verification: unit tests + full suite + app build + installer). Internal v2.3.5; wraps into v2.4.0.

---

## Risks & notes

- **No Save-As anywhere in sensory/detailed.** The three dialog sites are removed; the path is deterministic. Same test+tester(+round) → same file → overwrite (intended dedup). Cross-day re-runs reuse the filename but create a distinct DB row (date in the key) — pre-existing, flagged.
- **Required fields are a correctness fix, not just UX:** the DB unique index is `(session_name, tester_name, date)`; empty title+tester collapses to `('','',date)` and trips `UniqueViolation`. Blocking empties prevents that whole failure class.
- **Background auto-save never blocks.** Only `flushPending==true` (Ctrl+U / program-close) surfaces the one-shot "needs a name" summary; the 5 s timer silently `continue`s past incomplete sessions.
- **No data loss for incomplete sessions:** they aren't persisted until named, but the existing recovery snapshot (`closeEvent` → `m_recovery->flushNow(true)`) captures in-memory state, so they reappear next launch.
- **Removed feature:** Sensory's `resolveTestName`/`nextDefaultTestName` auto-name-on-blank is gone (superseded by the hard requirement). Confirm no other caller of `nextDefaultTestName` before deleting it.
- **TPM mode untouched.** Its write-back is to the source workbook, not a Save-As; out of scope.
