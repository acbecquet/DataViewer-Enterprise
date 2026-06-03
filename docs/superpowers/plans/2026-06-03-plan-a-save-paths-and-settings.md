# Plan A — Save-Path Defaults + Per-Mode Output Settings — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Reports save to a predictable, configurable folder with a consistent `<file>_report.pptx` name; every save/export dialog defaults to the user's Documents folder; a new **Settings** ribbon tab sets three per-mode output folders that persist across restarts.

**Architecture:** A new `src/utils/OutputPaths` helper is the single source of truth for the three persisted per-mode folders, the directory-precedence resolver (per-mode → last-used → Documents), and the `<base>_report.pptx` filename builder + sanitizer. Every save dialog is repointed at this helper, retiring the three divergent `lastBrowseDir()` fallbacks (and the hard-coded OneDrive path). A new `buildSettingsTab()` adds the ribbon tab.

**Tech Stack:** C++17 / Qt 6.10 (`QSettings`, `QStandardPaths`, `QFileDialog`, `QDir`/`QFileInfo`), namespace `DVE`, qmake + MinGW, `-Werror -Wall -Wextra`. Unit tests via `tests/tests.pro` SUBDIRS + QtTest.

**Design spec:** `docs/superpowers/specs/2026-06-03-plan-a-save-paths-and-settings-design.md`

---

## Conventions for EVERY build/test step (read once)

- **MIP decrypt before any C++ build/test:** run `python tools/decrypt_via_copy.py --apply` from the repo root first (idempotent).
- **Create new source files via Python delete-and-rewrite** (so they don't inherit MIP labels). Pattern:
  ```python
  import os
  path = "src/utils/OutputPaths.h"
  content = "..."   # full file contents
  if os.path.exists(path): os.remove(path)
  with open(path, "w", encoding="utf-8", newline="\n") as f: f.write(content)
  ```
- **Commit immediately after writing a new file** so git captures plaintext before MIP re-labels; verify with `git show HEAD:<path> | head -3`.
- **Run the unit suite** with `tests\run-tests.ps1` (auto-detects Qt + MinGW, builds incrementally). It compiles every SUBDIRS suite and runs them.
- **Compile the full app** (for UI tasks with no unit test) with the documented debug build:
  ```bat
  cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake -spec win32-g++ DataViewerEnterprise.pro && mingw32-make -j8"
  ```
  A clean `-Werror` compile is the gate. Watch for unused-parameter/unused-include warnings.
- Work happens on branch `dev` (current). No VERSION bump until the Release section at the end.

---

## Task 1: `OutputPaths` helper + unit tests (TDD)

**Files:**
- Create: `src/utils/OutputPaths.h`, `src/utils/OutputPaths.cpp`
- Create: `tests/tst_outputpaths/tst_outputpaths.pro`, `tests/tst_outputpaths/tst_outputpaths.cpp`
- Modify: `DataViewerEnterprise.pro`, `tests/tests.pro`

- [ ] **Step 1: Write the failing test** — create `tests/tst_outputpaths/tst_outputpaths.cpp` (Python delete-and-rewrite):

```cpp
#include <QtTest>
#include <QTemporaryDir>
#include <QStandardPaths>
#include <QSettings>
#include "utils/OutputPaths.h"

using namespace DVE;

class TestOutputPaths : public QObject
{
    Q_OBJECT
private slots:
    void initTestCase();
    void sanitize_strips_illegal();
    void sanitize_collapses_and_trims();
    void reportFileName_base_only();
    void reportFileName_with_sheet();
    void reportFileName_strips_existing_ext();
    void reportFileName_empty_base_fallback();
    void firstExistingDir_picks_first();
    void firstExistingDir_skips_missing_and_empty();
    void firstExistingDir_returns_fallback();
    void resolveReportDir_prefers_configured();
    void resolveReportDir_falls_to_lastused_when_unset();
    void resolveSaveDir_lastUsed_then_documents();
    void documentsDir_nonempty();
};

void TestOutputPaths::initTestCase()
{
    // Hermetic settings: redirect QSettings to a temp INI instead of the registry.
    QStandardPaths::setTestModeEnabled(true);
    QSettings::setDefaultFormat(QSettings::IniFormat);
}

void TestOutputPaths::sanitize_strips_illegal()
{
    QCOMPARE(OutputPaths::sanitize(QStringLiteral("a/b\\c:d*e?f\"g<h>i|j")),
             QStringLiteral("abcdefghij"));
}

void TestOutputPaths::sanitize_collapses_and_trims()
{
    QCOMPARE(OutputPaths::sanitize(QStringLiteral("  Foo   Bar  ")), QStringLiteral("Foo_Bar"));
    QCOMPARE(OutputPaths::sanitize(QStringLiteral("__x__")), QStringLiteral("x"));
}

void TestOutputPaths::reportFileName_base_only()
{
    QCOMPARE(OutputPaths::reportFileName(QStringLiteral("Foo")), QStringLiteral("Foo_report.pptx"));
}

void TestOutputPaths::reportFileName_with_sheet()
{
    QCOMPARE(OutputPaths::reportFileName(QStringLiteral("Foo"), QStringLiteral("Sheet 1")),
             QStringLiteral("Foo_Sheet_1_report.pptx"));
}

void TestOutputPaths::reportFileName_strips_existing_ext()
{
    QCOMPARE(OutputPaths::reportFileName(QStringLiteral("Foo.pptx")), QStringLiteral("Foo_report.pptx"));
}

void TestOutputPaths::reportFileName_empty_base_fallback()
{
    QCOMPARE(OutputPaths::reportFileName(QString()), QStringLiteral("untitled_report.pptx"));
}

void TestOutputPaths::firstExistingDir_picks_first()
{
    QTemporaryDir a, b;
    QCOMPARE(OutputPaths::firstExistingDir({a.path(), b.path()}, QStringLiteral("/fallback")),
             a.path());
}

void TestOutputPaths::firstExistingDir_skips_missing_and_empty()
{
    QTemporaryDir real;
    const QString missing = real.path() + QStringLiteral("/nope");
    QCOMPARE(OutputPaths::firstExistingDir({QString(), missing, real.path()},
                                           QStringLiteral("/fallback")),
             real.path());
}

void TestOutputPaths::firstExistingDir_returns_fallback()
{
    QCOMPARE(OutputPaths::firstExistingDir({QString(), QStringLiteral("/nope")},
                                           QStringLiteral("/fallback")),
             QStringLiteral("/fallback"));
}

void TestOutputPaths::resolveReportDir_prefers_configured()
{
    QTemporaryDir cfg, last;
    OutputPaths::setConfiguredDir(ReportMode::Tpm, cfg.path());
    QCOMPARE(OutputPaths::resolveReportDir(ReportMode::Tpm, last.path()), cfg.path());
}

void TestOutputPaths::resolveReportDir_falls_to_lastused_when_unset()
{
    QTemporaryDir last;
    OutputPaths::setConfiguredDir(ReportMode::Sensory, QString());   // unset
    QCOMPARE(OutputPaths::resolveReportDir(ReportMode::Sensory, last.path()), last.path());
}

void TestOutputPaths::resolveSaveDir_lastUsed_then_documents()
{
    QTemporaryDir last;
    QCOMPARE(OutputPaths::resolveSaveDir(last.path()), last.path());
    QCOMPARE(OutputPaths::resolveSaveDir(QString()), OutputPaths::documentsDir());
}

void TestOutputPaths::documentsDir_nonempty()
{
    QVERIFY(!OutputPaths::documentsDir().isEmpty());
}

QTEST_MAIN(TestOutputPaths)
#include "tst_outputpaths.moc"
```

- [ ] **Step 2: Create the test .pro** — `tests/tst_outputpaths/tst_outputpaths.pro` (Python delete-and-rewrite). Mirror an existing sibling (e.g. `tests/tst_sensorydataplaceholder/*.pro`) for the `testcase` CONFIG and any shared include; the helper only needs Qt core:

```pro
QT       += core testlib
QT       -= gui
CONFIG   += console c++17 testcase
CONFIG   -= app_bundle
TEMPLATE  = app
TARGET    = tst_outputpaths

INCLUDEPATH += ../../src ../../src/utils

SOURCES += tst_outputpaths.cpp \
           ../../src/utils/OutputPaths.cpp
HEADERS += ../../src/utils/OutputPaths.h
```

- [ ] **Step 3: Register the suite** — add `tst_outputpaths` to the `SUBDIRS` list in `tests/tests.pro` (read the file, append the entry following the existing formatting).

- [ ] **Step 4: Run the test to confirm it FAILS** — `python tools/decrypt_via_copy.py --apply` then `tests\run-tests.ps1`. Expected: `tst_outputpaths` fails to compile (no `OutputPaths` symbol). This proves the suite is wired in.

- [ ] **Step 5: Create the header** — `src/utils/OutputPaths.h` (Python delete-and-rewrite):

```cpp
#pragma once

#include <QString>
#include <QStringList>

namespace DVE {

enum class ReportMode { Tpm, Sensory, DetailedSensory };

// Single source of truth for report/save output directories + report filenames.
// Pure resolvers (firstExistingDir / resolve* / reportFileName / sanitize) are
// side-effect-free; only configuredDir/setConfiguredDir touch QSettings.
class OutputPaths {
public:
    // Persisted per-mode output folders. QSettings("SDR","DataViewerEnterprise"),
    // keys output/tpmDir, output/sensoryDir, output/detailedDir. "" when unset.
    static QString configuredDir(ReportMode mode);
    static void    setConfiguredDir(ReportMode mode, const QString& dir);

    // QStandardPaths Documents, falling back to ~/Documents.
    static QString documentsDir();

    // Report dialogs: configuredDir(mode) -> lastUsedDir -> Documents (first existing).
    static QString resolveReportDir(ReportMode mode, const QString& lastUsedDir);
    // Non-report dialogs: lastUsedDir -> Documents.
    static QString resolveSaveDir(const QString& lastUsedDir);

    // "Foo" -> "Foo_report.pptx"; ("Foo","Sheet 1") -> "Foo_Sheet_1_report.pptx".
    // A trailing ".pptx" on base is stripped; empty base -> "untitled".
    static QString reportFileName(const QString& base, const QString& sheet = QString());

    // Strip Windows-illegal chars (\ / : * ? " < > |), whitespace -> '_',
    // collapse runs of '_', trim leading/trailing '_'.
    static QString sanitize(const QString& raw);

    // First entry that is non-empty AND exists on disk; else fallback. Pure.
    static QString firstExistingDir(const QStringList& candidates, const QString& fallback);

private:
    static QString settingsKey(ReportMode mode);
};

} // namespace DVE
```

- [ ] **Step 6: Create the implementation** — `src/utils/OutputPaths.cpp` (Python delete-and-rewrite):

```cpp
#include "utils/OutputPaths.h"

#include <QSettings>
#include <QStandardPaths>
#include <QDir>

namespace DVE {

QString OutputPaths::settingsKey(ReportMode mode)
{
    switch (mode) {
    case ReportMode::Tpm:             return QStringLiteral("output/tpmDir");
    case ReportMode::Sensory:         return QStringLiteral("output/sensoryDir");
    case ReportMode::DetailedSensory: return QStringLiteral("output/detailedDir");
    }
    return QStringLiteral("output/tpmDir");
}

QString OutputPaths::configuredDir(ReportMode mode)
{
    QSettings s(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
    return s.value(settingsKey(mode)).toString();
}

void OutputPaths::setConfiguredDir(ReportMode mode, const QString& dir)
{
    QSettings s(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
    s.setValue(settingsKey(mode), dir);
}

QString OutputPaths::documentsDir()
{
    const QString docs = QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation);
    if (!docs.isEmpty() && QDir(docs).exists())
        return docs;
    return QDir::homePath() + QStringLiteral("/Documents");
}

QString OutputPaths::firstExistingDir(const QStringList& candidates, const QString& fallback)
{
    for (const QString& c : candidates) {
        if (!c.isEmpty() && QDir(c).exists())
            return c;
    }
    return fallback;
}

QString OutputPaths::resolveReportDir(ReportMode mode, const QString& lastUsedDir)
{
    return firstExistingDir({ configuredDir(mode), lastUsedDir }, documentsDir());
}

QString OutputPaths::resolveSaveDir(const QString& lastUsedDir)
{
    return firstExistingDir({ lastUsedDir }, documentsDir());
}

QString OutputPaths::sanitize(const QString& raw)
{
    static const QString illegal = QStringLiteral("\\/:*?\"<>|");
    QString out;
    out.reserve(raw.size());
    for (const QChar ch : raw) {
        if (illegal.contains(ch)) continue;
        out.append(ch.isSpace() ? QChar('_') : ch);
    }
    while (out.contains(QStringLiteral("__")))
        out.replace(QStringLiteral("__"), QStringLiteral("_"));
    while (out.startsWith(QLatin1Char('_'))) out.remove(0, 1);
    while (out.endsWith(QLatin1Char('_')))   out.chop(1);
    return out;
}

QString OutputPaths::reportFileName(const QString& base, const QString& sheet)
{
    QString b = base;
    if (b.endsWith(QStringLiteral(".pptx"), Qt::CaseInsensitive))
        b.chop(5);
    QString stem = sanitize(b);
    if (stem.isEmpty()) stem = QStringLiteral("untitled");
    if (!sheet.isEmpty()) {
        const QString s = sanitize(sheet);
        if (!s.isEmpty()) stem += QStringLiteral("_") + s;
    }
    return stem + QStringLiteral("_report.pptx");
}

} // namespace DVE
```

- [ ] **Step 7: Register in the app .pro** — in `DataViewerEnterprise.pro`, add `src/utils/OutputPaths.cpp` to `SOURCES` and `src/utils/OutputPaths.h` to `HEADERS` (next to the other `src/utils/*` entries).

- [ ] **Step 8: Run the test to confirm it PASSES** — `tests\run-tests.ps1`. Expected: `tst_outputpaths` all pass.

- [ ] **Step 9: Commit**
```bash
git add src/utils/OutputPaths.h src/utils/OutputPaths.cpp \
        tests/tst_outputpaths/ DataViewerEnterprise.pro tests/tests.pro
git commit -m "feat(utils): OutputPaths resolver + report-name builder with unit tests"
```

---

## Task 2: Settings ribbon tab

**Files:** Modify `src/MainWindow.h`, `src/MainWindow.cpp`. No unit test (UI/registry) — gate is a clean compile + manual check.

- [ ] **Step 1: Declare the build method** — in `src/MainWindow.h`, next to the existing `build*Tab` declarations, add:
```cpp
void buildSettingsTab(RibbonTab* tab);
```

- [ ] **Step 2: Implement `buildSettingsTab`** — in `src/MainWindow.cpp` (near `buildToolsTab`). Ensure includes exist at top: `#include "utils/OutputPaths.h"`, `<QFileDialog>`, `<QToolButton>`, `"utils/AppTheme.h"`, `"widgets/RibbonWidget.h"` (most already present):
```cpp
void MainWindow::buildSettingsTab(RibbonTab* tab)
{
    RibbonGroup* grp = tab->addGroup(QStringLiteral("Output Paths"));

    struct PathBtn { QString label; QString title; ReportMode mode; };
    const QVector<PathBtn> defs = {
        { QStringLiteral("Set TPM Output Path"),
          QStringLiteral("Select TPM Report Output Folder"),              ReportMode::Tpm },
        { QStringLiteral("Set Sensory Output Path"),
          QStringLiteral("Select Sensory Report Output Folder"),          ReportMode::Sensory },
        { QStringLiteral("Set Detailed Sensory Output Path"),
          QStringLiteral("Select Detailed Sensory Report Output Folder"), ReportMode::DetailedSensory },
    };

    auto tip = [](ReportMode m) -> QString {
        const QString d = OutputPaths::configuredDir(m);
        return d.isEmpty() ? QStringLiteral("Not set — defaults to Documents") : d;
    };

    for (const PathBtn& d : defs) {
        QToolButton* btn = grp->addLargeButton(d.label,
                                               AppTheme::icon(QStringLiteral("folder-open")),
                                               tip(d.mode));
        const ReportMode mode = d.mode;
        const QString title = d.title;
        connect(btn, &QToolButton::clicked, this, [this, btn, mode, title, tip]() {
            const QString cur = OutputPaths::configuredDir(mode);
            const QString start = cur.isEmpty() ? OutputPaths::documentsDir() : cur;
            const QString dir = QFileDialog::getExistingDirectory(this, title, start);
            if (!dir.isEmpty()) {
                OutputPaths::setConfiguredDir(mode, dir);
                btn->setToolTip(tip(mode));
            }
        });
    }
}
```

- [ ] **Step 3: Register the tab** — in `MainWindow::setupRibbon()`, immediately AFTER the `buildToolsTab(m_ribbon->addTab("Tools"));` line, add:
```cpp
buildSettingsTab(m_ribbon->addTab(QStringLiteral("Settings")));
```

- [ ] **Step 4: Compile the full app** (decrypt first; full `-Werror` build per Conventions). Expected: clean build.

- [ ] **Step 5: Manual check** — launch, confirm a **Settings** tab sits to the right of **Tools** with three word-wrapped buttons; clicking one opens a folder picker; after choosing, hover shows the path in the tooltip.

- [ ] **Step 6: Commit**
```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(ui): Settings ribbon tab with three per-mode output-path buttons"
```

---

## Task 3: Repoint TPM report save sites + retire MainWindow fallback

**Files:** Modify `src/MainWindow.cpp` (+`.h` if needed). Add `#include "utils/OutputPaths.h"` if absent.

For each site: read the current `QFileDialog::getSaveFileName` / `getExistingDirectory` call, replace its **directory** argument with the resolver, and set its **filename** via `reportFileName`. Use `m_lastBrowseDir` as the `lastUsedDir` argument. Derive the file base from `QFileInfo(file->filePath).completeBaseName()`.

- [ ] **Step 1: `onGenerateTestReport`** (≈ MainWindow.cpp:3192, TPM Test/per-sheet). Directory → `OutputPaths::resolveReportDir(ReportMode::Tpm, m_lastBrowseDir)`. Filename → `OutputPaths::reportFileName(QFileInfo(file->filePath).completeBaseName(), sheetName)` (keeps the sheet segment). Keep the post-save `setLastBrowseDir(...)`.

- [ ] **Step 2: `onGenerateFullReport`** (≈ MainWindow.cpp:3232). Single-file branch (≈:3286): dir → `resolveReportDir(Tpm, m_lastBrowseDir)`, name → `reportFileName(QFileInfo(files.first().filePath).completeBaseName())`. Multi-file branch (`getExistingDirectory` ≈:3298): dir seed → `resolveReportDir(Tpm, m_lastBrowseDir)`; per-file output name → `reportFileName(QFileInfo(f.filePath).completeBaseName())`; combined output name → `"Combined_" + QDate::currentDate().toString("yyyy-MM-dd") + "_report.pptx"`.

- [ ] **Step 3: Retire the fallback in `MainWindow::lastBrowseDir`** (≈:4439). Replace the body so it returns `m_lastBrowseDir` when it exists, else `OutputPaths::documentsDir()`. Delete the `QDir::homePath()+"/Documents"` literal. `setLastBrowseDir` is unchanged.

- [ ] **Step 4: Compile** (full `-Werror` build). Expected: clean.

- [ ] **Step 5: Manual check** — with no per-mode folder set, both TPM report dialogs open at Documents (or last-used after a prior save) and pre-fill `<file>_<sheet>_report.pptx` / `<file>_report.pptx`. Set the TPM output folder in Settings → dialogs now open there.

- [ ] **Step 6: Commit**
```bash
git add src/MainWindow.cpp src/MainWindow.h
git commit -m "feat(reports): TPM report dialogs use OutputPaths (per-mode dir + _report name)"
```

---

## Task 4: `IReportSource::suggestedReportBaseName` + Sensory report (site 4)

**Files:** Modify `src/reporting/IReportSource.h`, `src/reporting/SensoryReportSource.{h,cpp}`, `src/ui/ReportPreviewDialog.cpp`, and any other `IReportSource` implementor.

- [ ] **Step 1: Add the interface method** — in `src/reporting/IReportSource.h`:
```cpp
// Sanitized base for the suggested "<base>_report.pptx" filename.
virtual QString suggestedReportBaseName() const = 0;
```

- [ ] **Step 2: Find every implementor** — `grep -rn "public IReportSource" src/`. For each (at minimum `SensoryReportSource`), implement the method. `SensoryReportSource::suggestedReportBaseName()` returns the sanitized single-session `testTitle`, else `"sensory"`:
```cpp
QString SensoryReportSource::suggestedReportBaseName() const
{
    const QString label = sourceLabel();           // single testTitle, else "N sessions"
    const QString base  = OutputPaths::sanitize(label);
    return base.isEmpty() ? QStringLiteral("sensory") : base;
}
```
(Add `#include "utils/OutputPaths.h"` to `SensoryReportSource.cpp`.)

- [ ] **Step 3: Repoint `ReportPreviewDialog::onCreateReport`** (≈ ReportPreviewDialog.cpp:606) — replace the hard-coded `"report.pptx"` with a resolved path. Add `#include "utils/OutputPaths.h"`:
```cpp
const QString dir  = OutputPaths::resolveReportDir(ReportMode::Sensory, m_lastSaveDir /*or ""*/);
const QString name = OutputPaths::reportFileName(m_source->suggestedReportBaseName());
const QString fileName = QFileDialog::getSaveFileName(
    this, tr("Save Report"), dir + QStringLiteral("/") + name,
    tr("PowerPoint Presentation (*.pptx)"));
```
(If `ReportPreviewDialog` has no last-used member, pass `QString()` — resolver falls to Documents.)

- [ ] **Step 4: Compile** (full `-Werror` build). The pure-virtual addition forces every implementor to compile; fix any the grep surfaced.

- [ ] **Step 5: Manual check** — Sensory report save dialog now opens at the Sensory folder (or Documents) with `<sessionTitle>_report.pptx`, not the process working dir.

- [ ] **Step 6: Commit**
```bash
git add src/reporting/IReportSource.h src/reporting/SensoryReportSource.h \
        src/reporting/SensoryReportSource.cpp src/ui/ReportPreviewDialog.cpp
git commit -m "feat(reports): Sensory report dialog uses OutputPaths + suggestedReportBaseName"
```

---

## Task 5: Detailed-Sensory report (site 5) + DB-browser combined (sites 6-7) + retire fallback

**Files:** Modify `src/ui/DetailedSensoryPanel.cpp`, `src/ui/DatabaseBrowserDialog.cpp`. Add `#include "utils/OutputPaths.h"` to each.

- [ ] **Step 1: `DetailedSensoryPanel::generateReport`** (≈:1387). Dir → `OutputPaths::resolveReportDir(ReportMode::DetailedSensory, m_lastBrowseDir)`. Name → `OutputPaths::reportFileName(sess.testTitle.isEmpty() ? QStringLiteral("detailed_sensory") : sess.testTitle)`.

- [ ] **Step 2: Retire `DetailedSensoryPanel::lastBrowseDir`** (≈:1552) — return `m_lastBrowseDir` if it exists, else `OutputPaths::documentsDir()`. Delete the `m_db->getSetting(...)` and bare `homePath()` fallbacks.

- [ ] **Step 3: DB-browser combined Sensory** (≈ DatabaseBrowserDialog.cpp:721). Dir → `OutputPaths::resolveReportDir(ReportMode::Sensory, QString())`. Name → `"Combined_Sensory_report.pptx"`.

- [ ] **Step 4: DB-browser combined Detailed** (≈:892). Dir → `OutputPaths::resolveReportDir(ReportMode::DetailedSensory, QString())`. Name → `"Combined_Detailed_Sensory_report.pptx"`.

- [ ] **Step 5: Compile** (full `-Werror` build). Expected: clean.

- [ ] **Step 6: Manual check** — Detailed report + both DB-browser combined exports open at the right per-mode folder with the expected names.

- [ ] **Step 7: Commit**
```bash
git add src/ui/DetailedSensoryPanel.cpp src/ui/DatabaseBrowserDialog.cpp
git commit -m "feat(reports): Detailed + DB-browser combined reports use OutputPaths"
```

---

## Task 6: Non-report save/export sites + retire SensoryPanel fallback

**Files:** Modify `src/ui/SensoryPanel.cpp`, `src/plotting/PlotWidget.cpp`. Add `#include "utils/OutputPaths.h"` to each.

- [ ] **Step 1: `SensoryPanel::save`** (≈:1459) — dir seed → `OutputPaths::resolveSaveDir(m_lastBrowseDir)`; keep the existing session-title filename.

- [ ] **Step 2: `SensoryPanel::onSaveChart`** (≈:1356) and **`SensoryPanel::generateStats`** (≈:2164) — dir seed → `OutputPaths::resolveSaveDir(m_lastBrowseDir)`; keep existing names.

- [ ] **Step 3: Retire `SensoryPanel::lastBrowseDir`** (≈:1380) — return `m_lastBrowseDir` if it exists, else `OutputPaths::documentsDir()`. **Delete the hard-coded OneDrive `Weekly_Reports_Transfer` path.**

- [ ] **Step 4: `PlotWidget::onSaveImage`** (≈:362) — give it a directory seed of `OutputPaths::resolveSaveDir(m_lastSaveDir)`; if `PlotWidget` has no last-used member, pass `QString()` (→ Documents). Keep the existing image filename logic.

- [ ] **Step 5: Compile** (full `-Werror` build). Expected: clean.

- [ ] **Step 6: Manual check** — session save (xlsx/json), radar PNG, stats CSV, and TPM plot image all default to Documents on a cold start; OneDrive path no longer appears.

- [ ] **Step 7: Commit**
```bash
git add src/ui/SensoryPanel.cpp src/plotting/PlotWidget.cpp
git commit -m "feat(io): non-report save dialogs default to Documents via OutputPaths"
```

---

## Task 7: Full verification

- [ ] **Step 1: Decrypt + full clean build** — `python tools/decrypt_via_copy.py --apply`, then a clean `-Werror` release build:
  ```bat
  cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake -spec win32-g++ DataViewerEnterprise.pro CONFIG+=release && mingw32-make clean && mingw32-make -j8"
  ```
  Expected: clean build, `release\DataViewer.exe` produced.

- [ ] **Step 2: Run the full unit suite** — `tests\run-tests.ps1`. Expected: all suites pass, including `tst_outputpaths`. (DB-dependent suites skip cleanly when `DVE_TEST_PG_CONN` is unset.)

- [ ] **Step 3: Manual end-to-end checklist** (no per-mode folders set, then set):
  - Settings tab right of Tools; three buttons; each persists a folder and survives an app restart.
  - TPM Test/Full, Sensory, Detailed, and both DB-browser combined reports each open at their mode's folder (or Documents) with the correct `<…>_report.pptx` name.
  - Session saves / chart PNG / stats CSV / plot image default to Documents; OneDrive path gone.

---

## Release (after all tasks pass + user approval)

Not a TDD task — performed once the implementation is verified:
1. Bump `VERSION` in `DataViewerEnterprise.pro` (patch → `2.2.4`) and write `release_overview/release_overview_v_2_2_4.txt` (commit the overview atomically so git captures plaintext).
2. Build the installer via the **rebuild-dataviewer** skill (qmake → `mingw32-make clean` → `make` → `build_installer.bat`); confirm `dist\DataViewer-setup.exe` reports `2.2.4`.
3. **User eyeball-tests** the installed build. Only after approval: merge `dev` → `main`, push. **User performs the Synology drop** (never automated).
