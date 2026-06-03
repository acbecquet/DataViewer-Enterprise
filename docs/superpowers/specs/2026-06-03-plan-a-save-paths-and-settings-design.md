# Plan A — Save-Path Defaults + Per-Mode Output Settings Tab — Design

**Date:** 2026-06-03
**Status:** Approved (design); ready for implementation planning
**Part of:** the v2.2.x critical-bug initiative (Plan A of A/B/C). Sibling plans: **B** = DB-load render fix (Bug 2), **C** = auto-recovery + crash snapshot (Bug 1). Built one at a time; this is first.

---

## Goal

Make every report save to a predictable, configurable folder with a consistent `_report` filename, make every save/export dialog default to the active user's Documents folder, and add a **Settings** ribbon tab whose three buttons set per-mode default report-output folders that persist across restarts.

This rolls Bug 3 (save-path defaults + report naming) and the Settings-tab feature request into one plan because they share the same plumbing: "where does a report go, and what is it called."

## Architecture (2-3 sentences)

A new `src/utils/OutputPaths` helper becomes the single source of truth for (a) the three persisted per-mode output folders, (b) the directory-precedence resolver used to seed save dialogs, and (c) the `<base>_report.pptx` filename builder + sanitizer. Every existing save dialog is repointed at this helper, which **replaces the three divergent `lastBrowseDir()` implementations** (MainWindow, SensoryPanel, DetailedSensoryPanel) — eliminating the machine-specific hard-coded OneDrive fallback. A new `buildSettingsTab()` adds the ribbon tab whose handlers write the per-mode folders into the helper.

## Tech Stack

C++17 / Qt 6.10 (`QSettings` registry storage, `QStandardPaths`, `QFileDialog`, `QDir`/`QFileInfo`), namespace `DVE`, qmake + MinGW, `-Werror -Wall -Wextra`. New unit suite under `tests/` via the existing `tests/tests.pro` SUBDIRS template.

---

## Locked decisions (from brainstorming)

1. **Directory precedence for reports:** configured **per-mode folder → last-used dir (this session) → Documents**.
2. **Report filename:** lowercase `_report`. TPM **Test** (per-sheet) report keeps the sheet segment → `<file>_<sheet>_report.pptx`. Single-file / full reports → `<file>_report.pptx`.
3. **Scope of the Documents default:** **all** save/export dialogs default to Documents (or the per-mode folder for reports). The `<file>_report` naming convention applies to **reports only**; other exports keep their current name derivation.
4. **Refactor:** **unify** the three `lastBrowseDir()` helpers into one shared resolver (`OutputPaths`), per backlog items D001 / R-012; this removes the hard-coded OneDrive path.

### Defaults chosen by the implementer (flagged for confirmation in the spec review)
- Combined/multi-file reports (no single source file): TPM multi-file combined → `Combined_<date>_report.pptx`; DB-browser combined Sensory → `Combined_Sensory_report.pptx`; DB-browser combined Detailed → `Combined_Detailed_Sensory_report.pptx`.
- **Last-used dir remains session-only** (in-memory), not newly persisted to the registry. The persistent need is met by the per-mode Settings folders.
- Settings buttons reuse the already-shipped `folder-open.svg` icon (no new SVG → no risk of an unshipped resource showing blank in production).
- `QSettings("SDR","DataViewerEnterprise")` (the no-space form already used by window/report settings) is the persistence namespace.

---

## Component 1 — `src/utils/OutputPaths.{h,cpp}` (new)

Namespace `DVE`. Single source of truth. No widget dependencies (only Qt core + `QSettings`).

```cpp
namespace DVE {

enum class ReportMode { Tpm, Sensory, DetailedSensory };

class OutputPaths {
public:
    // ── Persisted per-mode output folders (QSettings "SDR"/"DataViewerEnterprise") ──
    //   keys: output/tpmDir, output/sensoryDir, output/detailedDir
    static QString configuredDir(ReportMode mode);              // "" if unset
    static void    setConfiguredDir(ReportMode mode, const QString& dir);

    // ── Documents fallback ──
    //   QStandardPaths::writableLocation(DocumentsLocation), else QDir::homePath()+"/Documents"
    static QString documentsDir();

    // ── Directory resolvers (return the first that exists) ──
    //   report:    configuredDir(mode) -> lastUsedDir -> documentsDir()
    //   non-report: lastUsedDir -> documentsDir()
    static QString resolveReportDir(ReportMode mode, const QString& lastUsedDir);
    static QString resolveSaveDir(const QString& lastUsedDir);

    // ── Filename builder ──
    //   reportFileName("Foo")            -> "Foo_report.pptx"
    //   reportFileName("Foo","Sheet 1")  -> "Foo_Sheet_1_report.pptx"   (sheet segment sanitized)
    //   base derived via QFileInfo(filePath).completeBaseName() at call sites (no .chopped(5))
    static QString reportFileName(const QString& base, const QString& sheet = QString());

    // ── Windows-safe sanitizer (used by reportFileName for both base and sheet) ──
    //   strips/replaces  \ / : * ? " < > |  and trims; collapses whitespace to '_'
    static QString sanitize(const QString& raw);
};

} // namespace DVE
```

**Testability:** `resolveReportDir`, `resolveSaveDir`, `reportFileName`, and `sanitize` are **pure** (inputs → outputs, no global state except the existence checks, which take explicit dir args). Only `configuredDir`/`setConfiguredDir` touch `QSettings`. Unit tests exercise the pure functions directly.

**Edge cases handled inside the helper:**
- A configured per-mode folder that no longer exists on disk → skipped (falls through to last-used → Documents). Use `QDir(path).exists()`.
- `DocumentsLocation` empty (rare) → `QDir::homePath() + "/Documents"`.
- Empty `base` (e.g. a sensory session with no title) → caller passes a fallback base (`"sensory"` / `"detailed_sensory"`); the helper does not invent names.

---

## Component 2 — Settings ribbon tab

**`MainWindow::setupRibbon()`** (≈ MainWindow.cpp:407-421; tabs currently Home / Reports / Tools): append after the `buildToolsTab(...)` line so the new tab renders to the **right of Tools** (tab order == call order):

```cpp
buildSettingsTab(m_ribbon->addTab("Settings"));
```

**New `MainWindow::buildSettingsTab(RibbonTab* tab)`** (declared in MainWindow.h next to the other `build*Tab` methods; modeled on the minimal `buildToolsTab` at MainWindow.cpp:596):

- One group `tab->addGroup("Output Paths")`.
- Three `group->addLargeButton(text, AppTheme::icon("folder-open"), tooltip)` buttons:
  - "Set TPM Output Path"
  - "Set Sensory Output Path"
  - "Set Detailed Sensory Output Path"
- Each `connect(btn, &QToolButton::clicked, this, …)` → a handler that:
  1. `QFileDialog::getExistingDirectory(this, title, OutputPaths::configuredDir(mode).isEmpty() ? OutputPaths::documentsDir() : OutputPaths::configuredDir(mode))`,
  2. on non-empty result, `OutputPaths::setConfiguredDir(mode, dir)`,
  3. refresh that button's tooltip to the current path.
- **Tooltip** of each button shows the current configured path, or `"Not set — defaults to Documents"` when empty. Set initial tooltips when building the tab; update on change. Keep button pointers (`m_tpmPathBtn` etc.) or look them up to update tooltips.

`addLargeButton` already enforces the 80×76 fixed size + word-wrap + 32×32 icon + QSS, so labels like "Set Detailed Sensory Output Path" wrap correctly with no extra handling.

---

## Component 3 — Save-site repointing (14 dialogs)

All line numbers are from the 2026-06-03 investigation and are guidance; re-locate during implementation. The transform is mechanical: replace the dialog's directory seed with an `OutputPaths` resolver call, and (for reports) set the filename via `OutputPaths::reportFileName(...)`.

### Reports — per-mode dir + `_report` naming

| # | Site (file:line) | Mode | Dir source | Filename |
|---|---|---|---|---|
| 1 | `MainWindow::onGenerateTestReport` (MainWindow.cpp:3192) | Tpm | `resolveReportDir(Tpm, lastUsed)` | `reportFileName(<fileBase>, <sheet>)` → `<file>_<sheet>_report.pptx` |
| 2 | `MainWindow::onGenerateFullReport` single-file (MainWindow.cpp:3286) | Tpm | `resolveReportDir(Tpm, lastUsed)` | `reportFileName(<fileBase>)` → `<file>_report.pptx` |
| 3 | `MainWindow::onGenerateFullReport` multi-file folder picker (MainWindow.cpp:3298) | Tpm | `resolveReportDir(Tpm, lastUsed)` | per-file `reportFileName(<fileBase>)`; combined → `Combined_<date>_report.pptx` |
| 4 | `ReportPreviewDialog::onCreateReport` (ReportPreviewDialog.cpp:606) — *today hard-codes `"report.pptx"`, no dir* | Sensory | `resolveReportDir(Sensory, lastUsed)` | `reportFileName(source->suggestedReportBaseName())` |
| 5 | `DetailedSensoryPanel::generateReport` (DetailedSensoryPanel.cpp:1387) | DetailedSensory | `resolveReportDir(DetailedSensory, lastUsed)` | `reportFileName(<sessionTitle or "detailed_sensory">)` |
| 6 | `DatabaseBrowserDialog` combined Sensory (DatabaseBrowserDialog.cpp:721) | Sensory | `resolveReportDir(Sensory, lastUsed)` | `Combined_Sensory_report.pptx` |
| 7 | `DatabaseBrowserDialog` combined Detailed (DatabaseBrowserDialog.cpp:892) | DetailedSensory | `resolveReportDir(DetailedSensory, lastUsed)` | `Combined_Detailed_Sensory_report.pptx` |

### Non-report saves/exports — dir defaults to Documents, names unchanged

| # | Site (file:line) | Dir source |
|---|---|---|
| 8 | `SensoryPanel::save` session xlsx/json (SensoryPanel.cpp:1459) | `resolveSaveDir(lastUsed)` |
| 9 | `SensoryPanel::onSaveChart` radar PNG (SensoryPanel.cpp:1356) | `resolveSaveDir(lastUsed)` |
| 10 | `SensoryPanel::generateStats` CSV (SensoryPanel.cpp:2164) | `resolveSaveDir(lastUsed)` |
| 11 | `DetailedSensoryPanel::save` session xlsx (DetailedSensoryPanel.cpp:1027) | `resolveSaveDir(lastUsed)` |
| 12 | `PlotWidget::onSaveImage` TPM plot image (PlotWidget.cpp:362) — *today empty default* | `resolveSaveDir(lastUsed)` (also start tracking last-used here) |

### Helper retirement

- `MainWindow::lastBrowseDir`/`setLastBrowseDir` (MainWindow.cpp:4439-4451), `SensoryPanel::lastBrowseDir` (SensoryPanel.cpp:1380), `DetailedSensoryPanel::lastBrowseDir` (DetailedSensoryPanel.cpp:1552): keep the per-widget `m_lastBrowseDir` **session memory** (updated by `setLastBrowseDir` after each pick, as today) but route the **fallback** through `OutputPaths`. Concretely, each `lastBrowseDir()` collapses to: return `m_lastBrowseDir` if it exists, else `OutputPaths::documentsDir()`. The OneDrive and `homePath()`-without-`/Documents` fallbacks are deleted. Open-file dialogs that call these helpers inherit the unified Documents fallback (no behavioral regression beyond dropping OneDrive).
- `ImageInboxDialog` watch-folder `getExistingDirectory` (ImageInboxDialog.cpp:592) is an **input** picker — left unchanged.

---

## Component 4 — `IReportSource::suggestedReportBaseName()` (interface addition)

`ReportPreviewDialog` (the Sensory full-report path) only receives an `IReportSource*` and currently has no file/session name to build `<base>_report` from. Add:

```cpp
// src/reporting/IReportSource.h
virtual QString suggestedReportBaseName() const = 0;   // sanitized base for "<base>_report.pptx"
```

- `SensoryReportSource::suggestedReportBaseName()` → the single session's `testTitle` (sanitized), else `"sensory"`.
- Any other `IReportSource` implementors get the method too (return a sensible sanitized label). This is the only interface churn in Plan A.

`ReportPreviewDialog::onCreateReport` then seeds the dialog with `OutputPaths::resolveReportDir(ReportMode::Sensory, lastUsed) + "/" + OutputPaths::reportFileName(source->suggestedReportBaseName())`. It already reads `QSettings` directly for `"preview/snap"`, so reading output settings in-place is consistent — but it goes through `OutputPaths` for the keys, not raw `QSettings`.

---

## Error handling & edge cases

- **Stale configured folder** (deleted/renamed) → resolver falls through; never opens a dialog at a non-existent path.
- **No Documents location** → `homePath()/Documents` fallback.
- **Empty/blank report base** → call sites pass a non-empty fallback base; sanitizer guarantees a valid Windows filename (no illegal chars, non-empty).
- **Sanitizer consistency** — one sanitizer for all report names replaces the three different ad-hoc sanitizers (`onGenerateTestReport`'s regex, the sensory "spaces→underscore", and `ReportGenerator`'s internal regex) so names are uniform.
- **`-Werror`** — ensure no unused parameters/includes in the new build method and helper.

## Testing

**Unit — new `tests/tst_outputpaths/` (added to `tests/tests.pro` SUBDIRS):**
- `resolveReportDir`: per-mode set & exists → returns it; per-mode set but missing → falls to last-used; per-mode unset → last-used; last-used missing/empty → Documents.
- `resolveSaveDir`: last-used exists → it; else Documents.
- `reportFileName`: `"Foo"` → `Foo_report.pptx`; `"Foo","Sheet 1"` → `Foo_Sheet_1_report.pptx`; already-`.pptx` base not double-suffixed.
- `sanitize`: strips `\ / : * ? " < > |`, collapses whitespace, non-empty guarantee.
- Hermetic: pure functions take explicit dir args (existence checks use temp dirs created by the test); no registry writes.

**Manual (registry + dialogs + ribbon can't be unit-tested):**
- Settings tab appears immediately right of Tools; three buttons present and labeled (word-wrapped).
- Each button opens a folder picker, persists the choice; tooltip updates; value survives an app restart (registry).
- For each report type, the save dialog opens at the mode's configured folder when set, else last-used, else Documents, and pre-fills the correct `<…>_report.pptx` name.
- With no per-mode folder set, all save/export dialogs open at Documents.

## Files

**Create:**
- `src/utils/OutputPaths.h`
- `src/utils/OutputPaths.cpp`
- `tests/tst_outputpaths/tst_outputpaths.pro`
- `tests/tst_outputpaths/tst_outputpaths.cpp`

**Modify:**
- `DataViewerEnterprise.pro` (add `src/utils/OutputPaths.{cpp,h}` to SOURCES/HEADERS)
- `tests/tests.pro` (add `tst_outputpaths` to SUBDIRS)
- `src/MainWindow.h` / `src/MainWindow.cpp` (`buildSettingsTab` + 3 handlers + button members; repoint sites 1-3; retire `lastBrowseDir` fallback)
- `src/ui/ReportPreviewDialog.cpp` (site 4)
- `src/ui/DetailedSensoryPanel.cpp` (site 5; retire `lastBrowseDir` fallback)
- `src/ui/DatabaseBrowserDialog.cpp` (sites 6-7)
- `src/ui/SensoryPanel.cpp` (sites 8-10; retire `lastBrowseDir` fallback)
- `src/plotting/PlotWidget.cpp` (site 12)
- `src/reporting/IReportSource.h` (add `suggestedReportBaseName`)
- `src/reporting/SensoryReportSource.{h,cpp}` (implement it)
- any other `IReportSource` implementor (implement it)

## Non-goals (explicitly out of scope for Plan A)

- Persisting last-used dir across restarts (kept session-only by decision).
- Bug 2 (DB-load) and Bug 1 (auto-recovery) — separate Plans B and C.
- Changing report *content* or layout; this is purely about output location + filename.
- A general Settings/Preferences dialog — the request is specifically three ribbon buttons → folder pickers.

## Build / version

No version bump for the spec/plan. The VERSION bump + installer build happens at the end of Plan A implementation (per the rebuild-dataviewer skill: clean rebuild after a VERSION change), followed by the user's installer eyeball-test before any merge to main.
