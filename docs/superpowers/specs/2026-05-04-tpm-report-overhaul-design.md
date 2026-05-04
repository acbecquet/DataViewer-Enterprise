# TPM Report Overhaul + Automated Test Runner

**Date:** 2026-05-04
**Status:** Spec — pending implementation
**Branch target:** `main` (single feature, no separate branch unless implementation is split)

## Summary

Two coupled changes to DataViewer Enterprise's TPM workflow:

1. **Report rendering and structure** — adaptive markers on plot data points, larger fonts on report plots, a fixed Y-axis scaling rule on TPM plots, three new slides in Full Report (Test Protocol, Test Overview, Conclusions), and a new combined-report mode that batches multiple files and produces individual reports plus a single combined report with section dividers and a cross-file Lifetime TPM comparison.

2. **Automated test runner** — a single PowerShell entry point on top of the existing 1,800-line Qt Test scaffold, replacing the bash + batch scripts and adding tests for every new helper introduced by the report work.

The two are bundled into one spec because the runner is implemented in service of the report work and lets us iterate with a fast local feedback loop. Either may ship first.

## 1. Plot rendering changes

### 1.1 Adaptive markers on every data point

Every series gets a marker at every data point. `ReportGenerator::buildPlots` no longer gates dots on `rows.size() <= 30`. The dot radius scales with point density to avoid overlap on lifetime tests:

```
adaptiveDotRadius(n):
  if n <= 30:   return 5
  if n >= 150:  return 2
  return round(linearInterp(30→5, 150→2, n))
```

The line is still drawn underneath at the existing `lineWidth = 2`.

### 1.2 Report-only font bumps

GUI plot fonts are unchanged. Report plots get a `PlotConfig` produced by a new helper `ReportGenerator::reportPlotConfig()`:

| Element | Current | Report |
|---|---|---|
| Title (`titleFont`) | 11 pt bold | **18 pt bold** |
| Axis label (`axisFont`) | 9 pt | **18 pt** |
| Tick / legend (`labelFont`) | 8 pt | **14 pt** |

The combined-report Lifetime Comparison slide uses an additional override pushing the legend font to **18 pt** for auditorium readability.

### 1.3 Fixed Y-axis scaling on TPM plots

Two new helpers on `ReportGenerator`:

```cpp
bool   isLongPuff(const SheetResult& sheet) const;
double computeTpmYMax(const SheetResult& sheet) const;
```

```
isLongPuff(sheet):
  return sheet.sheetName.contains("Long Puff", caseInsensitive)
      OR any sample.puffingRegime matches /\d+\s*mL\s*\/\s*10s\s*\/.*/

computeTpmYMax(sheet):
  let maxTPM = max(row.tpm for sample in sheet.samples for row in sample.rows)
  let avgTPM = mean(sample.averageTPM for sample in sheet.samples)
  if isLongPuff(sheet):
    return (15 <= maxTPM <= 25) ? 25 : maxTPM + 1
  else:
    return (avgTPM > 7) ? maxTPM + 1 : 7
```

The avg-vs-max asymmetry is intentional: for non-Long-Puff tests, the 0–7 frame holds even when one outlier puff blows past 7, as long as the test as a whole stays under that ceiling.

Applied to: TPM Trend line plot, Average TPM Bar Chart, and the new Lifetime Comparison bar chart. Draw Pressure stays auto-scaled — different metric (Pa) with no fixed reference range.

`PlotConfig.autoScale` is set to `false` and `yMin = 0`, `yMax = computeTpmYMax(sheet)` is set explicitly when these rules apply.

## 2. Slide flow

### 2.1 Single-file Full Report

```
1. Title slide                                 (existing)
2. Test Protocol                               NEW
3. Test Overview                               NEW
4..N. Data + image slides                      (existing)
N+1. Conclusions                               NEW
```

### 2.2 Multi-file (combined) Full Report

```
1.  Title slide  ("Combined Standard Test Report" + date)
2.  Test Protocol  (union of all tests across selected files)
3.  Test Overview  (combined — all files, all unique tests)
4.  Lifetime TPM Comparison  (cross-file bar chart)
5.  Section divider — File 1
6.  Test Overview  (just File 1)
7.. Data + image slides for File 1
M.  Section divider — File 2
M+1. Test Overview  (just File 2)
... etc
END. Conclusions
```

### 2.3 Slide details

**Title slide.** Single-file: existing template (`<filename> Standard Test Report` + date). Combined: title is `Combined Standard Test Report`, date stays. Subtitle is omitted for v1.

**Test Protocol slide.** 6-column table sourced from `resources/templates/Standardized Test Template - December 2025.xlsx`, "Test SOPs" sheet (loaded by the new `SopLoader`). Columns:

| Test | Objective | Pass Criteria | Equipment | Quantity | Est Duration |
|---|---|---|---|---|---|

`Est Duration` collapses the two source columns into `1mL: X / 2mL: Y`. Rows filtered to tests appearing in the report. Tests with no matching SOP entry get a row with `—` in the four content columns; the test name column still shows the test. SOPs that don't match a test in the report are excluded.

**Test Overview slide.** Title `"Test Overview"`. Body:

- Line 1 (auto-templated):
  - Single-file (and per-section): `"Standard performance evaluation of <DeviceName> across <N> tests."` `<DeviceName>` is the file's `completeBaseName`.
  - Combined-level: `"Combined performance evaluation across <N> files and <M> unique tests."`
- Bulleted list of test (sheet) names, in the order they appear in the file.
- Line 2: empty placeholder textbox — manually edited in PowerPoint.

The auto-templated description text is a single source-of-truth string in `ReportGenerator` and easy to refine over time without touching the slide layout.

**Section divider slide.** Reuses the existing cover-slide template (`PptxWriter::addCoverSlide`'s background + logo). The date placeholder is removed and the file's `completeBaseName` (no `.xlsx`) is centered both axes at large bold (~44 pt). Implementation: a new `PptxWriter::addSectionDividerSlide(QString filename)` that calls into the same XML scaffold as `addCoverSlide` but without the date shape and with the title-shape position centered.

**Lifetime TPM Comparison slide.** Combined-report only. Title `"Lifetime TPM Comparison"`. Body fills the slide with a single bar chart.

- One bar per sample drawn from each file's `"Lifetime Test"` sheet (case-insensitive exact match — `"Long Puff Lifetime Test"` and `"Rapid Puff Lifetime Test"` are different tests and excluded).
- Files without a `"Lifetime Test"` sheet are skipped silently.
- Bars colored by **file**, with progressive shading per **sample within that file**: each file gets a hue from the existing `kColors` palette (blue, orange, green, red, purple, teal). Within a file, samples are rendered with progressive shading from light to dark of that hue (HSV value bumped per sample, e.g. 60% → 100% across the in-file sample set).
- Sample names below each bar.
- Legend top-right, one row per file with a representative swatch + filename. Font 18 pt.
- Y-axis follows `computeTpmYMax` with the non-Long-Puff branch.

If the slide ends up with zero bars (no files have a Lifetime Test sheet), the slide is omitted entirely.

**Conclusions slide.** Title `"Conclusions"`. One large empty textbox positioned in the body where a paragraph would naturally go. No placeholder text.

## 3. Multi-file selection + combined-report assembly

### 3.1 Trigger and flow

`Reports → Full Report` button (no new ribbon button; existing button gains the multi-select capability). Click flow:

1. `QFileDialog::getOpenFileNames` (`.xlsx` filter, defaults to `lastBrowseDir`). User picks 1+ files.
2. If 1 file picked: existing single-file flow, `getSaveFileName` for output path, then `generateFullReport` with the new slides.
3. If 2+ files picked: `QFileDialog::getExistingDirectory` for the output folder. Then a batch loop produces all individual reports + the combined report into that folder.

This matches the multi-select pattern already established in sensory mode.

### 3.2 Output naming (multi-file)

```
<chosen_folder>/<file1>_Report.pptx
<chosen_folder>/<file2>_Report.pptx
...
<chosen_folder>/Combined_Report_<YYYY-MM-DD>.pptx
```

Filename collision: append `(2)`, `(3)`, … — never overwrite silently. (Implementation: in a helper `uniqueFilename(folder, base)`.)

### 3.3 Loading picked files

For each picked path:
1. If already in `m_loadedFiles`: reuse the in-memory `FileResult` (so any active `m_excludedRows` exclusions apply via `buildCleanedFile`).
2. Otherwise: `DataProcessor::processFile` synchronously inside a modal `QProgressDialog`. Multi-file load shows `Loading file X of N`.

### 3.4 Combined report assembly

```
generateCombinedFullReport(files: QVector<FileResult>,
                           config: ReportConfig,
                           outputPath: QString,
                           progress: ProgressFn) -> bool

  load SOPs once (via SopLoader)
  collect allTests = unique(sheet names that produce data, across files)

  PptxWriter w
  addTitleSlide(w, "Combined Standard Test Report", today)
  addTestProtocolSlide(w, filterSops(sopRows, allTests))
  addTestOverviewSlide(w, combinedDescription(files), allTests)
  addLifetimeComparisonSlide(w, files)        // skipped if no file has Lifetime Test

  for f in files:
      addSectionDividerSlide(w, QFileInfo(f.fileName).completeBaseName())
      addTestOverviewSlide(w, perFileDescription(f), f.testNames)
      for sheet in f.sheets:
          if !sheet.hasSamples(): continue
          addContentSlide(w, sheet, plots[sheet], buildTable(sheet, config))
          addImageSlides(w, sheet)            // existing image-slide logic

  addConclusionsSlide(w)
  return w.save(outputPath)
```

Individual reports are produced by the existing `generateFullReport`, extended with the same Test Protocol / Test Overview / Conclusions slides (no section dividers, no Lifetime Comparison).

### 3.5 Progress + cancellation

`ProgressFn` callback receives total-batch percentage: `((currentReport - 1) / totalReports) * 100 + intraReportPct / totalReports`. Modal `QProgressDialog` is shown for the duration of the batch, with a Cancel button. Cancel takes effect at the next file boundary; mid-file cancellation is not supported in v1 (matches existing behavior).

### 3.6 Failure handling

If any individual report fails (load error, PPTX save error, etc.), the loop continues for the remaining reports. The combined report failing does **not** prevent individual reports. At end, a summary dialog:

> Generated N of M reports.
> Failed: [filename1, filename2]
> See log for details.

The principle throughout: **never block the whole batch on a recoverable single-report problem.**

## 4. Component layout

| File | Change | Notes |
|---|---|---|
| `src/plotting/PlotEngine.{h,cpp}` | None | All styling driven via `PlotConfig`. |
| `src/reporting/ReportGenerator.{h,cpp}` | Major | New helpers + `generateCombinedFullReport`; existing `generateFullReport`/`generateTestReport` extended with new slides. |
| `src/reporting/PptxWriter.{h,cpp}` | Moderate | New: `addTestProtocolSlide`, `addTestOverviewSlide`, `addSectionDividerSlide`, `addLifetimeComparisonSlide`, `addConclusionsSlide`. Reuse existing `addContentSlide`/`makeTextBox`/`buildTableXml` infrastructure. |
| `src/utils/SopLoader.{h,cpp}` | **New** | Extracted from `SopDialog::loadFromExcel`. Returns `QVector<SopEntry>` from a template path. `SopDialog` refactored to consume it (no behavior change). |
| `src/MainWindow.cpp` | Moderate | `onGenerateFullReport` rewritten as multi-select → output-folder dialog → batch loop. |
| `src/MainWindow.h` | Minor | Helper signatures: `loadOrReuseFile(path)`, `generateBatchReports(...)`. |
| `DataViewerEnterprise.pro` | Minor | Add `SopLoader.{h,cpp}`. |

### 4.1 New ReportGenerator API

```cpp
class ReportGenerator : public QObject {
public:
    bool generateFullReport(const FileResult&, const ReportConfig&, ProgressFn);
    bool generateTestReport(const FileResult&, const QString&, const ReportConfig&, ProgressFn);

    // NEW
    bool generateCombinedFullReport(const QVector<FileResult>& files,
                                    const ReportConfig& config,
                                    const QString& outputPath,
                                    ProgressFn progress = nullptr);

private:
    // NEW helpers
    QVector<SopEntry> loadSopRows(const QStringList& reportTestNames) const;
    void addTitleSlide(PptxWriter&, const QString& title, const QString& date);
    void addTestProtocolSlide(PptxWriter&, const QVector<SopEntry>& filtered);
    void addTestOverviewSlide(PptxWriter&, const QString& description,
                              const QStringList& testNames);
    void addSectionDividerSlide(PptxWriter&, const QString& filename);
    void addLifetimeComparisonSlide(PptxWriter&, const QVector<FileResult>& files);
    void addConclusionsSlide(PptxWriter&);

    PlotConfig reportPlotConfig() const;
    int        adaptiveDotRadius(int pointCount) const;
    double     computeTpmYMax(const SheetResult&) const;
    bool       isLongPuff(const SheetResult&) const;

    /// Bar color for the Lifetime Comparison slide. fileIdx selects the hue
    /// from kColors; sampleIdx within the file shifts HSV value from 0.6→1.0.
    static QColor lifetimeBarColor(int fileIdx, int sampleIdx, int totalSamplesInFile);

    // Existing
    QVector<QByteArray> buildPlots(const SheetResult&, bool includeBarChart = true);
    SlideTable          buildTable(const SheetResult&, const ReportConfig&);
    QVector<QByteArray> collectImages(const SheetResult&);
};
```

### 4.2 SopLoader API (new file)

```cpp
namespace DVE {

struct SopEntry {
    QString test;
    QString sop;
    QString objective;
    QString passCriteria;
    QString equipment;
    QString quantity;
    QString estDuration1mL;
    QString estDuration2mL;
    QString note;
};

class SopLoader {
public:
    /// Load all SOP rows from the standardized template's "Test SOPs" sheet.
    /// Returns empty vector and logs at warning if the file is missing or
    /// unreadable — never throws, never blocks callers.
    static QVector<SopEntry> load(const QString& xlsxPath);
};

} // namespace DVE
```

`SopDialog::loadFromExcel` becomes a thin wrapper around `SopLoader::load`. This is a "good developer improving the code they're working in" cleanup, not unrelated refactor — without it, two places parse the same sheet with two slightly different code paths.

## 5. Verification

### 5.1 Manual checklist (work-machine, after deploy)

```
Single-file Full Report
[ ] Generates without error on a normal file (3+ sheets)
[ ] Title slide unchanged
[ ] Test Protocol slide shows 6 cols, only tests present in the file
[ ] Test Overview shows auto-templated line + bullets + empty trailing textbox
[ ] Each data slide: TPM Trend has markers, Y-axis matches the rule
[ ] Image slides unchanged
[ ] Conclusions slide: blank textbox

Multi-file Combined Report (3 files)
[ ] File picker accepts multi-select
[ ] Output folder dialog appears once
[ ] 3 individual reports + 1 combined report appear in chosen folder
[ ] Combined: Title → Protocol → Overview → Lifetime Comparison → 3 sections → Conclusions
[ ] Section divider slides: cover style, no date, filename centered
[ ] Lifetime Comparison: bars colored by file, shaded by sample within file
[ ] Files without "Lifetime Test" sheet are silently skipped on the comparison slide

Edge cases
[ ] Long Puff Lifetime Test sheet: Y-axis 0–25 default, falls back to maxTPM+1 when out of range
[ ] Sheet with avgTPM > 7: Y-axis bumps to maxTPM+1
[ ] Puffing regime "200mL/10s/60s": detected as Long Puff even on a non-Long-Puff sheet name
[ ] File picked that's already loaded: reused (no re-parse)
[ ] One report failing in a batch: other reports still generate, summary dialog shows failures
[ ] Filename collision in output dir: appends (2), no overwrite
```

The checklist gets appended to `tests/deployment/README.md` so it's run alongside the existing deployment self-test.

### 5.2 Error-handling matrix

| Failure | Behavior |
|---|---|
| Template SOPs file missing | Test Protocol slide skipped silently; logged at warning. Don't block report. |
| SOP entry missing for a test in the report | Row added with `—` for Objective/Pass/Equipment columns; test name still shown. |
| File in batch fails to load | Skipped from individual + combined; appended to failure summary. |
| Lifetime Test sheet absent in all files (combined) | Comparison slide omitted; report still completes. |
| Lifetime sheet present but no usable samples | Same as absent. |
| User cancels mid-batch | Stops at next file boundary; partial outputs remain on disk. |
| Output folder not writable | Surfaced via `QFileDialog::getExistingDirectory` (it won't accept). |
| PPTX save fails | That report skipped; failure summary shows it. The combined-report failure does not block individual reports. |

## 6. Automated test runner

### 6.1 What exists

The `tests/` directory already contains 11 Qt Test classes (~1,800 lines of real test code) and a `tests.pro` `SUBDIRS` template that builds them all. Two existing runner scripts (`build_all.bat`, `run_all_tests.sh`) work but are fragile: bash + batch on a Windows-only project, hardcoded Qt 6.10.1 (no longer present on the home machine), no `qmake` step in `build_all.bat`, an MSYS-pseudo-pipe workaround for `tst_reportgenerator`.

### 6.2 New entry point: `tests\run-tests.ps1`

Single PowerShell script. Replaces both legacy scripts.

```
1. Auto-detect Qt + MinGW
   - Probe C:\Qt\6.10.*\mingw_64\bin\qmake.exe; pick highest minor.
   - Probe C:\Qt\Tools\mingw1310_64\bin (or any mingw1*_64 if absent).
   - Fail with a clear error if either is missing.

2. Build phase
   - cd tests
   - <qmake> -recursive tests.pro
   - <make> -j<n>
   - Skip if all tst_*\release\tst_*.exe are newer than their sources.

3. Run phase
   - For each tst_*\release\tst_*.exe:
     - Run with -o <tmp>,txt
     - Parse "Totals: X passed, Y failed"
     - Print "  PASS (X passed)" or "  FAIL — Totals: ..."

4. Summary
   - "Results: P passed, F failed, S skipped"
   - On failure: print FAIL! lines from each failed test's text output
   - Exit code 0 iff all pass
```

The script does *not* attempt to be cross-platform. Windows-only project; PowerShell native is the right call.

### 6.3 Tests added in this work

These cover the new helpers introduced for the report changes:

- **`tst_sopLoader` (new test class)** — `SopLoader::load` against a fixture xlsx. Cases: file exists with 5 rows, file missing, malformed row.
- **`tst_reportgenerator` (existing — extend)**:
  - `isLongPuff` cases: sheet-name match, regime regex match, neither, edge cases (`"100mL/10s"` no trailing slash, `"200mL/3s/30s"` should not match).
  - `computeTpmYMax`: long puff in/out of range, non-Long-Puff with avgTPM under and over 7.
  - `adaptiveDotRadius`: 0 (degenerate), 30, 31, 100, 1000.
- **`tst_pptxwriter` (existing — extend)**:
  - `addTestProtocolSlide`, `addTestOverviewSlide`, `addSectionDividerSlide`, `addConclusionsSlide` produce well-formed XML. Verification depth: count expected tags (`<a:tbl>`, `<p:sp>`, etc.), check titles appear; do not full-validate XSD.

The combined-report end-to-end and Lifetime Comparison plot rendering remain in the manual checklist — visual output, not unit-testable cheaply.

### 6.4 Cleanup

`tests/build_all.bat` and `tests/run_all_tests.sh` are deleted once `run-tests.ps1` ships and is verified.

### 6.5 Documentation update

`CLAUDE.md` line "There are no automated tests yet" becomes:

> Run the test suite with `tests\run-tests.ps1`. 11 test classes covering pipeline, plotting, reporting, database, and zip/xml utilities. Builds incrementally (skipped if `tst_*\release\*.exe` newer than sources).

## Implementation order (suggested)

The two pieces are mostly independent. Suggested order:

1. **Test runner first** (Section 6) — small, foundational, gives us the feedback loop.
2. **`SopLoader` extraction** (Section 4) — small refactor, no behavior change.
3. **Plot rendering helpers** (Section 1) — `isLongPuff`, `computeTpmYMax`, `adaptiveDotRadius`, `reportPlotConfig`. Add tests as we go.
4. **New slide types** (Section 2) — Test Protocol, Test Overview, Conclusions, Section Divider. Wired into existing single-file Full Report first.
5. **Multi-file flow + Lifetime Comparison** (Section 3 + 2.3) — combined report assembly; the only piece that requires the work-machine setup to verify visually.

Steps 1–4 can be implemented and verified locally on the home machine. Step 5 needs the work machine for visual review of the combined report.

## Out of scope (explicit)

- Refining auto-templated overview text wording — will iterate after seeing real reports.
- Graphical CI on GitHub Actions — single-developer Windows-only project, not warranted yet.
- Changing GUI plot fonts or markers — only report plots change.
- Sensory-mode reports — this overhaul is TPM-only.
- Validating PPTX output by opening it in PowerPoint headlessly — manual visual review remains the validation step for slide content.
