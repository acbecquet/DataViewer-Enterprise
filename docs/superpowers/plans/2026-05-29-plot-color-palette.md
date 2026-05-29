# Distinct-Color Plot Palette Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make every plot in the app draw each series/bar/slice in a distinct, well-separated, non-yellow, projector-safe color — never reusing a color within a plot — sourced from one shared palette helper.

**Architecture:** A single palette helper on `AppTheme` (`seriesColor`/`seriesColors`/`shade`) becomes the one source of truth. It reuses the 20-color, no-yellow palette already vetted in `RadarChartWidget`, with golden-angle generation beyond 20 so it never repeats for any count. All plot sites — report charts, `PlotEngine`, on-screen `PlotWidget`, and `RadarChartWidget` — stop indexing private palettes with `% N` and call the helper. The Lifetime comparison chart keeps file grouping by giving each file a distinct base hue and shading samples within a projector-safe band, which provably yields no duplicate colors.

**Tech Stack:** C++17, Qt 6.10 (QtGui `QColor`/HSV), qmake + MinGW, Qt Test.

---

## Environment preamble (this machine)

- **Before any C++ build/test**, decrypt MIP-labeled sources first:
  ```bash
  python tools/decrypt_via_copy.py --apply
  ```
- **Create new source files via Python delete-and-rewrite** (so they don't inherit MIP labels), e.g.:
  ```python
  import os
  path = "tests/tst_apptheme/tst_apptheme.cpp"
  content = "..."
  if os.path.exists(path): os.remove(path)
  with open(path, "w", encoding="utf-8", newline="\n") as f: f.write(content)
  ```
- Build/test commands assume the repo root as CWD and Qt/MinGW per `CLAUDE.md`.
- Every commit message ends with the `Co-Authored-By` trailer from the harness instructions.

## File Structure

| File | Responsibility | Change |
|---|---|---|
| `src/utils/AppTheme.h` | Color/spacing/font tokens | **Add** `seriesColor` / `seriesColors` / `shade` declarations + `#include <QVector>` |
| `src/utils/AppTheme.cpp` | Token implementations | **Add** the three palette functions + `#include <cmath>` |
| `tests/tst_apptheme/` | Unit test for the palette helper | **Create** `.pro` + `.cpp`; **add** to `tests/tests.pro` SUBDIRS |
| `src/reporting/ReportGenerator.cpp/.h` | Report charts | Comparison chart → distinct file hues + `shade`; per-sheet plots → shared palette |
| `src/plotting/PlotEngine.cpp` | Shared chart renderer | Default fallback + per-sheet bar chart → `seriesColors` |
| `src/plotting/PlotWidget.cpp` | On-screen plots | All `% N` palettes → `AppTheme::seriesColor(idx)` |
| `src/ui/RadarChartWidget.cpp/.h` | Sensory radar | Drop private `kColors`; consume `AppTheme::seriesColor(idx)` |
| `tests/tst_*/…​.pro` | Link the helper | Add `AppTheme.cpp` to the 3 affected test projects |
| `CLAUDE.md` | Project conventions | **Add** the no-reuse rule |
| `tasks/lessons.md` | Lessons log | **Add** an entry |

---

## Task 1: `AppTheme` palette helper + unit test

**Files:**
- Modify: `src/utils/AppTheme.h`
- Modify: `src/utils/AppTheme.cpp`
- Create: `tests/tst_apptheme/tst_apptheme.pro`
- Create: `tests/tst_apptheme/tst_apptheme.cpp`
- Modify: `tests/tests.pro:38` (append to SUBDIRS)

- [ ] **Step 1: Declare the API in `AppTheme.h`**

Add `#include <QVector>` near the existing includes (after `#include <QColor>`):

```cpp
#include <QColor>
#include <QFont>
#include <QIcon>
#include <QVector>
```

Add these declarations inside `class AppTheme`, right after the `// ── Semantic ──` block (after `error()`):

```cpp
    // ── Plot series palette — single source of truth for ALL chart colors ───
    // Never reuse a color within a plot. Curated 20-color, no-yellow,
    // projector-safe palette (shared with the sensory radar); golden-angle
    // generation beyond 20 so it never repeats for any count.
    static QColor          seriesColor(int idx);   // the idx-th distinct color
    static QVector<QColor> seriesColors(int n);    // first n distinct colors
    // A distinct shade of a base hue for grouped charts (per-file families).
    // Only value/saturation vary (hue preserved), so shades of two different
    // base hues can never collide; the value band stays projector-safe.
    static QColor          shade(const QColor& base, int idx, int count);
```

- [ ] **Step 2: Write the failing test** — create `tests/tst_apptheme/tst_apptheme.cpp`

```cpp
#include <QtTest>
#include <QColor>
#include <QSet>
#include "AppTheme.h"

class TestAppTheme : public QObject
{
    Q_OBJECT
private slots:
    void testCountAndEmpty();
    void testAllUnique();
    void testSeriesColorMatchesVector();
    void testNoYellow();
    void testShadeDistinctAndBanded();
    void testShadeGrayStaysNeutral();
    void testGroupedComparisonUnique();
};

static int rgbKey(const QColor& c) { return (c.red() << 16) | (c.green() << 8) | c.blue(); }

void TestAppTheme::testCountAndEmpty()
{
    QVERIFY(AppTheme::seriesColors(0).isEmpty());
    QCOMPARE(AppTheme::seriesColors(5).size(), 5);
    QCOMPARE(AppTheme::seriesColors(25).size(), 25);
}

void TestAppTheme::testAllUnique()
{
    for (int n : {1, 5, 12, 20, 25, 40, 60}) {
        QSet<int> seen;
        const QVector<QColor> pal = AppTheme::seriesColors(n);
        for (const QColor& c : pal) seen.insert(rgbKey(c));
        QCOMPARE(seen.size(), n);   // no color repeats
    }
}

void TestAppTheme::testSeriesColorMatchesVector()
{
    const QVector<QColor> pal = AppTheme::seriesColors(30);
    for (int i = 0; i < pal.size(); ++i)
        QCOMPARE(AppTheme::seriesColor(i), pal[i]);
}

void TestAppTheme::testNoYellow()
{
    for (int i = 0; i < 60; ++i) {
        const int h = AppTheme::seriesColor(i).hue();   // -1 if achromatic
        if (h >= 0)
            QVERIFY2(!(h >= 45 && h <= 70),
                     qPrintable(QString("color %1 has yellow hue %2").arg(i).arg(h)));
    }
}

void TestAppTheme::testShadeDistinctAndBanded()
{
    const QColor base = AppTheme::seriesColor(0);   // chromatic (blue)
    const int count = 6;
    QSet<int> seen;
    int minV = 255, maxV = 0;
    for (int i = 0; i < count; ++i) {
        const QColor s = AppTheme::shade(base, i, count);
        seen.insert(rgbKey(s));
        QCOMPARE(s.hue(), base.hue());      // hue preserved → grouping intact
        minV = qMin(minV, s.value());
        maxV = qMax(maxV, s.value());
    }
    QCOMPARE(seen.size(), count);            // shades all distinct
    QVERIFY(maxV <= 217);                    // light end capped (projector-safe)
    QVERIFY(minV >= 140);                    // dark end floored
}

void TestAppTheme::testShadeGrayStaysNeutral()
{
    const QColor gray(127, 127, 127);
    QSet<int> seen;
    for (int i = 0; i < 4; ++i) {
        const QColor s = AppTheme::shade(gray, i, 4);
        QCOMPARE(s.saturation(), 0);         // stays neutral, no tint
        seen.insert(rgbKey(s));
    }
    QCOMPARE(seen.size(), 4);                // distinct grays
}

void TestAppTheme::testGroupedComparisonUnique()
{
    // The Lifetime comparison guarantee: distinct file hues + per-sample shades
    // ⇒ no two bars across the whole chart share a color.
    const int files = 4, samplesPerFile = 6;
    const QVector<QColor> fileHues = AppTheme::seriesColors(files);
    QSet<int> seen;
    for (int f = 0; f < files; ++f)
        for (int s = 0; s < samplesPerFile; ++s)
            seen.insert(rgbKey(AppTheme::shade(fileHues[f], s, samplesPerFile)));
    QCOMPARE(seen.size(), files * samplesPerFile);
}

QTEST_APPLESS_MAIN(TestAppTheme)
#include "tst_apptheme.moc"
```

- [ ] **Step 3: Create the test project** — `tests/tst_apptheme/tst_apptheme.pro`

```pro
QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle

TEMPLATE = app

INCLUDEPATH += ../../src ../../src/utils ../common

SOURCES += tst_apptheme.cpp \
           ../../src/utils/AppTheme.cpp
```

- [ ] **Step 4: Register the test** — append to `tests/tests.pro` SUBDIRS (after `tst_mainwindow_remotecell`)

```pro
    tst_mainwindow_remotecell \
    tst_apptheme
```

- [ ] **Step 5: Run the test to verify it FAILS to build**

```bash
python tools/decrypt_via_copy.py --apply
```
```powershell
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
```
Expected: `tst_apptheme` fails to compile — `seriesColor`/`seriesColors`/`shade` are not members of `AppTheme`.

- [ ] **Step 6: Implement the helper in `AppTheme.cpp`**

Add `#include <cmath>` to the include block (after `#include <QDebug>`). Then add these three functions at the end of the file (after `AppTheme::apply()`):

```cpp
QColor AppTheme::seriesColor(int idx)
{
    if (idx < 0) idx = 0;

    // Curated 20-color qualitative palette — NO yellow / yellow-adjacent hues
    // (they wash out on a projector), ordered most-distinct-first. Shared with
    // the sensory radar chart, which already vetted these for projector use.
    static const QColor kCurated[] = {
        QColor( 31, 119, 180), QColor(214,  39,  40), QColor( 44, 160,  44),
        QColor(148, 103, 189), QColor(255, 127,  14), QColor( 23, 190, 207),
        QColor(227, 119, 194), QColor(140,  86,  75), QColor(127, 127, 127),
        QColor(  0,   0, 139), QColor(139,   0,   0), QColor(  0, 100,   0),
        QColor( 75,   0, 130), QColor(255,   0, 255), QColor(  0, 191, 255),
        QColor(178,  34,  34), QColor( 70, 130, 180), QColor(255,  20, 147),
        QColor( 47,  79,  79), QColor(106,  90, 205),
    };
    static const int kN = int(sizeof(kCurated) / sizeof(kCurated[0]));

    if (idx < kN) return kCurated[idx];

    // Beyond the curated set: golden-angle hue rotation evenly spreads the rest
    // around the wheel. Skip the yellow band (45–70) and pin saturation/value
    // to a vivid, projector-safe band; nudge them each lap so colors stay
    // distinct even past a full turn.
    const int    k    = idx - kN;
    const double aDeg = std::fmod(210.0 + double(k + 1) * 137.508, 360.0);
    int hue = int(aDeg);
    if (hue >= 45 && hue <= 70) hue = (hue + 30) % 360;
    const int lap = k / 8;
    const int val = 196 - (lap % 3) * 28;    // 196 / 168 / 140
    const int sat = 178 + (lap % 2) * 40;    // 178 / 218
    return QColor::fromHsv(hue, qBound(0, sat, 255), qBound(0, val, 255));
}

QVector<QColor> AppTheme::seriesColors(int n)
{
    QVector<QColor> out;
    if (n <= 0) return out;
    out.reserve(n);
    for (int i = 0; i < n; ++i) out.append(seriesColor(i));
    return out;
}

QColor AppTheme::shade(const QColor& base, int idx, int count)
{
    if (count <= 1) return base;
    if (idx < 0) idx = 0;
    if (idx > count - 1) idx = count - 1;

    int h, s, v, a;
    base.getHsv(&h, &s, &v, &a);
    if (a < 0) a = 255;

    // Spread brightness across a projector-safe band (≈0.55→0.85 of full) so
    // even the lightest shade stays readable against white — never toward 1.0.
    const double t    = double(idx) / double(count - 1);   // 0 → 1
    const int    newV = int(std::lround(140.0 + t * (217.0 - 140.0)));

    if (s < 20)  // achromatic base (e.g. the gray family): stay neutral
        return QColor::fromHsv(0, 0, qBound(0, newV, 255), a);

    const int newS = qMax(s, 130);  // floor saturation so light shades aren't pastel
    return QColor::fromHsv(h, qBound(0, newS, 255), qBound(0, newV, 255), a);
}
```

- [ ] **Step 7: Run the test to verify it PASSES**

```powershell
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
```
Expected: `tst_apptheme` builds and all 7 slots pass; no other suite regresses.

- [ ] **Step 8: Commit**

```bash
git add src/utils/AppTheme.h src/utils/AppTheme.cpp tests/tst_apptheme tests/tests.pro
git commit -m "feat(theme): shared distinct-color plot palette (no-yellow, projector-safe)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: Report charts — comparison chart (grouped, no reuse) + per-sheet plots

**Files:**
- Modify: `src/reporting/ReportGenerator.h:43-45` (signature + doc)
- Modify: `src/reporting/ReportGenerator.cpp` (`lifetimeBarColor`, comparison loop ~837-859, `buildPlots` two palettes ~75-86 & ~155-166; add include)
- Modify: `tests/tst_reportgenerator/tst_reportgenerator.pro:21` (link `AppTheme.cpp`)

- [ ] **Step 1: Add the include** at the top of `src/reporting/ReportGenerator.cpp` (after the existing includes, alongside the other `../` includes):

```cpp
#include "../utils/AppTheme.h"
```

- [ ] **Step 2: Re-point `lifetimeBarColor` to the shared helper.** In `ReportGenerator.h`, replace lines 43-45:

```cpp
    /// Bar color for the Lifetime Comparison slide. fileIdx selects the hue
    /// from kFileHues; sampleIdx within the file shifts HSV value from 0.6→1.0.
    static QColor lifetimeBarColor(int fileIdx, int sampleIdx, int totalSamplesInFile);
```
with:
```cpp
    /// Bar color for the Lifetime Comparison slide. The file's distinct base
    /// hue (from AppTheme::seriesColors) is shaded per sample so every bar in
    /// the chart is a unique, projector-safe color while staying grouped by file.
    static QColor lifetimeBarColor(const QColor& fileBase, int sampleIdx, int totalSamplesInFile);
```

- [ ] **Step 3: Replace the body of `lifetimeBarColor`** in `ReportGenerator.cpp` (lines 46-65) with:

```cpp
QColor ReportGenerator::lifetimeBarColor(const QColor& fileBase, int sampleIdx, int totalSamplesInFile)
{
    return AppTheme::shade(fileBase, sampleIdx, totalSamplesInFile);
}
```

- [ ] **Step 4: Give each file a distinct base hue in the comparison loop.** In `ReportGenerator.cpp`, immediately before `for (int fi = 0; fi < files.size(); ++fi) {` (currently line 837), add:

```cpp
        // One distinct, well-separated base hue per file → strong between-file
        // contrast; shade() varies samples within each file.
        const QVector<QColor> fileHues = AppTheme::seriesColors(files.size());
```
Then change the color append (currently line 857) from:
```cpp
                colors.append(lifetimeBarColor(fi, sIdx, totalSamples));
```
to:
```cpp
                colors.append(lifetimeBarColor(fileHues[fi], sIdx, totalSamples));
```

- [ ] **Step 5: Route the per-sheet report plots through the shared palette.** In `buildPlots`, add one palette sized to the sheet's samples at the top of the function (just after `QVector<QByteArray> plots;`, line 71):

```cpp
    const QVector<QColor> palette = AppTheme::seriesColors(sheet.samples.size());
```
In **Plot 1 (TPM Trend)** delete the local array (lines 75-79):
```cpp
        static const QColor kColors[] = {
            QColor(0x00, 0x66, 0xCC), QColor(0xFF, 0x73, 0x00),
            QColor(0x00, 0xAA, 0x44), QColor(0xCC, 0x00, 0x00),
            QColor(0x99, 0x00, 0xCC), QColor(0x00, 0xAA, 0xCC),
        };
```
and change `ps.color = kColors[colorIdx++ % 6];` (line 86) to:
```cpp
            ps.color     = palette[colorIdx++];
```
In **Plot 3 (Draw Pressure)** delete the identical local array (lines 155-159) and change `ps.color = kColors[colorIdx++ % 6];` (line 166) to:
```cpp
            ps.color     = palette[colorIdx++];
```

- [ ] **Step 6: Link `AppTheme.cpp` into the test project.** In `tests/tst_reportgenerator/tst_reportgenerator.pro`, change line 21 from:
```pro
           ../../src/utils/SopLoader.cpp
```
to:
```pro
           ../../src/utils/SopLoader.cpp \
           ../../src/utils/AppTheme.cpp
```

- [ ] **Step 7: Build and run the report suite**

```bash
python tools/decrypt_via_copy.py --apply
```
```powershell
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
```
Expected: `tst_reportgenerator` and `tst_apptheme` pass; no regressions. (The comparison chart's no-reuse property is proven by `testGroupedComparisonUnique` in Task 1.)

- [ ] **Step 8: Commit**

```bash
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp tests/tst_reportgenerator/tst_reportgenerator.pro
git commit -m "feat(report): distinct colors per file/sample in report charts

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: `PlotEngine` — default fallback + per-sheet bar chart

**Files:**
- Modify: `src/plotting/PlotEngine.cpp` (fallback ~451-468; `renderTPMBarChart` palette ~989-1001; add include)
- Modify: `tests/tst_plotengine/tst_plotengine.pro:13` (link `AppTheme.cpp`)

- [ ] **Step 1: Add the include** at the top of `src/plotting/PlotEngine.cpp`:

```cpp
#include "../utils/AppTheme.h"
```

- [ ] **Step 2: Replace the legend fallback palette.** Replace lines 451-468:

```cpp
    // Build resolved color list for legend
    static const QColor kDefaultColors[] = {
        QColor(0x00, 0x66, 0xCC),
        QColor(0xFF, 0x73, 0x00),
        QColor(0x00, 0xAA, 0x44),
        QColor(0xCC, 0x00, 0x00),
        QColor(0x99, 0x00, 0xCC),
        QColor(0x00, 0xAA, 0xCC),
        QColor(0xCC, 0xAA, 0x00),
        QColor(0x66, 0x66, 0x66),
    };
    static const int kNDefaultColors = sizeof(kDefaultColors) / sizeof(kDefaultColors[0]);

    QVector<QColor> resolvedColors(n);
    for (int i = 0; i < n; ++i) {
        if (i < colors.size())
            resolvedColors[i] = colors[i];
        else
            resolvedColors[i] = kDefaultColors[i % kNDefaultColors];
    }
```
with:
```cpp
    // Build resolved color list for legend. Caller-supplied colors win; any
    // remainder comes from the shared distinct palette (no reuse, no yellow).
    const QVector<QColor> fallback = AppTheme::seriesColors(n);
    QVector<QColor> resolvedColors(n);
    for (int i = 0; i < n; ++i)
        resolvedColors[i] = (i < colors.size()) ? colors[i] : fallback[i];
```

- [ ] **Step 3: Replace the all-blue per-sheet bar palette.** Replace lines 990-1002:

```cpp
    // Professional blue palette, one shade per bar
    QVector<QColor> colors;
    for (int i = 0; i < sampleNames.size(); ++i) {
        // Cycle through a set of blues / teals
        static const QColor kCols[] = {
            QColor(0x00, 0x66, 0xCC),
            QColor(0x00, 0x8B, 0xD8),
            QColor(0x00, 0x6E, 0xA6),
            QColor(0x00, 0x9E, 0xC7),
            QColor(0x1F, 0x4E, 0x79),
            QColor(0x2E, 0x75, 0xB6),
        };
        colors.append(kCols[i % 6]);
    }
```
with:
```cpp
    // One distinct color per bar (matches the sample's trend-line color).
    const QVector<QColor> colors = AppTheme::seriesColors(sampleNames.size());
```

- [ ] **Step 4: Link `AppTheme.cpp` into the test project.** In `tests/tst_plotengine/tst_plotengine.pro`, change line 13 from:
```pro
SOURCES += ../../src/plotting/PlotEngine.cpp
```
to:
```pro
SOURCES += ../../src/plotting/PlotEngine.cpp
SOURCES += ../../src/utils/AppTheme.cpp
```

- [ ] **Step 5: Build and run the plot-engine suite**

```bash
python tools/decrypt_via_copy.py --apply
```
```powershell
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
```
Expected: `tst_plotengine` builds and passes; no regressions.

- [ ] **Step 6: Commit**

```bash
git add src/plotting/PlotEngine.cpp tests/tst_plotengine/tst_plotengine.pro
git commit -m "feat(plot): PlotEngine fallback + bar chart use shared distinct palette

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: On-screen `PlotWidget` — distinct colors, stable per sample

**Files:**
- Modify: `src/plotting/PlotWidget.cpp` (checkbox swatches ~163-223; TPM-trend ~406-466; Power Density ~512-526; Draw Pressure ~560-574; add include)

No unit test exists for `PlotWidget` (not in SUBDIRS); verify via the app build. Colors are keyed on the **stable sample index** so a sample keeps its color as others toggle and the oil overlay matches.

- [ ] **Step 1: Add the include** at the top of `src/plotting/PlotWidget.cpp`:

```cpp
#include "../utils/AppTheme.h"
```

- [ ] **Step 2: Replace the checkbox-swatch palette.** Delete lines 163-168:

```cpp
    static const QColor kCbColors[] = {
        QColor(0x00, 0x66, 0xCC), QColor(0xFF, 0x73, 0x00),
        QColor(0x00, 0xAA, 0x44), QColor(0xCC, 0x00, 0x00),
        QColor(0x99, 0x00, 0xCC), QColor(0x00, 0xAA, 0xCC),
    };
    static const int kNCbColors = sizeof(kCbColors) / sizeof(kCbColors[0]);
```
Change line 181 `QColor c = kCbColors[i % kNCbColors];` to:
```cpp
        QColor c = AppTheme::seriesColor(i);
```
Change line 223 `QColor oc = kCbColors[i % kNCbColors].darker(120);` to:
```cpp
        QColor oc = AppTheme::seriesColor(i).darker(120);
```

- [ ] **Step 3: Replace the TPM-Trend palette.** Delete lines 406-410:

```cpp
        static const QColor kColors[] = {
            QColor(0x00, 0x66, 0xCC), QColor(0xFF, 0x73, 0x00),
            QColor(0x00, 0xAA, 0x44), QColor(0xCC, 0x00, 0x00),
            QColor(0x99, 0x00, 0xCC), QColor(0x00, 0xAA, 0xCC),
        };
```
Change line 424 `ps.color = kColors[si % 6];  // index-based so oil overlay always matches` to:
```cpp
            ps.color     = AppTheme::seriesColor(si);  // stable per sample; oil overlay matches
```
Change line 466 `ps.color = kColors[si % 6];  // same color as this sample's TPM line` to:
```cpp
            ps.color     = AppTheme::seriesColor(si);  // same color as this sample's TPM line
```

- [ ] **Step 4: Replace the Power Density palette and its counter.** Delete lines 512-516:

```cpp
        static const QColor kColors[] = {
            QColor(0x00, 0x66, 0xCC), QColor(0xFF, 0x73, 0x00),
            QColor(0x00, 0xAA, 0x44), QColor(0xCC, 0x00, 0x00),
            QColor(0x99, 0x00, 0xCC), QColor(0x00, 0xAA, 0xCC),
        };
```
Delete line 518 `int colorIdx = 0;`. Change line 526 `ps.color = kColors[colorIdx++ % 6];` to:
```cpp
            ps.color     = AppTheme::seriesColor(si);  // stable per sample
```

- [ ] **Step 5: Replace the Draw Pressure palette and its counter.** Delete lines 560-564 (identical `kColors` array), delete line 566 `int colorIdx = 0;`, and change line 574 `ps.color = kColors[colorIdx++ % 6];` to:
```cpp
            ps.color     = AppTheme::seriesColor(si);  // stable per sample
```

- [ ] **Step 6: Build the app to verify it compiles (`-Werror` proves no orphaned `colorIdx`/arrays)**

```bash
python tools/decrypt_via_copy.py --apply
```
```bat
mkdir build & cd build
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro
mingw32-make -j8
```
Expected: clean build, no warnings/errors.

- [ ] **Step 7: Commit**

```bash
git add src/plotting/PlotWidget.cpp
git commit -m "feat(plot): on-screen plots use shared distinct palette (stable per sample)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: `RadarChartWidget` — consume the shared palette

**Files:**
- Modify: `src/ui/RadarChartWidget.h:40` (remove `kColors` member)
- Modify: `src/ui/RadarChartWidget.cpp` (remove `kColors` definition lines 11-37; replace all indexed uses; add include)
- Modify: `tests/tst_sensoryreportsource/tst_sensoryreportsource.pro:22` (link `AppTheme.cpp`)

- [ ] **Step 1: Add the include** at the top of `src/ui/RadarChartWidget.cpp` (after `#include <QtMath>`):

```cpp
#include "../utils/AppTheme.h"
```

- [ ] **Step 2: Remove the private palette definition.** Delete lines 11-37 of `RadarChartWidget.cpp` (the comment block + `const QList<QColor> RadarChartWidget::kColors = { … };`).

- [ ] **Step 3: Remove the member declaration.** In `RadarChartWidget.h`, delete line 40:
```cpp
    static const QList<QColor> kColors;
```

- [ ] **Step 4: Replace the two indexed-access forms with the shared accessor** (no `% size()` ⇒ never repeats). In `RadarChartWidget.cpp`:
  - Replace **all** occurrences of `kColors[colorIdx % kColors.size()]` with `AppTheme::seriesColor(colorIdx)`.
  - Replace **all** occurrences of `kColors[ci % kColors.size()]` with `AppTheme::seriesColor(ci)`.

(These two literal forms cover every use: lines ~368, ~377, ~445, ~455, ~570, ~577, ~611, ~618.)

- [ ] **Step 5: Link `AppTheme.cpp` into the test project.** In `tests/tst_sensoryreportsource/tst_sensoryreportsource.pro`, change line 22 from:
```pro
           ../../src/ui/RadarChartWidget.cpp
```
to:
```pro
           ../../src/ui/RadarChartWidget.cpp \
           ../../src/utils/AppTheme.cpp
```

- [ ] **Step 6: Build and run the sensory suite**

```bash
python tools/decrypt_via_copy.py --apply
```
```powershell
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
```
Expected: `tst_sensoryreportsource` builds and passes; no `kColors` references remain (a stray one would fail `-Werror`/link).

- [ ] **Step 7: Commit**

```bash
git add src/ui/RadarChartWidget.h src/ui/RadarChartWidget.cpp tests/tst_sensoryreportsource/tst_sensoryreportsource.pro
git commit -m "refactor(radar): consume shared AppTheme palette, drop private copy

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: Codify the rule + full verification

**Files:**
- Modify: `CLAUDE.md` (Conventions section)
- Modify: `tasks/lessons.md` (append entry)

- [ ] **Step 1: Add the rule to `CLAUDE.md`.** Under `## Conventions`, after the `tasks/lessons.md` bullet, add:

```markdown
- **Never reuse a color within a single plot.** Every series/bar/slice gets a
  distinct, well-separated color. Never index a fixed palette with `% N` (that
  wraparound is exactly what causes repeats) — size the palette to the element
  count via `AppTheme::seriesColors(n)` / `AppTheme::seriesColor(i)`, or
  `AppTheme::shade(base, i, count)` for grouped (per-file) charts. No yellow
  (hue ≈ 45°–70°); keep colors saturated and mid-dark so they read on a
  projector against white. Beyond the curated 20-color palette, colors are
  generated by golden-angle rotation — never fall back to repeating.
```

- [ ] **Step 2: Append a lessons entry to `tasks/lessons.md`**

```markdown
## Plot colors — never reuse within a plot (2026-05-29)

Plot palettes were small fixed arrays indexed with `% N`, so colors repeated
once series count exceeded the palette (the Lifetime comparison chart "only blue
and orange"). Fix: one shared `AppTheme::seriesColor/seriesColors/shade` helper,
sized to the element count, reusing the radar chart's vetted no-yellow palette
with golden-angle generation beyond 20. Rule: never `palette[i % N]` in a plot.
Yellow and pale/light colors wash out on projectors — keep series saturated and
mid-dark.
```

- [ ] **Step 3: Full clean build + full test suite**

```bash
python tools/decrypt_via_copy.py --apply
```
```powershell
powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1
```
Expected: every suite passes (incl. `tst_apptheme`, `tst_reportgenerator`, `tst_plotengine`, `tst_sensoryreportsource`).

- [ ] **Step 4: Visual confirmation** — open a multi-file dataset, generate a Combined Full Report, and confirm on the *Lifetime TPM Comparison* slide: each file is a distinct color family, every bar is a different shade, nothing yellow, nothing washed out. Spot-check a per-sheet TPM-trend/bar slide (now multi-color) and a sensory radar with several samples.

- [ ] **Step 5: Commit**

```bash
git add CLAUDE.md tasks/lessons.md
git commit -m "docs: codify never-reuse-a-color-in-a-plot rule + lesson

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Self-Review

**Spec coverage:**
- Invariant (no reuse, palette sized to count) → Task 1 helper + every call-site task.
- No yellow → curated palette has none; generated path skips 45–70 (Task 1 `testNoYellow`).
- Projector-safe → `shade` value band ≤217, saturation floor; generated sat/val pinned (Task 1).
- Comparison chart grouped + provably unique → Task 2 + `testGroupedComparisonUnique`.
- All other report plots → Task 2 Step 5; PlotEngine → Task 3.
- App-wide (GUI + radar) → Tasks 4 & 5.
- Rule in project CLAUDE.md → Task 6.
- Single source of truth (reuse radar palette) → Task 1 curated array + Task 5 drops the copy.

**Placeholder scan:** none — every code/command step is concrete.

**Type consistency:** `seriesColor(int)→QColor`, `seriesColors(int)→QVector<QColor>`, `shade(const QColor&,int,int)→QColor` used identically in Tasks 2–5; `lifetimeBarColor` new signature `(const QColor&,int,int)` matches its one caller.

**Build-graph check:** every test `.pro` that compiles a touched source now also compiles `AppTheme.cpp` (Tasks 1/2/3/5); the app `.pro` already compiles all of `src/`.

---

## Notes / risk

- `PlotWidget` has no unit test → relies on the `-Werror` app build (Task 4 Step 6) to catch orphaned `kColors`/`colorIdx`, plus the Task 6 visual check.
- The per-sheet bar chart changes from all-blue to multi-color by design (spec-approved).
- A few curated entries (cyan #6, deep-sky-blue #15, magenta #14) are brighter; they come from the team's vetted radar palette and only appear at higher series counts. Easy to swap a specific index later if any reads too light in the room.
