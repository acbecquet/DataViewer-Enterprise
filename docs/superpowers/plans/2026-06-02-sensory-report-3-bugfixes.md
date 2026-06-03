# Sensory-Report 3 Bugfixes Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix three bugs in the simple Sensory report workflow: (1) the generated `.pptx` should honor *every* element position the user arranges in the report Preview (full WYSIWYG), (2) add a `Round` (1/2/N/A) dropdown beside Tester that folds into the filename only, (3) raise the Smoothness/Burnt-Taste radar labels ~3 px.

**Architecture:** Bug 1 extends `PptxWriter`'s existing "null `QRectF` = keep baked-in position, non-null = override" convention (already used for table/radar) to the slide title, cover title/subtitle, divider title, and properties box — so the pptx renders from the same `ReportLayout` the preview canvas already uses. Bug 2 is UI-only: a `Round` combo whose value is folded into `SensorySession::testerName` (kept as the combined `"Charlie R1"` string) via two pure helper functions, split back out on load. Bug 3 nudges two label anchors in `RadarChartWidget`.

**Tech Stack:** C++17, Qt 6.10 (Widgets/Test/Core), qmake + MinGW 13.1, QtTest, `QZipReader` (`<private/qzipreader_p.h>`) for pptx XML assertions.

**Spec:** `docs/superpowers/specs/2026-06-02-sensory-report-3-bugfixes-design.md`

**Branch:** Work on `dev` (current branch). Commit after each task.

---

## Preflight (read once before starting)

**This is a Windows + MIP machine.** Two standing rules from `CLAUDE.md`:

1. **Create NEW source files via Python delete-and-rewrite** so they don't inherit a MIP label. Pattern (run through the Bash tool):

   ```python
   python -c "
   import os
   p = r'C:/full/path/to/File.h'
   content = '''...file contents...'''
   if os.path.exists(p): os.remove(p)
   open(p, 'w', encoding='utf-8', newline='\n').write(content)
   "
   ```
   Editing EXISTING files with the Edit tool is fine — just decrypt before building.

2. **Decrypt before every C++ build:** `python tools/decrypt_via_copy.py --apply` (idempotent; from repo root).

**Build / run commands** (the execution shell is Git Bash; call PowerShell for the runner):

- Run one test suite (builds incrementally, then runs the filtered test):
  ```bash
  python tools/decrypt_via_copy.py --apply
  powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter <name>
  ```
- After ADDING a new test dir to `tests/tests.pro`, force a re-qmake the first time:
  ```bash
  powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Rebuild -Filter <name>
  ```
- Build the full app (only needed for `SensoryPanel.cpp`, Task 3, and final Task 8):
  ```bash
  python tools/decrypt_via_copy.py --apply
  export PATH="/c/Qt/Tools/mingw1310_64/bin:/c/Qt/6.10.1/mingw_64/bin:$PATH"
  cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/build" \
    && /c/Qt/6.10.1/mingw_64/bin/qmake.exe -spec win32-g++ ../DataViewerEnterprise.pro \
    && mingw32-make -j8 2>&1 | tail -60
  ```
  (If `build/` does not exist: `mkdir build` first. Qt is pinned at 6.10.1 per `CLAUDE.md`; adjust the minor version only if that path is absent.)

## File Structure

**Created:**
- `src/ui/TesterRound.h` — header-only pure helpers `splitTesterRound()` / `combineTesterRound()` + `TesterRound` struct (Bug 2). Header-only so the unit test links it with zero extra sources.
- `tests/tst_testerround/tst_testerround.cpp` + `tests/tst_testerround/tst_testerround.pro` — unit tests for the helpers.

**Modified:**
- `src/ui/RadarChartWidget.cpp` — Bug 3 anchor nudge (1 line).
- `src/ui/SensoryPanel.h` / `src/ui/SensoryPanel.cpp` — Bug 2 Round combo + folding.
- `src/reporting/PptxWriter.h` / `.cpp` — Bug 1 title rect (content + cover/subtitle).
- `src/reporting/SensoryReportSource.h` / `.cpp` — Bug 1 properties-box helper + cover/divider rect wiring.
- `tests/tst_pptxwriter/tst_pptxwriter.cpp` — Bug 1 title/cover-rect assertions.
- `tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp` — Bug 1 properties-box assertions.
- `tests/tests.pro` — register `tst_testerround`.

---

## Task 1: Bug 3 — raise Smoothness & Burnt Taste radar labels ~3 px

**Files:**
- Modify: `src/ui/RadarChartWidget.cpp:211-213`

No automated test (sub-pixel visual tweak; the widget exposes no label-rect API). Verified by a compile + a user visual check.

- [ ] **Step 1: Make the edit**

In the `i == 1 || i == 4` branch of `drawAxisLabels`, change the anchor's y term from `+ 2.0` to `- 1.0` (a net −3 px; Qt's +y is down, so this raises both labels). Replace:

```cpp
            anchor = QPointF(
                center.x() + labelCenter.x(),
                center.y() + labelCenter.y() + labelHalfH + 2.0);
```

with:

```cpp
            anchor = QPointF(
                center.x() + labelCenter.x(),
                center.y() + labelCenter.y() + labelHalfH - 1.0);  // raise ~3px (was +2.0) so the Smoothness/Burnt Taste gap matches the bottom labels
```

- [ ] **Step 2: Compile-check (RadarChartWidget is built by the sensoryreportsource suite)**

Run:
```bash
python tools/decrypt_via_copy.py --apply
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter sensoryreportsource
```
Expected: builds clean; `tst_sensoryreportsource` runs (Postgres-dependent slots skip when `DVE_TEST_PG_CONN` is unset — that's fine). No failures.

- [ ] **Step 3: Commit**

```bash
git add src/ui/RadarChartWidget.cpp
git commit -m "fix(sensory): raise Smoothness/Burnt Taste radar labels ~3px to match bottom-label gap"
```

> **Manual check (defer to user at Task 8):** open Sensory mode; confirm the Smoothness/Burnt Taste gap from the pentagon visually matches Overall Flavor/Vapor Volume. If it's off, nudge the `- 1.0` by ±1.0.

---

## Task 2: Bug 2 — `TesterRound` split/combine helpers (TDD, new test suite)

**Files:**
- Create: `tests/tst_testerround/tst_testerround.cpp`
- Create: `tests/tst_testerround/tst_testerround.pro`
- Modify: `tests/tests.pro` (register the new suite)
- Create: `src/ui/TesterRound.h` (implementation — written AFTER the test fails)

- [ ] **Step 1: Write the failing test file** (create via Python delete-and-rewrite)

Write to `tests/tst_testerround/tst_testerround.cpp`:

```cpp
#include <QtTest>
#include "ui/TesterRound.h"

using namespace DVE;

class tst_TesterRound : public QObject
{
    Q_OBJECT
private slots:
    void splitParsesR1() {
        const TesterRound tr = splitTesterRound("Charlie R1");
        QCOMPARE(tr.tester, QString("Charlie"));
        QCOMPARE(tr.round,  QString("1"));
    }
    void splitParsesR2WithSpaceInName() {
        const TesterRound tr = splitTesterRound("Mary Jane R2");
        QCOMPARE(tr.tester, QString("Mary Jane"));
        QCOMPARE(tr.round,  QString("2"));
    }
    void splitPlainNameYieldsNA() {
        const TesterRound tr = splitTesterRound("Charlie");
        QCOMPARE(tr.tester, QString("Charlie"));
        QCOMPARE(tr.round,  QString("N/A"));
    }
    void splitUnknownRoundIsNotParsed() {
        // Only R1/R2 are round markers; "R3" stays part of the name.
        const TesterRound tr = splitTesterRound("Charlie R3");
        QCOMPARE(tr.tester, QString("Charlie R3"));
        QCOMPARE(tr.round,  QString("N/A"));
    }
    void splitEmptyYieldsNA() {
        const TesterRound tr = splitTesterRound("");
        QCOMPARE(tr.tester, QString(""));
        QCOMPARE(tr.round,  QString("N/A"));
    }
    void combineRound1AppendsSuffix() {
        QCOMPARE(combineTesterRound("Charlie", "1"), QString("Charlie R1"));
    }
    void combineRound2AppendsSuffix() {
        QCOMPARE(combineTesterRound("Charlie", "2"), QString("Charlie R2"));
    }
    void combineNAAppendsNothing() {
        QCOMPARE(combineTesterRound("Charlie", "N/A"), QString("Charlie"));
    }
    void combineTrimsAndKeepsEmptyEmpty() {
        QCOMPARE(combineTesterRound("  Charlie  ", "1"), QString("Charlie R1"));
        QCOMPARE(combineTesterRound("", "1"), QString(""));
        QCOMPARE(combineTesterRound("   ", "2"), QString(""));
    }
    void roundTripAllRounds() {
        const QStringList rounds{ "1", "2", "N/A" };
        for (const QString& r : rounds) {
            const QString combined = combineTesterRound("Charlie", r);
            const TesterRound tr = splitTesterRound(combined);
            QCOMPARE(tr.tester, QString("Charlie"));
            QCOMPARE(tr.round,  r);
        }
    }
};

QTEST_APPLESS_MAIN(tst_TesterRound)
#include "tst_testerround.moc"
```

Write to `tests/tst_testerround/tst_testerround.pro`:

```qmake
QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src

SOURCES += tst_testerround.cpp

HEADERS += ../../src/ui/TesterRound.h
```

- [ ] **Step 2: Register the suite in `tests/tests.pro`** (Edit tool)

Add `tst_testerround` to the `SUBDIRS` list. Change the last line:

```qmake
    tst_apptheme \
    tst_regimeutils
```

to:

```qmake
    tst_apptheme \
    tst_regimeutils \
    tst_testerround
```

- [ ] **Step 3: Run the test — verify it FAILS to build**

Run:
```bash
python tools/decrypt_via_copy.py --apply
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Rebuild -Filter testerround
```
Expected: BUILD FAILS — `ui/TesterRound.h: No such file or directory` (the header doesn't exist yet).

- [ ] **Step 4: Create the implementation** `src/ui/TesterRound.h` (via Python delete-and-rewrite)

```cpp
#pragma once
#include <QString>
#include <QRegularExpression>

namespace DVE {

// A sensory "tester" string split into a display name + round marker.
// Double-blind rounds 1/2 are encoded as a trailing " R1"/" R2" suffix on the
// stored testerName (the long-standing hand-typed convention). Anything else
// yields round == "N/A".
struct TesterRound {
    QString tester;   // name without the round suffix
    QString round;    // "1", "2", or "N/A"
};

// Parse a stored testerName. Trailing " R1"/" R2" (after at least one
// non-space char) is treated as the round; otherwise round = "N/A".
inline TesterRound splitTesterRound(const QString& stored)
{
    static const QRegularExpression re(QStringLiteral("^(.*\\S)\\s+R([12])$"));
    const QRegularExpressionMatch m = re.match(stored);
    if (m.hasMatch())
        return { m.captured(1), m.captured(2) };
    return { stored, QStringLiteral("N/A") };
}

// Recombine into the stored testerName. Empty tester stays empty (round is
// meaningless without a tester, and SensoryPanel::sessionLabel() must still
// fall back to the assessor). Round "N/A" (or anything other than 1/2)
// appends nothing.
inline QString combineTesterRound(const QString& tester, const QString& round)
{
    const QString t = tester.trimmed();
    if (t.isEmpty()) return t;
    if (round == QLatin1String("1")) return t + QStringLiteral(" R1");
    if (round == QLatin1String("2")) return t + QStringLiteral(" R2");
    return t;
}

} // namespace DVE
```

- [ ] **Step 5: Run the test — verify it PASSES**

Run:
```bash
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter testerround
```
Expected: PASS (10 slots, 0 failed).

- [ ] **Step 6: Commit**

```bash
git add src/ui/TesterRound.h tests/tst_testerround/ tests/tests.pro
git commit -m "feat(sensory): add TesterRound split/combine helpers + unit tests"
```

---

## Task 3: Bug 2 — wire the `Round` combo into `SensoryPanel`

**Files:**
- Modify: `src/ui/SensoryPanel.h:228` (add member)
- Modify: `src/ui/SensoryPanel.cpp` (include, header row, buildSession, applySession, two resets, local NoWheelComboBox)

No unit test (UI wiring; `SensoryPanel` has no test harness). Verified by an app compile + user UI check. The folding logic itself is already covered by Task 2.

- [ ] **Step 1: Add the member** to `src/ui/SensoryPanel.h` (Edit tool). `QComboBox` is already included (line 14). Replace:

```cpp
    QLineEdit*        m_testerEdit;
    QLineEdit*        m_mediaEdit;
```

with:

```cpp
    QLineEdit*        m_testerEdit;
    QComboBox*        m_roundCombo = nullptr;   // Bug 2: double-blind round (1/2/N/A)
    QLineEdit*        m_mediaEdit;
```

- [ ] **Step 2: Add the TesterRound include + a wheel-safe combo** in `src/ui/SensoryPanel.cpp` (Edit tool). `<QWheelEvent>` (line 21) and `<QComboBox>` (via the header) are already available, and the file already defines a `NoWheelDoubleSpinBox` (DVE namespace, ~line 109) — mirror it.

First add the helper include to the second include group. Replace:

```cpp
#include "reporting/SensoryReportSource.h"
#include "ReportPreviewDialog.h"
```

with:

```cpp
#include "reporting/SensoryReportSource.h"
#include "ReportPreviewDialog.h"
#include "ui/TesterRound.h"
```

Then add `NoWheelComboBox` immediately after the existing `NoWheelDoubleSpinBox` class. Replace:

```cpp
    void wheelEvent(QWheelEvent* e) override {
        e->ignore();  // never consume scroll — always let parent scroll area handle it
    }
};

// ─────────────────────────────────────────────────────────────────────────────
// FlowLayout — cards wrap left-to-right, then down
```

with:

```cpp
    void wheelEvent(QWheelEvent* e) override {
        e->ignore();  // never consume scroll — always let parent scroll area handle it
    }
};

// Ignores wheel events unless focused — prevents accidental Round changes
// when scrolling past the header row.
class NoWheelComboBox : public QComboBox
{
public:
    explicit NoWheelComboBox(QWidget* parent = nullptr) : QComboBox(parent) {
        setFocusPolicy(Qt::StrongFocus);
    }
protected:
    void wheelEvent(QWheelEvent* e) override {
        e->ignore();
    }
};

// ─────────────────────────────────────────────────────────────────────────────
// FlowLayout — cards wrap left-to-right, then down
```

- [ ] **Step 3: Insert the Round field** in `buildHeaderRow`, between Tester and Media (Edit tool). Replace:

```cpp
    addField("Tester:",    m_testerEdit);
    addField("Media:",     m_mediaEdit);
```

with:

```cpp
    addField("Tester:",    m_testerEdit);

    // Bug 2: Round selector for double-blind R1/R2. Folded into testerName via
    // combineTesterRound() at buildSession() time; split back out in
    // applySession(). Default "1" (most tests are double-blind round 1 first).
    layout->addWidget(new QLabel("Round:"));
    m_roundCombo = new NoWheelComboBox;
    m_roundCombo->addItems({QStringLiteral("1"), QStringLiteral("2"), QStringLiteral("N/A")});
    m_roundCombo->setCurrentIndex(0);
    m_roundCombo->setFixedWidth(60);
    layout->addWidget(m_roundCombo);

    addField("Media:",     m_mediaEdit);
```

- [ ] **Step 4: Fold Round into testerName** in `buildSession` (Edit tool). Replace line 891:

```cpp
    sess.testerName   = m_testerEdit->text().trimmed();
```

with:

```cpp
    sess.testerName   = combineTesterRound(m_testerEdit->text(), m_roundCombo->currentText());
```

- [ ] **Step 5: Split Round back out** in `applySession` (Edit tool). Replace line 947:

```cpp
    m_testerEdit->setText(session.testerName);
```

with:

```cpp
    const TesterRound tr = splitTesterRound(session.testerName);
    m_testerEdit->setText(tr.tester);
    if (m_roundCombo) m_roundCombo->setCurrentText(tr.round);
```

- [ ] **Step 6: Reset Round on new session** (`newSession`, Edit tool). Replace:

```cpp
    m_testTitleEdit->clear();
    m_assessorEdit->clear();
    m_testerEdit->clear();
    m_mediaEdit->clear();
    m_dateLabel->setText(empty.date);
```

with:

```cpp
    m_testTitleEdit->clear();
    m_assessorEdit->clear();
    m_testerEdit->clear();
    if (m_roundCombo) m_roundCombo->setCurrentIndex(0);   // default round "1"
    m_mediaEdit->clear();
    m_dateLabel->setText(empty.date);
```

- [ ] **Step 7: Reset Round on close-all** (`closeSessions`, Edit tool). Replace:

```cpp
        m_testTitleEdit->clear();
        m_assessorEdit->clear();
        m_testerEdit->clear();
        m_mediaEdit->clear();
        m_dateLabel->clear();
```

with:

```cpp
        m_testTitleEdit->clear();
        m_assessorEdit->clear();
        m_testerEdit->clear();
        if (m_roundCombo) m_roundCombo->setCurrentIndex(0);   // default round "1"
        m_mediaEdit->clear();
        m_dateLabel->clear();
```

> **Note — no remote-path change needed.** `onRemoteCellChanged` (SensoryPanel.cpp:2289) early-returns unless the column starts with `json_path:samples[` (line 2294); it never writes header fields, so `tester_name` only ever reaches the UI through `applySession` (Step 5). No other edit required.

- [ ] **Step 8: Build the full app to compile-check**

Run the full-app build from Preflight. Expected: links clean, produces `DataViewer.exe`. Fix any compile errors before committing.

- [ ] **Step 9: Commit**

```bash
git add src/ui/SensoryPanel.h src/ui/SensoryPanel.cpp
git commit -m "feat(sensory): add Round (1/2/N/A) field; fold into testerName for filename only"
```

> **Manual check (defer to Task 8):** header shows `Round:` between Tester and Media, default "1". Type Tester "Charlie", Round "1", Save → filename ends `- Charlie R1`. Reopen a legacy session whose tester was `Charlie R2` → Tester shows "Charlie", Round shows "2".

---

## Task 4: Bug 1 — `PptxWriter` honors the content-slide title rect (TDD)

**Files:**
- Modify: `tests/tst_pptxwriter/tst_pptxwriter.cpp` (new test slot)
- Modify: `src/reporting/PptxWriter.h:183-190` (buildContentSlideXml signature)
- Modify: `src/reporting/PptxWriter.cpp` (buildContentSlideXml body; both addContentSlide overloads)

- [ ] **Step 1: Write the failing test** — add this slot to `tst_pptxwriter.cpp`, just before the closing `};` of the class (Edit tool):

```cpp
    // ── Bug 1: layout.title rect lands in slide XML (content-slide title) ──
    void addContentSlide_titleRectAppliesToEmuPosition()
    {
        DVE::PptxWriter w;
        DVE::SlideTable table = makeSimpleTable(2, 2);
        QVector<DVE::SlideImage> noPlots;

        DVE::ContentSlideLayout layout;
        // 2.0" x, 0.6" y → 1828800 / 548640 EMU. Distinct from the hardcoded
        // title box (0.4", 0.1"), so a pass proves the rect was honored.
        layout.title = QRectF(2.0, 0.6, 9.0, 0.7);

        w.addContentSlide("Title Override", table, noPlots, layout);
        const QString path = tempPath("title_override.pptx");
        QVERIFY(w.save(path));
        QZipReader zr(path);
        const QByteArray slideXml =
            zr.fileData(QStringLiteral("ppt/slides/slide1.xml"));
        QVERIFY(!slideXml.isEmpty());
        QVERIFY2(slideXml.contains("1828800"),
                 "title.x override (2.0\" = 1828800 EMU) missing — title rect not honored");
        QVERIFY2(slideXml.contains("548640"),
                 "title.y override (0.6\" = 548640 EMU) missing — title rect not honored");
    }
```

- [ ] **Step 2: Run — verify it FAILS**

Run:
```bash
python tools/decrypt_via_copy.py --apply
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter pptxwriter
```
Expected: `addContentSlide_titleRectAppliesToEmuPosition` FAILS (title still emitted at the hardcoded 0.4"/0.1"; `1828800` absent).

- [ ] **Step 3: Add the `titleRect` parameter to `buildContentSlideXml`** in `PptxWriter.h` (Edit tool). Replace:

```cpp
    QString buildContentSlideXml(const QString& title,
                                 const SlideTable& table,
                                 const QVector<SlideImage>& plots,
                                 const QString& bgRid,
                                 const QString& logoRid,
                                 const QMap<QString, QString>& plotRids,
                                 const QString& extraShapesXml = QString(),
                                 int titleFontPt = 0) const;
```

with:

```cpp
    QString buildContentSlideXml(const QString& title,
                                 const SlideTable& table,
                                 const QVector<SlideImage>& plots,
                                 const QString& bgRid,
                                 const QString& logoRid,
                                 const QMap<QString, QString>& plotRids,
                                 const QString& extraShapesXml = QString(),
                                 int titleFontPt = 0,
                                 const QRectF& titleRect = QRectF()) const;
```

- [ ] **Step 4: Use the rect in `buildContentSlideXml`** in `PptxWriter.cpp` (Edit tool). Update the signature line `int titleFontPt) const` for this definition to `int titleFontPt, const QRectF& titleRect) const`, then replace the title-emit block:

```cpp
    int titleLines = qMax(1, (title.length() + 37) / 38);
    double titleH = titleLines * 0.50;
    shapes += makeTextBox(id++,
                          0.4, 0.1, 11.0, titleH,
                          title,
                          QStringLiteral("Montserrat"),
                          titleSz100, true, QStringLiteral("1F497D"),
                          QStringLiteral("l"));
```

with:

```cpp
    int titleLines = qMax(1, (title.length() + 37) / 38);
    double titleH = titleLines * 0.50;
    // Bug 1: a non-null layout rect (user-moved or canvas default) overrides
    // the legacy hardcoded title box, so the pptx matches the Preview canvas.
    double titleX = 0.4, titleY = 0.1, titleW = 11.0;
    if (!titleRect.isNull()) {
        titleX = titleRect.x();
        titleY = titleRect.y();
        titleW = titleRect.width();
        titleH = titleRect.height();
    }
    shapes += makeTextBox(id++,
                          titleX, titleY, titleW, titleH,
                          title,
                          QStringLiteral("Montserrat"),
                          titleSz100, true, QStringLiteral("1F497D"),
                          QStringLiteral("l"));
```

- [ ] **Step 5: Thread `layout.title` through the 5-arg `addContentSlide`** in `PptxWriter.cpp` (Edit tool). Replace:

```cpp
    slide.xml = buildContentSlideXml(sheetTitle, table, plots,
                                     bgRid, logoRid, plotRids,
                                     extraShapesXml, layout.titleFontPt);
```

with:

```cpp
    slide.xml = buildContentSlideXml(sheetTitle, table, plots,
                                     bgRid, logoRid, plotRids,
                                     extraShapesXml, layout.titleFontPt,
                                     layout.title);
```

- [ ] **Step 6: Keep the 4-arg `addContentSlide` title hardcoded** (preserve TPM/other-report behavior) in `PptxWriter.cpp` (Edit tool). Replace:

```cpp
    ContentSlideLayout dl;
    dl.title = QRectF(0.32, 0.10, 12.7, 0.55);  // legacy default; not currently
                                                // wired through buildContentSlideXml.
```

with:

```cpp
    ContentSlideLayout dl;
    dl.title = QRectF();   // null → buildContentSlideXml keeps the hardcoded
                           // title position for 4-arg (non-layout) callers.
```

- [ ] **Step 7: Run — verify it PASSES (and no regressions)**

Run:
```bash
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter pptxwriter
```
Expected: PASS, including the existing `addContentSlide_layoutOverrideAppliesToEmuPositions` and `addContentSlide_layoutFontsAppearInSlideXml` (they set no `title`, so the default-null path keeps the hardcoded title).

- [ ] **Step 8: Commit**

```bash
git add src/reporting/PptxWriter.h src/reporting/PptxWriter.cpp tests/tst_pptxwriter/tst_pptxwriter.cpp
git commit -m "fix(report): content-slide title honors layout rect in pptx (WYSIWYG)"
```

---

## Task 5: Bug 1 — `PptxWriter` honors cover title/subtitle rects (TDD)

**Files:**
- Modify: `tests/tst_pptxwriter/tst_pptxwriter.cpp` (new test slot)
- Modify: `src/reporting/PptxWriter.h:74-75` (addCoverSlide) and `:174-179` (buildCoverSlideXml)
- Modify: `src/reporting/PptxWriter.cpp` (addCoverSlide 4-arg body; buildCoverSlideXml body)

- [ ] **Step 1: Write the failing test** — add this slot to `tst_pptxwriter.cpp` before the closing `};` (Edit tool):

```cpp
    // ── Bug 1: cover title/subtitle rects land in slide XML ──────────────
    void addCoverSlide_titleSubtitleRectsApplyToEmuPositions()
    {
        DVE::PptxWriter w;
        w.setResourcePath(QStringLiteral("../../resources/images"));
        // title y = 1.5" → 1371600 EMU; subtitle y = 4.0" → 3657600 EMU.
        // Distinct from the hardcoded cover (title 2.0", date 4.6").
        w.addCoverSlide("Cover", "2026-06-02", /*titleFontPt=*/0, /*dateFontPt=*/0,
                        QRectF(1.0, 1.5, 10.0, 1.2), QRectF(1.0, 4.0, 10.0, 0.6));
        const QString path = tempPath("cover_rects.pptx");
        QVERIFY(w.save(path));
        QZipReader zr(path);
        const QByteArray slideXml =
            zr.fileData(QStringLiteral("ppt/slides/slide1.xml"));
        QVERIFY(!slideXml.isEmpty());
        QVERIFY2(slideXml.contains("1371600"),
                 "cover title.y (1.5\" = 1371600 EMU) missing — rect not honored");
        QVERIFY2(slideXml.contains("3657600"),
                 "cover subtitle.y (4.0\" = 3657600 EMU) missing — rect not honored");
    }
```

- [ ] **Step 2: Run — verify it FAILS to build**

Run:
```bash
python tools/decrypt_via_copy.py --apply
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter pptxwriter
```
Expected: BUILD FAILS — no `addCoverSlide` overload takes 6 arguments yet.

- [ ] **Step 3: Add rect params to the declarations** in `PptxWriter.h` (Edit tool).

Replace the font-overload declaration:

```cpp
    void addCoverSlide(const QString& title, const QString& dateStr,
                       int titleFontPt, int dateFontPt);
```

with:

```cpp
    void addCoverSlide(const QString& title, const QString& dateStr,
                       int titleFontPt, int dateFontPt,
                       const QRectF& titleRect = QRectF(),
                       const QRectF& subtitleRect = QRectF());
```

Replace the `buildCoverSlideXml` declaration:

```cpp
    QString buildCoverSlideXml(const QString& title,
                               const QString& date,
                               const QString& bgRid,
                               const QString& logoRid,
                               int titleFontPt = 0,
                               int dateFontPt  = 0) const;
```

with:

```cpp
    QString buildCoverSlideXml(const QString& title,
                               const QString& date,
                               const QString& bgRid,
                               const QString& logoRid,
                               int titleFontPt = 0,
                               int dateFontPt  = 0,
                               const QRectF& titleRect = QRectF(),
                               const QRectF& subtitleRect = QRectF()) const;
```

- [ ] **Step 4: Thread the rects through `addCoverSlide`** in `PptxWriter.cpp` (Edit tool). Replace:

```cpp
void PptxWriter::addCoverSlide(const QString& title, const QString& dateStr,
                                int titleFontPt, int dateFontPt)
{
    Slide slide;

    QByteArray bgData   = loadResourceImage(QStringLiteral("Cover_Page_Logo.jpg"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full_white.png"));

    // rId assignment: bg → rId1, logo → rId2.
    QString bgRid   = addMedia(bgData,   QStringLiteral("jpg"),  slide.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"),  slide.media);

    slide.xml = buildCoverSlideXml(title, dateStr, bgRid, logoRid,
                                    titleFontPt, dateFontPt);
    m_slides.append(slide);
}
```

with:

```cpp
void PptxWriter::addCoverSlide(const QString& title, const QString& dateStr,
                                int titleFontPt, int dateFontPt,
                                const QRectF& titleRect,
                                const QRectF& subtitleRect)
{
    Slide slide;

    QByteArray bgData   = loadResourceImage(QStringLiteral("Cover_Page_Logo.jpg"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full_white.png"));

    // rId assignment: bg → rId1, logo → rId2.
    QString bgRid   = addMedia(bgData,   QStringLiteral("jpg"),  slide.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"),  slide.media);

    slide.xml = buildCoverSlideXml(title, dateStr, bgRid, logoRid,
                                    titleFontPt, dateFontPt,
                                    titleRect, subtitleRect);
    m_slides.append(slide);
}
```

- [ ] **Step 5: Use the rects in `buildCoverSlideXml`** in `PptxWriter.cpp` (Edit tool). Update this definition's signature to add `, const QRectF& titleRect, const QRectF& subtitleRect` after `int dateFontPt`, then replace the title + date emit blocks:

```cpp
    const int titleSz100 = (titleFontPt > 0 ? titleFontPt : 46) * 100;
    shapes += makeTextBox(id++,
                          0.5, 2.0, 12.3, 1.5,
                          title,
                          QStringLiteral("Montserrat"),
                          titleSz100, true, QStringLiteral("FFFFFF"),
                          QStringLiteral("ctr"));

    // Date: Montserrat 24pt white, centred (omitted when date is empty).
    if (!date.isEmpty()) {
        const int dateSz100 = (dateFontPt > 0 ? dateFontPt : 24) * 100;
        shapes += makeTextBox(id++,
                              0.5, 4.6, 12.3, 0.7,
                              date,
                              QStringLiteral("Montserrat"),
                              dateSz100, false, QStringLiteral("FFFFFF"),
                              QStringLiteral("ctr"));
    }
```

with:

```cpp
    const int titleSz100 = (titleFontPt > 0 ? titleFontPt : 46) * 100;
    // Bug 1: honor a non-null layout rect; else legacy hardcoded box.
    double tX = 0.5, tY = 2.0, tW = 12.3, tH = 1.5;
    if (!titleRect.isNull()) {
        tX = titleRect.x(); tY = titleRect.y();
        tW = titleRect.width(); tH = titleRect.height();
    }
    shapes += makeTextBox(id++,
                          tX, tY, tW, tH,
                          title,
                          QStringLiteral("Montserrat"),
                          titleSz100, true, QStringLiteral("FFFFFF"),
                          QStringLiteral("ctr"));

    // Date: Montserrat 24pt white, centred (omitted when date is empty).
    if (!date.isEmpty()) {
        const int dateSz100 = (dateFontPt > 0 ? dateFontPt : 24) * 100;
        double dX = 0.5, dY = 4.6, dW = 12.3, dH = 0.7;
        if (!subtitleRect.isNull()) {
            dX = subtitleRect.x(); dY = subtitleRect.y();
            dW = subtitleRect.width(); dH = subtitleRect.height();
        }
        shapes += makeTextBox(id++,
                              dX, dY, dW, dH,
                              date,
                              QStringLiteral("Montserrat"),
                              dateSz100, false, QStringLiteral("FFFFFF"),
                              QStringLiteral("ctr"));
    }
```

- [ ] **Step 6: Run — verify it PASSES**

Run:
```bash
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter pptxwriter
```
Expected: PASS, including the existing cover-font tests (`addCoverSlide_legacyTwoArgEmitsHardcodedFontSizes`, `addCoverSlide_fontOverloadEmitsRequestedSizes`) — they pass null rects, keeping the hardcoded positions. `addSectionDividerSlide` (which calls the 4-arg `addCoverSlide`) is unaffected via the rect defaults.

- [ ] **Step 7: Commit**

```bash
git add src/reporting/PptxWriter.h src/reporting/PptxWriter.cpp tests/tst_pptxwriter/tst_pptxwriter.cpp
git commit -m "fix(report): cover title/subtitle honor layout rects in pptx (WYSIWYG)"
```

---

## Task 6: Bug 1 — extract & test the properties-box XML helper (TDD)

The properties textbox is built inline in `writeSensoryPptx` with a hardcoded bottom-right anchor. Extract it into a pure static helper that honors a rect override, and unit-test it directly (no radar render needed).

**Files:**
- Modify: `tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp` (new slots)
- Modify: `src/reporting/SensoryReportSource.h` (declare static helper)
- Modify: `src/reporting/SensoryReportSource.cpp` (define helper; call it from `writeSensoryPptx`)

- [ ] **Step 1: Write the failing tests** — add to `tst_sensoryreportsource.cpp` before the closing `};` (Edit tool):

```cpp
    // ── Bug 1: properties box honors a layout rect override ──────────────
    void propertiesBoxXml_honorsRectOverride()
    {
        const QStringList lines{ "Media: X", "Control: Y" };
        // Override at 1.0",2.0" → x=914400, y=1828800 EMU.
        const QString xml = DVE::SensoryReportSource::buildPropertiesBoxXml(
            lines, QRectF(1.0, 2.0, 3.0, 4.0), /*fontPt=*/0, 13.33, 7.5);
        QVERIFY(!xml.isEmpty());
        QVERIFY2(xml.contains(QStringLiteral("x=\"914400\"")),
                 "props box x override (1.0\") missing");
        QVERIFY2(xml.contains(QStringLiteral("y=\"1828800\"")),
                 "props box y override (2.0\") missing");
    }
    void propertiesBoxXml_nullRectAnchorsBottomRight()
    {
        const QStringList lines{ "Media: X" };
        const QString xml = DVE::SensoryReportSource::buildPropertiesBoxXml(
            lines, QRectF(), /*fontPt=*/0, 13.33, 7.5);
        QVERIFY(!xml.isEmpty());
        QVERIFY(xml.contains(QStringLiteral("name=\"PropsBox\"")));
        // Bottom-right legacy anchor: tbX = 13.33 - 3.17 - 0.05 = 10.11" → 9244584 EMU.
        QVERIFY2(xml.contains(QStringLiteral("9244584")),
                 "null rect should keep the bottom-right legacy x position");
    }
    void propertiesBoxXml_emptyLinesYieldEmpty()
    {
        QVERIFY(DVE::SensoryReportSource::buildPropertiesBoxXml(
            {}, QRectF(), 0, 13.33, 7.5).isEmpty());
    }
```

- [ ] **Step 2: Run — verify it FAILS to build**

Run:
```bash
python tools/decrypt_via_copy.py --apply
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter sensoryreportsource
```
Expected: BUILD FAILS — `buildPropertiesBoxXml` is not a member of `SensoryReportSource`.

- [ ] **Step 3: Declare the helper** in `SensoryReportSource.h` (Edit tool). Add this declaration in the `public:` section next to the other static helpers (e.g. immediately after the `writeSensoryPptx(...)` static declaration):

```cpp
    // Build the "properties" textbox XML for a sensory content slide. When
    // rectOverride is non-null it positions the box (WYSIWYG from the Preview);
    // otherwise the box is anchored bottom-right with a height derived from the
    // line count. fontPt 0 = legacy 16 pt. Returns empty when propLines empty.
    [[nodiscard]] static QString buildPropertiesBoxXml(const QStringList& propLines,
                                                       const QRectF& rectOverride,
                                                       int fontPt,
                                                       double slideW,
                                                       double slideH);
```

Ensure `<QRectF>` and `<QStringList>` are available (they are, transitively via the existing reporting headers; add `#include <QRectF>` to `SensoryReportSource.h` if the build complains).

- [ ] **Step 4: Define the helper** in `SensoryReportSource.cpp` (Edit tool). Add this function (e.g. just above `SensoryReportSource::writeSensoryPptx`):

```cpp
QString SensoryReportSource::buildPropertiesBoxXml(const QStringList& propLines,
                                                   const QRectF& rectOverride,
                                                   int fontPt,
                                                   double slideW,
                                                   double slideH)
{
    if (propLines.isEmpty()) return QString();

    const int propsFontPt = fontPt > 0 ? fontPt : 16;
    const int propsSz100  = propsFontPt * 100;

    double tbW = 3.17;
    const double scaleVsLegacy = propsFontPt / 16.0;
    const int charsPerLine = qMax(6, int(18.0 / scaleVsLegacy));
    int wrappedLines = 0;
    for (const QString& line : propLines) {
        int extraLines = qMax(0, (line.length() - charsPerLine) / charsPerLine);
        wrappedLines += extraLines;
    }
    double tbH = qMax(2.0, (2.0 + wrappedLines * 0.20) * scaleVsLegacy);
    double tbX = slideW - tbW - 0.05;
    double tbY = slideH - tbH - 0.05;
    // Bug 1: a non-null layout rect (user-moved or canvas default) wins.
    if (!rectOverride.isNull()) {
        tbX = rectOverride.x();
        tbY = rectOverride.y();
        tbW = rectOverride.width();
        tbH = rectOverride.height();
    }

    QString paras;
    for (const QString& line : propLines) {
        QString safe = line;
        safe.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;");
        paras += QStringLiteral(
            R"(<a:p><a:pPr algn="l"/>)"
            R"(<a:r><a:rPr lang="en-US" sz="%2" b="0" dirty="0">)"
            R"(<a:solidFill><a:srgbClr val="333333"/></a:solidFill>)"
            R"(<a:latin typeface="Calibri"/>)"
            R"(</a:rPr><a:t>%1</a:t></a:r></a:p>)")
            .arg(safe).arg(propsSz100);
    }

    auto toEmu = [](double in) { return QString::number(qRound64(in * 914400.0)); };
    return QStringLiteral(
        R"(<p:sp><p:nvSpPr>)"
        R"(<p:cNvPr id="90" name="PropsBox"/>)"
        R"(<p:cNvSpPr txBox="1"/><p:nvPr/>)"
        R"(</p:nvSpPr><p:spPr>)"
        R"(<a:xfrm><a:off x="%1" y="%2"/><a:ext cx="%3" cy="%4"/></a:xfrm>)"
        R"(<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>)"
        R"(<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>)"
        R"(<a:ln w="0"><a:noFill/></a:ln>)"
        R"(</p:spPr><p:txBody>)"
        R"(<a:bodyPr wrap="square" lIns="36000" tIns="18000" rIns="36000" bIns="18000" rtlCol="0"/>)"
        R"(<a:lstStyle/>%5)"
        R"(</p:txBody></p:sp>)")
        .arg(toEmu(tbX), toEmu(tbY), toEmu(tbW), toEmu(tbH))
        .arg(paras);
}
```

- [ ] **Step 5: Replace the inline block in `writeSensoryPptx` with a call** (Edit tool). Replace the whole inline properties-box block (the `QString extraXml; if (!propLines.isEmpty()) { ... }` that spans roughly SensoryReportSource.cpp:719-776, ending just before the `// slideLayout was resolved...` comment) with:

```cpp
            QString extraXml = buildPropertiesBoxXml(
                propLines, slideLayout.propertiesBox.rect,
                slideLayout.propertiesBox.fontPt, slideW, slideH);
```

(`slideW`/`slideH` are already in scope from the chart-placement block above; `propLines`, `slideLayout` are unchanged.)

- [ ] **Step 6: Run — verify it PASSES**

Run:
```bash
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter sensoryreportsource
```
Expected: PASS (new props-box slots + all existing slots; Postgres slots skip if `DVE_TEST_PG_CONN` unset).

- [ ] **Step 7: Commit**

```bash
git add src/reporting/SensoryReportSource.h src/reporting/SensoryReportSource.cpp tests/tst_sensoryreportsource/tst_sensoryreportsource.cpp
git commit -m "fix(report): properties box honors layout rect; extract testable helper"
```

---

## Task 7: Bug 1 — wire cover & divider rects into `writeSensoryPptx`

The content-slide title (Task 4), cumulative title (Task 4, via `addContentSlide(..., layout.cumulative)`), and properties box (Task 6) now flow automatically. The cover and section-divider titles still pass only fonts — wire their rects.

**Files:**
- Modify: `src/reporting/SensoryReportSource.cpp` (the cover call ~:513 and the divider call ~:541)

No new unit test (full `writeSensoryPptx` rendering needs the radar widget; the `addCoverSlide` rect mechanism is already proven in Task 5). Verified by compile + Task 8 manual check.

- [ ] **Step 1: Pass cover rects** (Edit tool). Replace:

```cpp
    pptx.addCoverSlide(coverTitle, coverDate,
                        layout.coverTitleFontPt, layout.coverSubtitleFontPt);
```

with:

```cpp
    pptx.addCoverSlide(coverTitle, coverDate,
                        layout.coverTitleFontPt, layout.coverSubtitleFontPt,
                        layout.coverTitle, layout.coverSubtitle);
```

- [ ] **Step 2: Pass divider title rect** (Edit tool). Replace:

```cpp
            pptx.addCoverSlide(groupTitle, groupDate,
                                dividerFontPt, /*dateFontPt=*/0);
```

with:

```cpp
            pptx.addCoverSlide(groupTitle, groupDate,
                                dividerFontPt, /*dateFontPt=*/0,
                                layout.dividerTitles.value(dividerKey), QRectF());
```

(`dividerKey` is already defined just above as `QStringLiteral("divider_%1").arg(firstIdx)`.)

- [ ] **Step 3: Compile-check + run the suite**

Run:
```bash
python tools/decrypt_via_copy.py --apply
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Filter sensoryreportsource
```
Expected: builds clean; all slots pass/skip.

- [ ] **Step 4: Commit**

```bash
git add src/reporting/SensoryReportSource.cpp
git commit -m "fix(report): cover & divider titles honor layout rects in pptx (WYSIWYG)"
```

---

## Task 8: Final integration — full build + full suite + manual verification

**Files:** none (verification only).

- [ ] **Step 1: Decrypt + full app build**

Run the full-app build from Preflight. Expected: clean link, `DataViewer.exe` produced.

- [ ] **Step 2: Full test suite**

Run:
```bash
python tools/decrypt_via_copy.py --apply
powershell.exe -ExecutionPolicy Bypass -File tests/run-tests.ps1 -Rebuild
```
Expected: all suites build; every test passes or skips (DB suites skip cleanly without `DVE_TEST_PG_CONN`). Investigate any failure before declaring done.

- [ ] **Step 3: Manual verification (the parts no test covers — run the app)**

- **Bug 3:** Sensory mode → confirm Smoothness/Burnt-Taste label gap matches Overall Flavor/Vapor Volume. Adjust the `- 1.0` (Task 1) by ±1 px if needed and re-commit.
- **Bug 2:** header shows `Round:` between Tester and Media, default "1"; mouse-wheel over it does nothing. Set Tester "Charlie" + Round "1" → Save → filename ends `- Charlie R1`. Reopen a legacy `Charlie R2` session → Tester "Charlie", Round "2". Set Round "N/A" → filename has no suffix.
- **Bug 1 (the headline):** open Sensory Report (ribbon) → Preview. Drag/resize the slide **title**, the **properties box**, the **cover title**, and a **section-divider title**. Create Report. Open the `.pptx` → every moved element sits where the Preview showed it. (Table + radar already worked; now title/props/cover/divider match too.) Note: an *unmodified* report's title now renders full-width to match the canvas default — expected per the spec's "preview is authoritative" decision.

- [ ] **Step 4: Final commit (only if Step 3 required a tweak)**

```bash
git add -A
git commit -m "fix(sensory): post-verification tweaks for 3-bugfix set"
```

---

## Self-Review checklist (run before handing off)

- [ ] **Spec coverage:** Bug 1 title (T4), cover/subtitle (T5), properties box (T6), divider + cover wiring (T7), cumulative (T4 auto); Bug 2 helpers (T2) + UI fold (T3); Bug 3 (T1). All spec sections mapped.
- [ ] **Placeholders:** none — every step has exact paths + full code.
- [ ] **Type/name consistency:** `splitTesterRound`/`combineTesterRound`/`TesterRound{tester,round}`, `m_roundCombo`, `buildPropertiesBoxXml(propLines, rectOverride, fontPt, slideW, slideH)`, `buildContentSlideXml(..., titleRect)`, `addCoverSlide(..., titleRect, subtitleRect)`, `buildCoverSlideXml(..., titleRect, subtitleRect)` — used identically across tasks.
