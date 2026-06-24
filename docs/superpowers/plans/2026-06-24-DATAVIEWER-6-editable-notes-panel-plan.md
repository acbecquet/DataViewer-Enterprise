# Editable TPM Notes/Story Panel — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the TPM-mode `m_dataTable` grid *widget* with a right-side, per-sample, scrollable, editable "story" panel — note-bearing rows shown as detailed editable cards, the runs between them as expandable summaries — while keeping every existing edit → recalc → Excel write-back → Postgres `data_rows` → LiveSync behaviour.

**Architecture:** A new pure segmentation model (`NotesStory`) turns a `SampleResult` into an ordered list of Note / Summary segments. A new `NotesStoryPanel` widget (a `QScrollArea` of composed input widgets — no custom painting) renders those segments with real Qt editors and emits `cellEdited(dataRowIndex, col, text)`. `MainWindow` re-orients its central splitter to horizontal `Plot | Panel`, populates the panel from the **cleaned** sample, and routes `cellEdited` through the same mutate→recalc→Excel+DB+LiveSync path the table used. `PlotEngine` gains per-point rings + an emphasized point for v1 note→plot linking.

**Tech Stack:** C++17, Qt 6.10 Widgets, qmake + MinGW. TDD with Qt Test. `-Werror -Wall -Wextra`.

---

## Preconditions & sequencing

- **DATAVIEWER-16 (Phase 2b.0 cleanup fixes) HAS LANDED** on `feature/v2.6.0-backlog-batch` — commits `4954369` (Phase-1 MS-1/2/3) + `02afc16` (cleanup GAP-A/B/F + persist + Undo-all). **Drift to absorb:** the cleanup extracted a pure, GUI-free `src/pipeline/DataCleanup.{h,cpp}` — `buildCleanedSheet` / `exclusionsFor` / `buildCleanedFile(fileIdx)` now live there (MainWindow delegates). Use those symbols (the panel's Task 7 still calls them via the MainWindow wrappers `exclusionsFor(...)` / `currentSheetHasCleanup()`, but `buildCleanedSheet` is now `DataCleanup::buildCleanedSheet`). Also: `closeFile`→`onCloseFile`; `DataCleanupDialog` is in `src/ui/` (not `src/widgets/`). **Re-confirm every `MainWindow.cpp` line reference below by grep before editing** (the file churned — anchor on function names, not line numbers).
- Tasks 1–5 (the new `NotesStory` model + `NotesStoryPanel` widget) are **greenfield, independent of 2b.0**, and can start immediately.
- **MIP:** before any build run `python tools/decrypt_via_copy.py --apply`. Create new source files via Python delete-and-rewrite (see CLAUDE.md), not the Write tool. Build: `qmake CONFIG+=release -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8` under `C:\Qt\6.10.1`. Test exes: windows-subsystem, run with `-o out.txt,txt`, need Qt `bin` on PATH + `QT_QPA_PLATFORM=offscreen`.
- Reference spec: `docs/superpowers/specs/2026-06-24-DATAVIEWER-6-MS4-editable-notes-panel-design.md`. Teardown inventory: `docs/superpowers/specs/2026-06-24-DATAVIEWER-6-MS4-optionB-spike.md` §5.
- Guardrail: run `/ponytail-review` on the diff before each commit — reuse existing plumbing, no speculative abstractions, no new deps; never trim validation.

## File structure

- **Create** `src/pipeline/NotesStory.h` / `NotesStory.cpp` — pure segmentation model (`buildStory`). Lives with the pipeline because it's data-shaping, no Qt widgets.
- **Create** `src/widgets/NotesStoryPanel.h` / `NotesStoryPanel.cpp` — the editable panel widget.
- **Create** `tests/tst_notesstory/tst_notesstory.pro` + `tst_notesstory.cpp` — model unit tests (no GUI).
- **Modify** `src/plotting/PlotEngine.h` / `PlotEngine.cpp` — per-point rings + emphasized index.
- **Modify** `src/plotting/PlotWidget.h` / `PlotWidget.cpp` — thread note-puffs/selected through to `PlotSeries`; `selectPuff(int)` slot; `notePointClicked` deferred to DATAVIEWER-18.
- **Modify** `src/MainWindow.h` / `MainWindow.cpp` — central layout swap, panel construction, populate, edit routing, `onPropCellChanged` rework, table-widget removal.
- **Modify** `DataViewerEnterprise.pro` and `tests/tests.pro` — add the new sources + the `tst_notesstory` SUBDIRS entry.

---

### Task 1: Story segmentation model

**Files:**
- Create: `src/pipeline/NotesStory.h`, `src/pipeline/NotesStory.cpp`
- Create: `tests/tst_notesstory/tst_notesstory.cpp`, `tests/tst_notesstory/tst_notesstory.pro`
- Modify: `tests/tests.pro` (add `tst_notesstory` to SUBDIRS), `DataViewerEnterprise.pro` (add `src/pipeline/NotesStory.cpp`)

The model walks the sample's **visible** rows (skip rows where `beforeWeight==0 || afterWeight==0`, matching the table/plot filter) in order, emitting one `Note` segment per row whose `notes` is non-empty, and grouping each contiguous run of note-less visible rows into one `Summary`. Excluded rows stay in the story (marked) but are **excluded from summary aggregates**.

- [ ] **Step 1: Write the failing test** — `tests/tst_notesstory/tst_notesstory.cpp`

```cpp
#include <QtTest>
#include "../../src/pipeline/NotesStory.h"
using namespace DVE;

static DataRow row(double puffs, double tpm, double var, const QString& notes,
                   double before = 1.0, double after = 0.9) {
    DataRow r; r.puffs = puffs; r.tpm = tpm; r.variationTPM = var; r.notes = notes;
    r.beforeWeight = before; r.afterWeight = after; return r;
}

class TstNotesStory : public QObject { Q_OBJECT
private slots:
    void segmentsSplitOnNoteRows();
    void summaryAggregatesOverIncludedRowsOnly();
    void zeroWeightRowsAreSkipped();
};

void TstNotesStory::segmentsSplitOnNoteRows() {
    SampleResult s;
    s.rows = { row(1,3.6,0.1,""), row(2,3.5,0.1,""), row(3,3.4,0.2,"burnt onset"),
               row(4,3.3,0.2,""), row(5,3.1,0.2,"clog") };
    const QVector<StorySegment> segs = NotesStory::build(s, /*excluded=*/{});
    QCOMPARE(segs.size(), 4);
    QCOMPARE(segs[0].kind, StorySegment::Summary);      // puffs 1-2
    QCOMPARE(segs[0].summary.count, 2);
    QCOMPARE(segs[1].kind, StorySegment::Note);         // puff 3
    QCOMPARE(segs[1].rowIndex, 2);
    QCOMPARE(segs[2].kind, StorySegment::Summary);      // puff 4
    QCOMPARE(segs[3].kind, StorySegment::Note);         // puff 5
}

void TstNotesStory::summaryAggregatesOverIncludedRowsOnly() {
    SampleResult s;
    s.rows = { row(1,4.0,0.0,""), row(2,2.0,0.0,""), row(3,3.0,0.0,"") };
    const QVector<StorySegment> segs = NotesStory::build(s, /*excluded=*/{1}); // drop puff 2
    QCOMPARE(segs.size(), 1);
    QCOMPARE(segs[0].kind, StorySegment::Summary);
    QCOMPARE(segs[0].summary.count, 2);                 // 2 included rows
    QCOMPARE(segs[0].summary.avgTpm, 3.5);              // (4.0 + 3.0) / 2
}

void TstNotesStory::zeroWeightRowsAreSkipped() {
    SampleResult s;
    s.rows = { row(1,3.0,0.1,""), row(2,0,0,"", 0.0, 0.0), row(3,3.0,0.1,"note") };
    const QVector<StorySegment> segs = NotesStory::build(s, {});
    QCOMPARE(segs.size(), 2);                            // empty row dropped: 1 summary + 1 note
    QCOMPARE(segs[0].summary.count, 1);
    QCOMPARE(segs[1].kind, StorySegment::Note);
    QCOMPARE(segs[1].rowIndex, 2);
}

QTEST_MAIN(TstNotesStory)
#include "tst_notesstory.moc"
```

- [ ] **Step 2: Write `tst_notesstory.pro`**

```pro
QT += testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
SOURCES += tst_notesstory.cpp ../../src/pipeline/NotesStory.cpp
INCLUDEPATH += ../../src
```

- [ ] **Step 3: Run the test to verify it fails** — Run: `cd tests/tst_notesstory && qmake && mingw32-make && QT_QPA_PLATFORM=offscreen ./release/tst_notesstory.exe -o out.txt,txt`. Expected: FAIL to compile (`NotesStory.h` not found).

- [ ] **Step 4: Write `src/pipeline/NotesStory.h`**

```cpp
#pragma once
#include "ReportData.h"
#include <QSet>
#include <QVector>

namespace DVE {

struct StorySummary {
    int    puffStart = 0;   // first visible row's puffs
    int    puffEnd   = 0;   // last visible row's puffs
    int    count     = 0;   // included visible rows in the run
    double avgTpm    = 0.0; // mean tpm over included rows
    double varTpm    = 0.0; // mean variationTPM over included rows
};

struct StorySegment {
    enum Kind { Note, Summary } kind = Summary;
    int          rowIndex = -1;     // index into SampleResult::rows (Note only)
    bool         excluded = false;  // Note only: row is in the exclusion set
    StorySummary summary;           // Summary only
};

namespace NotesStory {
    // Ordered story over the sample's VISIBLE rows (zero-weight rows skipped).
    // Note rows = non-empty trimmed notes; runs between them collapse to one
    // Summary. Summary aggregates use INCLUDED rows only (excluded rows still
    // appear as their own marked Note segment if they carry a note).
    QVector<StorySegment> build(const SampleResult& sample, const QSet<int>& excluded);
}

} // namespace DVE
```

- [ ] **Step 5: Write `src/pipeline/NotesStory.cpp`**

```cpp
#include "NotesStory.h"

namespace DVE {

static bool isVisible(const DataRow& r) {
    return r.beforeWeight != 0.0 && r.afterWeight != 0.0;
}

QVector<StorySegment> NotesStory::build(const SampleResult& sample, const QSet<int>& excluded) {
    QVector<StorySegment> out;

    StorySummary run;
    int    runIncluded = 0;
    double sumTpm = 0.0, sumVar = 0.0;
    bool   runOpen = false;

    auto flushRun = [&]() {
        if (!runOpen) return;
        run.count  = runIncluded;
        run.avgTpm = runIncluded ? sumTpm / runIncluded : 0.0;
        run.varTpm = runIncluded ? sumVar / runIncluded : 0.0;
        StorySegment seg; seg.kind = StorySegment::Summary; seg.summary = run;
        out.append(seg);
        run = StorySummary{}; runIncluded = 0; sumTpm = sumVar = 0.0; runOpen = false;
    };

    for (int i = 0; i < sample.rows.size(); ++i) {
        const DataRow& r = sample.rows[i];
        if (!isVisible(r)) continue;

        if (!r.notes.trimmed().isEmpty()) {
            flushRun();
            StorySegment seg;
            seg.kind = StorySegment::Note;
            seg.rowIndex = i;
            seg.excluded = excluded.contains(i);
            out.append(seg);
            continue;
        }

        // note-less visible row → accumulate into the current run
        if (!runOpen) { run.puffStart = int(r.puffs); runOpen = true; }
        run.puffEnd = int(r.puffs);
        if (!excluded.contains(i)) {
            ++runIncluded; sumTpm += r.tpm; sumVar += r.variationTPM;
        }
    }
    flushRun();
    return out;
}

} // namespace DVE
```

- [ ] **Step 6: Run the test to verify it passes** — Run as Step 3. Expected: PASS (3 tests).

- [ ] **Step 7: Add the test to the suite + the source to the app** — In `tests/tests.pro` add `tst_notesstory` to `SUBDIRS`. In `DataViewerEnterprise.pro` add `src/pipeline/NotesStory.cpp` to `SOURCES` and `src/pipeline/NotesStory.h` to `HEADERS`.

- [ ] **Step 8: Commit**

```bash
git add src/pipeline/NotesStory.h src/pipeline/NotesStory.cpp tests/tst_notesstory/ tests/tests.pro DataViewerEnterprise.pro
git commit -m "feat(tpm): story segmentation model for the notes panel (DATAVIEWER-6)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 2: NotesStoryPanel widget skeleton + signal contract

**Files:**
- Create: `src/widgets/NotesStoryPanel.h`, `src/widgets/NotesStoryPanel.cpp`
- Modify: `DataViewerEnterprise.pro` (add the new sources)

The panel is a `QScrollArea` (vertical scroll only; **horizontal scrollbar OFF**; `setWidgetResizable(true)`) wrapping a content `QWidget` + `QVBoxLayout`. `setSample()` rebuilds the layout from `NotesStory::build`. Editors emit `cellEdited(dataRowIndex, col, text)` where `col` is a `DVE::Cols` value (`NOTES`, `SMELL`, `CLOG`, or `RESISTANCE` when the sheet is per-row-regime). Quantitative values render as read-only `QLabel`s.

- [ ] **Step 1: Write `src/widgets/NotesStoryPanel.h`**

```cpp
#pragma once
#include "../pipeline/ReportData.h"
#include "../pipeline/NotesStory.h"
#include <QWidget>
#include <QSet>
class QVBoxLayout; class QScrollArea;

namespace DVE {

class NotesStoryPanel : public QWidget {
    Q_OBJECT
public:
    explicit NotesStoryPanel(QWidget* parent = nullptr);

    // Rebuild for one sample. `excluded` = exclusion row-index set for this
    // sample; `perRowRegime` selects the regime editor vs a read-only resistance.
    void setSample(const SampleResult& sample, const QSet<int>& excluded, bool perRowRegime);
    void clear();

public slots:
    // v1 plot->note hook (DATAVIEWER-18 will call this); scrolls to + flashes
    // the card for the given sample-row index, no-op if not a note row.
    void highlightRow(int dataRowIndex);

signals:
    // Emitted on a qualitative edit. col is a DVE::Cols value (NOTES/SMELL/
    // CLOG/RESISTANCE). MainWindow mirrors the old onTableCellChanged path.
    void cellEdited(int dataRowIndex, int col, const QString& text);
    // Emitted when a note card is clicked (v1 note->plot emphasis).
    void noteActivated(int dataRowIndex);

private:
    QWidget*     buildNoteCard(const SampleResult& s, int rowIndex, bool excluded, bool perRowRegime);
    QWidget*     buildSummaryBar(const SampleResult& s, const StorySummary& sum, int segIndex, bool perRowRegime);
    void         rebuild();

    SampleResult m_sample;
    QSet<int>    m_excluded;
    bool         m_perRowRegime = false;
    QScrollArea* m_scroll  = nullptr;
    QWidget*     m_content = nullptr;
    QVBoxLayout* m_vbox    = nullptr;
    QHash<int, QWidget*> m_cardByRow;   // rowIndex -> card, for highlightRow
};

} // namespace DVE
```

- [ ] **Step 2: Write the constructor + scroll scaffold + clear/rebuild dispatch in `src/widgets/NotesStoryPanel.cpp`**

```cpp
#include "NotesStoryPanel.h"
#include "../utils/AppTheme.h"
#include <QVBoxLayout>
#include <QScrollArea>
#include <QScrollBar>

namespace DVE {

NotesStoryPanel::NotesStoryPanel(QWidget* parent) : QWidget(parent) {
    auto* outer = new QVBoxLayout(this);
    outer->setContentsMargins(0, 0, 0, 0);
    m_scroll = new QScrollArea(this);
    m_scroll->setWidgetResizable(true);
    m_scroll->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);  // no horizontal overflow
    m_scroll->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_scroll->setFrameShape(QFrame::NoFrame);
    m_content = new QWidget(m_scroll);
    m_vbox = new QVBoxLayout(m_content);
    m_vbox->setContentsMargins(8, 8, 8, 8);
    m_vbox->setSpacing(8);
    m_vbox->addStretch(1);
    m_scroll->setWidget(m_content);
    outer->addWidget(m_scroll);
    setMinimumWidth(280);
}

void NotesStoryPanel::clear() { m_sample = SampleResult{}; m_excluded.clear(); rebuild(); }

void NotesStoryPanel::setSample(const SampleResult& sample, const QSet<int>& excluded, bool perRowRegime) {
    m_sample = sample; m_excluded = excluded; m_perRowRegime = perRowRegime; rebuild();
}

void NotesStoryPanel::rebuild() {
    m_cardByRow.clear();
    QLayoutItem* item;
    while ((item = m_vbox->takeAt(0)) != nullptr) {
        if (QWidget* w = item->widget()) w->deleteLater();
        delete item;
    }
    const QVector<StorySegment> segs = NotesStory::build(m_sample, m_excluded);
    for (int i = 0; i < segs.size(); ++i) {
        const StorySegment& seg = segs[i];
        QWidget* w = (seg.kind == StorySegment::Note)
            ? buildNoteCard(m_sample, seg.rowIndex, seg.excluded, m_perRowRegime)
            : buildSummaryBar(m_sample, seg.summary, i, m_perRowRegime);
        m_vbox->addWidget(w);
        if (seg.kind == StorySegment::Note) m_cardByRow.insert(seg.rowIndex, w);
    }
    m_vbox->addStretch(1);
}

} // namespace DVE
```

- [ ] **Step 3: Build to verify it compiles** — add the sources to `DataViewerEnterprise.pro`, then `qmake CONFIG+=release && mingw32-make -j8`. Expected: compiles clean (`buildNoteCard`/`buildSummaryBar`/`highlightRow` defined as stubs returning `new QWidget` for now so it links). Add stubs, then implement in Tasks 3–5.

- [ ] **Step 4: Commit**

```bash
git add src/widgets/NotesStoryPanel.h src/widgets/NotesStoryPanel.cpp DataViewerEnterprise.pro
git commit -m "feat(tpm): NotesStoryPanel scroll scaffold + signal contract (DATAVIEWER-6)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 3: Note card — read-only context + qualitative editors wired to `cellEdited`

**Files:** Modify `src/widgets/NotesStoryPanel.cpp` (implement `buildNoteCard`).

Card contents per the spec §4.1: header `Puff N` + read-only `TPM <tpm> · avg <averageTPM> · var <variationTPM>`; an editable `QPlainTextEdit` for the note; read-only context chips `TPM before` / `TPM after` (neighbouring visible rows' `tpm`) / `Draw <drawPressure>`; editable `QSpinBox` (0–4) for smell, `QComboBox{N,Y}` for clog, and a regime editor (`QComboBox` editable) when `perRowRegime`. Every editor connects to `emit cellEdited(rowIndex, col, value)` and `mousePressEvent`/a click on the card emits `noteActivated(rowIndex)`.

- [ ] **Step 1: Implement `buildNoteCard`** (representative full code; word-wrap + max-width keep it inside the panel — no horizontal overflow):

```cpp
#include <QFrame>
#include <QLabel>
#include <QPlainTextEdit>
#include <QSpinBox>
#include <QComboBox>
#include <QGridLayout>

QWidget* NotesStoryPanel::buildNoteCard(const SampleResult& s, int rowIndex, bool excluded, bool perRowRegime) {
    const DataRow& dr = s.rows[rowIndex];
    auto* card = new QFrame;
    card->setObjectName("storyCard");
    card->setStyleSheet(QString("#storyCard{border:1px solid %1;border-radius:6px;background:%2;}")
        .arg(AppTheme::border().name(), excluded ? "#F8F0F0" : "#FFFFFF"));
    auto* v = new QVBoxLayout(card);
    v->setContentsMargins(8, 8, 8, 8); v->setSpacing(6);

    auto* hdr = new QLabel(QString("Puff %1%2").arg(int(dr.puffs))
        .arg(excluded ? "   • excluded" : ""));
    hdr->setStyleSheet(QString("font-weight:600;color:%1;").arg(AppTheme::accent().name()));
    v->addWidget(hdr);

    auto* ctx = new QLabel(QString("TPM %1 · avg %2 · var %3")
        .arg(dr.tpm,0,'f',2).arg(s.averageTPM,0,'f',2).arg(dr.variationTPM,0,'f',2));
    ctx->setStyleSheet("color:#666;font-size:8pt;");
    v->addWidget(ctx);

    auto* note = new QPlainTextEdit(dr.notes);
    note->setPlaceholderText("Add a note…");
    note->setFixedHeight(48);
    note->setLineWrapMode(QPlainTextEdit::WidgetWidth);
    connect(note, &QPlainTextEdit::textChanged, this, [this, rowIndex, note]() {
        emit cellEdited(rowIndex, Cols::NOTES, note->toPlainText());
    });
    v->addWidget(note);

    // Read-only context chips (find neighbouring visible rows for before/after TPM)
    auto neighbourTpm = [&](int dir)->QString {
        for (int i = rowIndex + dir; i >= 0 && i < s.rows.size(); i += dir)
            if (s.rows[i].beforeWeight != 0.0 && s.rows[i].afterWeight != 0.0)
                return QString::number(s.rows[i].tpm, 'f', 2);
        return "—";
    };
    auto* chips = new QHBoxLayout; chips->setSpacing(6);
    auto roChip = [&](const QString& t){ auto* l=new QLabel(t); l->setStyleSheet(
        "background:#F0F0F0;color:#555;border-radius:6px;padding:2px 6px;font-size:8pt;"); return l; };
    chips->addWidget(roChip(QString("TPM before %1").arg(neighbourTpm(-1))));
    chips->addWidget(roChip(QString("TPM after %1").arg(neighbourTpm(+1))));
    chips->addWidget(roChip(QString("Draw %1").arg(dr.drawPressure,0,'f',1)));
    chips->addStretch(1);
    v->addLayout(chips);

    // Editable qualitative row
    auto* edits = new QHBoxLayout; edits->setSpacing(6);
    auto* smell = new QSpinBox; smell->setRange(0,4); smell->setValue(dr.smell.toInt());
    smell->setPrefix("Smell ");
    connect(smell, QOverload<int>::of(&QSpinBox::valueChanged), this,
        [this, rowIndex](int val){ emit cellEdited(rowIndex, Cols::SMELL, QString::number(val)); });
    auto* clog = new QComboBox; clog->addItems({"N","Y"});
    clog->setCurrentText(dr.clog.trimmed().toUpper() == "Y" ? "Y" : "N");
    connect(clog, &QComboBox::currentTextChanged, this,
        [this, rowIndex](const QString& t){ emit cellEdited(rowIndex, Cols::CLOG, t); });
    edits->addWidget(new QLabel("Clog")); edits->addWidget(clog);
    edits->addWidget(smell);
    if (perRowRegime) {
        auto* regime = new QComboBox; regime->setEditable(true);
        regime->setCurrentText(dr.puffingRegime);
        connect(regime, &QComboBox::currentTextChanged, this,
            [this, rowIndex](const QString& t){ emit cellEdited(rowIndex, Cols::RESISTANCE, t); });
        edits->addWidget(new QLabel("Regime")); edits->addWidget(regime);
    }
    edits->addStretch(1);
    v->addLayout(edits);
    return card;
}
```

- [ ] **Step 2: Build to verify it compiles** — `qmake CONFIG+=release && mingw32-make -j8`. Expected: clean.

- [ ] **Step 3: Commit**

```bash
git add src/widgets/NotesStoryPanel.cpp
git commit -m "feat(tpm): editable note cards with qualitative editors + read-only context (DATAVIEWER-6)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 4: Summary bar + expand-to-compact-editable-table

**Files:** Modify `src/widgets/NotesStoryPanel.cpp` (implement `buildSummaryBar`).

Collapsed: a `QToolButton` (checkable, `ToolButtonTextBesideIcon`, chevron) reading `Puffs A–B · avg X · var Y`. On toggle, build/show a compact `QGridLayout` of the run's individual visible rows: columns `Puff | TPM(read-only) | Draw(read-only) | Smell(QSpinBox 0–4) | Clog(combo) | add-note`. Each editable cell emits `cellEdited(rowIndex, col, value)`; "add-note" emits `noteActivated(rowIndex)` (promotion handled by Task 8/host re-populate). Fixed small column widths keep it within the panel.

- [ ] **Step 1: Implement `buildSummaryBar`** — collapsed `QToolButton` + a hidden child grid container; on `toggled(bool)` lazily populate the grid from the run's visible, in-range rows (re-derive the run by re-walking `m_sample` between the summary's `puffStart`/`puffEnd`), wiring smell/clog editors exactly as in Task 3. Use `setMaximumWidth` on inputs and `grid->setColumnStretch(lastCol,1)` so nothing exceeds the panel width.

- [ ] **Step 2: Build to verify it compiles + manual smoke** — run the app, load a TPM file (after Task 6 wires the panel), expand a summary, confirm no horizontal scrollbar appears. Expected: clean build.

- [ ] **Step 3: Commit**

```bash
git add src/widgets/NotesStoryPanel.cpp
git commit -m "feat(tpm): expandable summary bars with compact editable rows (DATAVIEWER-6)

Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>"
```

---

### Task 5: Excluded-row marking + `highlightRow`

**Files:** Modify `src/widgets/NotesStoryPanel.cpp`.

- [ ] **Step 1:** Note cards already tint `#F8F0F0` + show "• excluded" (Task 3). Add the same tint + an "excluded" tag to excluded rows inside the expanded summary grid (Task 4), keeping all text at full readability (no strikethrough — owner wants to read the note explaining *why*).
- [ ] **Step 2:** Implement `highlightRow(int dataRowIndex)`: if `m_cardByRow.contains(dataRowIndex)`, `m_scroll->ensureWidgetVisible(card)` and flash a 1.2s border via a `QPropertyAnimation`/stylesheet swap.
- [ ] **Step 3: Commit** `git commit -m "feat(tpm): excluded-row marking (readable) + highlightRow scroll/flash (DATAVIEWER-6)"` (with the Co-Authored-By trailer).

---

### Task 6: MainWindow — horizontal `Plot | Panel` layout + construct the panel

> **Gate:** start only after DATAVIEWER-16 (2b.0) has merged. Re-grep all line numbers.

**Files:** Modify `src/MainWindow.h` (add `NotesStoryPanel* m_storyPanel = nullptr;`, forward-declare), `src/MainWindow.cpp` (`setupCentralWidget`).

- [ ] **Step 1:** In `setupCentralWidget` change `m_centralSplitter` to `Qt::Horizontal`. Add `#include "widgets/NotesStoryPanel.h"`.
- [ ] **Step 2:** Construct `m_storyPanel = new NotesStoryPanel(this);` Keep `m_plotWidget` as the dominant pane: `m_centralSplitter->addWidget(m_plotWidget); m_centralSplitter->addWidget(m_storyPanel); m_centralSplitter->setStretchFactor(0,70); m_centralSplitter->setStretchFactor(1,30); m_storyPanel->setMaximumWidth(460);`
- [ ] **Step 3:** Re-parent the sample-nav bar (`m_sampleNavBar`) above the story panel (it stays — samples still cycle). Leave `m_dataTable` constructed-but-not-added-to-the-visible-layout for now (removed in Task 10) so behaviour can be diffed. Keep `m_centralStack->addWidget(m_centralSplitter)` (index 0) so the mode-return lines `setCurrentWidget(m_centralSplitter)` are unchanged.
- [ ] **Step 4:** Connect `m_storyPanel`'s signals to new MainWindow slots (Task 8): `connect(m_storyPanel,&NotesStoryPanel::cellEdited,this,&MainWindow::onStoryCellEdited); connect(m_storyPanel,&NotesStoryPanel::noteActivated,this,&MainWindow::onStoryNoteActivated);`
- [ ] **Step 5: Build, run, smoke** — TPM file shows plot-left + (empty) panel-right; mode switches still work. **Commit.**

---

### Task 7: Populate the panel from the cleaned sample in `displayCurrentSample`

**Files:** Modify `src/MainWindow.cpp` (`displayCurrentSample`).

- [ ] **Step 1:** After the existing cleaned/raw branch computes the sample used for plot+properties, feed the **same** sample to the panel:

```cpp
const QSet<int> excl = exclusionsFor(m_currentFileIndex, m_currentSheetIndex, m_currentSampleIndex);
if (currentSheetHasCleanup()) {
    const SheetResult cleaned = buildCleanedSheet(*sheet, m_currentFileIndex, m_currentSheetIndex);
    m_plotWidget->setSheetData(cleaned);
    updateProperties(cleaned.samples[m_currentSampleIndex]);
    m_storyPanel->setSample(sheet->samples[m_currentSampleIndex], excl, perRowRegime); // raw rows; excl marks them
} else {
    m_plotWidget->setSheetData(*sheet);
    updateProperties(sample);
    m_storyPanel->setSample(sample, excl, perRowRegime);
}
```

  (The panel takes the **raw** sample + the exclusion set so it can SHOW excluded rows marked, while its summary aggregates exclude them — matching `NotesStory::build`. Plot/properties keep using the cleaned sample.)
- [ ] **Step 2:** In the empty-sheet branch and the raw/SOP branch, call `m_storyPanel->clear()`.
- [ ] **Step 3: Build + smoke** — load a TPM file with notes; confirm note cards + summaries render and match the plot. **Commit.**

---

### Task 8: Route panel edits through the existing mutate→recalc→Excel+DB+LiveSync path

**Files:** Modify `src/MainWindow.h` (declare `void onStoryCellEdited(int,int,const QString&); void onStoryNoteActivated(int);`), `src/MainWindow.cpp`.

- [ ] **Step 1: Write `onStoryCellEdited`** — mirrors `onTableCellChanged` but indexes the data row directly (no visible-row mapping) and adds the LiveSync commit inline:

```cpp
void MainWindow::onStoryCellEdited(int dataRow, int col, const QString& text) {
    SheetResult* sheet = currentSheet();
    FileResult*  file  = currentFile();
    if (!sheet || !file || sheet->samples.isEmpty()) return;
    if (m_currentSampleIndex >= sheet->samples.size()) return;
    SampleResult& sample = sheet->samples[m_currentSampleIndex];
    if (dataRow < 0 || dataRow >= sample.rows.size()) return;
    DataRow& dr = sample.rows[dataRow];

    switch (col) {                                  // qualitative only
        case Cols::RESISTANCE: if (sheet->hasPerRowRegime) dr.puffingRegime = text; break;
        case Cols::SMELL: dr.smell = text; break;
        case Cols::CLOG:  dr.clog  = text; break;
        case Cols::NOTES: dr.notes = text; break;
        default: return;
    }
    recalculateSampleMetrics(*sheet);
    if (currentSheetHasCleanup())
        m_plotWidget->setSheetData(buildCleanedSheet(*sheet, m_currentFileIndex, m_currentSheetIndex));
    else
        m_plotWidget->setSheetData(*sheet);
    markFileModified();
    if (sheet->hasPerRowRegime) refreshPlotRegimes();

    // LiveSync per-cell (mirror onDataTableItemChanged); liveColumnForDataCol maps col->DB column
    if (m_liveSync && dr.id > 0) {
        const QString column = liveColumnForDataCol(col);
        if (!column.isEmpty()) m_liveSync->commitCell(QStringLiteral("data_rows"), dr.id, column, text);
    }
    // Excel write-back (same formula as onTableCellChanged)
    int excelRow = dataRow + 5;
    int excelCol = m_currentSampleIndex * 12 + col + 1;
    queueExcelWrite(file->filePath, sheet->sheetName, excelRow, excelCol, text);
    // Offline retry capture: reuse the existing PendingEdit block from onTableCellChanged.
}
```

  Note: if 2b.0/Task 10 removed `liveColumnForDataCol`, restore the 4 needed mappings inline (`NOTES->"notes"`, `SMELL->"smell"`, `CLOG->"clog"`, `RESISTANCE->"puffing_regime"`); confirm the actual `data_rows` column names against `LiveSync`'s allow-list (`LiveSync.cpp:137-148`).
- [ ] **Step 2: Write `onStoryNoteActivated(int dataRow)`** — v1 note→plot: `m_plotWidget->selectPuff(int(sample.rows[dataRow].puffs));` (slot added in Task 11).
- [ ] **Step 3: Build + smoke** — edit a note/smell/clog in the panel; confirm the Excel file updates (open it) and, with a test Postgres up, that `data_rows` receives the commit. **Commit.**

---

### Task 9: Rework `onPropCellChanged` to refresh the panel, not the table

**Files:** Modify `src/MainWindow.cpp` (`onPropCellChanged`, TPM branch ~lines 2169-2201 pre-2b.0).

- [ ] **Step 1:** Delete the `m_dataTable` calc-column refresh loop in the TPM branch; replace with `m_storyPanel->setSample(s, exclusionsFor(...), sheet->hasPerRowRegime);` after `recalculateSampleMetrics`. Keep the cleanup-aware `m_plotWidget->setSheetData(...)` push and `queueExcelWrite`. Sensory branch unchanged.
- [ ] **Step 2: Build + smoke** — edit a Sample-Property that affects power (e.g. Voltage); confirm panel TPM context updates. **Commit.**

---

### Task 10: Remove the `m_dataTable` widget surface (keep the plumbing)

**Files:** Modify `src/MainWindow.cpp`/`.h`; delete the two delegates + `tst_cellfocusdelegate`; per spike §5 + owner decision #3 (KEEP `RemoteCellHelpers` + `tst_mainwindow_remotecell`).

- [ ] **Step 1:** Remove the `m_dataTable` construction block in `setupCentralWidget`, its 6 connects, `onTableCellChanged`, `onAddRow`/`onRemoveRow`, `onDataTableItemChanged`/`Clicked`, the raw/SOP in-table render + strikethrough in `displayCurrentSample`, the `CellFocusDelegate` + `RegimeComboDelegate` (`.cpp/.h` + `tst_cellfocusdelegate` + its `tests/tests.pro` entry), and the orphaned members/helpers (`m_applyingRemote`, `dataTableHeaders`, `liveColumnForDataCol`/`dataColForLiveColumn` — but first inline the 4 mappings into Task 8, `kDataTableColumns` & friends). Split `handleRemoteRowChange` (keep the T19 row-deleted banner head; drop the T18 `data_rows` tail). Delete the dead `buildViewTab`/`onViewDataTable/Plots/Both`/`onZoomIn/Out/FitToWindow` (already-dead "View" tab). **Keep** `m_liveSync`, `onPropCellChanged` (reworked), `m_propTable`, the sample-nav widgets, and the whole-file `data_rows` DB save.
- [ ] **Step 2: Build `-Werror -Wall -Wextra`** — fix every `-Wunused` orphan (the `m_pendingEdits`/`flushPendingEdits`/`PendingEdit` unit is a delete-together set incl. its `onConnectionCameOnline` call site, OR keep it if Task 8 reuses the offline capture). Run full suite: `tests\run-tests.ps1`. Expected: green (minus known-flaky `tst_responsivelayout`).
- [ ] **Step 3: Commit** — one atomic `refactor(tpm): remove m_dataTable widget; story panel is the TPM edit surface (DATAVIEWER-6)`.

---

### Task 11: PlotEngine — per-point rings + emphasized point (v1 note→plot)

**Files:** Modify `src/plotting/PlotEngine.h`/`.cpp`, `src/plotting/PlotWidget.h`/`.cpp`.

- [ ] **Step 1: Write the failing test** in `tests/tst_plotengine/tst_plotengine.cpp`: render a `PlotSeries` with `ringed={1}` and assert the returned pixmap is non-null and differs from the same series without rings (pixel-diff at the ringed point's neighbourhood). Run → FAIL.
- [ ] **Step 2: Extend `PlotSeries`** with `QVector<int> ringed;` and `int emphasized = -1;` (indices into `x`/`y`). In `renderLinePlot`, after drawing dots, draw an amber ring (`QColor(0xBA,0x75,0x17)`, 2px, radius `dotRadius+5`) at each `ringed` index, and for `emphasized` add a larger ring + a dashed vertical guide to the x-axis. Round all drawn coordinates.
- [ ] **Step 3: PlotWidget** — add `void selectPuff(int puffs);` storing `m_selectedPuff`; when building the TPM-trend `PlotSeries` for the current sample, set `ps.ringed` to the indices of rows whose `notes` is non-empty and `ps.emphasized` to the index matching `m_selectedPuff`. Re-render. (Multi-sample plots: ring only on each sample's own series.)
- [ ] **Step 4: Run tests** → PASS. Build app. Smoke: clicking a note card rings + guides its plot point. **Commit.**

---

### Task 12: Raw/SOP empty-state

**Files:** Modify `src/MainWindow.cpp` (`displayCurrentSample` raw/SOP branch) + `NotesStoryPanel`.

- [ ] **Step 1:** When a raw/SOP sheet (or no samples) is selected, `m_storyPanel->clear()` and show a hint label ("Use **View Raw Data** to open this sheet in Excel") — add `NotesStoryPanel::showHint(const QString&)`.
- [ ] **Step 2: Build + smoke + Commit.**

---

### Task 13: Final verification

- [ ] **Step 1:** `python tools/decrypt_via_copy.py --apply`; `qmake CONFIG+=release && mingw32-make clean && mingw32-make -j8` — `-Werror` clean.
- [ ] **Step 2:** `tests\run-tests.ps1` — full suite green except known-flaky `tst_responsivelayout`; `tst_notesstory` + the new `tst_plotengine` ring assertion pass.
- [ ] **Step 3:** Manual smoke per spec §11 acceptance: panel fits with zero horizontal overflow at narrow widths; edit note/smell/clog/regime → Excel + DB update; excluded rows shown+marked+readable; aggregates match plot/report; note ring + note→plot emphasis; sample-nav + mode-switch intact; raw/SOP hint.
- [ ] **Step 4:** Append `tasks/lessons.md`: "the TPM grid was the edit + live-collab + offline-capture + raw-render + cleanup-strikethrough surface, and `onPropCellChanged` also wrote it — Option B kept all that behaviour and re-pointed it at NotesStoryPanel."
- [ ] **Step 5: Commit** the lessons + any final fixups.

---

## Self-review notes

- **Spec coverage:** §2 re-presentation → Tasks 6–10; §4.1 note card → Task 3; §4.2 summaries → Task 4; §4.3 live edit → Task 8; §5 excluded rows → Tasks 5,7; §6 v1 linking → Tasks 8,11 (Stage-2 = DATAVIEWER-18, out of scope); §7 View Raw Data = DATAVIEWER-17 (separate); §3 no-overflow → Task 2 scaffold; §9 phasing → task order. Cleanup (§2b.0) = DATAVIEWER-16 (parallel, gates Task 6).
- **Type consistency:** `cellEdited(int,int,QString)` col is a `DVE::Cols` value used identically in Tasks 2/3/4/8; `selectPuff(int)`/`highlightRow(int)`/`noteActivated(int)` signatures consistent across Tasks 2/5/8/11.
- **Open (spec §12):** smell editor = `QSpinBox` (chosen); regime editor = editable `QComboBox` (chosen); promote-on-add-note = host re-populate after the note becomes non-empty (Task 8 emits, `displayCurrentSample` rebuilds).
