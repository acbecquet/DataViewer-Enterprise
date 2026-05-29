# Per-Row Puffing Regime (v2.2.1) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the unused per-row *Resistance* column with a per-row *Puffing Regime* string column, propagated through the pipeline, PostgreSQL + offline SQLite, the editable data table, the live plot (new regime filter), and PowerPoint reports (per-regime plot/slide fan-out) — with old-template files continuing to work unchanged.

**Architecture:** Additive & non-destructive. `DataRow` gains `QString puffingRegime` beside the preserved `double resistance`; `data_rows` gains a `puffing_regime TEXT` column beside `resistance`. A per-sheet `SheetResult::hasPerRowRegime` flag — set from the column-E header at Excel ingest, derived from non-NULL regime data on DB load — switches one physical UI column (index 4) between Resistance (number) and Puffing Regime (string) behaviour. A new pure-function module `RegimeUtils` centralizes regime detection, unique-extraction, and filter-and-recompute; the plot widget and report generator consume it.

**Tech Stack:** C++17, Qt 6.10 (Widgets/SQL/Concurrent), qmake + MinGW 13.1, PostgreSQL 16 (QPSQL) + SQLite snapshot, Qt Test, openpyxl (bundled Python 3.11 / MIP-allowlisted Python 3.13 for tooling).

---

## Conventions for every task

- **MIP / file creation (this machine):** Before *any* C++ build, run `python tools/decrypt_via_copy.py --apply` from the repo root. Create **new** `.cpp`/`.h` files with Python's delete-and-rewrite pattern (see project `CLAUDE.md`), never the Edit/Write create path, so they don't inherit MIP labels:
  ```python
  import os
  path = "src/pipeline/RegimeUtils.h"
  content = r"""...full file contents..."""
  if os.path.exists(path): os.remove(path)
  with open(path, "w", encoding="utf-8", newline="\n") as f: f.write(content)
  ```
  Editing *existing* files with the Edit tool is fine (re-run the decrypt script first if a read shows `%TSD-Header-###%`).
- **Build the app (verification):** out-of-tree per `CLAUDE.md`:
  ```bat
  set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
  cd build && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8
  ```
  Build is `-Werror -Wall -Wextra` — do not downgrade warnings; fix the code.
- **Run the unit suite:** `tests\run-tests.ps1` (PowerShell) auto-detects Qt+MinGW, builds incrementally via `tests/tests.pro`, runs all suites. DB suites need `tests\start-test-postgres.ps1` first (sets `DVE_TEST_PG_CONN`); they skip cleanly when it's unset.
- **Commit cadence:** one commit per completed step group as written. Branch is `feature/per-row-puffing-regime` (already created).
- **Regenerate fixtures:** `python tests/generate_fixtures.py` writes `tests/data/*.xlsx`.

---

## Task 1: `RegimeUtils` pure-function module

Centralizes all regime logic so the pipeline, plot, and reports share one tested implementation.

**Files:**
- Create: `src/pipeline/RegimeUtils.h`
- Create: `src/pipeline/RegimeUtils.cpp`
- Create: `tests/tst_regimeutils/tst_regimeutils.pro`
- Create: `tests/tst_regimeutils/tst_regimeutils.cpp`
- Modify: `tests/tests.pro` (add `tst_regimeutils` to SUBDIRS)
- Modify: `DataViewerEnterprise.pro` (add `RegimeUtils.cpp`/`.h` to SOURCES/HEADERS)

### API (what RegimeUtils provides)

```cpp
namespace DVE::RegimeUtils {
    // "(unspecified)" bucket label for blank per-row regimes.
    QString unspecifiedLabel();

    // The regime key used for matching/bucketing one row: trimmed
    // puffingRegime, or unspecifiedLabel() when blank.
    QString regimeKey(const DataRow& row);

    // True if a column-E header string designates a per-row regime column
    // (case-insensitive contains "puffing" or "regime").
    bool isRegimeHeader(const QString& colEHeader);

    // Ordered (first-seen) unique NON-BLANK regimes across all rows of a sheet
    // / file. Used to populate the plot regime picker. Empty if no row carries
    // a non-blank regime.
    QStringList uniqueRegimes(const SheetResult& sheet);
    QStringList uniqueRegimes(const FileResult& file);

    // Like uniqueRegimes(sheet) but includes the "(unspecified)" bucket when
    // some rows are blank — drives report fan-out so no rows are dropped.
    QStringList uniqueRegimeKeys(const SheetResult& sheet);

    // Whether any row in the sheet carries a non-blank per-row regime.
    bool sheetHasRegimeData(const SheetResult& sheet);

    // Return a copy of `sheet` keeping only rows whose regimeKey == regime,
    // with per-sample and sheet metrics recomputed; samples that end up with
    // zero rows are dropped. Used by report fan-out (and as the reference
    // implementation the live plot mirrors inline).
    SheetResult filterByRegime(const SheetResult& sheet, const QString& regime);
}
```

- [ ] **Step 0: Add the data-model fields RegimeUtils depends on** — `src/pipeline/ReportData.h`

These two fields are consumed by RegimeUtils (this task) and every later task, so they go in first.
In `struct DataRow`, after the `resistance` line (line 19):
```cpp
    double resistance      = 0.0;
    QString puffingRegime;          // per-row regime (new template); empty on old template
```
In `struct SheetResult`, after `QString templateVersion;` (line 94):
```cpp
    QString templateVersion;   // "new" | "old"
    bool    hasPerRowRegime = false;   // true => column index 4 is a per-row Puffing Regime string, not Resistance
```
Leave `namespace Cols` / `COUNT = 12` unchanged — index 4 is dual-purpose, not a new column.

- [ ] **Step 1: Write the failing test** — `tests/tst_regimeutils/tst_regimeutils.cpp`

```cpp
#include <QtTest>
#include "../../src/pipeline/RegimeUtils.h"
#include "../../src/pipeline/ReportData.h"

using namespace DVE;

class TstRegimeUtils : public QObject {
    Q_OBJECT
private:
    static DataRow row(double puffs, double before, double after, const QString& regime) {
        DataRow r; r.puffs = puffs; r.beforeWeight = before; r.afterWeight = after;
        r.puffingRegime = regime; return r;
    }
    static SheetResult twoRegimeSheet() {
        SheetResult s; s.sheetName = "Lifetime Test"; s.hasPerRowRegime = true;
        SampleResult sm; sm.sampleName = "S1"; sm.power = 4.0;
        sm.rows << row(10, 25.10, 25.06, "60mL/3s/30s")
                << row(20, 25.06, 25.02, "60mL/3s/30s")
                << row(30, 25.02, 24.97, "200mL/9s/300s")
                << row(40, 24.97, 24.93, "200mL/9s/300s");
        s.samples << sm;
        return s;
    }
private slots:
    void isRegimeHeader_matches() {
        QVERIFY(RegimeUtils::isRegimeHeader("Puffing Regime"));
        QVERIFY(RegimeUtils::isRegimeHeader("puffing regime"));
        QVERIFY(!RegimeUtils::isRegimeHeader("Resistance (Ω)"));
        QVERIFY(!RegimeUtils::isRegimeHeader("Resistance"));
    }
    void regimeKey_blankBecomesUnspecified() {
        QCOMPARE(RegimeUtils::regimeKey(row(1,0,0,"  ")), RegimeUtils::unspecifiedLabel());
        QCOMPARE(RegimeUtils::regimeKey(row(1,0,0,"60mL/3s/30s")), QString("60mL/3s/30s"));
    }
    void uniqueRegimes_firstSeenOrder() {
        const QStringList u = RegimeUtils::uniqueRegimes(twoRegimeSheet());
        QCOMPARE(u, (QStringList{"60mL/3s/30s", "200mL/9s/300s"}));
    }
    void uniqueRegimeKeys_includesUnspecifiedForBlanks() {
        SheetResult s; SampleResult sm;
        sm.rows << row(10,25.10,25.06,"60mL/3s/30s") << row(20,25.06,25.02,"");
        s.samples << sm;
        QCOMPARE(RegimeUtils::uniqueRegimeKeys(s),
                 (QStringList{"60mL/3s/30s", RegimeUtils::unspecifiedLabel()}));
        QVERIFY(RegimeUtils::uniqueRegimes(s) == QStringList{"60mL/3s/30s"});  // picker omits blank
    }
    void sheetHasRegimeData_true() {
        QVERIFY(RegimeUtils::sheetHasRegimeData(twoRegimeSheet()));
        SheetResult plain; SampleResult sm; sm.rows << row(10,25.1,25.06,"");
        plain.samples << sm;
        QVERIFY(!RegimeUtils::sheetHasRegimeData(plain));
    }
    void filterByRegime_keepsOnlyMatchingRowsAndRecomputes() {
        const SheetResult f = RegimeUtils::filterByRegime(twoRegimeSheet(), "200mL/9s/300s");
        QCOMPARE(f.samples.size(), 1);
        QCOMPARE(f.samples[0].rows.size(), 2);
        for (const DataRow& r : f.samples[0].rows)
            QCOMPARE(r.puffingRegime, QString("200mL/9s/300s"));
        // averageTPM recomputed over the 2 kept rows (both ~ (before-after)*1000/interval)
        QVERIFY(f.samples[0].averageTPM > 0.0);
    }
    void filterByRegime_dropsEmptySamples() {
        const SheetResult f = RegimeUtils::filterByRegime(twoRegimeSheet(), "nonexistent");
        QCOMPARE(f.samples.size(), 0);
    }
};
QTEST_MAIN(TstRegimeUtils)
#include "tst_regimeutils.moc"
```

- [ ] **Step 2: Create the test `.pro`** — first `Read tests/tst_sheetprocessors/tst_sheetprocessors.pro` to copy the exact `QT`/`CONFIG`/`INCLUDEPATH` lines, then create `tests/tst_regimeutils/tst_regimeutils.pro` mirroring it with these sources:

```pro
QT       += testlib core
CONFIG   += console c++17 testcase
CONFIG   -= app_bundle
TEMPLATE  = app
TARGET    = tst_regimeutils

SOURCES += tst_regimeutils.cpp \
    ../../src/pipeline/RegimeUtils.cpp \
    ../../src/pipeline/SheetProcessors.cpp \
    ../../src/pipeline/TpmCalculator.cpp \
    ../../src/ExcelReader.cpp
HEADERS += ../../src/pipeline/RegimeUtils.h
```
(Reconcile `QT`/`INCLUDEPATH`/extra sources with the sibling `.pro` if it differs — `ExcelReader.cpp` is needed because `SheetProcessors.h` includes `ExcelReader.h`.)

- [ ] **Step 3: Add to the test runner** — in `tests/tests.pro`, add `tst_regimeutils` to the `SUBDIRS` list (mirror an existing entry).

- [ ] **Step 4: Run the test, verify it fails to build/link** (RegimeUtils not yet created)

Run: `tests\run-tests.ps1`
Expected: `tst_regimeutils` fails — `RegimeUtils.h: No such file` / unresolved `DVE::RegimeUtils::*`.

- [ ] **Step 5: Create `src/pipeline/RegimeUtils.h`** (via Python delete-and-rewrite)

```cpp
#pragma once

#include "ReportData.h"
#include <QString>
#include <QStringList>

namespace DVE {
namespace RegimeUtils {

QString     unspecifiedLabel();
QString     regimeKey(const DataRow& row);
bool        isRegimeHeader(const QString& colEHeader);
QStringList uniqueRegimes(const SheetResult& sheet);     // non-blank, for the picker
QStringList uniqueRegimes(const FileResult& file);
QStringList uniqueRegimeKeys(const SheetResult& sheet);  // incl. "(unspecified)", for report fan-out
bool        sheetHasRegimeData(const SheetResult& sheet);
SheetResult filterByRegime(const SheetResult& sheet, const QString& regime);

} // namespace RegimeUtils
} // namespace DVE
```

- [ ] **Step 6: Create `src/pipeline/RegimeUtils.cpp`** (via Python delete-and-rewrite)

```cpp
#include "RegimeUtils.h"
#include "SheetProcessors.h"

#include <QSet>

namespace DVE {
namespace RegimeUtils {

QString unspecifiedLabel() { return QStringLiteral("(unspecified)"); }

QString regimeKey(const DataRow& row)
{
    const QString t = row.puffingRegime.trimmed();
    return t.isEmpty() ? unspecifiedLabel() : t;
}

bool isRegimeHeader(const QString& colEHeader)
{
    const QString h = colEHeader.trimmed();
    return h.contains(QStringLiteral("puffing"), Qt::CaseInsensitive)
        || h.contains(QStringLiteral("regime"),  Qt::CaseInsensitive);
}

bool sheetHasRegimeData(const SheetResult& sheet)
{
    for (const SampleResult& s : sheet.samples)
        for (const DataRow& r : s.rows)
            if (!r.puffingRegime.trimmed().isEmpty())
                return true;
    return false;
}

QStringList uniqueRegimes(const SheetResult& sheet)
{
    QStringList ordered;
    QSet<QString> seen;
    for (const SampleResult& s : sheet.samples) {
        for (const DataRow& r : s.rows) {
            const QString t = r.puffingRegime.trimmed();
            if (t.isEmpty()) continue;            // blanks never populate the picker
            if (!seen.contains(t)) { seen.insert(t); ordered << t; }
        }
    }
    return ordered;
}

QStringList uniqueRegimes(const FileResult& file)
{
    QStringList ordered;
    QSet<QString> seen;
    for (const SheetResult& sheet : file.sheets) {
        for (const QString& t : uniqueRegimes(sheet)) {
            if (!seen.contains(t)) { seen.insert(t); ordered << t; }
        }
    }
    return ordered;
}

QStringList uniqueRegimeKeys(const SheetResult& sheet)
{
    QStringList ordered;
    QSet<QString> seen;
    for (const SampleResult& s : sheet.samples) {
        for (const DataRow& r : s.rows) {
            const QString k = regimeKey(r);     // blanks bucket as "(unspecified)"
            if (!seen.contains(k)) { seen.insert(k); ordered << k; }
        }
    }
    return ordered;
}

SheetResult filterByRegime(const SheetResult& sheet, const QString& regime)
{
    GenericSheetProcessor proc;          // for calculateMetrics / computeSheetAggregates
    SheetResult out = sheet;             // copy headers, flags, etc.
    out.samples.clear();

    for (const SampleResult& src : sheet.samples) {
        SampleResult s = src;
        s.rows.clear();
        for (const DataRow& r : src.rows)
            if (regimeKey(r) == regime)
                s.rows.append(r);
        if (s.rows.isEmpty())
            continue;                    // drop samples with no matching rows
        proc.calculateMetrics(s);        // recompute tpm/avg/stddev/etc. over kept rows
        out.samples.append(s);
    }
    proc.computeSheetAggregates(out);    // recompute sheet trend / overall aggregates
    return out;
}

} // namespace RegimeUtils
} // namespace DVE
```

- [ ] **Step 7: Add to the main `.pro`** — in `DataViewerEnterprise.pro`, add `src/pipeline/RegimeUtils.cpp` to `SOURCES` and `src/pipeline/RegimeUtils.h` to `HEADERS` (next to the other `pipeline/` entries).

- [ ] **Step 8: Run the test, verify it passes**

Run: `tests\run-tests.ps1`
Expected: `tst_regimeutils` PASS (all 6 functions).

- [ ] **Step 9: Commit**

```bash
git add src/pipeline/RegimeUtils.h src/pipeline/RegimeUtils.cpp \
        tests/tst_regimeutils/ tests/tests.pro DataViewerEnterprise.pro
git commit -m "feat(pipeline): RegimeUtils — per-row regime detect/extract/filter helpers"
```

---

## Task 2: Data model + pipeline parsing

Add the per-row field + per-sheet flag and make column-index 4 parse as a regime string for new-template sheets.

**Files:**
- Modify: `src/pipeline/ReportData.h` (DataRow + SheetResult)
- Modify: `src/pipeline/SheetProcessors.h` (setter on base) and `.cpp` (dual-mode read)
- Modify: `src/ExcelReader.cpp` (`getColumnHeaders` recognises the regime header)
- Modify: `src/pipeline/DataProcessor.cpp` (compute + propagate the flag)
- Test: `tests/tst_sheetprocessors/tst_sheetprocessors.cpp`

- [ ] **Step 1: Write the failing test** — append to `tests/tst_sheetprocessors/tst_sheetprocessors.cpp` (inside the existing test class; `Read` the file first to match its sample-building helpers/macros):

```cpp
void perRowRegime_newTemplate_readsStringIntoPuffingRegime() {
    ExcelReader::SampleData raw;
    raw.metadata.sampleID = "S1";
    // 12-col data row: col 4 holds a regime string instead of resistance.
    QVector<QVariant> r0 = {10, 25.10, 25.06, 0.45, QString("60mL/3s/30s"),
                            "", "", "", 0, 0, 0, 0};
    QVector<QVariant> r1 = {20, 25.06, 25.02, 0.42, QString("200mL/9s/300s"),
                            "", "", "", 0, 0, 0, 0};
    raw.dataRows << r0 << r1;

    GenericSheetProcessor proc;
    proc.setPerRowRegime(true);
    SheetResult sheet = proc.process({raw}, "Lifetime Test", "new");

    QCOMPARE(sheet.hasPerRowRegime, true);
    QCOMPARE(sheet.samples[0].rows[0].puffingRegime, QString("60mL/3s/30s"));
    QCOMPARE(sheet.samples[0].rows[1].puffingRegime, QString("200mL/9s/300s"));
    QCOMPARE(sheet.samples[0].rows[0].resistance, 0.0);   // not parsed as a number
}

void perRowRegime_oldTemplate_readsResistanceNumber() {
    ExcelReader::SampleData raw;
    raw.metadata.sampleID = "S1";
    QVector<QVariant> r0 = {10, 25.10, 25.06, 0.45, 1.25 /*resistance*/,
                            "", "", "", 0, 0, 0, 0};
    raw.dataRows << r0;

    GenericSheetProcessor proc;          // setPerRowRegime defaults false
    SheetResult sheet = proc.process({raw}, "Lifetime Test", "old");

    QCOMPARE(sheet.hasPerRowRegime, false);
    QCOMPARE(sheet.samples[0].rows[0].resistance, 1.25);
    QVERIFY(sheet.samples[0].rows[0].puffingRegime.isEmpty());
}
```

- [ ] **Step 2: Run, verify it fails to build** (`setPerRowRegime`, `hasPerRowRegime`, `DataRow::puffingRegime` don't exist)

Run: `tests\run-tests.ps1`
Expected: compile error — no member `setPerRowRegime` / `hasPerRowRegime` / `puffingRegime`.

- [ ] **Step 3: Confirm the data-model fields exist** — `src/pipeline/ReportData.h`

`DataRow::puffingRegime` and `SheetResult::hasPerRowRegime` were added in **Task 1 Step 0**. Confirm they're present (they're consumed below) and do **not** re-add them. `namespace Cols` / `COUNT = 12` stay unchanged — index 4 is dual-purpose, not a new column.

- [ ] **Step 4: Add the processor flag** — `src/pipeline/SheetProcessors.h`

In `class SheetProcessor`, add a public setter + protected member:
```cpp
    void calculateMetrics(SampleResult& s);
    void computeSheetAggregates(SheetResult& sheet);

    // When true, per-row column index 4 is read as a Puffing Regime string
    // (new template) instead of a Resistance double. Set by DataProcessor
    // from the column-E header before process() is called.
    void setPerRowRegime(bool v) { m_perRowRegime = v; }

protected:
    bool m_perRowRegime = false;
    SampleResult buildSampleResult(const ExcelReader::SampleData& raw,
                                   int sampleIndex,
                                   const QString& templateVersion);
```

- [ ] **Step 5: Dual-mode read + flag propagation** — `src/pipeline/SheetProcessors.cpp`

In `buildSampleResult`, replace the resistance read (line 114):
```cpp
        dr.drawPressure  = varToDouble(cell(ColIdx::DRAW_PRESSURE));
        if (m_perRowRegime)
            dr.puffingRegime = varToString(cell(ColIdx::RESISTANCE));   // col 4 = regime string
        else
            dr.resistance    = varToDouble(cell(ColIdx::RESISTANCE));   // col 4 = resistance number
        dr.smell         = varToString(cell(ColIdx::SMELL));
```
In `GenericSheetProcessor::process` (after `sheet.templateVersion = templateVersion;`, line 351):
```cpp
    sheet.templateVersion = templateVersion;
    sheet.hasPerRowRegime = m_perRowRegime;
```
(Add the identical `sheet.hasPerRowRegime = m_perRowRegime;` is **not** needed in the concrete processors — they all delegate to `GenericSheetProcessor::process`, which carries the flag. But `m_perRowRegime` must propagate to the delegate: in each concrete `process` that does `GenericSheetProcessor generic; return generic.process(...)`, add `generic.setPerRowRegime(m_perRowRegime);` before the call.)

For every concrete processor body in `SheetProcessors.cpp` (UserTestSim, LongPuff, RapidPuff, TempCycling, DeviceLife, ExtendedTest), change:
```cpp
    GenericSheetProcessor generic;
    generic.setPerRowRegime(m_perRowRegime);
    return generic.process(rawSamples, sheetName, templateVersion);
```

- [ ] **Step 6: Keep the regime header verbatim** — `src/ExcelReader.cpp`, `getColumnHeaders()`. No behaviour change is needed: a `"Puffing Regime"` header is already returned unchanged; only the legacy bare `"Resistance"` is normalised to `"Resistance (Ω)"`. Add a clarifying comment after that `if` (line 307-308) so the verbatim pass-through isn't "helpfully" normalised away later:
```cpp
        if (h.compare(QStringLiteral("Resistance"), Qt::CaseInsensitive) == 0)
            h = QStringLiteral("Resistance (Ω)");
        // New (v2.2.1) template: column-4 header is "Puffing Regime" — returned
        // verbatim; DataProcessor uses it to set SheetResult::hasPerRowRegime.
        headers << h;
```

- [ ] **Step 7: Compute + propagate the flag** — `src/pipeline/DataProcessor.cpp`

Add include at top:
```cpp
#include "RegimeUtils.h"
```
In `processSheet`, where the processor is created (after line 203 `std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));`):
```cpp
    std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));
    // Column index 4's header decides per-row Resistance vs Puffing Regime.
    const QStringList hdrs = reader.getColumnHeaders();
    const bool perRowRegime =
        (hdrs.size() > 4) && RegimeUtils::isRegimeHeader(hdrs.at(4));
    processor->setPerRowRegime(perRowRegime);
```
The flag now flows into `SheetResult::hasPerRowRegime` via `process()`. (`processFile` already sets `sheetResult.columnHeaders = reader.getColumnHeaders()` afterwards — leave it; the SOP/raw-table early-return path needs no flag.)

- [ ] **Step 8: Run, verify the two new tests pass**

Run: `tests\run-tests.ps1`
Expected: `tst_sheetprocessors` PASS including the two new functions; all previously-passing suites still PASS.

- [ ] **Step 9: Commit**

```bash
git add src/pipeline/ReportData.h src/pipeline/SheetProcessors.h \
        src/pipeline/SheetProcessors.cpp src/pipeline/DataProcessor.cpp \
        src/ExcelReader.cpp tests/tst_sheetprocessors/tst_sheetprocessors.cpp
git commit -m "feat(pipeline): per-row puffing regime field + per-sheet detection"
```

---

## Task 3: Database — additive column, migration, snapshot, live sync

Persist `data_rows.puffing_regime`. Old rows stay NULL; new-template rows write a (possibly empty) string. `hasPerRowRegime` is derived on load from non-NULL presence.

**Files:**
- Modify: `deploy/postgres/init.sql` (add column)
- Create: `deploy/postgres/migrations/2026-05-29-v2.2.1-per-row-regime.sql`
- Modify: `src/database/DatabaseManager.cpp` (INSERT/UPDATE/SELECT + bindings + load flag)
- Modify: `src/database/OfflineSnapshot.cpp` (schema, regenerate, loadFile read + flag)
- Modify: `src/database/LiveSync.cpp` (allowlist)
- Modify: `src/database/MigrationTool.cpp` (`kColsDataRows`)
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp`

- [ ] **Step 1: Write the failing test** — append to `tests/tst_databasemanager/tst_databasemanager.cpp` (Read it first for the existing connect/skip-if-no-PG harness and helpers). The test round-trips a new-template file and an old-template file:

```cpp
void perRowRegime_roundTrip_newAndOld() {
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
    DatabaseManager db;
    QVERIFY(db.connectUsing(qEnvironmentVariable("DVE_TEST_PG_CONN")));  // match existing helper

    // New-template file: rows carry regimes; sheet flagged.
    FileResult fNew; fNew.filePath = "regime_new.xlsx"; fNew.fileName = "regime_new.xlsx";
    SheetResult shNew; shNew.sheetName = "Lifetime Test"; shNew.hasPerRowRegime = true;
    SampleResult smNew; smNew.sampleName = "S1";
    DataRow a; a.puffs=10; a.beforeWeight=25.1; a.afterWeight=25.06; a.puffingRegime="60mL/3s/30s";
    DataRow b; b.puffs=20; b.beforeWeight=25.06; b.afterWeight=25.02; b.puffingRegime="200mL/9s/300s";
    smNew.rows << a << b; shNew.samples << smNew; fNew.sheets << shNew;
    QVERIFY(db.saveFile(fNew));

    FileResult loaded = db.loadFileByPath("regime_new.xlsx");
    QCOMPARE(loaded.sheets.size(), 1);
    QCOMPARE(loaded.sheets[0].hasPerRowRegime, true);
    QCOMPARE(loaded.sheets[0].samples[0].rows[0].puffingRegime, QString("60mL/3s/30s"));
    QCOMPARE(loaded.sheets[0].samples[0].rows[1].puffingRegime, QString("200mL/9s/300s"));

    // Old-template file: no regime; column stays NULL → flag false on load.
    FileResult fOld; fOld.filePath = "regime_old.xlsx"; fOld.fileName = "regime_old.xlsx";
    SheetResult shOld; shOld.sheetName = "Lifetime Test"; shOld.hasPerRowRegime = false;
    SampleResult smOld; smOld.sampleName = "S1";
    DataRow c; c.puffs=10; c.beforeWeight=25.1; c.afterWeight=25.06; c.resistance=1.25;
    smOld.rows << c; shOld.samples << smOld; fOld.sheets << shOld;
    QVERIFY(db.saveFile(fOld));

    FileResult loadedOld = db.loadFileByPath("regime_old.xlsx");
    QCOMPARE(loadedOld.sheets[0].hasPerRowRegime, false);
    QVERIFY(loadedOld.sheets[0].samples[0].rows[0].puffingRegime.isEmpty());
    QCOMPARE(loadedOld.sheets[0].samples[0].rows[0].resistance, 1.25);
}
```
(Adapt `connectUsing`/`saveFile` to the exact names the suite already uses.)

- [ ] **Step 2: Apply the schema column to the test DB and run, verify failure**

First make the column exist for the ephemeral test container by adding it to `init.sql` (Step 3) and re-running `tests\start-test-postgres.ps1`. Then:
Run: `tests\run-tests.ps1`
Expected: the new test FAILS — `puffingRegime` comes back empty for the new file (DAO not yet wired), or a SQL error if the column is missing.

- [ ] **Step 3: Add the column to the canonical schema** — `deploy/postgres/init.sql`, in `CREATE TABLE IF NOT EXISTS data_rows`, after the `resistance` line (line 80):
```sql
    resistance        DOUBLE PRECISION DEFAULT 0.0,
    puffing_regime    TEXT,
```

- [ ] **Step 4: Create the migration** — `deploy/postgres/migrations/2026-05-29-v2.2.1-per-row-regime.sql`:
```sql
-- v2.2.1 — per-row Puffing Regime column on data_rows.
-- Additive & non-destructive: old rows keep puffing_regime = NULL, which the
-- app reads as "old template (per-row Resistance)". New-template rows store a
-- (possibly empty) string. Safe to re-run.
BEGIN;
ALTER TABLE data_rows ADD COLUMN IF NOT EXISTS puffing_regime TEXT;
INSERT INTO schema_meta(key, value) VALUES ('schema_version', '3')
    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;
COMMIT;
```

- [ ] **Step 5: Wire the DAO** — `src/database/DatabaseManager.cpp`

(a) `data_rows` UPDATE prepare (line 416): insert `puffing_regime = ?` right before `updated_by = ?`:
```cpp
            "variation_tpm = ?, oil_consumed = ?, puffing_regime = ?, updated_by = ? "
            "WHERE id = ? AND version = ? RETURNING version") ||
```
(b) `data_rows` INSERT prepare (line 421): add the column + one placeholder:
```cpp
            "INSERT INTO data_rows (sample_id, sort_order, puffs, before_weight, after_weight, "
            "draw_pressure, resistance, smell, clog, notes, tpm, tpm_power_density, "
            "variation_tpm, oil_consumed, puffing_regime, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id, version")) {
```
(c) UPDATE binds (lines 606-609) — insert puffing_regime at index 14 and renumber `who`/`id`/`version`:
```cpp
                    updateRow.bindValue(13, dr.oilConsumed);
                    updateRow.bindValue(14, sheet.hasPerRowRegime ? QVariant(dr.puffingRegime) : QVariant());
                    updateRow.bindValue(15, who);
                    updateRow.bindValue(16, static_cast<qlonglong>(dr.id));
                    updateRow.bindValue(17, dr.version);
```
(d) INSERT binds (lines 640-641) — insert at index 14, renumber `who`:
```cpp
                    insertRow.bindValue(13, dr.oilConsumed);
                    insertRow.bindValue(14, sheet.hasPerRowRegime ? QVariant(dr.puffingRegime) : QVariant());
                    insertRow.bindValue(15, who);
```
(`sheet` is in scope — the outer loop var at line 452. `QVariant()` binds SQL NULL for old-template rows.)
(e) Bulk SELECT (line 949): append `dr.puffing_regime`:
```cpp
        q.prepare("SELECT dr.id, dr.sample_id, dr.version, dr.puffs, "
                  "dr.before_weight, dr.after_weight, dr.draw_pressure, "
                  "dr.resistance, dr.smell, dr.clog, dr.notes, dr.tpm, "
                  "dr.tpm_power_density, dr.variation_tpm, dr.oil_consumed, "
                  "dr.puffing_regime "
                  "FROM data_rows dr ...");          // rest unchanged
```
(f) Read it + track per-sample regime presence. Declare a set above the read loop (near line 946) and populate it in the loop:
```cpp
    QHash<qint64, QVector<DataRow>> rowsBySample;
    QSet<qint64> samplesWithRegime;     // sample ids that have a non-NULL puffing_regime
    {
        ...
            while (q.next()) {
                DataRow dr;
                ...
                dr.oilConsumed     = q.value(14).toDouble();
                const QVariant pr  = q.value(15);
                if (!pr.isNull()) { dr.puffingRegime = pr.toString(); samplesWithRegime.insert(sId); }
                rowsBySample[sId].append(dr);
            }
```
(g) Set the sheet flag during assembly (the `for (const TestInfo& ti : tests)` loop, ~line 1042). After the sample loop fills `sheet.samples`:
```cpp
        for (SampleResult& sr : samples) {
            sr.rows = rowsBySample.value(sr.id);
            if (samplesWithRegime.contains(sr.id)) sheet.hasPerRowRegime = true;
            ...
            sheet.samples.append(sr);
        }
```

- [ ] **Step 6: Mirror in the offline snapshot** — `src/database/OfflineSnapshot.cpp`

(a) Hardcoded schema (line 107), add after `resistance REAL DEFAULT 0.0,`:
```cpp
        resistance        REAL DEFAULT 0.0,
        puffing_regime    TEXT,
```
(b) `regenerate()` data_rows copy (lines 497-515): add `puffing_regime` to BOTH the source SELECT and the dest INSERT (append after `oil_consumed`, before `updated_at`) and bump the loop to 19:
```cpp
            if (!src.exec("SELECT id, sample_id, sort_order, puffs, before_weight, "
                          "after_weight, draw_pressure, resistance, smell, clog, notes, "
                          "tpm, tpm_power_density, variation_tpm, oil_consumed, "
                          "puffing_regime, updated_at, updated_by, version "
                          "FROM data_rows ORDER BY id")) { ... }
            ...
            dst.prepare("INSERT INTO data_rows (id, sample_id, sort_order, puffs, "
                        "before_weight, after_weight, draw_pressure, resistance, smell, "
                        "clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed, "
                        "puffing_regime, updated_at, updated_by, version) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            while (src.next()) {
                for (int c = 0; c < 19; ++c) dst.bindValue(c, src.value(c));
                ...
```
(c) Offline `loadFile` read (lines 1071-1098): append `puffing_regime` to the SELECT (becomes value(13)) and read + flag. Declare a sheet-scope `bool sheetHasRegime = false;` before the sample loop, set it in the row loop, and assign after:
```cpp
                q.prepare("SELECT id, puffs, before_weight, after_weight, draw_pressure, "
                          "resistance, smell, clog, notes, tpm, tpm_power_density, "
                          "variation_tpm, oil_consumed, puffing_regime "
                          "FROM data_rows WHERE sample_id = ? ORDER BY sort_order");
                ...
                while (q.next()) {
                    DataRow dr;
                    ...
                    dr.oilConsumed     = q.value(12).toDouble();
                    const QVariant pr  = q.value(13);
                    if (!pr.isNull()) { dr.puffingRegime = pr.toString(); sheetHasRegime = true; }
                    sr.rows.append(dr);
                }
```
After the sample loop, before `result.sheets.append(sheet);`:
```cpp
        sheet.hasPerRowRegime = sheetHasRegime;
        result.sheets.append(sheet);
```
(Place `bool sheetHasRegime = false;` at the top of the per-sheet block and reset per sheet.)

- [ ] **Step 7: Allow live sync of the new column** — `src/database/LiveSync.cpp`, in the `data_rows` allowlist set (line 99), add:
```cpp
            QStringLiteral("variation_tpm"), QStringLiteral("oil_consumed"),
            QStringLiteral("puffing_regime")
```

- [ ] **Step 8: Add to the migration-tool column list** — `src/database/MigrationTool.cpp`, `kColsDataRows` (line 98), append `"puffing_regime"`:
```cpp
static const QStringList kColsDataRows = {
    "id", "sample_id", "sort_order", "puffs", "before_weight", "after_weight",
    "draw_pressure", "resistance", "smell", "clog", "notes", "tpm",
    "tpm_power_density", "variation_tpm", "oil_consumed", "puffing_regime"
};
```

- [ ] **Step 9: Recreate the test DB and run, verify pass**

```powershell
docker rm -f dve-test-pg ; .\tests\start-test-postgres.ps1   # rebuilds with new init.sql
.\tests\run-tests.ps1
```
Expected: `tst_databasemanager` PASS including `perRowRegime_roundTrip_newAndOld`; other DB-dependent suites still PASS.

- [ ] **Step 10: Commit**

```bash
git add deploy/postgres/init.sql deploy/postgres/migrations/2026-05-29-v2.2.1-per-row-regime.sql \
        src/database/DatabaseManager.cpp src/database/OfflineSnapshot.cpp \
        src/database/LiveSync.cpp src/database/MigrationTool.cpp \
        tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "feat(db): data_rows.puffing_regime — additive column, migration, snapshot, livesync"
```

---

## Task 4: Editable data table (TPM mode)

Column index 4 becomes "Puffing Regime" (string, combo editor) for new-template sheets; stays "Resistance" (number) for old. LiveSync col↔DB-column map becomes sheet-aware.

**Files:**
- Create: `src/widgets/RegimeComboDelegate.h` / `.cpp`
- Modify: `src/MainWindow.h` (members + helper declarations)
- Modify: `src/MainWindow.cpp` (headers, populate, edit, write-back, column map, delegate wiring)
- Modify: `DataViewerEnterprise.pro` (add the delegate sources)

> **Testing note:** This task is UI wiring with no clean Qt Test seam; verify by build + the manual checklist in Step 9. The regime *logic* is already covered by `tst_regimeutils`.

- [ ] **Step 1: Make the data-table headers sheet-aware** — `src/MainWindow.cpp`

Change `dataTableHeaders()` (line 5232) to take the flag, and add the wrapper used at construction:
```cpp
QStringList MainWindow::dataTableHeaders(bool perRowRegime)
{
    return {"Puffs","Before (g)","After (g)","Pressure",
            perRowRegime ? "Puffing Regime" : "Resistance",
            "Smell","Clog","Notes","TPM (mg/puff)","TPM Pwr Density","Variation (%)","Oil Consumed (mg)"};
}
```
Update the declaration in `src/MainWindow.h` to `QStringList dataTableHeaders(bool perRowRegime = false);`.
The construction-time calls (lines 682-683, 2864-2865, 2877-2880) keep using `dataTableHeaders()` (defaults to the "Resistance" labels — the per-sheet relabel happens in `displayCurrentSample`, Step 4). Column *count* is identical either way (12), so `setColumnCount(dataTableHeaders().size())` is unaffected.

- [ ] **Step 2: Sheet-aware LiveSync column map** — `src/MainWindow.cpp` + `.h`

Add two private members (declare in `MainWindow.h`):
```cpp
    QString liveColumnForDataCol(int col) const;
    int     dataColForLiveColumn(const QString& dbColumn) const;
```
Implement (put near the other helpers, e.g. after `currentFile()`):
```cpp
QString MainWindow::liveColumnForDataCol(int col) const
{
    if (col == DVE::Cols::RESISTANCE) {   // dual-purpose column 4
        const SheetResult* s = currentSheet();
        return (s && s->hasPerRowRegime) ? QStringLiteral("puffing_regime")
                                         : QStringLiteral("resistance");
    }
    return columnNameForDataTableColumn(col);
}

int MainWindow::dataColForLiveColumn(const QString& dbColumn) const
{
    if (dbColumn == QLatin1String("puffing_regime")) return DVE::Cols::RESISTANCE;
    return dataTableColumnForColumnName(dbColumn);
}
```
Replace the call sites:
- line 722 (focus broadcast): `const QString column = liveColumnForDataCol(curr->column());`
- line 2544 (live commit on edit): `const QString column = liveColumnForDataCol(it->column());`
- line 2746 and line 2771 (`onRemoteCellChanged`): `const int col = dataColForLiveColumn(column);` (both occurrences)

- [ ] **Step 3: Edit handler — store regime string for new-template** — `src/MainWindow.cpp`, `onTableCellChanged` (line 1668):
```cpp
        case 4:
            if (sheet->hasPerRowRegime) dr.puffingRegime = text;
            else                         dr.resistance   = text.toDouble();
            break;
```
(The `col < 0 || col > 7` guard and the rest of the handler — recalc, plot refresh, write-back — are unchanged. The write-back at line 1748 already queues `text` (the string) to Excel cell `sampleIdx*12 + col + 1`. Live-refreshing the plot's regime picker after a regime edit is wired in Task 5 Step 5, once `refreshPlotRegimes()` exists.)

- [ ] **Step 4: Populate — render the regime cell + relabel header** — `src/MainWindow.cpp`, `displayCurrentSample`

Near the top of the function where `sheet`/`sample` are obtained, compute the flag and relabel the header once:
```cpp
    const bool perRowRegime = sheet && sheet->hasPerRowRegime;
    m_dataTable->setHorizontalHeaderLabels(dataTableHeaders(perRowRegime));
```
In the populate loop, replace the column-4 line (line 2955):
```cpp
        (dr.drawPressure == 0.0) ? setEmpty() : setNum(dr.drawPressure, 2);
        if (perRowRegime) setStr(dr.puffingRegime);
        else (dr.resistance == 0.0) ? setEmpty() : setNum(dr.resistance, 3);
        setStr(dr.smell);
```
(Both branches advance `col` by exactly 1, keeping alignment.)

- [ ] **Step 5: Create the combo delegate** — `src/widgets/RegimeComboDelegate.h` (Python delete-and-rewrite):
```cpp
#pragma once
#include "CellFocusDelegate.h"
#include <QStringList>

namespace DVE {

// Editable-combo editor for the per-row Puffing Regime column. Subclasses
// CellFocusDelegate so the remote-focus border / flash painting is preserved.
// When inactive (old-template sheet), falls back to the base (plain) editor.
class RegimeComboDelegate : public CellFocusDelegate {
    Q_OBJECT
public:
    explicit RegimeComboDelegate(QObject* parent = nullptr);
    void setActive(bool a) { m_active = a; }
    void setRegimes(const QStringList& r) { m_regimes = r; }

    QWidget* createEditor(QWidget* parent, const QStyleOptionViewItem& opt,
                          const QModelIndex& idx) const override;
    void setEditorData(QWidget* editor, const QModelIndex& idx) const override;
    void setModelData(QWidget* editor, QAbstractItemModel* model,
                      const QModelIndex& idx) const override;
private:
    bool        m_active = false;
    QStringList m_regimes;
};

} // namespace DVE
```
`src/widgets/RegimeComboDelegate.cpp`:
```cpp
#include "RegimeComboDelegate.h"
#include <QComboBox>

namespace DVE {

// Common regimes offered alongside the file's own, so the XXmL/YYs/ZZs format
// stays consistent. Free text is still allowed (editable combo).
static const QStringList kPresetRegimes = {
    "60mL/3s/30s", "200mL/10s/60s", "100mL/2.5s/15s", "60mL/2s/5s", "200mL/3s/30s"
};

RegimeComboDelegate::RegimeComboDelegate(QObject* parent)
    : CellFocusDelegate(parent) {}

QWidget* RegimeComboDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem& opt,
                                           const QModelIndex& idx) const
{
    if (!m_active)
        return CellFocusDelegate::createEditor(parent, opt, idx);  // plain editor for old template
    QComboBox* cb = new QComboBox(parent);
    cb->setEditable(true);
    QStringList items = m_regimes;
    for (const QString& p : kPresetRegimes)
        if (!items.contains(p)) items << p;
    cb->addItems(items);
    return cb;
}

void RegimeComboDelegate::setEditorData(QWidget* editor, const QModelIndex& idx) const
{
    if (auto* cb = qobject_cast<QComboBox*>(editor)) {
        cb->setCurrentText(idx.data(Qt::EditRole).toString());
        return;
    }
    CellFocusDelegate::setEditorData(editor, idx);
}

void RegimeComboDelegate::setModelData(QWidget* editor, QAbstractItemModel* model,
                                       const QModelIndex& idx) const
{
    if (auto* cb = qobject_cast<QComboBox*>(editor)) {
        model->setData(idx, cb->currentText().trimmed(), Qt::EditRole);
        return;
    }
    CellFocusDelegate::setModelData(editor, model, idx);
}

} // namespace DVE
```
(First `Read src/widgets/CellFocusDelegate.h` to confirm the class name, namespace, and that `createEditor`/`setEditorData`/`setModelData` are `virtual`/overridable. If `CellFocusDelegate` does not override `createEditor` itself, the base is `QStyledItemDelegate` and these overrides still work.)

- [ ] **Step 6: Wire the delegate + add the file-regimes helper** — `src/MainWindow.cpp` + `.h`

Add `#include "pipeline/RegimeUtils.h"` at the top of `MainWindow.cpp` (if not already present) and a private helper (declare `QStringList currentFileRegimes() const;` in `MainWindow.h`):
```cpp
QStringList MainWindow::currentFileRegimes() const
{
    const FileResult* f = const_cast<MainWindow*>(this)->currentFile();
    return f ? DVE::RegimeUtils::uniqueRegimes(*f) : QStringList();
}
```
Declare the delegate member in `MainWindow.h`: `DVE::RegimeComboDelegate* m_regimeDelegate = nullptr;` and forward-include the delegate header. In the data-table construction (after line 686 where `m_cellFocusDelegate` is set):
```cpp
    m_regimeDelegate = new DVE::RegimeComboDelegate(this);
    m_dataTable->setItemDelegateForColumn(DVE::Cols::RESISTANCE, m_regimeDelegate);
```
In `displayCurrentSample` (Step 4 area), keep the delegate in sync:
```cpp
    if (m_regimeDelegate) {
        m_regimeDelegate->setActive(perRowRegime);
        m_regimeDelegate->setRegimes(currentFileRegimes());
    }
```

- [ ] **Step 7: Add delegate to the `.pro`** — `DataViewerEnterprise.pro`: add `src/widgets/RegimeComboDelegate.cpp` to SOURCES and `.h` to HEADERS (next to `CellFocusDelegate`).

- [ ] **Step 8: Decrypt + build**

```bash
python tools/decrypt_via_copy.py --apply
```
```bat
cd build && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8
```
Expected: clean build, no warnings.

- [ ] **Step 9: Manual verification** (run `release\DataViewer.exe` or use the `run` skill)
  - Open a **new-template** file (after Task 7, or hand-edit a fixture): column 5 header reads **"Puffing Regime"**; cells show regime strings; double-click → editable dropdown with presets + free text; edit persists to the table and (if online) to the DB.
  - Open an **old-template** file: column 5 still reads **"Resistance"**, numeric, plain editor.

- [ ] **Step 10: Commit**

```bash
git add src/widgets/RegimeComboDelegate.h src/widgets/RegimeComboDelegate.cpp \
        src/MainWindow.h src/MainWindow.cpp DataViewerEnterprise.pro
git commit -m "feat(ui): data-table column 4 dual-mode (Resistance | Puffing Regime) + combo editor"
```

---

## Task 5: Live plot — regime picker + filtering

Add a regime dropdown between the plot-type combo and the Save button; filter plotted rows by the selected regime (correctly recomputing the bar chart).

**Files:**
- Modify: `src/plotting/PlotWidget.h` / `.cpp`
- Modify: `src/MainWindow.h` / `.cpp` (compute file regimes, wire the picker)

> **Testing note:** UI wiring — verify by build + manual checklist (Step 7). The row-filter/recompute semantics are covered by `tst_regimeutils::filterByRegime_*`.

- [ ] **Step 1: PlotWidget — add the regime combo** — `src/plotting/PlotWidget.h`

Add members:
```cpp
    QComboBox*   m_plotTypeCombo;
    QComboBox*   m_regimeCombo  = nullptr;   // "All regimes" + each unique per-row regime
    QLabel*      m_regimeLabel  = nullptr;
    QFrame*      m_regimeSep    = nullptr;
    QPushButton* m_saveBtn;
```
Add a public method + private slot:
```cpp
public:
    // Populate the regime picker with the file's unique per-row regimes.
    // Empty list hides the picker (old-template files). Preserves the current
    // selection if still present.
    void setAvailableRegimes(const QStringList& regimes);
private slots:
    void onRegimeChanged(int index);
```

- [ ] **Step 2: PlotWidget — build + place the combo** — `src/plotting/PlotWidget.cpp` ctor, between the plot-type combo and the separator (lines 71-73):
```cpp
    topLayout->addWidget(typeLabel);
    topLayout->addWidget(m_plotTypeCombo);

    m_regimeLabel = new QLabel("Regime:", topBar);
    m_regimeLabel->setStyleSheet("font-weight: 600; font-size: 9pt;");
    m_regimeCombo = new QComboBox(topBar);
    m_regimeCombo->addItem("All regimes");
    m_regimeCombo->setMinimumWidth(130);
    m_regimeCombo->setMaximumWidth(200);
    m_regimeLabel->setVisible(false);          // hidden until a file has regimes
    m_regimeCombo->setVisible(false);

    topLayout->addWidget(m_regimeLabel);
    topLayout->addWidget(m_regimeCombo);
    topLayout->addWidget(sep);
    topLayout->addWidget(m_saveBtn);
    topLayout->addStretch(1);
```
Connect in the ctor (with the other connects, line 131):
```cpp
    connect(m_regimeCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &PlotWidget::onRegimeChanged);
```

- [ ] **Step 3: PlotWidget — implement the slot + setter** — `src/plotting/PlotWidget.cpp`:
```cpp
void PlotWidget::setAvailableRegimes(const QStringList& regimes)
{
    const QString prev = m_regimeCombo ? m_regimeCombo->currentText() : QString();
    m_regimeCombo->blockSignals(true);
    m_regimeCombo->clear();
    m_regimeCombo->addItem("All regimes");
    m_regimeCombo->addItems(regimes);
    int idx = m_regimeCombo->findText(prev);
    m_regimeCombo->setCurrentIndex(idx >= 0 ? idx : 0);     // preserve selection
    m_regimeCombo->blockSignals(false);

    const bool show = !regimes.isEmpty();
    m_regimeLabel->setVisible(show);
    m_regimeCombo->setVisible(show);
    updatePlot();
}

void PlotWidget::onRegimeChanged(int /*index*/) { updatePlot(); }
```

- [ ] **Step 4: PlotWidget — filter rows by regime in render** — `src/plotting/PlotWidget.cpp`, `renderCurrentPlot`

Add the include at top: `#include "../pipeline/RegimeUtils.h"` and `#include "../pipeline/TpmCalculator.h"`.
At the start of `renderCurrentPlot` (after computing `plotType`), determine the active filter:
```cpp
    const QString selRegime = (m_regimeCombo && m_regimeCombo->isVisible())
                              ? m_regimeCombo->currentText() : QString();
    const bool filterRegime = !selRegime.isEmpty() && selRegime != "All regimes";
    auto rowMatches = [&](const DataRow& r) {
        return !filterRegime || RegimeUtils::regimeKey(r) == selRegime;
    };
```
In the **TPM Trend** fast-path guard (line 390), also require no regime filter:
```cpp
        if (!filterRegime && !m_currentSheet.tpmTrend.isEmpty() &&
            !m_currentSheet.puffCounts.isEmpty() &&
            (visIdx.isEmpty() || m_currentSheet.samples.size() <= 1))
```
In each per-row loop that builds series (TPM Trend primary line ~418, oil overlay ~461, Power Density ~513, Draw Pressure ~554), add `if (!rowMatches(row)) continue;` as the first line of the loop body (alongside the existing weight/pressure skips).
For the **TPM Bar Chart** (lines 481-488), recompute the average/stddev over matching rows instead of using the precomputed `sr.averageTPM`:
```cpp
        for (int si : visIdx) {
            if (si >= m_currentSheet.samples.size()) continue;
            const SampleResult& sr = m_currentSheet.samples[si];
            QString nm = sr.sampleName.isEmpty() ? QString("S%1").arg(si + 1) : sr.sampleName;
            if (!filterRegime) {
                names.append(nm); avgTPM.append(sr.averageTPM); stdDev.append(sr.stdDevTPM);
            } else {
                QVector<double> t;
                for (const DataRow& r : sr.rows) {
                    if (r.beforeWeight == 0.0 || r.afterWeight == 0.0) continue;
                    if (!rowMatches(r)) continue;
                    t.append(r.tpm);
                }
                if (t.isEmpty()) continue;   // sample has no rows for this regime
                names.append(nm);
                avgTPM.append(TpmCalculator::average(t));
                stdDev.append(TpmCalculator::stddev(t));
            }
        }
```
(Add `PlotWidget.cpp` already links `TpmCalculator` via the test/app `.pro`; confirm `TpmCalculator.cpp` is in `DataViewerEnterprise.pro` SOURCES — it is, as the pipeline uses it.)

- [ ] **Step 5: MainWindow — wire the picker** — `src/MainWindow.h` / `.cpp`

`currentFileRegimes()` already exists (Task 4 Step 6). Add `refreshPlotRegimes()` (declare `void refreshPlotRegimes();` in `.h`):
```cpp
void MainWindow::refreshPlotRegimes()
{
    if (m_plotWidget) m_plotWidget->setAvailableRegimes(currentFileRegimes());
}
```
Call `refreshPlotRegimes();` after a file becomes current: at the end of `onFileSelected`/`onSheetSelected` (wherever `displayCurrentSample()` is invoked on a file/sheet switch) and at the end of `onFileLoadFinished` (line ~2264, after the active file is set). Also add it to `onTableCellChanged` so a newly-typed regime appears in the picker — near the end (after `markFileModified();`, before the offline block):
```cpp
    if (sheet->hasPerRowRegime) refreshPlotRegimes();
```

- [ ] **Step 6: Decrypt + build**
```bash
python tools/decrypt_via_copy.py --apply
```
```bat
cd build && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8
```
Expected: clean build.

- [ ] **Step 7: Manual verification**
  - New-template file with ≥2 regimes: "Regime:" dropdown appears between Plot Type and Save, listing "All regimes" + each unique regime. Selecting one filters every plot type to those rows; TPM Bar Chart bars reflect the regime's rows only. "All regimes" restores the full plot. Sample checkboxes/zoom are preserved across regime changes.
  - Old-template file: no regime dropdown (hidden).

- [ ] **Step 8: Commit**
```bash
git add src/plotting/PlotWidget.h src/plotting/PlotWidget.cpp src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(plot): live regime picker filtering all plot types (bar chart recomputed)"
```

---

## Task 6: Reports — per-regime plot/slide fan-out

For each sheet, emit one content slide per unique per-row regime (filtered + recomputed), titled "{sheet} — {regime}". Old/no-regime sheets emit exactly one slide as today. De-duplicate the triplicated `kPlotLayout` + extract a shared emit helper.

**Files:**
- Modify: `src/reporting/ReportGenerator.h` (declare helper + layout constant)
- Modify: `src/reporting/ReportGenerator.cpp`
- Test: `tests/tst_reportgenerator/tst_reportgenerator.cpp`

- [ ] **Step 1: Write the failing test** — append to `tests/tst_reportgenerator/tst_reportgenerator.cpp` (Read it first for the existing harness; many report tests count slides via `PptxWriter` or by inspecting the saved `.pptx`). Add a count helper on `ReportGenerator` to make this testable without unzipping:

Add to `ReportGenerator` (public, declared in `.h`):
```cpp
    // Number of content slides a sheet would emit (1 per unique regime, or 1
    // for old/no-regime sheets). Pure function — used by reports and tests.
    static int regimeSlideCount(const SheetResult& sheet);
```
Test:
```cpp
void regimeSlideCount_fansOutByRegime() {
    SheetResult two; two.sheetName = "Lifetime Test"; two.hasPerRowRegime = true;
    SampleResult s; 
    DataRow a; a.beforeWeight=25.1; a.afterWeight=25.06; a.puffingRegime="60mL/3s/30s";
    DataRow b; b.beforeWeight=25.06; b.afterWeight=25.02; b.puffingRegime="200mL/9s/300s";
    s.rows << a << b; two.samples << s;
    QCOMPARE(ReportGenerator::regimeSlideCount(two), 2);

    SheetResult old; old.sheetName = "Lifetime Test";
    SampleResult so; DataRow c; c.beforeWeight=25.1; c.afterWeight=25.06; so.rows << c;
    old.samples << so;
    QCOMPARE(ReportGenerator::regimeSlideCount(old), 1);   // no regime → single slide
}
```

- [ ] **Step 2: Run, verify failure** (`regimeSlideCount` undefined)
Run: `tests\run-tests.ps1` → `tst_reportgenerator` fails to compile/link.

- [ ] **Step 3: Add the shared layout constant + helpers** — `src/reporting/ReportGenerator.cpp`

At file scope (after the includes, before the ctor), replace the three local `kPlotLayout` arrays with one shared definition:
```cpp
namespace {
// Fixed positions for the three report plots (inches): TPM Trend, Avg TPM Bar, Draw Pressure.
struct PlotSlot { double x, y, w, h; };
static const PlotSlot kPlotLayout[3] = {
    { 0.10, 3.25, 4.32, 3.20 },
    { 4.51, 3.25, 4.32, 3.20 },
    { 8.91, 3.25, 4.32, 3.20 },
};
QVector<SlideImage> layoutPlots(QVector<QByteArray> pngs) {
    QVector<SlideImage> imgs;
    for (int pi = 0; pi < pngs.size() && pi < 3; ++pi) {
        SlideImage img; img.pngData = std::move(pngs[pi]);
        img.x = kPlotLayout[pi].x; img.y = kPlotLayout[pi].y;
        img.w = kPlotLayout[pi].w; img.h = kPlotLayout[pi].h;
        imgs.append(std::move(img));
    }
    return imgs;
}
} // namespace
```
Add `#include "../pipeline/RegimeUtils.h"` at the top.

Add the slide-count helper + the fan-out emitter (declare both in `.h`; `emitSheetContentSlides` is a private member, `regimeSlideCount` is `static`):
`addContentSlide`'s signature is `addContentSlide(const QString& title, const SlideTable&, const QVector<SlideImage>&)` (confirm by reading `src/reporting/PptxWriter.h`). Both branches below pass the table.
```cpp
int ReportGenerator::regimeSlideCount(const SheetResult& sheet)
{
    if (!DVE::RegimeUtils::sheetHasRegimeData(sheet)) return 1;   // old / no regime
    return DVE::RegimeUtils::uniqueRegimeKeys(sheet).size();
}

// Emit one content slide per unique per-row regime key (filtered + recomputed,
// including an "(unspecified)" slide when some rows are blank so no rows drop),
// or a single unchanged slide for old / no-regime sheets. `sheet` must already
// be empty-sample-filtered by the caller.
void ReportGenerator::emitSheetContentSlides(PptxWriter& writer,
                                             const SheetResult& sheet,
                                             const ReportConfig& config,
                                             bool includeBarChart)
{
    if (!DVE::RegimeUtils::sheetHasRegimeData(sheet)) {           // old / no per-row regime → unchanged
        SlideTable tbl = buildTable(sheet, config);
        QVector<QByteArray> pngs = config.includePlots ? buildPlots(sheet, includeBarChart)
                                                       : QVector<QByteArray>{};
        writer.addContentSlide(sheet.sheetName, tbl, layoutPlots(std::move(pngs)));
        return;
    }
    for (const QString& key : DVE::RegimeUtils::uniqueRegimeKeys(sheet)) {
        const SheetResult fs = DVE::RegimeUtils::filterByRegime(sheet, key);
        if (!fs.hasSamples()) continue;
        SlideTable tbl = buildTable(fs, config);
        QVector<QByteArray> pngs = config.includePlots ? buildPlots(fs, includeBarChart)
                                                       : QVector<QByteArray>{};
        writer.addContentSlide(sheet.sheetName + QStringLiteral(" – ") + key,
                               tbl, layoutPlots(std::move(pngs)));
    }
}
```

- [ ] **Step 4: Call the helper from all three report paths** — `src/reporting/ReportGenerator.cpp`

In `generateFullReport` replace lines 436-468 (the buildTable + local kPlotLayout + addContentSlide block) with:
```cpp
        emitSheetContentSlides(writer, sheet, config, /*includeBarChart=*/true);
```
In `generateTestReport` replace lines 589-621 with:
```cpp
        reportProgress(progress, 50, "Building slides...");
        emitSheetContentSlides(writer, *target, config, /*includeBarChart=*/target->samples.size() > 1);
```
In `generateCombinedFullReport` replace lines 904-920 with:
```cpp
            emitSheetContentSlides(writer, sheet, ReportConfig{}, /*includeBarChart=*/true);
```
(Leave the image-slide blocks that follow each — they iterate `sheet.samples` and are regime-agnostic, so per-sample image slides still emit once. The local `kPlotLayout` arrays are now removed; the shared one is used inside the helper.)

- [ ] **Step 5: Run, verify pass**
Run: `tests\run-tests.ps1`
Expected: `tst_reportgenerator` PASS including `regimeSlideCount_fansOutByRegime`; existing report tests still PASS (old/single-regime path unchanged).

- [ ] **Step 6: Build the app to confirm the three call-site edits compile**
```bash
python tools/decrypt_via_copy.py --apply
```
```bat
cd build && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8
```

- [ ] **Step 7: Commit**
```bash
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp \
        tests/tst_reportgenerator/tst_reportgenerator.cpp
git commit -m "feat(report): per-regime plot/slide fan-out + dedupe kPlotLayout"
```

---

## Task 7: Template `.xlsx`, fixtures, version bump

Relabel the 36 `Resistance (Ω)` headers to `Puffing Regime` and pre-fill each block's data rows from its header regime; add a new-template unit fixture; bump VERSION.

**Files:**
- Modify (binary): `resources/templates/Standardized Test Template - December 2025.xlsx`
- Create: `tools/relabel_template_regime.py` (one-shot, kept for provenance)
- Modify: `tests/generate_fixtures.py` (new-template fixture)
- Modify: `DataViewerEnterprise.pro` (VERSION)

- [ ] **Step 1: Inspect `Temperature Cycling Test #1`** (the non-standard sheet). Run:
```bash
PYTHONIOENCODING=utf-8 PYTHONUTF8=1 python - <<'PY'
import openpyxl
wb = openpyxl.load_workbook("resources/templates/Standardized Test Template - December 2025.xlsx")
ws = wb["Temperature Cycling Test #1"]
for r in range(1, 12):
    print(r, [ws.cell(row=r, column=c).value for c in range(1, 14)])
PY
```
Decide from the output whether it has the standard 12-col blocks with a "Resistance (Ω)" header somewhere. The relabel script (Step 2) locates headers by content, so it will handle this sheet if present and skip it (logging) if not.

- [ ] **Step 2: Write the relabel/pre-fill script** — `tools/relabel_template_regime.py`:
```python
#!/usr/bin/env python3
"""v2.2.1 one-shot: relabel per-row 'Resistance (Ω)' -> 'Puffing Regime' and
pre-fill each sample block's regime column (Excel col E of the block) with that
block's header 'Puffing Regime:' value. Preserves cell styles; touches nothing
else. Idempotent."""
import sys, openpyxl
from copy import copy

PATH = "resources/templates/Standardized Test Template - December 2025.xlsx"
wb = openpyxl.load_workbook(PATH)   # keep_vba defaults False; this is .xlsx
relabelled = prefilled = 0

for ws in wb.worksheets:
    if "sop" in ws.title.lower():
        continue
    nblocks = max(1, ws.max_column // 12)
    for b in range(nblocks):
        base = b * 12               # 0-based block start
        # locate the per-row header row: the row whose block col A == "puffs"
        hdr_row = None
        for r in range(1, 13):
            v = ws.cell(row=r, column=base + 1).value
            if v and str(v).strip().lower() == "puffs":
                hdr_row = r; break
        if hdr_row is None:
            print(f"  SKIP {ws.title!r} block {b}: no 'puffs' header row")
            continue
        res_cell = ws.cell(row=hdr_row, column=base + 5)   # col E of the block
        if not (res_cell.value and "resist" in str(res_cell.value).lower()):
            # already relabelled or non-standard — skip quietly
            continue
        res_cell.value = "Puffing Regime"
        relabelled += 1
        # header regime value: row 2, block col 8 (1-based base+8); label at base+7
        regime = ws.cell(row=2, column=base + 8).value
        if regime:
            # pre-fill every data row that has a puff value in col A
            r = hdr_row + 1
            while True:
                puff = ws.cell(row=r, column=base + 1).value
                if puff is None or str(puff).strip() == "":
                    break
                cell = ws.cell(row=r, column=base + 5)
                if cell.value in (None, ""):
                    cell.value = regime
                    prefilled += 1
                r += 1

wb.save(PATH)
print(f"Relabelled {relabelled} header cell(s); pre-filled {prefilled} data cell(s).")
```
(Note: `res_cell.value = ...` keeps the existing cell *style* — openpyxl only replaces the value. If a verification shows the dark-blue header style was lost, copy it explicitly with `copy(res_cell._style)` before/after — but value-only assignment preserves style.)

- [ ] **Step 3: Run the relabel script and verify**
```bash
PYTHONIOENCODING=utf-8 PYTHONUTF8=1 python tools/relabel_template_regime.py
PYTHONIOENCODING=utf-8 PYTHONUTF8=1 python - <<'PY'
import openpyxl
wb = openpyxl.load_workbook("resources/templates/Standardized Test Template - December 2025.xlsx")
bad = 0
for ws in wb.worksheets:
    if "sop" in ws.title.lower(): continue
    for b in range(max(1, ws.max_column//12)):
        for r in range(1,13):
            if str(ws.cell(row=r, column=b*12+1).value).strip().lower()=="puffs":
                h = ws.cell(row=r, column=b*12+5).value
                if h != "Puffing Regime":
                    print("LEFTOVER", ws.title, b, h); bad += 1
print("remaining non-relabelled standard headers:", bad)
PY
```
Expected: relabelled 36 (minus any genuinely-non-standard block it skipped); `remaining... : 0`.
Then load the template in the app (or run `tst_excelreader` against it if a test references it) to confirm it still opens and column E reads as a regime.

- [ ] **Step 4: Add a new-template unit fixture** — `tests/generate_fixtures.py`

Add a generator (and call it from `__main__`) that writes column-E header `"Puffing Regime"` and per-row regime values including a mid-session change:
```python
def gen_format_e_regime():
    """New template (v2.2.1): per-row Puffing Regime column instead of Resistance."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Lifetime Test"
    write_format_e_headers(ws, 0, "Regime-1")
    # override col-E (index 5, 1-based) header to the new label
    ws.cell(row=4, column=5, value="Puffing Regime")
    write_data_rows(ws, 0, PUFFS, BEFORE_W, AFTER_W, DRAW_P)
    # per-row regimes: first 3 rows one regime, last 2 another (mid-session change)
    regimes = ["60mL/3s/30s", "60mL/3s/30s", "60mL/3s/30s", "200mL/9s/300s", "200mL/9s/300s"]
    for i, rg in enumerate(regimes):
        ws.cell(row=5 + i, column=5, value=rg)
    wb.save(os.path.join(DATA_DIR, "format_e_regime.xlsx"))
    print("  format_e_regime.xlsx")
```
Add `gen_format_e_regime()` to the `__main__` block. Run `python tests/generate_fixtures.py`.

- [ ] **Step 5: Add an ExcelReader→pipeline integration test** — append to `tests/tst_excelreader/tst_excelreader.cpp` (Read it first for its load helper):
```cpp
void newTemplate_perRowRegime_endToEnd() {
    DataProcessor dp;
    FileResult fr = dp.processFile(dataPath("format_e_regime.xlsx"));
    QVERIFY(!fr.sheets.isEmpty());
    const SheetResult& sh = fr.sheets[0];
    QVERIFY(sh.hasPerRowRegime);
    QCOMPARE(sh.samples[0].rows.first().puffingRegime, QString("60mL/3s/30s"));
    QCOMPARE(sh.samples[0].rows.last().puffingRegime,  QString("200mL/9s/300s"));
}
```
(Use whatever `dataPath()`/`DataProcessor` include the suite already uses; if `tst_excelreader` doesn't link `DataProcessor`, place this in `tst_sheetprocessors` or `tst_dataprocessor` instead, or add the needed sources to `tst_excelreader.pro`.)

- [ ] **Step 6: Run the full suite**
Run: `tests\run-tests.ps1`
Expected: all suites PASS, including the new fixture-backed test.

- [ ] **Step 7: Bump VERSION + clean rebuild** — `DataViewerEnterprise.pro` line 14:
```pro
VERSION = 2.2.1
```
Then (VERSION changes need a clean rebuild so `main.o` re-embeds it):
```bash
python tools/decrypt_via_copy.py --apply
```
```bat
cd build && mingw32-make clean && C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j8
```
Confirm `release\DataViewer.exe` (or the debug binary) reports 2.2.1 (Help/About or file properties).

- [ ] **Step 8: Commit**
```bash
git add "resources/templates/Standardized Test Template - December 2025.xlsx" \
        tools/relabel_template_regime.py tests/generate_fixtures.py \
        tests/tst_excelreader/tst_excelreader.cpp DataViewerEnterprise.pro
git commit -m "feat(template): relabel per-row column to Puffing Regime + prefill; v2.2.1"
```

---

## Final verification (whole feature)

- [ ] Run the full unit suite green: `tests\run-tests.ps1` (start `tests\start-test-postgres.ps1` first for DB suites).
- [ ] Build the release binary clean (Task 7 Step 7) with no `-Werror` warnings.
- [ ] Manual end-to-end on the **new** template:
  - Load it → TPM mode shows "Puffing Regime" column with pre-filled regimes; edit via dropdown.
  - Plot regime picker lists the unique regimes; filtering scopes all four plot types; bar chart recomputes.
  - Generate a Full report → a sheet with N regimes yields N content slides titled "{sheet} — {regime}".
  - Save → reopen (DB round-trip) → regime column + flag survive.
- [ ] Manual regression on an **old** template file: column 4 = Resistance, no regime picker, single slide per sheet — unchanged.
- [ ] Offline mode: close cleanly online (regenerates snapshot), go offline, reopen a new-template file from the snapshot → regime column present and read-only.
- [ ] Stop in-repo. Do **not** build the installer or touch Synology — that is the user's manual release step (`CLAUDE.md` release workflow).
- [ ] Update `tasks/lessons.md` if any UI/architectural assumption was corrected during implementation.

## Out of scope (per spec)
- The `.xlsm` "Automated Testing Template - DVE" (user-owned).
- Installer build, deployment self-test, Synology transfer.
- Per-row regime in the Properties dock or any new plot *types*.
