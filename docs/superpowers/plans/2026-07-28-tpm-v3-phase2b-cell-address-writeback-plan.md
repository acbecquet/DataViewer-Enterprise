# TPM v3 Phase 2b (Write Provenance + Address-Mapped Write-Back) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the hard-coded `sampleIndex*12 + col + 1` / prop-switch write-back math with addresses derived from recorded parse provenance, re-enable editing on inferred (13/8-col) layouts, and fix the latent Cart/Project header-write corruption - gated by a round-trip harness proving derived addresses byte-agree with the legacy math on standard sheets and point at the true source cells on real files.

**Architecture:** Compact per-sheet/per-sample **write provenance** (`startColumn`, `blockCols`, `dataStartRow`, `columnKeys`, `headerCells`) is recorded at parse time from the schema the v3 reader already resolved (Standard/Cart/Project/inferred), serialized through recovery JSON, and consumed by pure helper functions in a new `src/pipeline/CellAddress.{h,cpp}` that MainWindow's two live edit paths call with graceful fallback to the legacy math when provenance is absent (old recovery snapshots, DB-cache fallbacks).
Postgres/offline-snapshot persistence is deliberately OUT: DB-browser loads re-run `processFile` when the source file exists (fresh provenance), and when it does not exist there is no workbook to write to.
Spec anchors: design spec section 7 (write-back), registry doc (naming), explorer write-back map 2026-07-28 (in this plan's Reference section).

**Tech Stack:** C++17 / Qt 6.10 (qmake + MinGW), Qt Test, existing suites + new `tst_v3roundtrip`.

---

## Machine + repo rules (read first)

- This machine (S1134987) MIP-labels files written by trusted Python; the Write/Edit tools do NOT label. **Create all new source files with the Write tool** - never python file writes, never heredoc/echo.
- Ciphertext (`%TSD-Header-###%`) in a file: read via `git show HEAD:<path>`; run `python tools/decrypt_via_copy.py --apply` from repo root before any build.
- The repo is PUBLIC. `tests/corpus/` is gitignored; never commit real workbooks or results.txt artifacts.
- Branch `worktree-tpm-template-v3-research`. Commit per task; plain dashes; NO Co-Authored-By.
- Qt Test stdout is INVISIBLE - always `-o results.txt,txt` and read the file.
- Suite inner loop (from suite dir): `export PATH="/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH"`, `/c/Qt/6.10.1/mingw_64/bin/qmake.exe`, `mingw32-make`, `release/<suite>.exe -o results.txt,txt` (or debug/). Qt bin MUST be on PATH or the exe dies silently without writing results.
- Full-suite gate from repo root: `powershell -ExecutionPolicy Bypass -File tests/run-tests.ps1`.
- -Werror -Wall -Wextra.

## Reference: current write-back map (verified 2026-07-28)

- `onStoryCellEdited(int dataRow, int col, const QString& text)` at `MainWindow.cpp:3383`; math at `:3437-3439`: `excelRow = dataRow + 5; excelCol = m_currentSampleIndex * 12 + col + 1;` then `queueExcelWrite(file->filePath, sheet->sheetName, excelRow, excelCol, text)`. Column switch `:3416-3424` maps `DVE::Cols` RESISTANCE(4, per-row regime only when `hasPerRowRegime`)/SMELL(5)/CLOG(6)/NOTES(7).
- `onPropCellChanged(int row, int col)` at `MainWindow.cpp:2118`; TPM branch from `:2170`; `int off = m_currentSampleIndex * 12;` at `:2194`; field switch `:2200-2230` (row 1 sample_id -> (1, off+6), 3 tester -> (3, off+4), 5 media -> (2, off+2), 6 viscosity -> (3, off+2), 7 resistance -> (2, off+4), 8 voltage -> (3, off+6), 10 heating_technology -> (1, off+8), 12 puffing_regime -> (2, off+8), 13 initial_oil_mass -> (3, off+8)); dispatch `:2260-2261`.
- v2.10.2 guards: `onEditHeaders` `:2012-2018` (handler is UNWIRED - no caller; leave it), `onPropCellChanged` `:2182-2187`, `onStoryCellEdited` `:3405-3412`.
- Payload: `ExcelCellWrite{int row; int col; QString value;}` 1-based, batch is single-file/single-sheet (`queueExcelWrite` `:7116`).
- The prop switch matches ONLY the new/old standardized layout; Cart-format sheets store e.g. sample id at block-relative (2,2) and Project-format at assembled (1,7)+(1,9) - header edits on those sheets write to WRONG cells today (latent corruption, unguarded because they pass `standardFits`).
- Provenance predecessors: `ExcelReader::Sample::startColumn` (reader-internal, dropped), `TemplateSchema::blockCols` + `LegacyAdapter` block math (`startColumn = blockIndex * schema.blockCols`) - never surfaced onto `SheetResult`/`SampleResult`.
- Write-back reaches Excel-loaded, DB-loaded (re-parsed when source exists; DB cache fallback when not), and recovery-restored files; target is always `FileResult::filePath`.

## Non-goals

- NO `_dve_schema` manifest, NO NameFirst activation, NO SchemaResolver unification (Phase 2c).
- NO Postgres/offline-snapshot persistence of provenance (rationale in Architecture; recovery JSON only).
- NO row insert/delete write-back (none exists in production today).
- NO rewiring of the dead `onEditHeaders` handler (unwired; update only its guard comment if touched by search-replace, otherwise leave byte-identical).
- NO change to what values are written or to the queue/flush machinery - only WHERE cells land.
- Editing the ASSEMBLED sample id on Project-layout sheets stays rejected (its value spans two source cells); all other Project/Cart header fields become correctly editable.

---

### Task 1: Provenance fields + CellAddress helpers

**Files:**
- Modify: `src/pipeline/ReportData.h` (SampleResult + SheetResult fields)
- Create: `src/pipeline/CellAddress.h`, `src/pipeline/CellAddress.cpp`
- Modify: `DataViewerEnterprise.pro`
- Create: `tests/tst_v3roundtrip/tst_v3roundtrip.pro`, `tests/tst_v3roundtrip/tst_v3roundtrip.cpp`
- Modify: `tests/tests.pro` (add SUBDIRS entry)

- [ ] **Step 1: Add the model fields**

In `src/pipeline/ReportData.h`, add to `SampleResult` (next to the identity fields, keeping struct layout conventions):

```cpp
    // Write provenance (Phase 2b): 0-based physical column where this sample's
    // block starts in the source sheet. -1 = unknown (pre-2b recovery
    // snapshots, DB-cache fallbacks) - write-back then uses the legacy
    // sampleIndex*12 math via CellAddress fallback.
    int startColumn = -1;
```

and to `SheetResult` (next to `columnHeaders`):

```cpp
    // Write provenance (Phase 2b), recorded from the schema the reader
    // resolved. Empty/0 = unknown (see SampleResult::startColumn).
    int                  blockCols = 0;       // physical block width (12/13/8)
    int                  dataStartRow = 0;    // 1-based Excel row of data row 0 (5 on every known layout)
    QStringList          columnKeys;          // canonical metric key per physical column slot
    QMap<QString, QPoint> headerCells;        // header key -> (x=1-based block-relative col of the VALUE cell, y=1-based Excel row)
```

(`#include <QPoint>` and `<QMap>` as needed - QMap is already used by SampleResult::extra.)
Convention fixed here once: `QPoint::x()` = block-relative 1-based column of the value cell, `QPoint::y()` = 1-based Excel row. Document exactly that in the comment.

- [ ] **Step 2: Write the failing helper tests**

`tests/tst_v3roundtrip/tst_v3roundtrip.cpp` (new suite; model the `.pro` on `tests/tst_v3model/tst_v3model.pro`, sources: `CellAddress.cpp` only for now, INCLUDEPATH `../../src` and `../../src/pipeline`):

```cpp
#include <QtTest>
#include "CellAddress.h"
#include "ReportData.h"

using namespace DVE;

class TestV3RoundTrip : public QObject {
    Q_OBJECT
private slots:
    void dataCellMatchesLegacyMathOnStandardSheets();
    void dataCellUsesColumnKeysOnInferredSheets();
    void dataCellRejectsMissingColumn();
    void headerCellMatchesLegacyPropSwitch();
    void headerCellRejectsUnknownKey();
    void fallbackWhenProvenanceAbsent();
};

namespace {
SheetResult standardSheet()
{
    SheetResult sh;
    sh.blockCols = 12;
    sh.dataStartRow = 5;
    sh.columnKeys = {"puffs","before_weight","after_weight","draw_pressure","resistance",
                     "smell","clog","notes","tpm","tpm_power_density","variation_tpm","oil_consumed"};
    // Standard layout header value cells (block-relative col, Excel row).
    sh.headerCells.insert("sample_id",          QPoint(6, 1));
    sh.headerCells.insert("heating_technology", QPoint(8, 1));
    sh.headerCells.insert("media",              QPoint(2, 2));
    sh.headerCells.insert("resistance",         QPoint(4, 2));
    sh.headerCells.insert("puffing_regime",     QPoint(8, 2));
    sh.headerCells.insert("viscosity",          QPoint(2, 3));
    sh.headerCells.insert("tester",             QPoint(4, 3));
    sh.headerCells.insert("voltage",            QPoint(6, 3));
    sh.headerCells.insert("initial_oil_mass",   QPoint(8, 3));
    return sh;
}
SampleResult sampleAt(int startCol) { SampleResult s; s.startColumn = startCol; return s; }
} // namespace

void TestV3RoundTrip::dataCellMatchesLegacyMathOnStandardSheets()
{
    const SheetResult sh = standardSheet();
    // Legacy: excelRow = dataRow + 5; excelCol = sampleIndex*12 + col + 1.
    for (int sampleIndex : {0, 1, 3}) {
        const SampleResult s = sampleAt(sampleIndex * 12);
        for (int col : {4, 5, 6, 7}) {              // regime/smell/clog/notes slots
            for (int dataRow : {0, 2, 9}) {
                const CellAddress a = CellAddress::dataCell(sh, s, sh.columnKeys[col], dataRow);
                QVERIFY(a.valid);
                QCOMPARE(a.row, dataRow + 5);
                QCOMPARE(a.col, sampleIndex * 12 + col + 1);
            }
        }
    }
}

void TestV3RoundTrip::dataCellUsesColumnKeysOnInferredSheets()
{
    SheetResult sh;
    sh.blockCols = 13;                               // S26 Cart-era shape
    sh.dataStartRow = 5;
    sh.columnKeys = {"puffs","before_weight","after_weight","pv1","pv2","pv3","pv4","pv5",
                     "resistance","smell","clog","notes","tpm"};
    const SampleResult s = sampleAt(13);             // block 2
    const CellAddress a = CellAddress::dataCell(sh, s, QStringLiteral("smell"), 3);
    QVERIFY(a.valid);
    QCOMPARE(a.row, 8);                              // 5 + 3
    QCOMPARE(a.col, 13 + 9 + 1);                     // startColumn + slot(smell)=9 + 1-based
}

void TestV3RoundTrip::dataCellRejectsMissingColumn()
{
    SheetResult sh;
    sh.blockCols = 8;                                // UserSim shape: no smell column
    sh.dataStartRow = 5;
    sh.columnKeys = {"chronology","puffs","before_weight","after_weight",
                     "draw_pressure","failure","notes","tpm"};
    const CellAddress a = CellAddress::dataCell(sh, sampleAt(0), QStringLiteral("smell"), 0);
    QVERIFY(!a.valid);
}

void TestV3RoundTrip::headerCellMatchesLegacyPropSwitch()
{
    const SheetResult sh = standardSheet();
    // Legacy prop switch, off = sampleIndex*12 (0-based): the exact table from
    // MainWindow.cpp:2200-2230.
    const struct { const char* key; int row; int colOff; } legacy[] = {
        {"sample_id", 1, 6}, {"tester", 3, 4}, {"media", 2, 2}, {"viscosity", 3, 2},
        {"resistance", 2, 4}, {"voltage", 3, 6}, {"heating_technology", 1, 8},
        {"puffing_regime", 2, 8}, {"initial_oil_mass", 3, 8},
    };
    for (int sampleIndex : {0, 2}) {
        const SampleResult s = sampleAt(sampleIndex * 12);
        for (const auto& e : legacy) {
            const CellAddress a = CellAddress::headerCell(sh, s, QString::fromUtf8(e.key));
            QVERIFY2(a.valid, e.key);
            QCOMPARE(a.row, e.row);
            QCOMPARE(a.col, sampleIndex * 12 + e.colOff);
        }
    }
}

void TestV3RoundTrip::headerCellRejectsUnknownKey()
{
    const SheetResult sh = standardSheet();          // no project_name in the map
    QVERIFY(!CellAddress::headerCell(sh, sampleAt(0), QStringLiteral("project_name")).valid);
    QVERIFY(!CellAddress::headerCell(sh, sampleAt(-1), QStringLiteral("media")).valid); // no provenance
}

void TestV3RoundTrip::fallbackWhenProvenanceAbsent()
{
    SheetResult sh;                                  // defaults: no provenance
    QVERIFY(!CellAddress::hasProvenance(sh, sampleAt(-1)));
    QVERIFY(CellAddress::hasProvenance(standardSheet(), sampleAt(0)));
    QVERIFY(!CellAddress::hasProvenance(standardSheet(), sampleAt(-1)));
}

QTEST_MAIN(TestV3RoundTrip)
#include "tst_v3roundtrip.moc"
```

- [ ] **Step 3: Run to verify red**

Add the suite to `tests/tests.pro` SUBDIRS. Inner loop in `tests/tst_v3roundtrip`. Expected: compile failure on `CellAddress.h`.

- [ ] **Step 4: Implement**

`src/pipeline/CellAddress.h`:

```cpp
#pragma once
#include "ReportData.h"

namespace DVE {

// A 1-based Excel cell target for write-back. invalid => do not write (the
// field/column does not exist in this sheet's layout, or provenance is absent).
struct CellAddress {
    bool valid = false;
    int  row = 0;    // 1-based Excel row
    int  col = 0;    // 1-based Excel column

    // True when both sheet and sample carry Phase-2b write provenance.
    static bool hasProvenance(const SheetResult& sheet, const SampleResult& sample);

    // Address of data row `dataRow` (0-based) of the metric `columnKey`.
    static CellAddress dataCell(const SheetResult& sheet, const SampleResult& sample,
                                const QString& columnKey, int dataRow);

    // Address of the header-band VALUE cell for `headerKey`.
    static CellAddress headerCell(const SheetResult& sheet, const SampleResult& sample,
                                  const QString& headerKey);
};

} // namespace DVE
```

`src/pipeline/CellAddress.cpp`:

```cpp
#include "CellAddress.h"

namespace DVE {

bool CellAddress::hasProvenance(const SheetResult& sheet, const SampleResult& sample)
{
    return sheet.blockCols > 0 && sheet.dataStartRow > 0
        && !sheet.columnKeys.isEmpty() && sample.startColumn >= 0;
}

CellAddress CellAddress::dataCell(const SheetResult& sheet, const SampleResult& sample,
                                  const QString& columnKey, int dataRow)
{
    CellAddress a;
    if (!hasProvenance(sheet, sample) || dataRow < 0)
        return a;
    const int slot = sheet.columnKeys.indexOf(columnKey);
    if (slot < 0)
        return a;                          // column not present in this layout
    a.valid = true;
    a.row = sheet.dataStartRow + dataRow;
    a.col = sample.startColumn + slot + 1; // 0-based origin + slot -> 1-based
    return a;
}

CellAddress CellAddress::headerCell(const SheetResult& sheet, const SampleResult& sample,
                                    const QString& headerKey)
{
    CellAddress a;
    if (!hasProvenance(sheet, sample))
        return a;
    const auto it = sheet.headerCells.constFind(headerKey);
    if (it == sheet.headerCells.constEnd())
        return a;                          // field not present in this layout
    a.valid = true;
    a.row = it.value().y();                // 1-based Excel row
    a.col = sample.startColumn + it.value().x(); // block-relative 1-based col
    return a;
}

} // namespace DVE
```

Register `CellAddress.cpp/.h` in `DataViewerEnterprise.pro`.

- [ ] **Step 5: Run to verify green, then commit**

```bash
git add src/pipeline/ReportData.h src/pipeline/CellAddress.h src/pipeline/CellAddress.cpp DataViewerEnterprise.pro tests/tests.pro tests/tst_v3roundtrip
git commit -m "feat(v3): write provenance fields + CellAddress helpers - derived addresses match legacy math on standard layouts"
```

---

### Task 2: Populate provenance on both parse paths

**Files:**
- Modify: `src/pipeline/DataProcessor.cpp` (standard schema-driven path)
- Modify: `src/pipeline/SheetProcessors.cpp` (copy `SampleData::startColumn` into `SampleResult`)
- Modify: `src/model/LegacyAdapter.cpp` (`lowerInferredSheet` sheet+sample provenance)
- Test: `tests/tst_v3inference/tst_v3inference.cpp`, `tests/tst_dataprocessor/tst_dataprocessor.cpp` (or the existing suite that parses standard fixtures end-to-end - read both first and put the standard-path test where fixture parsing already happens)

- [ ] **Step 1: Read the seams**

Read `DataProcessor::processSheet` (the production schema-driven path - it resolves the schema variant incl. Cart/Project landmark sniff and lowers each `model::Sample` via `LegacyAdapter::lowerSample` into `ExcelReader::SampleData`, then `SheetProcessors` builds the `SheetResult`), `SheetProcessors::buildSampleResult` (or equivalent - find where `SampleData` becomes `SampleResult`), and `LegacyAdapter::lowerInferredSheet`. Identify:
- where the resolved `TemplateSchema` is in scope alongside the produced `SheetResult` (standard path), and
- where `SampleData::startColumn` is dropped today.

- [ ] **Step 2: Write the failing tests**

Standard path (place in the suite that already parses fixtures through `DataProcessor::processFile` - `tst_dataprocessor` if it does, else the shadow/inference suite's E2E pattern; read first, reuse its fixture-loading helper):

```cpp
void TestDataProcessor::writeProvenanceRecordedOnStandardParse()
{
    // Any standard fixture with >= 2 samples; new_format.xlsx / the multi-sample
    // fixture the suite already uses.
    FileResult f = processFixture("new_format.xlsx");   // suite's existing helper
    const SheetResult& sh = f.sheets[0];
    QCOMPARE(sh.blockCols, 12);
    QCOMPARE(sh.dataStartRow, 5);
    QCOMPARE(sh.columnKeys.size(), 12);
    QCOMPARE(sh.columnKeys[0], QStringLiteral("puffs"));
    QCOMPARE(sh.columnKeys[5], QStringLiteral("smell"));
    QVERIFY(sh.headerCells.contains(QStringLiteral("media")));
    QCOMPARE(sh.headerCells.value(QStringLiteral("media")), QPoint(2, 2));
    for (int i = 0; i < sh.samples.size(); ++i)
        QCOMPARE(sh.samples[i].startColumn, i * 12);
}

void TestDataProcessor::writeProvenanceRecordsProjectLayoutHeaders()
{
    // Project-format fixture (the one the shadow suite parses byte-identically).
    FileResult f = processFixture("project_format.xlsx");
    const SheetResult& sh = f.sheets[0];
    // Project layout: media (2,2); tester on ROW 1 col 5 value cell; NO sample_id
    // entry (assembled from project+sample - not single-cell addressable).
    QCOMPARE(sh.headerCells.value(QStringLiteral("media")), QPoint(2, 2));
    QCOMPARE(sh.headerCells.value(QStringLiteral("tester")), QPoint(5, 1));
    QVERIFY(!sh.headerCells.contains(QStringLiteral("sample_id")));
    QVERIFY(sh.headerCells.contains(QStringLiteral("project_name")));
}
```

(Adjust fixture names to what actually exists in `tests/data/` - `git ls-files tests/data` and the suite's existing tests name them; the Project fixture MUST be the one whose header band is the Project variant. The expected QPoints come from `StandardSchema.cpp`'s layout blocks: hf(key, row, col) means value cell = (col, row) in our (x=col, y=row) convention - VERIFY against StandardSchema.cpp before hardcoding, e.g. Project layout has `hf("tester", ..., 1, 5)` = row 1, col 5.)

Inferred path (append to `tst_v3inference.cpp`, reusing `makeS26Grid`):

```cpp
void TestV3Inference::writeProvenanceRecordedOnInferredParse()
{
    QVector<QVector<QVariant>> cells = makeS26Grid();
    const TemplateSchema schema = SchemaInference::inferSchema(cells, QStringLiteral("t"));
    const Sheet sheet = SchemaDrivenReader::parseSheet(cells, QStringLiteral("t"), schema,
                                                       false, ColumnResolution::NameFirst);
    const SheetResult sr = LegacyAdapter::lowerInferredSheet(sheet, QStringLiteral("t"),
                                                             QStringLiteral("new"));
    QCOMPARE(sr.blockCols, 13);
    QCOMPARE(sr.dataStartRow, 5);
    QCOMPARE(sr.columnKeys.size(), 13);
    QVERIFY(sr.columnKeys.contains(QStringLiteral("smell")));
    QVERIFY(sr.headerCells.contains(QStringLiteral("media")));
    for (int i = 0; i < sr.samples.size(); ++i)
        QCOMPARE(sr.samples[i].startColumn, i * 13);
}
```

(Mirror the suite's existing call style; if `makeS26Grid` yields 1 sample the loop still verifies `startColumn == 0`.)

- [ ] **Step 3: Run to verify red, then implement**

Standard path - in `DataProcessor::processSheet`, after the `SheetResult` is built and the resolved schema + per-sample block origins are known, set:

```cpp
    // Write provenance (Phase 2b): record where this sheet's cells came from
    // so write-back derives addresses instead of assuming the 12-wide
    // standardized layout (which silently corrupts Cart/Project header cells).
    sheetResult.blockCols    = schema.blockCols;
    sheetResult.dataStartRow = schema.dataStartRow;
    for (const model::MetricDef& c : schema.columns)
        sheetResult.columnKeys.append(c.key);
    for (const model::HeaderFieldDef& hf : schema.headerFields)
        sheetResult.headerCells.insert(hf.key, QPoint(hf.col, hf.row));
```

and per sample `sheetResult.samples[i].startColumn = i * schema.blockCols;` UNLESS `SampleData::startColumn` is already flowing (preferred: copy it in `SheetProcessors` where `SampleData` becomes `SampleResult` - one line, `result.startColumn = data.startColumn;` - because the reader's own origin is ground truth). Choose ONE source of truth: copy from `SampleData` in SheetProcessors AND set the sheet-level fields in DataProcessor. Note: the raw-table branch (`isRawTable`) and sensory paths must NOT get provenance (leave defaults).
IMPORTANT byte-safety: these fields are new plain members; nothing else reads them yet. `tst_v3shadow` compares serialized JSON of FileResults - Task 3 adds them to JSON, so at THIS task's point the shadow suite still serializes the old shape and must stay 21/0/3. Run it to confirm.

Inferred path - in `LegacyAdapter::lowerInferredSheet`, same recording from the inferred `sheet.schema`, and per-sample `startColumn = blockIndex * schema.blockCols` (the block loop already computes the origin - reuse its variable).

Project-layout nuance requires NO code: the Project variant's headerFields simply do not include `sample_id` (they carry `project_name` + `sample_suffix`), so the `headerCells` map naturally lacks it and `CellAddress::headerCell` rejects sample-id edits on those sheets.

- [ ] **Step 4: Gates**

tst_dataprocessor (or host suite) green with the 2 new tests; tst_v3inference green (+1); **tst_v3shadow 21/0/3 unchanged**; tst_v3model green.

- [ ] **Step 5: Commit**

```bash
git add src/pipeline/DataProcessor.cpp src/pipeline/SheetProcessors.cpp src/model/LegacyAdapter.cpp tests/
git commit -m "feat(v3): record write provenance on standard + inferred parse paths"
```

(Stage tests explicitly if untracked artifacts are present - never blanket-add results.txt.)

---

### Task 3: Recovery-JSON round-trip for provenance

**Files:**
- Modify: `src/pipeline/ReportDataJson.cpp`
- Test: `tests/tst_reportdatajson/tst_reportdatajson.cpp`

- [ ] **Step 1: Write the failing tests** (follow the suite's `makeFile()` pattern)

```cpp
void TstReportDataJson::writeProvenanceRoundTrips()
{
    FileResult f = makeFile();
    SheetResult& sh = f.sheets[0];
    sh.blockCols = 13;
    sh.dataStartRow = 5;
    sh.columnKeys = QStringList{"puffs", "before_weight"};
    sh.headerCells.insert(QStringLiteral("media"), QPoint(2, 2));
    sh.samples[0].startColumn = 13;

    const FileResult back = fileResultFromJson(fileResultToJson(f));
    QCOMPARE(back.sheets[0].blockCols, 13);
    QCOMPARE(back.sheets[0].dataStartRow, 5);
    QCOMPARE(back.sheets[0].columnKeys, sh.columnKeys);
    QCOMPARE(back.sheets[0].headerCells.value(QStringLiteral("media")), QPoint(2, 2));
    QCOMPARE(back.sheets[0].samples[0].startColumn, 13);
}

void TstReportDataJson::preProvenanceSnapshotsDecodeToNoProvenance()
{
    // A JSON produced BEFORE 2b (no provenance keys) must decode to the
    // "unknown" defaults so write-back falls back to legacy behavior.
    FileResult f = makeFile();                      // defaults: no provenance
    QJsonObject j = fileResultToJson(f);
    const FileResult back = fileResultFromJson(j);
    QCOMPARE(back.sheets[0].blockCols, 0);
    QCOMPARE(back.sheets[0].samples[0].startColumn, -1);
}
```

- [ ] **Step 2: Run red, then implement**

In the sheet serializer: emit `block_cols`, `data_start_row`, `column_keys` (string array), `header_cells` (object key -> `{"c": col, "r": row}`) ONLY when provenance is present (`blockCols > 0`) - absent keys keep pre-2b snapshots byte-stable, mirroring the `extra`-map convention. In the sample serializer: emit `start_column` only when `>= 0`. Readers default to `0` / `-1` when keys are absent.
Update the JSON parity tripwire comment/test if the suite has one (it counted serialized keys in 2a-era work - read the suite's tripwire test and extend its expected key lists).
IMPORTANT: `tst_v3shadow` serializes parse output to JSON for the byte-identity diff; after Task 2 the standard path POPULATES provenance, so from this task on the shadow JSON grows the new keys on BOTH sides of its comparison (production vs legacy referee)... but the legacy referee path (`processFileLegacy`) does NOT populate provenance. THIS WOULD BREAK THE SHADOW GATE. Resolution (decided here, not optional): the referee's output must be compared on the legacy-visible domain - the shadow harness's diff must ignore the provenance keys. Do this by having the shadow suite strip the new keys before diffing (`JsonDiff` call site in `tst_v3shadow.cpp` - add a small `stripProvenance(QJsonObject&)` helper there removing `block_cols`/`data_start_row`/`column_keys`/`header_cells`/`start_column` recursively). The gate then still proves every legacy-era byte identical. Add a comment in the harness explaining why.
Run `tst_v3shadow` (21/0/3) after the strip helper lands.

- [ ] **Step 3: Gates + commit**

tst_reportdatajson green (+2); tst_v3shadow 21/0/3; tst_recoverymanager (if it exists - `ls tests/ | grep -i recover`) green.

```bash
git add src/pipeline/ReportDataJson.cpp tests/tst_reportdatajson tests/tst_v3shadow
git commit -m "feat(v3): write provenance rides recovery JSON - absent keys keep pre-2b snapshots stable; shadow diff masks provenance"
```

---

### Task 4: Round-trip harness over real grids

**Files:**
- Modify: `tests/tst_v3roundtrip/tst_v3roundtrip.cpp` + its `.pro` (now links the parse chain: model sources + pipeline sources + CorpusUtils, mirroring `tst_v3inference.pro`'s source list + python-runner config if it uses ExcelReader; PREFER building grids via the fixture .xlsx files through the same loading path tst_v3inference uses - read that suite's fixture-loading mechanics first)

- [ ] **Step 1: Write the harness tests**

Two data-driven checks over every fixture in `tests/data/` (and corpus files when `DVE_TEST_CORPUS_DIR` is set - reuse `DVE::testutil::corpusFiles()`):

```cpp
// 1. Mapped-domain identity: for every sheet with provenance, every
//    qualitative/measured column key and every data row, the ORIGINAL grid
//    cell at the derived address must equal the value the model stores (text
//    columns compared as trimmed strings; numeric as doubles with fuzzy
//    compare; empty cells skip). Header fields likewise via headerCell().
// 2. Legacy-math equivalence: on sheets where templateVersion is "new"/"old"
//    AND blockCols == 12 AND the header layout is the standardized one
//    (headerCells.value("media") == QPoint(2,2) && headerCells.contains("sample_id")),
//    derived addresses for the prop-switch keys and the story columns must
//    equal the legacy formulas exactly (same table as Task 1's tests, but
//    executed against REAL parsed provenance, not hand-built structs).
```

Implementation notes for the executor:
- The original grid comes from the same cells the parser consumed: `ExcelReader::currentSheetCells()` per sheet - mirror how `tst_v3shadow`/`tst_v3inference` obtain grids for fixtures (python-subprocess reader; skip cleanly with QSKIP when python is unavailable, same guard those suites use).
- Compare data cells only for keys in `columnKeys` whose role is not derived: use `MetricRegistry::metric(key)` and skip `Role::Derived` (derived cells hold formulas the reader cached differently). Also skip `puffs`/`before_weight` repair-rule cells: the adapter fills forward/extrapolates zeros, so compare those two keys only when the ORIGINAL cell is non-empty.
- For header fields compare only when the original cell is non-empty; numeric header fields (`viscosity` etc.) compare `toDouble` with `qFuzzyCompare`-style tolerance; text trimmed-equal. Skip `power` (formula cell).
- Sensory sheets / raw tables / sheets without provenance: skip silently (count them, `qDebug` the tally so the test output shows coverage).
- The corpus pass MUST pass on T58G (Project layout - proves the Cart/Project header fix: `tester` resolves to (5,1)-style Project cells and matches the grid) and on S26 (inferred 13-wide: `smell` at slot 9 matches).

- [ ] **Step 2: Run over fixtures (red until wiring bugs are fixed), then over corpus**

Fixtures run: all green. Then `DVE_TEST_CORPUS_DIR=<worktree>/tests/corpus` run: green (this is the Phase 2b gate). Report exact totals.

- [ ] **Step 3: Commit**

```bash
git add tests/tst_v3roundtrip
git commit -m "test(v3): round-trip harness - derived addresses identity-checked against source grids + legacy-math equivalence"
```

---

### Task 5: MainWindow rewire + guard lift

**Files:**
- Modify: `src/MainWindow.cpp` (`onPropCellChanged` ~2170-2261, `onStoryCellEdited` ~3405-3439)

- [ ] **Step 1: Read both functions fully**, then implement:

`onPropCellChanged`: replace the inferred guard + switch's excelRow/Col assignments:
- The switch keeps its model mutations exactly, but instead of assigning `excelRow`/`excelCol` it assigns a `headerKey` string per row: 1 -> "sample_id", 3 -> "tester", 5 -> "media", 6 -> "viscosity", 7 -> "resistance", 8 -> "voltage", 10 -> "heating_technology", 12 -> "puffing_regime", 13 -> "initial_oil_mass" (same `affectsPower` flags).
- After the switch: `const CellAddress addr = CellAddress::hasProvenance(*sheet, s) ? CellAddress::headerCell(*sheet, s, headerKey) : legacyPropAddress(headerKey, m_currentSampleIndex);` where `legacyPropAddress` is a small file-local helper reproducing the OLD table verbatim (kept for provenance-less files: old recovery snapshots, DB-cache fallbacks with standard layout).
- Provenance-less INFERRED sheets keep the guard: replace the current guard condition `sheet->fromInferredSchema` with `sheet->fromInferredSchema && !CellAddress::hasProvenance(*sheet, s)` (same revert + message).
- If `addr.valid` is false WITH provenance (field not in this layout - e.g. sample_id on Project sheets): revert the visible cell via `updateProperties(s)` and `updateStatusBar(tr("This field is not a single cell in this sheet's layout - edit it in the source file."))`, and DO NOT queue a write, but keep the in-memory model mutation reverted too (return before mutating: restructure so the mutation happens only when the write target resolves OR provenance is absent-and-legacy-fallback applies). CAREFUL: today the model mutates before queueing; preserve that order for valid targets to keep behavior identical.
- Valid: `queueExcelWrite(file->filePath, sheet->sheetName, addr.row, addr.col, text);` as today.

`onStoryCellEdited`: same pattern:
- Column switch maps col -> key: 4 -> (`hasPerRowRegime` ? "puffing_regime" : return as today), 5 -> "smell", 6 -> "clog", 7 -> "notes" (keep the model mutations).
- Guard becomes `sheet->fromInferredSchema && !CellAddress::hasProvenance(*sheet, *sample)`.
- Address: provenance ? `CellAddress::dataCell(*sheet, *sample, key, dataRow)` : legacy `{dataRow + 5, m_currentSampleIndex * 12 + col + 1}`.
- Invalid-with-provenance (column absent, e.g. smell on UserSim 8-col): revert via the same `setSample` re-population used by the guard + status message `tr("This column does not exist in this sheet's layout.")`; no write.
- Valid: queue as today.

Also update the two guard messages to drop "until v3 write-back" (it has arrived): e.g. `tr("This file predates layout tracking - reopen it from the source .xlsx to enable editing.")`.

- [ ] **Step 2: Build the app** (-Werror) and run the FULL suite. The two touched functions have no direct unit suite; the round-trip harness (Task 4) plus the CellAddress unit tests are the coverage, and behavior on provenance-less files is byte-preserved by the legacy fallbacks.

- [ ] **Step 3: Commit**

```bash
git add src/MainWindow.cpp
git commit -m "feat(v3): write-back targets CellAddress-derived cells - inferred layouts editable, Cart/Project header corruption fixed, legacy fallback for provenance-less files"
```

---

### Task 6: Gates, version 2.10.3, installer, docs

- [ ] **Step 1:** Full suite green; corpus round-trip green; tst_v3shadow 21/0/3 + corpus 25/0/5; -Werror build clean.
- [ ] **Step 2:** Bump `VERSION = 2.10.3` in `DataViewerEnterprise.pro`. Clean rebuild REQUIRED (qmake does not detect VERSION changes): in-tree ROOT release build (`build_installer.bat` packages ROOT `release\DataViewer.exe`, NOT `build\release\`): from repo root `python tools/decrypt_via_copy.py --apply`, then qmake CONFIG+=release in-tree, `mingw32-make clean`, `mingw32-make -j8`. Then `MSYS_NO_PATHCONV=1 cmd /c '.\build_installer.bat' < /dev/null` (needs release\python_bundle.zip - if absent, run `tools\prepare_python_embed.bat` first; check `ls release/python_bundle.zip`). Verify `dist\DataViewer-setup.exe` ProductVersion 2.10.3.
- [ ] **Step 3:** Write `release_overview/release_overview_v_2_10_3.txt` (internal build; 2a+2b combined: registry vocabulary, canonical titles on inferred sheets, editing re-enabled on old-format files, Cart/Project header-write fix, per-puff pressure preservation) with an owner smoke checklist (edit cells on S26 + UserSim sheets and verify the right source cells changed; edit a T58G header and verify it lands in the correct Project-layout cell; standard files unchanged).
- [ ] **Step 4:** Registry doc Phase log line, sprint tracker, commit docs. Copy installer + overview to the main repo dist\/release_overview\ (overwriting the v2.10.2 copies).

---

## Self-review (done at authoring time)

- Spec coverage: design spec section 7 write-back (Tasks 1/2/5), round-trip gate (Task 4), inferred-edit re-enablement (Task 5 guard change), recovery persistence (Task 3). Explicitly deferred: manifest/NameFirst (2c), DB persistence (architecture note), row insert/delete (does not exist today).
- Placeholder scan: Task 2 Step 3 and Task 5 Step 1 are precise-seam instructions with the decision logic fully specified (the 2a pattern that worked); all new-file code is complete. Task 4's checks are specified as behavior contracts with explicit skip rules rather than literal code because the grid-loading mechanics must mirror the host suite - the executor is told exactly which suite to mirror.
- Type consistency: `CellAddress{valid,row,col}` + `hasProvenance`/`dataCell`/`headerCell` used identically in Tasks 1/4/5; provenance field names (`blockCols`,`dataStartRow`,`columnKeys`,`headerCells`,`startColumn`) and JSON keys (`block_cols`,`data_start_row`,`column_keys`,`header_cells`,`start_column`) consistent across Tasks 2/3.
- Shadow-gate interaction analyzed and resolved in Task 3 (provenance masking in the harness diff) - the one place 2b touches the byte-identity contract.
