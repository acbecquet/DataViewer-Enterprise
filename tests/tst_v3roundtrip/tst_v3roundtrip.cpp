#include <QtTest>
#include "CellAddress.h"
#include "ReportData.h"

using namespace DVE;

// Phase 2b round-trip suite. Task 1 scope: pure unit tests over the
// CellAddress helpers with hand-built provenance - derived addresses must
// reproduce the legacy write-back math exactly on standard layouts, use
// columnKeys slots on inferred layouts, and reject fields/columns a layout
// does not carry. Task 4 extends this suite into the real-grid harness.
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
    // Standard layout header value cells (block-relative col, Excel row) -
    // verified against StandardSchema.cpp's standardized hf(...) table:
    // hf(key, display, type, row, col) -> QPoint(x=col, y=row).
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
