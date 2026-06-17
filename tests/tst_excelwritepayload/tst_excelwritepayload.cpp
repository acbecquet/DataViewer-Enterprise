// tst_excelwritepayload — SP3-T4 (R6) payload-equivalence guard.
//
// The Excel write-back subprocess was moved off the UI thread. The off-thread
// worker and the synchronous close-path fallback both drive openpyxl via the
// shared builders in ExcelWritePayload. This test pins that those builders emit
// the EXACT python source + argv the original *synchronous* MainWindow code
// produced before the move, so the threading relocation provably cannot change
// what is written to the user's workbook (the keystone lesson: test the real
// payload, not a parallel reimplementation).
//
// The reference literals below are copied verbatim from the pre-SP3-T4
// MainWindow::writeCellsToExcel (kWriteCells) and ::deleteRowFromExcel
// (kDeleteRow), and the reference argv mirrors the original inline construction
// (path, sheet, then row/col/value triplets). If the shared builder ever drifts
// from these references, this test goes RED — which is exactly the guard we want.

#include <QtTest/QtTest>
#include <QString>
#include <QStringList>
#include <QVector>

#include "utils/ExcelWritePayload.h"

using DVE::ExcelCellWrite;

namespace {

// Verbatim copy of the original synchronous kWriteCells literal (pre-SP3-T4).
const char* refWriteCellsScript()
{
    static const char* k = R"PY(
import os
import sys
from openpyxl import load_workbook
path, sheet = sys.argv[1], sys.argv[2]
wb = load_workbook(path)
ws = wb[sheet]
args = sys.argv[3:]
i = 0
while i + 2 < len(args):
    r, c, val = int(args[i]), int(args[i+1]), args[i+2]
    try:
        ws.cell(row=r, column=c).value = float(val) if val.strip() else None
    except ValueError:
        ws.cell(row=r, column=c).value = val if val.strip() else None
    i += 3
tmp = path + ".dve_tmp"
wb.save(tmp)
os.replace(tmp, path)
print("OK")
)PY";
    return k;
}

// Verbatim copy of the original synchronous kDeleteRow literal (pre-SP3-T4).
const char* refDeleteRowScript()
{
    static const char* k = R"PY(
import os
import sys
from openpyxl import load_workbook
path, sheet, row_s = sys.argv[1], sys.argv[2], sys.argv[3]
wb = load_workbook(path)
ws = wb[sheet]
ws.delete_rows(int(row_s), 1)
tmp = path + ".dve_tmp"
wb.save(tmp)
os.replace(tmp, path)
print("OK")
)PY";
    return k;
}

// Verbatim copy of the original synchronous argv construction for a cell batch.
QStringList refWriteCellsArgs(const QString& filePath,
                              const QString& sheetName,
                              const QVector<ExcelCellWrite>& cells)
{
    QStringList args = { filePath, sheetName };
    for (const ExcelCellWrite& cw : cells) {
        args << QString::number(cw.row) << QString::number(cw.col) << cw.value;
    }
    return args;
}

} // namespace

class TstExcelWritePayload : public QObject
{
    Q_OBJECT
private slots:
    void writeCellsScript_isByteIdentical();
    void deleteRowScript_isByteIdentical();
    void writeCellsArgs_emptyBatch();
    void writeCellsArgs_singleCell();
    void writeCellsArgs_multiCellWithNumericAndTextAndEmpty();
    void deleteRowArgs_matchesReference();
    // SP3-T4 (R6 fix): the failed-flush re-merge helper shared by
    // onExcelFlushFinished + finishExcelWritesBlocking.
    void mergePending_emptyPending_keepsInFlight();
    void mergePending_emptyInFlight_keepsPending();
    void mergePending_disjointCells_appendsNewerAfterOlder();
    void mergePending_sameCell_newerOverlaysInPlace();
    void mergePending_doesNotMutateInputs();
};

void TstExcelWritePayload::writeCellsScript_isByteIdentical()
{
    // QByteArray compare so a stray whitespace / newline difference fails loudly.
    QCOMPARE(QByteArray(DVE::excelWriteCellsScript()),
             QByteArray(refWriteCellsScript()));
}

void TstExcelWritePayload::deleteRowScript_isByteIdentical()
{
    QCOMPARE(QByteArray(DVE::excelDeleteRowScript()),
             QByteArray(refDeleteRowScript()));
}

void TstExcelWritePayload::writeCellsArgs_emptyBatch()
{
    const QString f = "C:/data/wb.xlsx";
    const QString s = "Sheet1";
    QVector<ExcelCellWrite> cells;   // empty
    QCOMPARE(DVE::buildWriteCellsArgs(f, s, cells),
             refWriteCellsArgs(f, s, cells));
    // Empty batch still carries path + sheet.
    QCOMPARE(DVE::buildWriteCellsArgs(f, s, cells), (QStringList{ f, s }));
}

void TstExcelWritePayload::writeCellsArgs_singleCell()
{
    const QString f = "C:/data/My Workbook.xlsx";
    const QString s = "TPM Data";
    QVector<ExcelCellWrite> cells = { { 5, 2, QStringLiteral("12.345") } };
    QCOMPARE(DVE::buildWriteCellsArgs(f, s, cells),
             refWriteCellsArgs(f, s, cells));
    QCOMPARE(DVE::buildWriteCellsArgs(f, s, cells),
             (QStringList{ f, s, "5", "2", "12.345" }));
}

void TstExcelWritePayload::writeCellsArgs_multiCellWithNumericAndTextAndEmpty()
{
    const QString f = "C:/data/wb.xlsx";
    const QString s = "Sheet With Spaces";
    // Mix numeric, free text (with spaces/punctuation), and an empty value —
    // the cases the openpyxl script branches on (float vs str vs None).
    QVector<ExcelCellWrite> cells = {
        { 5,  1, QStringLiteral("100") },
        { 6,  3, QStringLiteral("note, with comma") },
        { 7, 13, QString() },                 // empty -> None in the script
        { 8,  4, QStringLiteral("0.0001") },
    };
    QCOMPARE(DVE::buildWriteCellsArgs(f, s, cells),
             refWriteCellsArgs(f, s, cells));
    QCOMPARE(DVE::buildWriteCellsArgs(f, s, cells),
             (QStringList{ f, s,
                           "5", "1", "100",
                           "6", "3", "note, with comma",
                           "7", "13", QString(),
                           "8", "4", "0.0001" }));
}

void TstExcelWritePayload::deleteRowArgs_matchesReference()
{
    const QString f = "C:/data/wb.xlsx";
    const QString s = "Sheet1";
    QCOMPARE(DVE::buildDeleteRowArgs(f, s, 5),
             (QStringList{ f, s, "5" }));
    QCOMPARE(DVE::buildDeleteRowArgs(f, s, 42),
             (QStringList{ f, s, "42" }));
}

// ─── mergePendingWithInFlight — failed-flush re-merge (newest wins per cell) ───
//
// Reproduces the exact semantics of the loop this helper replaced in
// onExcelFlushFinished + finishExcelWritesBlocking: seed with the OLDER
// in-flight cells (order preserved), then overlay each NEWER pending cell —
// same (row,col) overwrites in place, a fresh cell is appended.

namespace {
bool sameCells(const QVector<ExcelCellWrite>& a, const QVector<ExcelCellWrite>& b)
{
    if (a.size() != b.size()) return false;
    for (int i = 0; i < a.size(); ++i) {
        if (a[i].row != b[i].row || a[i].col != b[i].col || a[i].value != b[i].value)
            return false;
    }
    return true;
}
} // namespace

void TstExcelWritePayload::mergePending_emptyPending_keepsInFlight()
{
    QVector<ExcelCellWrite> inFlight = {
        { 5, 1, "100" }, { 6, 2, "200" },
    };
    QVector<ExcelCellWrite> pending; // empty
    const auto merged = DVE::mergePendingWithInFlight(inFlight, pending);
    QVERIFY(sameCells(merged, inFlight));
}

void TstExcelWritePayload::mergePending_emptyInFlight_keepsPending()
{
    QVector<ExcelCellWrite> inFlight; // empty
    QVector<ExcelCellWrite> pending = {
        { 7, 3, "abc" }, { 8, 4, "def" },
    };
    const auto merged = DVE::mergePendingWithInFlight(inFlight, pending);
    QVERIFY(sameCells(merged, pending));
}

void TstExcelWritePayload::mergePending_disjointCells_appendsNewerAfterOlder()
{
    QVector<ExcelCellWrite> inFlight = { { 5, 1, "old1" }, { 6, 2, "old2" } };
    QVector<ExcelCellWrite> pending  = { { 7, 3, "new3" }, { 8, 4, "new4" } };
    const auto merged = DVE::mergePendingWithInFlight(inFlight, pending);
    // Older in-flight first (order preserved), then newer pending appended.
    const QVector<ExcelCellWrite> expected = {
        { 5, 1, "old1" }, { 6, 2, "old2" }, { 7, 3, "new3" }, { 8, 4, "new4" },
    };
    QVERIFY(sameCells(merged, expected));
}

void TstExcelWritePayload::mergePending_sameCell_newerOverlaysInPlace()
{
    QVector<ExcelCellWrite> inFlight = {
        { 5, 1, "stale" }, { 6, 2, "keep" },
    };
    QVector<ExcelCellWrite> pending = {
        { 5, 1, "fresh" },   // same cell as inFlight[0] → overwrites IN PLACE
        { 9, 9, "added" },   // fresh cell → appended
    };
    const auto merged = DVE::mergePendingWithInFlight(inFlight, pending);
    // (5,1) value updated in place (position preserved), (6,2) untouched,
    // (9,9) appended. Newest wins per cell, no duplicate (5,1).
    const QVector<ExcelCellWrite> expected = {
        { 5, 1, "fresh" }, { 6, 2, "keep" }, { 9, 9, "added" },
    };
    QVERIFY(sameCells(merged, expected));
}

void TstExcelWritePayload::mergePending_doesNotMutateInputs()
{
    QVector<ExcelCellWrite> inFlight = { { 5, 1, "old" } };
    QVector<ExcelCellWrite> pending  = { { 5, 1, "new" } };
    const QVector<ExcelCellWrite> inFlightBefore = inFlight;
    const QVector<ExcelCellWrite> pendingBefore  = pending;
    (void)DVE::mergePendingWithInFlight(inFlight, pending);
    QVERIFY(sameCells(inFlight, inFlightBefore));   // inputs untouched
    QVERIFY(sameCells(pending,  pendingBefore));
}

QTEST_MAIN(TstExcelWritePayload)
#include "tst_excelwritepayload.moc"
