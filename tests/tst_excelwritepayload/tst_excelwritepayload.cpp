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

QTEST_MAIN(TstExcelWritePayload)
#include "tst_excelwritepayload.moc"
