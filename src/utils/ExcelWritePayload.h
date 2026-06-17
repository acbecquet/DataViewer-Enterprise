#pragma once

#include <QString>
#include <QStringList>
#include <QVector>

// ─────────────────────────────────────────────────────────────────────────────
// ExcelWritePayload — the PURE builders for the openpyxl write-back subprocess.
//
// SP3-T4 (R6): the Excel write-back was moved off the UI thread. The off-thread
// worker and the synchronous close path must drive the EXACT same openpyxl
// invocation, and a unit test must be able to prove the payload didn't change
// when the call moved threads. To make that possible the "what we run" decision
// — the python source text + the argv after the script path — lives here as
// free functions with NO QProcess, NO QWidget, NO MainWindow state. The worker,
// the synchronous fallback, and the equivalence test all call these same builders,
// so the threading move provably cannot alter what gets written to the workbook.
// (Keystone lesson from SP2: test the real payload builder, not a parallel copy.)
// ─────────────────────────────────────────────────────────────────────────────

namespace DVE {

// One pending cell edit (Excel 1-based row/col). This is the single definition;
// MainWindow::CellWrite is a `using` alias of it so m_pendingWrites flows into
// the builder with no conversion (no opportunity for drift).
struct ExcelCellWrite {
    int     row;
    int     col;
    QString value;
};

// The python source text. Exposed so MainWindow's synchronous path and the
// off-thread worker reference the identical script (no duplicated literal that
// could silently diverge). Defined once in ExcelWritePayload.cpp.
const char* excelWriteCellsScript();
const char* excelDeleteRowScript();

// Build the argv (AFTER the script path) for writing `cells` into `sheetName`
// of `filePath`. Triplets: path, sheet, then row col value per cell — byte-for-byte
// what the original inline code produced.
QStringList buildWriteCellsArgs(const QString& filePath,
                                const QString& sheetName,
                                const QVector<ExcelCellWrite>& cells);

// Build the argv (AFTER the script path) for deleting Excel 1-based row
// `excelRow1` from `sheetName` of `filePath`.
QStringList buildDeleteRowArgs(const QString& filePath,
                               const QString& sheetName,
                               int excelRow1);

// SP3-T4 (R6 fix): merge a failed in-flight batch back together with the cells
// that accumulated while it was running, when both target the SAME file/sheet.
// The in-flight cells are the OLDER edits, so they seed the result (preserving
// their order); each newer pending cell then OVERLAYS by (row,col) — a same-cell
// pending edit overwrites the in-flight value in place, and a pending edit to a
// fresh cell is appended. The result is the de-duplicated union with newest-wins
// per cell. Pure list logic shared by onExcelFlushFinished() and
// finishExcelWritesBlocking() so the two retry paths can't drift; unit-tested
// directly. Does NOT mutate its inputs.
QVector<ExcelCellWrite> mergePendingWithInFlight(
    const QVector<ExcelCellWrite>& inFlight,
    const QVector<ExcelCellWrite>& pending);

} // namespace DVE
