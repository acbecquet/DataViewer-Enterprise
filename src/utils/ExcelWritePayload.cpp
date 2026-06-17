#include "utils/ExcelWritePayload.h"

namespace DVE {

// ─── The openpyxl scripts (single source of truth) ────────────────────────────
//
// These were inline `static const char*` literals in MainWindow::writeCellsToExcel
// and ::deleteRowFromExcel before SP3-T4. Moving them here means the synchronous
// path, the off-thread worker, and the equivalence test all reference the SAME
// bytes — there is no second copy to drift out of sync. The leading/trailing
// newlines of the raw string are part of the payload and MUST NOT change.

const char* excelWriteCellsScript()
{
    // v2.0.2 M3: atomic save. wb.save() writes the entire OOXML zip
    // by truncating the target file first, so a crash mid-write (power
    // loss, process kill, antivirus interruption) leaves the workbook
    // truncated to zero bytes and the user's data unrecoverable. The
    // tmp + os.replace pattern is atomic on both NTFS and POSIX — the
    // original file is unchanged until the move succeeds.
    static const char* kWriteCells = R"PY(
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
    return kWriteCells;
}

const char* excelDeleteRowScript()
{
    // v2.4.2 R6: atomic save (tmp + os.replace) — the same crash-safety
    // kWriteCells documents. wb.save(path) truncates the target first, so a
    // kill mid-write (power loss, AV, process kill) left the user's source
    // workbook truncated to zero bytes. Keep this tail identical to
    // the write path so the two delete/write paths can't drift again.
    static const char* kDeleteRow = R"PY(
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
    return kDeleteRow;
}

QStringList buildWriteCellsArgs(const QString& filePath,
                                const QString& sheetName,
                                const QVector<ExcelCellWrite>& cells)
{
    QStringList args = { filePath, sheetName };
    for (const ExcelCellWrite& cw : cells) {
        args << QString::number(cw.row) << QString::number(cw.col) << cw.value;
    }
    return args;
}

QStringList buildDeleteRowArgs(const QString& filePath,
                               const QString& sheetName,
                               int excelRow1)
{
    return { filePath, sheetName, QString::number(excelRow1) };
}

} // namespace DVE
