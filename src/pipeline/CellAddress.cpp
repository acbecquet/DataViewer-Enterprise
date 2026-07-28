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
