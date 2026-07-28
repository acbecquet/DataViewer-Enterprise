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
