#pragma once
#include "SchemaDrivenReader.h"
#include "../ExcelReader.h"
#include "../pipeline/ReportData.h"

namespace DVE { namespace model {

// Phase-1 strangler shell: lowers the v3 model to the legacy raw shape so the
// UNCHANGED SheetProcessor chain computes every metric - production output is
// byte-identical by construction. Removed in Phase 4.
class LegacyAdapter {
public:
    static ExcelReader::SampleData lowerSample(const Sample& s,
                                               const TemplateSchema& schema,
                                               int blockIndex);

    // Smoke-fix batch: lowers an INFERRED (non-standard) sheet directly to a
    // SheetResult - the inferred schema's block width is 13/8, not 12, so there
    // is no faithful 12-wide SampleData detour. Known metric keys map to the
    // fixed DataRow / SampleResult fields (mirroring SheetProcessor::
    // buildSampleResult's coercions + repair rules); every unknown per-row
    // metric rides in DataRow::extra and every unknown header field in
    // SampleResult::extra. Derived columns from the sheet are dropped and the
    // tpm chain is recomputed by GenericSheetProcessor::calculateMetrics /
    // computeSheetAggregates, exactly as the standard path overwrites them.
    // Sets SheetResult::fromInferredSchema = true.
    static SheetResult lowerInferredSheet(const Sheet& sheet,
                                          const QString& sheetName,
                                          const QString& templateVersion);
};

}} // namespace DVE::model
