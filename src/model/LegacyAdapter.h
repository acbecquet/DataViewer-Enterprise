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

    // Smoke-fix batch, generalized in Phase 2c: lowers a schema-described
    // sheet (header-INFERRED or MANIFEST-declared) directly to a SheetResult -
    // these block widths (13/8/manifest-declared) have no faithful 12-wide
    // SampleData detour. Known metric keys map to the fixed DataRow /
    // SampleResult fields (mirroring SheetProcessor::buildSampleResult's
    // coercions + repair rules); a per-row puffing_regime series lands in
    // DataRow::puffingRegime; every unknown per-row metric rides in
    // DataRow::extra and every unknown header field in SampleResult::extra.
    // Derived columns from the sheet are dropped and the tpm chain is
    // recomputed by GenericSheetProcessor::calculateMetrics /
    // computeSheetAggregates, exactly as the standard path overwrites them.
    // Write-provenance columnKeys are recorded in RESOLVED physical slot
    // order (Sheet::columnSlots inversion), so NameFirst sheets with
    // reordered columns write back to the right cells; identity resolutions
    // (inference / positional) keep those outputs byte-unchanged.
    static SheetResult lowerSchemaSheet(const Sheet& sheet,
                                        const QString& sheetName,
                                        const QString& templateVersion,
                                        bool fromInference,
                                        bool perRowRegime);

    // Thin forwarder (pre-2c signature, call sites + tests untouched):
    // inference sheets lower with fromInference=true and no per-row regime.
    static SheetResult lowerInferredSheet(const Sheet& sheet,
                                          const QString& sheetName,
                                          const QString& templateVersion);
};

}} // namespace DVE::model
