#pragma once
#include "SchemaDrivenReader.h"
#include "../ExcelReader.h"

namespace DVE { namespace model {

// Phase-1 strangler shell: lowers the v3 model to the legacy raw shape so the
// UNCHANGED SheetProcessor chain computes every metric - production output is
// byte-identical by construction. Removed in Phase 4.
class LegacyAdapter {
public:
    static ExcelReader::SampleData lowerSample(const Sample& s,
                                               const TemplateSchema& schema,
                                               int blockIndex);
};

}} // namespace DVE::model
