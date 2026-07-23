#pragma once
#include "TemplateSchema.h"

namespace DVE { namespace model {

// Compiled-in schema reproducing the current standardized template exactly
// (spec section 6). perRowRegime selects the col-5 variant: the new-template
// per-row Puffing Regime text column vs the old-template Resistance column
// (same rule DataProcessor::processSheet applies via RegimeUtils today).
TemplateSchema standardV1(bool perRowRegime);

}} // namespace DVE::model
