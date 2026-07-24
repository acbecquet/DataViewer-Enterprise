#pragma once
#include "TemplateSchema.h"

namespace DVE { namespace model {

// Header-band layout for the compiled standard schema. All three variants share
// IDENTICAL data columns + aggregates; only the header-field cell references
// differ, mirroring the three metadata branches of ExcelReader::extractMetadata
// exactly (spec section 6). The Phase-1 strangler sniffs the same landmark cells
// extractMetadata sniffs and picks the matching layout per block:
//   Standard - current standardized template (Dec-2025 "new" / Jan-2025 "old").
//              NB the "old" template writes no Heating Technology cell;
//              extractMetadata's old branch leaves it empty, so the strangler
//              drops that key when tv != "new" (see DataProcessor::processSheet).
//   Cart     - legacy "Cart #" sheets (Formats A/B): sample id / resistance /
//              viscosity on row 2; media / puff-regime / voltage on row 3. No
//              test name / date / tester / heating tech / initial oil mass.
//   Project  - legacy "Project:" sheets (Format C): sample id is ASSEMBLED from a
//              project-name + sample-suffix cell pair (joined in processSheet,
//              matching extractMetadata); date + tester live on row 1.
enum class HeaderLayout { Standard, Cart, Project };

// Compiled-in schema reproducing the current standardized template exactly
// (spec section 6). perRowRegime selects the col-5 variant: the new-template
// per-row Puffing Regime text column vs the old-template Resistance column
// (same rule DataProcessor::processSheet applies via RegimeUtils today).
// `layout` selects the header-band variant; data columns + aggregates are the
// same for every layout.
TemplateSchema standardV1(bool perRowRegime, HeaderLayout layout = HeaderLayout::Standard);

}} // namespace DVE::model
