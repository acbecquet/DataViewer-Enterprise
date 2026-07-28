#pragma once
#include <QString>

namespace DVE { namespace model {

// Parsed puffing-regime string ("60mL/3s/30s" or "60mL/3s/30s/5minute").
// Registry contract (RATIFIED 2026-07-27, section 8.2/9.2): the canonical
// representation is these four values; the composite string is a legacy
// source encoding. 3-part regimes default sessionRestS to 0.
struct RegimeParts {
    bool   valid        = false;
    double puffVolumeMl = 0.0;   // part 1, mL
    double puffTimeS    = 0.0;   // part 2 (always the middle value), s
    double puffRestS    = 0.0;   // part 3, s
    double sessionRestS = 0.0;   // optional part 4, s (minutes converted)
};

class RegimeParser {
public:
    static RegimeParts parse(const QString& text);
};

}} // namespace DVE::model
