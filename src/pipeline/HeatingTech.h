#pragma once
#include <QString>

namespace DVE {

// Resistance offset (ohms) added in the power formula P = V^2 / (R + offset),
// keyed by heating technology. Matches the Excel SWITCH formula. S17B/S25B/S25B1
// are specific T58G variants (DV-19) and share its 0.78 offset. Single source of
// truth - all power-calc sites call this instead of inlining the branch.
inline double heatingTechResistanceOffset(const QString& tech)
{
    const QString t = tech.trimmed().toUpper();
    if (t == QLatin1String("CCELL3.0") || t == QLatin1String("CCELL 3.0")
        || t == QLatin1String("T58G")  || t == QLatin1String("S17B")
        || t == QLatin1String("S25B")  || t == QLatin1String("S25B1"))
        return 0.78;
    if (t == QLatin1String("T51"))
        return 0.25;
    return 0.0;
}

// Default resistance (ohms) seeded when the user picks a T58G variant from the
// heating-technology dropdown (DV-19). Editable afterwards. Returns true and sets
// `out` for a known variant; returns false and leaves `out` untouched otherwise.
inline bool heatingTechDefaultResistance(const QString& tech, double& out)
{
    const QString t = tech.trimmed().toUpper();
    if (t == QLatin1String("S17B"))  { out = 1.3; return true; }
    if (t == QLatin1String("S25B"))  { out = 1.1; return true; }
    if (t == QLatin1String("S25B1")) { out = 1.0; return true; }
    return false;
}

} // namespace DVE
