#pragma once
#include <QString>
#include <QRegularExpression>

namespace DVE {

// A sensory "tester" string split into a display name + round marker.
// Rounds 1-4 are encoded as a trailing " R1".." R4" suffix on the stored
// testerName (the double-blind convention; R1/R2 are long-standing, R3/R4 added
// universally in v2.5.13). Anything else yields round == "N/A".
struct TesterRound {
    QString tester;   // name without the round suffix
    QString round;    // "1".."4", or "N/A"
};

// Parse a stored testerName. A trailing " R1".." R4" (after at least one
// non-space char) is treated as the round; otherwise round = "N/A".
inline TesterRound splitTesterRound(const QString& stored)
{
    static const QRegularExpression re(QStringLiteral("^(.*\\S)\\s+R([1-4])$"));
    const QRegularExpressionMatch m = re.match(stored);
    if (m.hasMatch())
        return { m.captured(1), m.captured(2) };
    return { stored, QStringLiteral("N/A") };
}

// Recombine into the stored testerName. Empty tester stays empty (round is
// meaningless without a tester, and SensoryPanel::sessionLabel() must still
// fall back to the assessor). Round "N/A" (or anything other than 1/2)
// appends nothing.
inline QString combineTesterRound(const QString& tester, const QString& round)
{
    const QString t = tester.trimmed();
    if (t.isEmpty()) return t;
    // Rounds 1-4 fold into the testerName as " R1".." R4"; anything else
    // (e.g. "N/A") appends nothing. R1/R2 behaviour is unchanged.
    if (round == QLatin1String("1") || round == QLatin1String("2")
     || round == QLatin1String("3") || round == QLatin1String("4"))
        return t + QStringLiteral(" R") + round;
    return t;
}

// DATAVIEWER-15: the canonical session_name (the natural-key TEXT column) is the
// trimmed Test Title ALONE -- the single shape the live-save path and both import
// paths must agree on. Routing all three through this keeps a re-saved imported
// session from forking its key from "title - tester" back to "title". The DB
// natural key carries the tester (incl. round) in tester_name, not here.
inline QString canonicalSensorySessionName(const QString& testTitle)
{
    return testTitle.trimmed();
}

} // namespace DVE
