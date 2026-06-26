#pragma once
#include <QString>
#include <QRegularExpression>

namespace DVE {

// A sensory "tester" string split into a display name + round marker.
// Double-blind rounds 1/2 are encoded as a trailing " R1"/" R2" suffix on the
// stored testerName (the long-standing hand-typed convention). Anything else
// yields round == "N/A".
struct TesterRound {
    QString tester;   // name without the round suffix
    QString round;    // "1", "2", or "N/A"
};

// Parse a stored testerName. Trailing " R1"/" R2" (after at least one
// non-space char) is treated as the round; otherwise round = "N/A".
inline TesterRound splitTesterRound(const QString& stored)
{
    static const QRegularExpression re(QStringLiteral("^(.*\\S)\\s+R([12])$"));
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
    if (round == QLatin1String("1")) return t + QStringLiteral(" R1");
    if (round == QLatin1String("2")) return t + QStringLiteral(" R2");
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
