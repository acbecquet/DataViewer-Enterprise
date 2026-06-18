#include "CompatClassifier.h"

#include <QVersionNumber>
#include <QList>

namespace DVE {

QString unstampedEraLabel() { return QStringLiteral("pre-v2.4.2 (unstamped)"); }

namespace {

// Strip a "DataViewer/" prefix and a leading 'v'/'V' so QVersionNumber can parse
// "DataViewer/2.2.5", "v2.2.5", or "2.2.5" uniformly. Mirrors UpdateChecker's
// version-parse idiom.
QString normalizeVersion(QString s)
{
    s = s.trimmed();
    const int slash = s.lastIndexOf('/');
    if (slash >= 0) s = s.mid(slash + 1);
    if (!s.isEmpty() && (s.at(0) == QLatin1Char('v') || s.at(0) == QLatin1Char('V')))
        s = s.mid(1);
    return s.trimmed();
}

QString eraLabelFor(int major, int minor)
{
    return QStringLiteral("v%1.%2.x").arg(major).arg(minor);
}

// Coarse era start dates (first release of each v2 minor line), from git tags +
// release commits. Only consulted for unstamped (NULL app_version) rows, to
// infer an approximate era from a row's creation date.
struct EraStart { int major; int minor; QDate date; };
const QList<EraStart>& eraTable()
{
    static const QList<EraStart> t = {
        { 2, 0, QDate(2026, 5, 13) },  // v2.0.0
        { 2, 1, QDate(2026, 5, 27) },  // v2.1.0
        { 2, 2, QDate(2026, 5, 29) },  // v2.2.0
        { 2, 3, QDate(2026, 6, 5)  },  // v2.3.0 / v2.3.1
        { 2, 4, QDate(2026, 6, 8)  },  // v2.4.0
    };
    return t;
}

} // namespace

CompatClass classifyEra(const QString& appVersion, const QDate& creation)
{
    CompatClass c;

    // 1. Stamped: derive the era directly from the version string.
    const QString norm = normalizeVersion(appVersion);
    if (!norm.isEmpty()) {
        qsizetype suffixIndex = 0;
        const QVersionNumber v = QVersionNumber::fromString(norm, &suffixIndex);
        if (!v.isNull() && suffixIndex > 0) {
            c.eraLabel = eraLabelFor(v.majorVersion(), v.minorVersion());
            c.approx = false;
            return c;
        }
    }

    // 2. Unstamped but datable: infer the newest era whose start <= creation.
    if (creation.isValid()) {
        const QList<EraStart>& tbl = eraTable();
        for (int i = tbl.size() - 1; i >= 0; --i) {
            if (tbl.at(i).date <= creation) {
                c.eraLabel = eraLabelFor(tbl.at(i).major, tbl.at(i).minor);
                c.approx = true;
                return c;
            }
        }
    }

    // 3. Unstamped and undatable (or predates v2.0.0): unknown bucket.
    c.eraLabel = unstampedEraLabel();
    c.approx = false;
    return c;
}

QStringList sensoryHealth(bool hasLegacyStringScores, bool isPlaceholder, int sampleCount)
{
    QStringList h;
    if (hasLegacyStringScores) h << QStringLiteral("Legacy string scores");
    if (isPlaceholder || sampleCount <= 0) h << QStringLiteral("Junk candidate");
    return h;
}

QStringList fileHealth(int sampleCount, bool missingRegimes)
{
    QStringList h;
    if (sampleCount <= 0)  h << QStringLiteral("No samples");
    if (missingRegimes)    h << QStringLiteral("Missing puff regimes");
    return h;
}

} // namespace DVE
