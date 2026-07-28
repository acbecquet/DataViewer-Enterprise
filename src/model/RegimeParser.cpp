#include "RegimeParser.h"
#include <QRegularExpression>
#include <QStringList>

namespace DVE { namespace model {

namespace {

// "60mL" -> 60 given suffix ml; "2.5s" -> 2.5 given suffix s; strict about the
// unit family but tolerant of spacing and case. Returns false on mismatch.
bool numberWithUnit(const QString& raw, const QStringList& units, double* out)
{
    static const QRegularExpression re(
        QStringLiteral("^\\s*([0-9]+(?:\\.[0-9]+)?)\\s*([a-zA-Z]*)\\s*$"));
    const auto m = re.match(raw);
    if (!m.hasMatch()) return false;
    const QString unit = m.captured(2).toLower();
    if (!units.contains(unit)) return false;
    *out = m.captured(1).toDouble();
    return true;
}

} // namespace

RegimeParts RegimeParser::parse(const QString& text)
{
    RegimeParts r;
    const QStringList parts = text.split(QLatin1Char('/'));
    if (parts.size() < 3 || parts.size() > 4) return r;

    if (!numberWithUnit(parts[0], {QStringLiteral("ml")}, &r.puffVolumeMl)) return r;
    if (!numberWithUnit(parts[1], {QStringLiteral("s"), QStringLiteral("sec")}, &r.puffTimeS)) return r;
    if (!numberWithUnit(parts[2], {QStringLiteral("s"), QStringLiteral("sec")}, &r.puffRestS)) return r;

    if (parts.size() == 4) {
        double v = 0.0;
        if (numberWithUnit(parts[3], {QStringLiteral("s"), QStringLiteral("sec"),
                                      QStringLiteral("second"), QStringLiteral("seconds")}, &v)) {
            r.sessionRestS = v;
        } else if (numberWithUnit(parts[3], {QStringLiteral("min"), QStringLiteral("minute"),
                                             QStringLiteral("minutes"), QStringLiteral("m")}, &v)) {
            r.sessionRestS = v * 60.0;
        } else {
            return r;
        }
    }
    r.valid = true;
    return r;
}

}} // namespace DVE::model
