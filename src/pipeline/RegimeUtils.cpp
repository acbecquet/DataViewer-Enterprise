#include "RegimeUtils.h"
#include "SheetProcessors.h"

#include <QSet>
#include <QRegularExpression>

namespace DVE {
namespace RegimeUtils {

QString unspecifiedLabel() { return QStringLiteral("(unspecified)"); }

bool isStandardRegimeFormat(const QString& text)
{
    // <vol>mL/<puff>s/<interval>s — mL prefix and surrounding whitespace optional,
    // decimals allowed. Empty is treated as standard (the "(unspecified)" state).
    static const QRegularExpression kStd(
        QStringLiteral("^(\\d+(?:\\.\\d+)?\\s*mL\\s*/\\s*)?\\d+(?:\\.\\d+)?\\s*s\\s*/\\s*\\d+(?:\\.\\d+)?\\s*s$"),
        QRegularExpression::CaseInsensitiveOption);
    const QString t = text.trimmed();
    return t.isEmpty() || kStd.match(t).hasMatch();
}

QString canonicalRegime(const QString& text)
{
    const QString t = text.trimmed();
    // "CORESTA" is an accepted alias for the parametric standard label.
    if (t.compare(QLatin1String("CORESTA"), Qt::CaseInsensitive) == 0)
        return QStringLiteral("60mL/3s/30s");
    return t;
}

QString regimeKey(const DataRow& row)
{
    const QString t = row.puffingRegime.trimmed();
    return t.isEmpty() ? unspecifiedLabel() : t;
}

bool isRegimeHeader(const QString& colEHeader)
{
    // Intentionally substring-based (not equality): real-world templates vary
    // the exact header text. Do not tighten to ==; the negative tests confirm
    // "Resistance"/"Resistance (Ω)" do not match.
    const QString h = colEHeader.trimmed();
    return h.contains(QStringLiteral("puffing"), Qt::CaseInsensitive)
        || h.contains(QStringLiteral("regime"),  Qt::CaseInsensitive);
}

bool sheetHasRegimeData(const SheetResult& sheet)
{
    for (const SampleResult& s : sheet.samples)
        for (const DataRow& r : s.rows)
            if (!r.puffingRegime.trimmed().isEmpty())
                return true;
    return false;
}

QStringList uniqueRegimes(const SheetResult& sheet)
{
    QStringList ordered;
    QSet<QString> seen;
    for (const SampleResult& s : sheet.samples) {
        for (const DataRow& r : s.rows) {
            const QString t = r.puffingRegime.trimmed();
            if (t.isEmpty()) continue;
            if (!seen.contains(t)) { seen.insert(t); ordered << t; }
        }
    }
    return ordered;
}

QStringList uniqueRegimes(const FileResult& file)
{
    QStringList ordered;
    QSet<QString> seen;
    for (const SheetResult& sheet : file.sheets) {
        for (const QString& t : uniqueRegimes(sheet)) {
            if (!seen.contains(t)) { seen.insert(t); ordered << t; }
        }
    }
    return ordered;
}

QStringList uniqueRegimeKeys(const SheetResult& sheet)
{
    QStringList ordered;
    QSet<QString> seen;
    for (const SampleResult& s : sheet.samples) {
        for (const DataRow& r : s.rows) {
            const QString k = regimeKey(r);
            if (!seen.contains(k)) { seen.insert(k); ordered << k; }
        }
    }
    return ordered;
}

SheetResult filterByRegime(const SheetResult& sheet, const QString& regime)
{
    GenericSheetProcessor proc;
    SheetResult out = sheet;
    out.samples.clear();

    for (const SampleResult& src : sheet.samples) {
        SampleResult s = src;
        s.rows.clear();
        for (const DataRow& r : src.rows)
            if (regimeKey(r) == regime)
                s.rows.append(r);
        if (s.rows.isEmpty())
            continue;
        // Every kept row carries this regime, so reflect it at the sample level
        // too: keeps a fan-out slide's report table column and long-puff axis
        // detection consistent with the slide's regime (not the block header's).
        s.puffingRegime = regime;
        proc.calculateMetrics(s);
        out.samples.append(s);
    }
    proc.computeSheetAggregates(out);
    return out;
}

} // namespace RegimeUtils
} // namespace DVE
