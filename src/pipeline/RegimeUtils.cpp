#include "RegimeUtils.h"
#include "SheetProcessors.h"

#include <QSet>

namespace DVE {
namespace RegimeUtils {

QString unspecifiedLabel() { return QStringLiteral("(unspecified)"); }

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
        proc.calculateMetrics(s);
        out.samples.append(s);
    }
    proc.computeSheetAggregates(out);
    return out;
}

} // namespace RegimeUtils
} // namespace DVE
