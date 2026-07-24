#include "SchemaInference.h"
#include "SchemaDrivenReader.h"
#include "StandardSchema.h"
#include <QMap>
#include <QSet>

namespace DVE { namespace model {

namespace {

QVariant cellAt(const QVector<QVector<QVariant>>& g, int r, int c)
{
    if (r < 0 || r >= g.size() || c < 0 || c >= g[r].size()) return {};
    return g[r][c];
}

bool cellTextEmpty(const QVariant& v)
{
    return !v.isValid() || v.toString().trimmed().isEmpty();
}

// Tries to parse a cell as a plain double (no unit-stripping tolerance -
// column/header type inference is a simple "does it look like a number").
bool parsesAsDouble(const QVariant& v, double* out)
{
    if (!v.isValid()) return false;
    switch (v.typeId()) {
        case QMetaType::Double:
        case QMetaType::Int:
        case QMetaType::LongLong:
        case QMetaType::UInt:
        case QMetaType::ULongLong:
            if (out) *out = v.toDouble();
            return true;
        default:
            break;
    }
    const QString s = v.toString().trimmed();
    if (s.isEmpty()) return false;
    bool ok = false;
    const double d = s.toDouble(&ok);
    if (ok && out) *out = d;
    return ok;
}

// lowercase; runs of non-alnum characters collapse to a single '_'; no
// leading/trailing '_'. "PV1"->"pv1", "Failure? (if yes, when)"->
// "failure_if_yes_when".
QString snakeCase(const QString& raw)
{
    QString out;
    bool pendingUnderscore = true;   // true = swallow leading separators
    for (const QChar c : raw.toLower()) {
        if (c.isLetterOrNumber()) {
            out.append(c);
            pendingUnderscore = false;
        } else if (!pendingUnderscore) {
            out.append(QLatin1Char('_'));
            pendingUnderscore = true;
        }
    }
    while (out.endsWith(QLatin1Char('_')))
        out.chop(1);
    return out;
}

// Every header alias a MetricDef can be recognized by: displayName, key,
// and headerAliases, all normalized.
QStringList wantedAliases(const MetricDef& m)
{
    QStringList w{SchemaDrivenReader::normalizeHeader(m.displayName),
                  SchemaDrivenReader::normalizeHeader(m.key)};
    for (const QString& a : m.headerAliases)
        w << SchemaDrivenReader::normalizeHeader(a);
    return w;
}

// Known-metric knowledge base = standardV1(false).columns (the non-perRow-
// regime column set, i.e. "resistance" at slot 4) + standardV1(true).columns
// [4] (the "puffing_regime" variant of that same slot) + the smoke-fix
// alias additions below. standardV1's own shipped aliases are untouched -
// these copies are local to SchemaInference.
const QVector<MetricDef>& knownMetricsKnowledgeBase()
{
    static const QVector<MetricDef> kb = [] {
        QVector<MetricDef> v = standardV1(/*perRowRegime=*/false).columns;
        v.append(standardV1(/*perRowRegime=*/true).columns[4]);   // puffing_regime
        for (MetricDef& m : v) {
            if (m.key == QLatin1String("smell"))
                m.headerAliases << QStringLiteral("Smell (0-4)");
            else if (m.key == QLatin1String("clog"))
                m.headerAliases << QStringLiteral("Clog?");
            else if (m.key == QLatin1String("before_weight"))
                m.headerAliases << QStringLiteral("Before Weight/g") << QStringLiteral("Before weight/g");
            else if (m.key == QLatin1String("after_weight"))
                m.headerAliases << QStringLiteral("After Weight/g") << QStringLiteral("After weight/g");
        }
        return v;
    }();
    return kb;
}

// Known header-band LABEL text -> canonical header-field key. Keyed by
// normalized label text so both punctuated ("Voltage:") and bare ("Voltage")
// spellings resolve to the same entry.
const QMap<QString, QString>& labelAliasTable()
{
    static const QMap<QString, QString> table = [] {
        QMap<QString, QString> t;
        auto add = [&](const QString& label, const QString& key) {
            t.insert(SchemaDrivenReader::normalizeHeader(label), key);
        };
        add(QStringLiteral("Cart #"),               QStringLiteral("sample_id"));
        add(QStringLiteral("Ri (Ohms)"),             QStringLiteral("resistance"));
        add(QStringLiteral("Rf (Ohms)"),             QStringLiteral("rf_ohms"));
        add(QStringLiteral("Power"),                 QStringLiteral("power"));
        add(QStringLiteral("Puff Regime"),           QStringLiteral("puffing_regime"));
        add(QStringLiteral("Puffing Regime:"),       QStringLiteral("puffing_regime"));
        add(QStringLiteral("Viscosity"),             QStringLiteral("viscosity"));
        add(QStringLiteral("Viscosity:"),            QStringLiteral("viscosity"));
        add(QStringLiteral("Voltage"),                QStringLiteral("voltage"));
        add(QStringLiteral("Voltage:"),               QStringLiteral("voltage"));
        add(QStringLiteral("Media"),                  QStringLiteral("media"));
        add(QStringLiteral("Media:"),                 QStringLiteral("media"));
        add(QStringLiteral("Date:"),                  QStringLiteral("date"));
        add(QStringLiteral("Sample ID:"),             QStringLiteral("sample_id"));
        add(QStringLiteral("Resistance:"),            QStringLiteral("resistance"));
        add(QStringLiteral("Resistance (Ohms):"),     QStringLiteral("resistance"));
        add(QStringLiteral("Tester:"),                QStringLiteral("tester"));
        add(QStringLiteral("Initial Oil Mass:"),      QStringLiteral("initial_oil_mass"));
        add(QStringLiteral("Test Name"),              QStringLiteral("test_name"));
        return t;
    }();
    return table;
}

// Header-field keys the standard schema itself knows about, across all
// three header-band layouts (Standard/Cart/Project) - used to let an
// aggregate's "header:<key>" input pass even when this sheet's own label
// scan didn't happen to find that field (e.g. "power" is frequently derived
// rather than labelled).
const QSet<QString>& knownStandardHeaderKeys()
{
    static const QSet<QString> keys = [] {
        QSet<QString> s;
        const HeaderLayout layouts[] = {HeaderLayout::Standard, HeaderLayout::Cart, HeaderLayout::Project};
        for (HeaderLayout layout : layouts) {
            const TemplateSchema t = standardV1(/*perRowRegime=*/false, layout);
            for (const HeaderFieldDef& h : t.headerFields)
                s.insert(h.key);
        }
        return s;
    }();
    return keys;
}

} // namespace

int SchemaInference::inferBlockCols(const QVector<QVariant>& headerRow)
{
    int firstIdx = -1;
    QString firstNorm;
    for (int c = 0; c < headerRow.size(); ++c) {
        const QString norm = SchemaDrivenReader::normalizeHeader(headerRow[c].toString());
        if (!norm.isEmpty()) { firstIdx = c; firstNorm = norm; break; }
    }
    if (firstIdx < 0) return 0;   // no header text at all

    for (int c = firstIdx + 1; c < headerRow.size(); ++c) {
        if (SchemaDrivenReader::normalizeHeader(headerRow[c].toString()) == firstNorm)
            return c - firstIdx;
    }
    return 0;   // no repeat found
}

bool SchemaInference::standardFits(const QVector<QVector<QVariant>>& cells,
                                   const TemplateSchema& standard)
{
    const int hdrRowIdx = standard.columnHeaderRow - 1;
    const QVector<QVariant> headerRow =
        (hdrRowIdx >= 0 && hdrRowIdx < cells.size()) ? cells[hdrRowIdx] : QVector<QVariant>();

    // An all-empty header row cannot be inferred from at all (blank row 4 on
    // a legacy positional-only file) - treat it as fitting the standard so
    // those files keep the existing byte-identical path.
    bool anyText = false;
    for (const QVariant& v : headerRow) {
        if (!SchemaDrivenReader::normalizeHeader(v.toString()).isEmpty()) { anyText = true; break; }
    }
    if (!anyText) return true;

    const int width      = headerRow.size();
    const int blockCols  = inferBlockCols(headerRow);
    const bool widthFits = (blockCols == standard.blockCols) ||
                           (blockCols == 0 && width <= standard.blockCols);
    if (!widthFits) return false;

    for (int i = 0; i < 3 && i < standard.columns.size(); ++i) {
        const QStringList wanted = wantedAliases(standard.columns[i]);
        const QString cellNorm = (i < headerRow.size())
            ? SchemaDrivenReader::normalizeHeader(headerRow[i].toString()) : QString();
        if (cellNorm.isEmpty() || !wanted.contains(cellNorm))
            return false;
    }
    return true;
}

TemplateSchema SchemaInference::inferSchema(const QVector<QVector<QVariant>>& cells,
                                            const QString& sheetName)
{
    TemplateSchema s;
    s.schemaId        = QStringLiteral("inferred:%1").arg(sheetName);
    s.version         = 1;
    s.headerRows      = 3;
    s.columnHeaderRow = 4;
    s.dataStartRow    = 5;

    const int hdrRowIdx = s.columnHeaderRow - 1;
    const QVector<QVariant> headerRow =
        (hdrRowIdx >= 0 && hdrRowIdx < cells.size()) ? cells[hdrRowIdx] : QVector<QVariant>();

    int blockCols = inferBlockCols(headerRow);
    if (blockCols <= 0) {
        // Single block: span the header row's non-empty width.
        int w = 0;
        for (int c = headerRow.size() - 1; c >= 0; --c) {
            if (!SchemaDrivenReader::normalizeHeader(headerRow[c].toString()).isEmpty()) { w = c + 1; break; }
        }
        blockCols = w;
    }
    s.blockCols = blockCols;

    // ── Columns: block-1 column-header cells matched against the known
    //    metric knowledge base; unmatched -> new open metric. ──
    const QVector<MetricDef>& kb = knownMetricsKnowledgeBase();
    for (int c = 0; c < blockCols; ++c) {
        const QString headerText = (c < headerRow.size()) ? headerRow[c].toString().trimmed() : QString();
        const QString norm = SchemaDrivenReader::normalizeHeader(headerText);

        const MetricDef* match = nullptr;
        if (!norm.isEmpty()) {
            for (const MetricDef& m : kb) {
                if (wantedAliases(m).contains(norm)) { match = &m; break; }
            }
        }

        if (match) {
            s.columns.append(*match);
            continue;
        }

        MetricDef nm;
        nm.key         = headerText.isEmpty() ? QStringLiteral("col_%1").arg(c + 1) : snakeCase(headerText);
        nm.displayName = headerText;
        bool numeric = false;
        for (int r = s.dataStartRow - 1; r < cells.size(); ++r) {
            const QVariant v = cellAt(cells, r, c);
            if (cellTextEmpty(v)) continue;
            double d;
            numeric = parsesAsDouble(v, &d);
            break;
        }
        nm.type     = numeric ? ValueType::Number : ValueType::Text;
        nm.role     = numeric ? Role::Measured : Role::Qualitative;
        nm.editable = (nm.role == Role::Qualitative);
        s.columns.append(nm);
    }

    // ── Header fields: label:value pairs scanned across block-1 rows 1..3. ──
    for (int r = 0; r < 3 && r < cells.size(); ++r) {
        const QVector<QVariant>& row = cells[r];
        for (int c = 0; c < blockCols; ) {
            const QString raw = (c < row.size()) ? row[c].toString().trimmed() : QString();
            if (raw.isEmpty()) { ++c; continue; }

            const QString norm      = SchemaDrivenReader::normalizeHeader(raw);
            const QString canonical = labelAliasTable().value(norm);
            const bool punctLabel   = raw.endsWith(QLatin1Char(':')) || raw.endsWith(QLatin1Char('?'));
            if (canonical.isEmpty() && !punctLabel) { ++c; continue; }   // pure value / unrelated text

            const QString key = !canonical.isEmpty() ? canonical : snakeCase(raw);
            const QVariant value = (c + 1 < blockCols) ? cellAt(cells, r, c + 1) : QVariant();

            if (!s.headerField(key)) {   // first occurrence wins
                HeaderFieldDef field;
                field.key         = key;
                field.displayName = raw;
                double d;
                field.type = parsesAsDouble(value, &d) ? ValueType::Number : ValueType::Text;
                field.row  = r + 1;
                field.col  = c + 2;      // 1-based column of the VALUE cell
                s.headerFields.append(field);
            }
            c += 2;   // skip the label AND its value cell
        }
    }

    // ── Aggregates: copy the standard set only when every input already
    //    resolves - against inferred columns, aggregates accepted earlier in
    //    this same pass (standardV1's list is dependency-ordered), or, for
    //    "header:" inputs, either a found header field or a key the standard
    //    schema itself defines (power/initial_oil_mass etc. may be derived
    //    later even when not literally labelled on this sheet). ──
    QSet<QString> available;
    for (const MetricDef& m : s.columns) available.insert(m.key);

    const TemplateSchema std0 = standardV1(/*perRowRegime=*/false);
    for (const AggregateDef& a : std0.aggregates) {
        bool ok = true;
        for (const QString& inp : a.inputs) {
            if (inp.startsWith(QLatin1String("header:"))) {
                const QString key = inp.mid(7);
                if (!s.headerField(key) && !knownStandardHeaderKeys().contains(key)) { ok = false; break; }
            } else if (!available.contains(inp)) {
                ok = false; break;
            }
        }
        if (ok) {
            s.aggregates.append(a);
            available.insert(a.key);
        }
    }

    return s;
}

}} // namespace DVE::model
