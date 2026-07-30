#include "Manifest.h"
#include "MetricRegistry.h"
#include <QSet>

namespace DVE { namespace model {

const QString Manifest::kSheetName = QStringLiteral("_dve_schema");

namespace {

// ── Canonical enum spellings - ONE table per enum, used by both gridFor and
//    parse so serialization round-trips by construction. ──

struct TypeName { ValueType type; const char* name; };
const TypeName kTypeNames[] = {
    {ValueType::Number,     "number"},
    {ValueType::Text,       "text"},
    {ValueType::Bool,       "bool"},
    {ValueType::Mixed,      "mixed"},
    {ValueType::NumberList, "numberlist"},
    {ValueType::Image,      "image"},
};

struct RoleName { Role role; const char* name; };
const RoleName kRoleNames[] = {
    {Role::Measured,    "measured"},
    {Role::Qualitative, "qualitative"},
    {Role::Derived,     "derived"},
    {Role::Identity,    "identity"},
};

QString typeName(ValueType t)
{
    for (const TypeName& e : kTypeNames)
        if (e.type == t) return QLatin1String(e.name);
    return QStringLiteral("text");
}

QString roleName(Role r)
{
    for (const RoleName& e : kRoleNames)
        if (e.role == r) return QLatin1String(e.name);
    return QStringLiteral("measured");
}

bool parseTypeName(const QString& s, ValueType* out)
{
    for (const TypeName& e : kTypeNames)
        if (s.compare(QLatin1String(e.name), Qt::CaseInsensitive) == 0) {
            *out = e.type;
            return true;
        }
    return false;
}

bool parseRoleName(const QString& s, Role* out)
{
    for (const RoleName& e : kRoleNames)
        if (s.compare(QLatin1String(e.name), Qt::CaseInsensitive) == 0) {
            *out = e.role;
            return true;
        }
    return false;
}

// ── Cell / list helpers ──

QString cellStr(const QVector<QVariant>& row, int i)
{
    return row.value(i).toString().trimmed();
}

// Comma-separated list (sheets, inputs). Empty parts dropped.
QStringList splitList(const QString& s)
{
    QStringList out;
    const QStringList parts = s.split(QLatin1Char(','), Qt::SkipEmptyParts);
    for (const QString& p : parts) {
        const QString t = p.trimmed();
        if (!t.isEmpty()) out.append(t);
    }
    return out;
}

// Tags serialize as "k=v;k2=v2". Registry tag VALUES contain literal ';' and
// '=' (e.g. smell's scale note), so the writer backslash-escapes both (and
// '\' itself) and the parser splits on UNESCAPED separators only - one escape
// convention shared by both directions. Round-trip is exact up to boundary
// whitespace: segment keys/values are trimmed on parse (grid cells are
// hand-editable; stray spaces are noise, not data).
QString escapeTagText(QString s)
{
    s.replace(QLatin1String("\\"), QLatin1String("\\\\"));
    s.replace(QLatin1String(";"), QLatin1String("\\;"));
    s.replace(QLatin1String("="), QLatin1String("\\="));
    return s;
}

QString unescapeTagText(const QString& s)
{
    QString out;
    out.reserve(s.size());
    for (int i = 0; i < s.size(); ++i) {
        if (s.at(i) == QLatin1Char('\\') && i + 1 < s.size()) ++i;
        out.append(s.at(i));
    }
    return out;
}

QString joinTags(const QMap<QString, QString>& tags)
{
    QStringList pairs;
    for (auto it = tags.constBegin(); it != tags.constEnd(); ++it)
        pairs.append(escapeTagText(it.key()) + QLatin1Char('=') + escapeTagText(it.value()));
    return pairs.join(QLatin1Char(';'));
}

// First position of `sep` in `s` that is not backslash-escaped, or -1.
int indexOfUnescaped(const QString& s, QChar sep, int from = 0)
{
    for (int i = from; i < s.size(); ++i) {
        if (s.at(i) == QLatin1Char('\\')) { ++i; continue; }
        if (s.at(i) == sep) return i;
    }
    return -1;
}

void parseTags(const QString& s, QMap<QString, QString>* tags, QStringList* warnings)
{
    QStringList segments;
    int start = 0;
    while (true) {
        const int semi = indexOfUnescaped(s, QLatin1Char(';'), start);
        segments.append(semi < 0 ? s.mid(start) : s.mid(start, semi - start));
        if (semi < 0) break;
        start = semi + 1;
    }
    for (const QString& seg : segments) {
        if (seg.trimmed().isEmpty()) continue;
        const int eq = indexOfUnescaped(seg, QLatin1Char('='));
        const QString k = eq < 0 ? QString() : unescapeTagText(seg.left(eq).trimmed());
        if (eq <= 0 || k.isEmpty()) {
            warnings->append(QStringLiteral("manifest tag '%1' is not k=v - ignored").arg(seg.trimmed()));
            continue;
        }
        // Merge over inherited registry tags, per-key override.
        tags->insert(k, unescapeTagText(seg.mid(eq + 1).trimmed()));
    }
}

} // namespace

QVector<QVector<QVariant>> Manifest::gridFor(const TemplateSchema& schema,
                                             const QStringList& sheets)
{
    QVector<QVector<QVariant>> g;
    auto row = [&g](const QVector<QVariant>& cells) { g.append(cells); };

    row({QStringLiteral("[schema]")});
    row({QStringLiteral("id"), schema.schemaId});
    row({QStringLiteral("version"), schema.version});
    row({QStringLiteral("sheets"), sheets.join(QStringLiteral(", "))});
    row({QStringLiteral("block_cols"), schema.blockCols});
    row({QStringLiteral("column_header_row"), schema.columnHeaderRow});
    row({QStringLiteral("data_start_row"), schema.dataStartRow});
    row({QStringLiteral("header_rows"), schema.headerRows});

    row({QStringLiteral("[header]")});
    for (const HeaderFieldDef& h : schema.headerFields)
        row({h.key, h.displayName, h.row, h.col, typeName(h.type), h.unit});

    row({QStringLiteral("[column]")});
    for (const MetricDef& c : schema.columns)
        row({c.key, c.displayName, typeName(c.type), c.unit, roleName(c.role),
             c.calculator, c.inputs.join(QStringLiteral(", ")), joinTags(c.tags)});

    return g;
}

Manifest::ParseResult Manifest::parse(const QVector<QVector<QVariant>>& grid)
{
    ParseResult pr;
    enum class Section { Schema, Header, Column };
    Section section = Section::Schema;
    bool blockOpen = false;
    Block cur;
    QSet<QString> columnKeys;   // dup detection, per block; metrics and headers
    QSet<QString> headerKeys;   // are separate namespaces (naming policy)

    auto flush = [&pr, &cur, &columnKeys, &headerKeys, &blockOpen]() {
        if (!blockOpen) return;
        if (cur.schema.columns.isEmpty()) {
            pr.warnings.append(QStringLiteral(
                "manifest block '%1' has no [column] rows - block ignored")
                .arg(cur.schema.schemaId));
        } else {
            if (cur.schema.columns.size() > cur.schema.blockCols) {
                pr.warnings.append(QStringLiteral(
                    "manifest block '%1' declares block_cols %2 but lists %3 columns - block_cols raised to fit")
                    .arg(cur.schema.schemaId).arg(cur.schema.blockCols).arg(cur.schema.columns.size()));
                cur.schema.blockCols = int(cur.schema.columns.size());
            }
            pr.blocks.append(cur);
        }
        cur = Block();
        columnKeys.clear();
        headerKeys.clear();
        blockOpen = false;
    };

    for (const QVector<QVariant>& row : grid) {
        const QString a = cellStr(row, 0);
        if (a.isEmpty()) continue;                       // blank rows ignored
        const QString tag = a.toLower();                 // tags are case-insensitive

        if (tag == QLatin1String("[schema]")) {
            flush();
            blockOpen = true;
            section = Section::Schema;
            continue;
        }
        if (tag == QLatin1String("[header]")) { section = Section::Header; continue; }
        if (tag == QLatin1String("[column]")) { section = Section::Column; continue; }
        if (!blockOpen) continue;   // decorative rows before the first [schema] - ignored

        if (section == Section::Schema) {
            const QVariant val = row.value(1);
            const QString sval = val.toString().trimmed();
            auto intCell = [&val, &sval, &pr, &tag](int* dst) {
                bool ok = false;
                const int n = val.toInt(&ok);
                if (ok) *dst = n;
                else pr.warnings.append(QStringLiteral(
                    "manifest [schema] %1 value '%2' is not a number - kept %3")
                    .arg(tag, sval).arg(*dst));
            };
            if      (tag == QLatin1String("id"))                cur.schema.schemaId = sval;
            else if (tag == QLatin1String("version"))           intCell(&cur.schema.version);
            else if (tag == QLatin1String("sheets"))            cur.sheets = splitList(sval);
            else if (tag == QLatin1String("block_cols"))        intCell(&cur.schema.blockCols);
            else if (tag == QLatin1String("column_header_row")) intCell(&cur.schema.columnHeaderRow);
            else if (tag == QLatin1String("data_start_row"))    intCell(&cur.schema.dataStartRow);
            else if (tag == QLatin1String("header_rows"))       intCell(&cur.schema.headerRows);
            else pr.warnings.append(QStringLiteral(
                "manifest [schema] tag '%1' unknown - ignored").arg(a));
        } else if (section == Section::Header) {
            // key | display | row | col | type | unit
            // Dup-check case-insensitively: "Puffs" vs "puffs" is authoring
            // error, not two fields (canonical keys are lowercase snake_case).
            if (headerKeys.contains(a.toLower())) {
                pr.warnings.append(QStringLiteral(
                    "manifest [header] key '%1' duplicated - second occurrence skipped").arg(a));
                continue;
            }
            headerKeys.insert(a.toLower());
            HeaderFieldDef def;
            if (const HeaderFieldDef* reg = MetricRegistry::headerField(a))
                def = *reg;                               // inherit type/unit/calculator/inputs
            def.key = a;
            const QString display = cellStr(row, 1);
            if (!display.isEmpty()) def.displayName = display;
            bool ok = false;
            const int r = row.value(2).toInt(&ok);
            if (ok && r > 0) def.row = r;
            const int c = row.value(3).toInt(&ok);
            if (ok && c > 0) def.col = c;
            const QString type = cellStr(row, 4);
            if (!type.isEmpty() && !parseTypeName(type, &def.type)) {
                pr.warnings.append(QStringLiteral(
                    "manifest [header] '%1' type '%2' unknown - defaulting to text").arg(a, type));
                def.type = ValueType::Text;
            }
            const QString unit = cellStr(row, 5);
            if (!unit.isEmpty()) def.unit = unit;
            cur.schema.headerFields.append(def);
        } else {
            // key | display | type | unit | role | calculator | inputs | tags
            // Case-insensitive dup-check - see the [header] twin above.
            if (columnKeys.contains(a.toLower())) {
                pr.warnings.append(QStringLiteral(
                    "manifest [column] key '%1' duplicated - second occurrence skipped").arg(a));
                continue;
            }
            columnKeys.insert(a.toLower());
            MetricDef def;
            if (const MetricDef* reg = MetricRegistry::metric(a))
                def = *reg;                               // inherit aliases/tags/calculator
            def.key = a;
            const QString display = cellStr(row, 1);
            if (!display.isEmpty()) def.displayName = display;
            const QString type = cellStr(row, 2);
            if (!type.isEmpty() && !parseTypeName(type, &def.type)) {
                pr.warnings.append(QStringLiteral(
                    "manifest [column] '%1' type '%2' unknown - defaulting to text").arg(a, type));
                def.type = ValueType::Text;
            }
            const QString unit = cellStr(row, 3);
            if (!unit.isEmpty()) def.unit = unit;
            const QString role = cellStr(row, 4);
            if (!role.isEmpty() && !parseRoleName(role, &def.role)) {
                pr.warnings.append(QStringLiteral(
                    "manifest [column] '%1' role '%2' unknown - defaulting to measured").arg(a, role));
                def.role = Role::Measured;
            }
            const QString calc = cellStr(row, 5);
            if (!calc.isEmpty()) def.calculator = calc;
            const QString inputs = cellStr(row, 6);
            if (!inputs.isEmpty()) def.inputs = splitList(inputs);
            const QString tags = cellStr(row, 7);
            if (!tags.isEmpty()) parseTags(tags, &def.tags, &pr.warnings);
            cur.schema.columns.append(def);
        }
    }
    flush();
    return pr;
}

const Manifest::Block* Manifest::blockForSheet(const ParseResult& pr, const QString& sheetName)
{
    const QString want = sheetName.trimmed();
    for (const Block& b : pr.blocks)
        for (const QString& s : b.sheets)
            if (s != QLatin1String("*") && s.trimmed().compare(want, Qt::CaseInsensitive) == 0)
                return &b;
    for (const Block& b : pr.blocks)
        if (b.sheets.contains(QStringLiteral("*")))
            return &b;
    return nullptr;
}

}} // namespace DVE::model
