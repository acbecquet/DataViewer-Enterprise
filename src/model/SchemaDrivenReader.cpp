#include "SchemaDrivenReader.h"
#include <QDebug>
#include <QtGlobal>

namespace DVE { namespace model {

QString SchemaDrivenReader::normalizeHeader(const QString& s)
{
    QString out;
    for (const QChar c : s.toLower())
        if (c.isLetterOrNumber()) out.append(c);
    return out;
}

static QVariant cellAt(const QVector<QVector<QVariant>>& g, int r, int c)
{
    if (r < 0 || r >= g.size() || c < 0 || c >= g[r].size()) return {};
    return g[r][c];
}

static bool cellEmpty(const QVariant& v)
{
    return !v.isValid() || v.toString().trimmed().isEmpty();
}

// True when every header cell and every data value of the sample is empty -
// the shape a padded-out block (no real content) parses into.
static bool sampleIsBlank(const Sample& s)
{
    for (auto it = s.headers.constBegin(); it != s.headers.constEnd(); ++it)
        if (!cellEmpty(it.value())) return false;
    for (const MetricSeries& m : s.data)
        for (const QVariant& v : m.values)
            if (!cellEmpty(v)) return false;
    return true;
}

// Per metric: the ABSOLUTE grid column it reads from, name-first then position.
// `collided` is set (never cleared) when a pass-2 fallback found its default
// slot already claimed by a name match and had to move - the caller warns
// once per parse.
static QVector<int> resolveColumns(const QVector<QVector<QVariant>>& g,
                                   const TemplateSchema& s, int off,
                                   ColumnResolution resolution,
                                   bool* collided)
{
    // Positional: metric i reads physical column off+i, byte-for-byte the same
    // as legacy ExcelReader::extractRow(row, off, blockCols). No header lookup,
    // so misaligned / renamed historical layouts read exactly as legacy did.
    if (resolution == ColumnResolution::Positional) {
        QVector<int> map(s.columns.size());
        for (int i = 0; i < s.columns.size(); ++i)
            map[i] = off + i;
        return map;
    }

    const int hdrRow = s.columnHeaderRow - 1;
    QVector<int> map(s.columns.size());
    QVector<bool> taken(s.blockCols, false);
    // pass 1: name matches
    QVector<int> unresolved;
    for (int i = 0; i < s.columns.size(); ++i) {
        const MetricDef& m = s.columns[i];
        QStringList wanted{SchemaDrivenReader::normalizeHeader(m.displayName),
                           SchemaDrivenReader::normalizeHeader(m.key)};
        for (const QString& a : m.headerAliases)
            wanted << SchemaDrivenReader::normalizeHeader(a);
        int found = -1;
        for (int c = 0; c < s.blockCols; ++c) {
            if (taken[c]) continue;
            const QString h = SchemaDrivenReader::normalizeHeader(
                cellAt(g, hdrRow, off + c).toString());
            if (!h.isEmpty() && wanted.contains(h)) { found = c; break; }
        }
        if (found >= 0) { map[i] = off + found; taken[found] = true; }
        else            { map[i] = -1; unresolved << i; }
    }
    // pass 2: positional fallback for the rest (schema index == default
    // position). Collision poka-yoke (registry naming-policy rule 5 - warn,
    // never fail): when another metric's name match already claimed the
    // default slot, take the nearest FREE slot at-or-after it, wrapping to
    // earlier free slots when none remain after - never share a taken one.
    // A free slot always exists while columns.size() <= blockCols (pass 1
    // takes at most one slot per metric); the identity assignment below is
    // release-mode defense for the Q_ASSERT-guarded malformed case.
    for (int i : unresolved) {
        int slot = -1;
        for (int c = i; c < s.blockCols; ++c)
            if (!taken[c]) { slot = c; break; }
        if (slot < 0)
            for (int c = 0; c < i && c < s.blockCols; ++c)
                if (!taken[c]) { slot = c; break; }
        if (slot < 0) slot = i;
        if (slot != i && collided) *collided = true;
        if (slot < s.blockCols) taken[slot] = true;
        map[i] = off + slot;
    }
    return map;
}

Sheet SchemaDrivenReader::parseSheet(const QVector<QVector<QVariant>>& g,
                                     const QString& sheetName,
                                     const TemplateSchema& s,
                                     bool perRowRegime,
                                     ColumnResolution resolution)
{
    // A schema declaring more metrics than blockCols would silently read into
    // the next block's cells - fail fast in debug rather than emit corrupt data.
    Q_ASSERT(s.columns.size() <= s.blockCols);

    Sheet out;
    out.sheetName    = sheetName;
    out.schema       = s;
    out.perRowRegime = perRowRegime;

    const int hdrRow   = s.columnHeaderRow - 1;
    const int hdrWidth = (hdrRow >= 0 && hdrRow < g.size()) ? g[hdrRow].size() : 0;
    const int blocks   = hdrWidth / s.blockCols;      // same floor rule as countSamples()

    bool collided = false;
    for (int b = 0; b < blocks; ++b) {
        const int off = b * s.blockCols;
        Sample sample;
        for (const HeaderFieldDef& h : s.headerFields) {
            const QVariant v = cellAt(g, h.row - 1, off + h.col - 1);
            sample.headers.insert(h.key, v);          // raw; typing at lowering
        }
        const QVector<int> colMap = resolveColumns(g, s, off, resolution, &collided);
        if (b == 0) {
            // Expose block 0's resolution as block-relative slots - see
            // Sheet::columnSlots. Identity under Positional by construction.
            out.columnSlots.resize(colMap.size());
            for (int i = 0; i < colMap.size(); ++i)
                out.columnSlots[i] = colMap[i] - off;
        }
        for (const MetricDef& m : s.columns)
            sample.data.append(MetricSeries{m.key, {}});
        for (int r = s.dataStartRow - 1; r < g.size(); ++r) {
            bool allEmpty = true;
            // Scans the schema's RESOLVED metric columns, not the full physical
            // block width like legacy getSample - identical for standard-v1
            // (a bijective 12-col map onto blockCols=12), a conscious choice for
            // narrower future schemas that don't claim every physical column.
            for (int i = 0; i < colMap.size(); ++i)
                if (!cellEmpty(cellAt(g, r, colMap[i]))) { allEmpty = false; break; }
            if (allEmpty) break;                      // legacy stop rule
            for (int i = 0; i < colMap.size(); ++i)
                sample.data[i].values.append(cellAt(g, r, colMap[i]));
        }
        out.samples.append(sample);
    }
    // Phantom trailing samples - schema (NameFirst) paths only: rows come
    // back padded to the sheet-wide max width, so hdrWidth / blockCols can
    // over-count blocks and grow trailing samples with no content at all.
    // Drop them. Positional keeps the legacy phantom behavior untouched
    // (byte-identity; the shadow suite is the referee).
    if (resolution == ColumnResolution::NameFirst)
        while (!out.samples.isEmpty() && sampleIsBlank(out.samples.last()))
            out.samples.removeLast();
    // Once per parse (registry naming-policy rule 5: warn, never fail).
    if (collided)
        qWarning().noquote()
            << "SchemaDrivenReader: header name-match collision on sheet"
            << sheetName
            << "- positional fallback moved to the nearest free slot";
    return out;
}

}} // namespace DVE::model
