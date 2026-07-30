#include "SchemaDrivenReader.h"
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

// Per metric: the ABSOLUTE grid column it reads from, name-first then position.
static QVector<int> resolveColumns(const QVector<QVector<QVariant>>& g,
                                   const TemplateSchema& s, int off,
                                   ColumnResolution resolution)
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
    // pass 2: positional fallback for the rest (schema index == default position).
    // Known gap: if another metric's name-match already claimed this default slot,
    // both metrics read the same physical column here - unreachable for standard
    // workbooks (every column either all-matches by name or all-falls-back
    // positionally), so left as-is; real collision handling lands with the
    // Phase 2 manifest, where columns carry explicit positions.
    for (int i : unresolved)
        map[i] = off + i;
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

    for (int b = 0; b < blocks; ++b) {
        const int off = b * s.blockCols;
        Sample sample;
        for (const HeaderFieldDef& h : s.headerFields) {
            const QVariant v = cellAt(g, h.row - 1, off + h.col - 1);
            sample.headers.insert(h.key, v);          // raw; typing at lowering
        }
        const QVector<int> colMap = resolveColumns(g, s, off, resolution);
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
    return out;
}

}} // namespace DVE::model
