#include "SchemaResolver.h"
#include "SchemaInference.h"
#include "../pipeline/RegimeUtils.h"

namespace DVE { namespace model {

SchemaResolver::Resolution SchemaResolver::resolve(
    const Manifest::ParseResult* manifest,
    const QString& sheetName,
    const QVector<QVector<QVariant>>& cells,
    const QString& templateVersion)
{
    // The ladder decides from cells alone today; templateVersion is applied
    // downstream in processSheet (the Standard-layout heating-technology drop)
    // and stays in the signature for future rungs.
    Q_UNUSED(templateVersion);

    Resolution res;

    // ── Rung 1: manifest. Fires only when the workbook carries a _dve_schema
    //    block naming (or defaulting to, via "*") this sheet. No existing
    //    workbook has one, so byte-identity of the ladder below is preserved
    //    by construction. ──
    if (manifest) {
        if (const Manifest::Block* block = Manifest::blockForSheet(*manifest, sheetName)) {
            res.schema           = block->schema;
            res.source           = Source::Manifest;
            res.columnResolution = ColumnResolution::NameFirst;
            // Manifest sheets are self-describing: a declared puffing_regime
            // column IS the per-row-regime signal.
            res.perRowRegime =
                res.schema.column(QStringLiteral("puffing_regime")) != nullptr;
            return res;
        }
    }

    // ── Rung 3 gate (smoke-fix batch): non-standard block layouts (e.g. the
    //    13-col S26 Cart-era sheets and 8-col UserSim sheets) don't fit the
    //    12-wide positional assumption, so their sample 2+ header reads land a
    //    column short. When the sheet does NOT fit the standard shape, infer a
    //    schema from its own header band + column-header row and parse by name.
    //    Standard-shaped sheets (incl. blank-header legacy files and padded
    //    rows) keep the EXISTING byte-identical positional path below. ──
    if (!SchemaInference::standardFits(cells, standardV1(false))) {
        res.schema           = SchemaInference::inferSchema(cells, sheetName);
        res.source           = Source::Inference;
        res.columnResolution = ColumnResolution::NameFirst;
        res.perRowRegime     = false;
        return res;
    }

    // ── Rung 2: compiled standard. ──
    // Column index 4's header decides per-row Resistance vs Puffing Regime -
    // same rule as processSheetLegacy. processSheet used to read this via
    // reader.getColumnHeaders(): getCellString(3, c) (trimmed text; null ->
    // empty) over a constant 12 columns, so the legacy `hdrs.size() > 4`
    // guard was always true, and its only rewrite ("Resistance" ->
    // "Resistance (Ω)") can never satisfy isRegimeHeader - the raw cell
    // text is byte-equivalent.
    const QString col4Header =
        (cells.size() > 3 && cells[3].size() > 4 && !cells[3][4].isNull())
            ? cells[3][4].toString().trimmed() : QString();
    const bool perRowRegime = RegimeUtils::isRegimeHeader(col4Header);

    res.schema           = standardV1(perRowRegime);
    res.source           = Source::Standard;
    res.columnResolution = ColumnResolution::Positional;
    res.perRowRegime     = perRowRegime;

    // ── Per-block metadata layout: sniff the SAME landmark cells
    //    ExcelReader::extractMetadata sniffs and pick the matching header
    //    layout. Block count mirrors SchemaDrivenReader::parseSheet exactly:
    //    header-row width / blockCols (the countSamples() floor rule), so
    //    blockLayouts lines up 1:1 with the samples parseSheet will produce. ──
    const int hdrRow   = res.schema.columnHeaderRow - 1;
    const int hdrWidth = (hdrRow >= 0 && hdrRow < cells.size()) ? cells[hdrRow].size() : 0;
    const int blocks   = hdrWidth / res.schema.blockCols;
    res.blockLayouts.reserve(blocks);

    for (int b = 0; b < blocks; ++b) {
        const int off = b * res.schema.blockCols;

        auto sniff = [&](int r, int c) -> QString {
            if (r < 0 || r >= cells.size() || (off + c) < 0 || (off + c) >= cells[r].size())
                return QString();
            return cells[r][off + c].toString().trimmed();
        };
        const bool isCart    = sniff(1, 0).contains(QStringLiteral("Cart"), Qt::CaseInsensitive);
        const bool isProject = sniff(0, 5).contains(QStringLiteral("Project:"), Qt::CaseInsensitive);
        res.blockLayouts.append(isCart ? HeaderLayout::Cart
                                       : isProject ? HeaderLayout::Project
                                                   : HeaderLayout::Standard);
    }

    return res;
}

}} // namespace DVE::model
