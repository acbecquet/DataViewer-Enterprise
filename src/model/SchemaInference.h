#pragma once
#include "TemplateSchema.h"
#include <QVariant>
#include <QVector>

namespace DVE { namespace model {

// Header-driven schema inference (smoke-fix batch): real historical
// workbooks whose block layout is NOT the standard 12-column shape (e.g. the
// 13-col S26 Cart-era sheets - PV1..PV5 inserted before Resistance - and the
// 8-col "User Test Simulation" sheets with a leading Chronology column and
// their own label layout). Builds a TemplateSchema straight from the sheet's
// own header band + column-header row so SchemaDrivenReader can parse it.
// Pure functions of their inputs - no I/O, fully unit-testable.
class SchemaInference {
public:
    // Repeat distance of the first non-empty normalized header token in the
    // column-header row. 0 = no repeat found (single block spanning the
    // row's non-empty width).
    static int inferBlockCols(const QVector<QVariant>& headerRow);

    // True when the sheet is standard-shaped: inferBlockCols is
    // standard.blockCols (or 0 with width <= standard.blockCols) AND the
    // first three header cells normalize-match the standard puffs/
    // before_weight/after_weight aliases at their standard positions.
    // An all-empty header row (no header text at all) also fits - it cannot
    // be inferred from, so legacy files with a blank row 4 keep the existing
    // byte-identical positional path. Standard-shaped sheets MUST keep that
    // path - this predicate is the gate.
    static bool standardFits(const QVector<QVector<QVariant>>& cells,
                             const TemplateSchema& standard);

    // Builds a schema from the sheet itself (columnHeaderRow=4,
    // dataStartRow=5 geometry assumed - every known historical layout shares
    // it):
    //  - columns: each block-1 header cell matched against the KNOWN metric
    //    knowledge base (standardV1 defs plus the alias additions below);
    //    matched -> copy of that MetricDef; unmatched -> NEW MetricDef (key
    //    = snake_case of the header text, type Number if the first non-empty
    //    data cell in that column parses numeric else Text).
    //  - headerFields: scan block-1 cells in rows 1..3; a cell is a LABEL if
    //    it matches the known label-alias table OR its trimmed text ends
    //    with ':' or '?'; the VALUE is the cell to its right; known labels
    //    map to canonical keys, unknown labels become new keys (snake_case).
    //  - aggregates: copy the standard set only when every input metric of
    //    an aggregate exists among the inferred columns/aggregates-so-far
    //    (header: inputs count as satisfied when the header field was found
    //    OR the key is one of the standard schema's own header-field keys).
    static TemplateSchema inferSchema(const QVector<QVector<QVariant>>& cells,
                                      const QString& sheetName);
};

}} // namespace DVE::model
