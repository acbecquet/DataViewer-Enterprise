#pragma once
#include "TemplateSchema.h"
#include <QVariant>
#include <QVector>

namespace DVE { namespace model {

// The _dve_schema manifest: an Excel-native grid describing the template.
// Parsed through the same cell path as every other sheet; validation is
// poka-yoke (warn + proceed, registry naming-policy rule 5) - a manifest can
// be wrong, never fatal. Known metric/header keys inherit their registry defs
// (aliases, tags, calculators); manifest cells override display/unit/type.
class Manifest {
public:
    static const QString kSheetName;   // "_dve_schema"

    struct Block {
        TemplateSchema schema;
        QStringList    sheets;         // sheet names this block applies to; "*" = default
    };
    struct ParseResult {
        QVector<Block> blocks;
        QStringList    warnings;
    };

    static ParseResult parse(const QVector<QVector<QVariant>>& grid);

    // First block naming the sheet, else the "*" block, else nullptr.
    static const Block* blockForSheet(const ParseResult& pr, const QString& sheetName);

    // Writer: serialize a schema to manifest grid rows (round-trips through
    // parse). Appendable - concatenate grids for multi-block manifests.
    static QVector<QVector<QVariant>> gridFor(const TemplateSchema& schema,
                                              const QStringList& sheets);
};

}} // namespace DVE::model
