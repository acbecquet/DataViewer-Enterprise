#pragma once
#include "Manifest.h"
#include "SchemaDrivenReader.h"
#include "StandardSchema.h"
#include <QVariant>
#include <QVector>

namespace DVE { namespace model {

// One routing ladder for every sheet: manifest -> compiled standard ->
// header-driven inference. Absorbs the standardFits fork and the per-block
// Cart/Project landmark sniff that DataProcessor carried inline. The resolver
// only DECIDES (schema, layouts, flags) - it is a pure function of the cells;
// applying the decision (header-variant swaps, sampleID join, lowering) stays
// in DataProcessor::processSheet, where the parsed data lives.
class SchemaResolver {
public:
    enum class Source { Manifest, Standard, Inference };

    struct Resolution {
        TemplateSchema           schema;
        Source                   source = Source::Standard;
        ColumnResolution         columnResolution = ColumnResolution::Positional;
        bool                     perRowRegime = false;
        // Standard source only: the layout each block's landmark sniff chose
        // (drives header-variant swaps + the provenance headerCells map).
        QVector<HeaderLayout>    blockLayouts;
    };

    // manifest may be null (workbook has no _dve_schema sheet).
    static Resolution resolve(const Manifest::ParseResult* manifest,
                              const QString& sheetName,
                              const QVector<QVector<QVariant>>& cells,
                              const QString& templateVersion);
};

}} // namespace DVE::model
