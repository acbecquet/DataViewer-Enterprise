#pragma once
#include "MetricSample.h"
#include "TemplateSchema.h"

namespace DVE { namespace model {

struct Sheet {
    QString         sheetName;
    TemplateSchema  schema;
    bool            perRowRegime = false;
    // Resolved block-relative physical slot per schema column:
    // schema.columns[i] was read from physical column columnSlots[i] of each
    // block (identity under Positional). Filled by parseSheet from block 0's
    // resolution - block 0 is authoritative; per-block permutation divergence
    // is out of scope (Phase 2c Non-goals: blocks are uniform in every known
    // and generated workbook, and the round-trip harness would flag a
    // violator). Empty when the grid had no blocks to resolve - consumers
    // (LegacyAdapter::lowerSchemaSheet) treat empty as identity.
    QVector<int>    columnSlots;
    QVector<Sample> samples;
};

struct File {
    QString        filePath;
    QString        fileName;
    QString        templateVersion;   // legacy "new"/"old" tag, adapter input
    QVector<Sheet> sheets;
};

// How parseSheet maps each schema metric onto a physical grid column.
//   NameFirst - match by header text (aliases included) first, fall back to
//               positional (schema index == default slot) for the rest. This is
//               the manifest-era behavior: it tracks columns that have been
//               reordered or renamed in the sheet.
//   Positional - schema metric i always reads physical column off+i, ignoring
//               header text entirely. This exactly reproduces the legacy
//               ExcelReader::extractRow (which reads a fixed 12-wide block by
//               position and never consults header names to place data). The
//               Phase-1 strangler MUST use this to stay byte-identical to the
//               legacy pipeline on the historical layouts (e.g. the PV1-5
//               "Cart" sheets and the 13-wide misaligned "Comparison Test"
//               sheet, where name-matching would shuffle data into the wrong
//               metric slots).
enum class ColumnResolution { NameFirst, Positional };

// Parses one worksheet grid (0-based [row][col] QVariant cells, exactly as
// ExcelReader stores them) into a model::Sheet using a TemplateSchema.
// Pure function of its inputs - no I/O, fully unit-testable.
// `perRowRegime` is a downstream pass-through tag stored on the returned Sheet
// (Sheet::perRowRegime) - it does not affect parsing here. It must agree with
// the schema variant the caller already chose (i.e. standardV1(perRowRegime));
// nothing in parseSheet cross-checks the flag against `schema`.
class SchemaDrivenReader {
public:
    static Sheet   parseSheet(const QVector<QVector<QVariant>>& cells,
                              const QString& sheetName,
                              const TemplateSchema& schema,
                              bool perRowRegime,
                              ColumnResolution resolution = ColumnResolution::NameFirst);
    // lowercase, [a-z0-9] only - the name-matching normalizer.
    static QString normalizeHeader(const QString& s);
};

}} // namespace DVE::model
