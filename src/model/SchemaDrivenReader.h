#pragma once
#include "MetricSample.h"
#include "TemplateSchema.h"

namespace DVE { namespace model {

struct Sheet {
    QString         sheetName;
    TemplateSchema  schema;
    bool            perRowRegime = false;
    QVector<Sample> samples;
};

struct File {
    QString        filePath;
    QString        fileName;
    QString        templateVersion;   // legacy "new"/"old" tag, adapter input
    QVector<Sheet> sheets;
};

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
                              bool perRowRegime);
    // lowercase, [a-z0-9] only - the name-matching normalizer.
    static QString normalizeHeader(const QString& s);
};

}} // namespace DVE::model
