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
