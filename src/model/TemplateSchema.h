#pragma once
#include "MetricDef.h"
#include <QVector>

namespace DVE { namespace model {

struct TemplateSchema {
    QString schemaId;             // "standard"
    int     version = 1;

    // Block geometry, 1-based template rows (physical sheet coordinates).
    int headerRows      = 3;
    int columnHeaderRow = 4;
    int dataStartRow    = 5;
    int blockCols       = 12;

    QVector<HeaderFieldDef> headerFields;
    QVector<MetricDef>      columns;      // order == physical column order
    QVector<AggregateDef>   aggregates;

    const MetricDef*      column(const QString& key) const;
    int                   columnPos(const QString& key) const;  // 0-based, -1 absent
    const HeaderFieldDef* headerField(const QString& key) const;
};

}} // namespace DVE::model
