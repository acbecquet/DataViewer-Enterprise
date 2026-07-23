#pragma once
#include "MetricDef.h"
#include <QMap>
#include <QVariant>
#include <QVector>

namespace DVE { namespace model {

struct MetricSeries {
    QString           key;
    QVector<QVariant> values;
};

// The owner's canonical shape: headers apply to all data equally; data is a
// set of free-standing named series.
struct Sample {
    QMap<QString, QVariant> headers;
    QVector<MetricSeries>   data;      // schema column order
    qint64                  id = -1;
    int                     version = 0;

    const MetricSeries* series(const QString& key) const {
        for (const MetricSeries& m : data) if (m.key == key) return &m;
        return nullptr;
    }
    int rowCount() const {
        int n = 0;
        for (const MetricSeries& m : data) n = qMax(n, int(m.values.size()));
        return n;
    }
};

}} // namespace DVE::model
