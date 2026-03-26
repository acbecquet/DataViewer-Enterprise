#pragma once

#include <QString>
#include <QStringList>
#include <QMap>
#include <QVector>
#include <QRectF>

namespace DVE {

// Data-entry / table order (matches the physical evaluation sheet)
static const QStringList kSensoryMetrics = {
    "Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness", "Overall Liking"
};

// Plot order — Overall Liking at the top (first axis = 12-o'clock)
static const QStringList kSensoryMetricsPlot = {
    "Overall Liking", "Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness"
};

struct SensorySample {
    QString name;
    QMap<QString, int> scores;   // metric → 1–9, default 5
    QString comments;
};

struct SensorySession {
    QString  sessionName;
    QString  testTitle;
    QString  assessorName;
    QString  testerName;
    QString  media;
    QString  puffLength;
    QString  burnStatus;
    QString  clogStatus;
    QString  leakStatus;
    QString  date;
    QVector<SensorySample> samples;
    QString  timestamp;       // ISO8601
    QString  sourceImagePath; // path of scanned form image (may be empty)

    // Images linked to this session (same pattern as TPM SampleResult)
    QStringList     imagePaths;
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
};

} // namespace DVE
