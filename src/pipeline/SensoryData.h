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
    QMap<QString, double> scores;   // metric → 1.0–9.0, default 5.0
    QString comments;

    // Per-sample device properties
    double  voltage    = 0.0;
    double  resistance = 0.0;
    double  power      = 0.0;       // calculated: V²/(R + Roffset)
    QString heatingTechnology;
};

struct SensorySession {
    QString  sessionName;
    QString  testTitle;
    QString  assessorName;
    QString  testerName;
    QString  media;
    QString  date;
    QVector<SensorySample> samples;
    QString  timestamp;       // ISO8601
    QString  sourceImagePath; // path of scanned form image (may be empty)

    // Session-level test properties
    QString  control;            // control sample name
    bool     isBlind = false;    // blind test?
    QString  primaryDifferences; // what is being tested

    // Legacy fields kept for backward compatibility with old JSON/DB
    QString  puffLength;
    QString  burnStatus;
    QString  clogStatus;
    QString  leakStatus;
    double   resistance = 0.0;
    double   voltage    = 0.0;
    double   power      = 0.0;
    QString  heatingTechnology;

    // Images linked to this session (same pattern as TPM SampleResult)
    QStringList     imagePaths;
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;

    // Persistence anchors (added for the report-preview feature, Phase 1A Task 4)
    int     id = -1;             // -1 if not yet persisted; populated by DB loaders
    QString sourceFilePath;      // the .xlsx file the session was loaded from (empty if DB-only)
    QString excelLayoutJson;     // dve_layout custom property pulled by ExcelReader
                                 // (empty until Phase 2 lands)
};

} // namespace DVE
