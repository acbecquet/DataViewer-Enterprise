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
    int     id      = -1;        // -1 if not yet persisted; populated by DB loaders
    int     version = 0;         // server-assigned row version; 0 if unknown
    QString sourceFilePath;      // the .xlsx file the session was loaded from (empty if DB-only)
    QString excelLayoutJson;     // dve_layout custom property pulled by ExcelReader
                                 // (empty until Phase 2 lands)
};

inline bool isPlaceholderSession(const SensorySession& s)
{
    // A session is a "placeholder" -- not yet meaningfully filled in --
    // when ALL of the following hold:
    //   1. It has never been persisted (id < 0). Once it has an id,
    //      subsequent saves must always go through so the user can
    //      legitimately blank a row out without it silently rolling
    //      back via NOTIFY.
    //   2. It still has the default new-session name.
    //   3. No header / session-level field has user content.
    //   4. No sample card has user content (name, comments, V/R,
    //      heating tech, image, or a non-default score).
    //
    // Without this gate, the 5-second auto-save persists the freshly-
    // created session to Postgres, which NOTIFY-broadcasts it to every
    // other client's navigator as a spurious "New Session" entry.
    if (s.id >= 0) return false;
    if (s.sessionName != QStringLiteral("New Session")) return false;

    const bool noHeaderContent = s.testerName.isEmpty()
                              && s.assessorName.isEmpty()
                              && s.media.isEmpty()
                              && s.testTitle.isEmpty()
                              && s.control.isEmpty()
                              && s.primaryDifferences.isEmpty()
                              && s.imagePaths.isEmpty();
    if (!noHeaderContent) return false;

    for (const SensorySample& sample : s.samples) {
        if (!sample.name.isEmpty())              return false;
        if (!sample.comments.isEmpty())          return false;
        if (sample.voltage    > 0.0)             return false;
        if (sample.resistance > 0.0)             return false;
        if (!sample.heatingTechnology.isEmpty()) return false;
        for (auto it = sample.scores.constBegin();
             it != sample.scores.constEnd(); ++it) {
            if (it.value() != 5.0) return false;
        }
    }
    return true;
}

} // namespace DVE
