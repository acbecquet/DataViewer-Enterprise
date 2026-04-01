#pragma once

#include <QString>
#include <QStringList>
#include <QMap>
#include <QVector>
#include <QRectF>

namespace DVE {

// ── Vapor Quality metrics (6 axes) ──────────────────────────────────────────
// Displayed on left radar chart. Stored at original scale.
inline const QStringList kDetailedVaporQualityMetrics = {
    "Burn Taste",
    "Flavor Intensity",
    "Throat Irritation",
    "Nasal Irritation",
    "Vapor Quality Overall",
    "Cough"
};

// ── Consistency metrics (5 axes) ────────────────────────────────────────────
// Displayed on right radar chart. Stored at original scale.
inline const QStringList kDetailedConsistencyMetrics = {
    "Volume Consistency",
    "Performance Consistency",
    "Vapor Temperature",
    "Vapor vs Oil",
    "Vapor Volume"
};

// ── All metrics in data-entry order (for tables/export) ─────────────────────
inline const QStringList kDetailedAllMetrics = {
    "Burn Taste", "Flavor Intensity", "Throat Irritation",
    "Nasal Irritation", "Vapor Quality Overall", "Cough",
    "Volume Consistency", "Performance Consistency",
    "Vapor Temperature", "Vapor vs Oil", "Vapor Volume"
};

// ── Radar axis labels with scale annotations ────────────────────────────────
inline const QMap<QString, QString> kDetailedAxisLabels = {
    {"Burn Taste",              "Burn Taste\n(1=none, 9=too much)"},
    {"Flavor Intensity",        "Flavor Intensity\n(5=ideal, 1-9)"},
    {"Throat Irritation",       "Throat Irritation\n(1=none, 9=very bad)"},
    {"Nasal Irritation",        "Nasal Irritation\n(1=none, 9=worst)"},
    {"Vapor Quality Overall",   "Vapor Quality Overall\n(9=best)"},
    {"Cough",                   "Cough\n(1=none, 4=worst)"},
    {"Volume Consistency",      "Volume Consistency\n(1=best, 4=worst)"},
    {"Performance Consistency", "Performance Consistency\n(1=best, 3=worst)"},
    {"Vapor Temperature",       "Vapor Temperature\n(1=good, 4=worst)"},
    {"Vapor vs Oil",            "Vapor vs Oil\n(1=true, 4=major issue)"},
    {"Vapor Volume",            "Vapor Volume\n(3=ok, 1-5)"}
};

// ── Multiple-choice option text (for QComboBox dropdowns) ───────────────────
struct ChoiceOption { int value; QString text; };

inline const QVector<ChoiceOption> kCoughOptions = {
    {1, "1 - No"},
    {2, "2 - Yes, but doesn't bother me"},
    {3, "3 - Yes, will avoid buying next time"},
    {4, "4 - Yes, will stop immediately"}
};

inline const QVector<ChoiceOption> kVaporVsOilOptions = {
    {1, "1 - Very true to original oil"},
    {2, "2 - Main features but not everything"},
    {3, "3 - Minor off taste/flavor issue"},
    {4, "4 - Major off taste/flavor issue"}
};

inline const QVector<ChoiceOption> kVolumeConsistencyOptions = {
    {1, "1 - Very consistent"},
    {2, "2 - Mostly yes, increased a little"},
    {3, "3 - No, increased obviously"},
    {4, "4 - No, jumped everywhere"}
};

inline const QVector<ChoiceOption> kVaporVolumeOptions = {
    {1, "1 - Too little"},
    {2, "2 - A bit little"},
    {3, "3 - Ok"},
    {4, "4 - Big"},
    {5, "5 - Very big"}
};

inline const QVector<ChoiceOption> kVaporTemperatureOptions = {
    {1, "1 - Always good range"},
    {2, "2 - Got hot after puffs"},
    {3, "3 - Always too hot"},
    {4, "4 - Too cold"}
};

inline const QVector<ChoiceOption> kPerformanceConsistencyOptions = {
    {1, "1 - Very consistent"},
    {2, "2 - Mostly yes, small changes"},
    {3, "3 - No, not stable"}
};

// ── Max score per metric (for normalization to 1-9) ─────────────────────────
inline const QMap<QString, int> kDetailedMetricMaxScore = {
    {"Burn Taste", 9}, {"Flavor Intensity", 9},
    {"Throat Irritation", 9}, {"Nasal Irritation", 9},
    {"Vapor Quality Overall", 9}, {"Cough", 4},
    {"Volume Consistency", 4}, {"Performance Consistency", 3},
    {"Vapor Temperature", 4}, {"Vapor vs Oil", 4},
    {"Vapor Volume", 5}
};

// Map a raw score to 1-9 for radar chart display.
// Linear: rawMin=1 -> 1, rawMax -> 9
inline double normalizeToRadar(const QString& metric, double raw)
{
    int maxScore = kDetailedMetricMaxScore.value(metric, 9);
    if (maxScore <= 1) return 1.0;
    if (maxScore == 9) return raw;  // already 1-9
    // Linear map: 1 -> 1, maxScore -> 9
    return 1.0 + (raw - 1.0) * 8.0 / (maxScore - 1.0);
}

// ── Data structs ────────────────────────────────────────────────────────────
struct DetailedSensorySample {
    QString name;
    QMap<QString, double> scores;   // metric key -> raw score
    QString comments;

    // Per-sample device properties
    double  voltage    = 0.0;
    double  resistance = 0.0;
    double  power      = 0.0;
    QString heatingTechnology;
};

struct DetailedSensorySession {
    QString sessionName;
    QString testTitle;
    QString assessorName;
    QString testerName;
    QString facilitatorName;
    QString facilitatorComment;
    QString media;
    QString date;
    QString timestamp;          // ISO8601

    QVector<DetailedSensorySample> samples;

    // Extra fields from S2-1 template
    int     oilSmellLiking = 3;  // 1-5
    bool    clog = false;
    QString clogOilLevel;
    QString mouthpieceNotes;     // mouthpiece/draw resistance notes
    QString deviceReturnDate;
    QString viscosity;

    // Images
    QStringList     imagePaths;
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
};

} // namespace DVE
