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

inline const QVector<ChoiceOption> kMouthpieceOptions = {
    {1, "1 - Very easy pull. Good design overall"},
    {2, "2 - Draw resistance fine, mouthpiece needs improvement"},
    {3, "3 - Mouthpiece fine, draw resistance too small"},
    {4, "4 - Mouthpiece fine, draw resistance too high"},
    {5, "5 - Mouthpiece and draw resistance made it very hard to puff"}
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

// Map a raw score to 1-9 for radar chart display (INVERTED).
// 1 (best) maps to 9 (outermost), maxScore (worst) maps to 1 (innermost).
// This way a good result fills the plot area.
inline double normalizeToRadar(const QString& metric, double raw)
{
    int maxScore = kDetailedMetricMaxScore.value(metric, 9);
    if (maxScore <= 1) return 9.0;
    // First normalize to 1-9 range
    double norm = (maxScore == 9) ? raw : 1.0 + (raw - 1.0) * 8.0 / (maxScore - 1.0);
    // Invert: 1 -> 9, 9 -> 1
    return 10.0 - norm;
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
    // C3: server-assigned image-row identities, parallel to imagePaths.
    QVector<qint64> imageIds;
    QVector<int>    imageVersions;

    // Optimistic-concurrency anchors (matches FileResult / SensorySession).
    // -1 / 0 means "not yet persisted" — tryWriteDetailedSensorySession will
    // INSERT rather than UPDATE.
    int id      = -1;
    int version = 0;
};

} // namespace DVE

class QJsonObject;

namespace DVE {

// Canonical JSON contract for DetailedSensorySession. All persistence
// paths (Postgres JSONB column, offline SQLite snapshot, and the Plan C
// recovery snapshot) route through these two helpers. Keep the field set
// + key names symmetric with the DatabaseManager / OfflineSnapshot copies:
// any new DetailedSensorySample / DetailedSensorySession field must be
// added to BOTH functions and to the round-trip test in
// tst_detailedsensoryjson. (Image vectors + id/version are intentionally
// NOT serialized here, matching the historical DB helpers.)
QJsonObject detailedSensorySessionToJson(const DetailedSensorySession& s);
DetailedSensorySession detailedSensorySessionFromJson(const QJsonObject& obj);

// True when the session is a freshly-created placeholder that the user has
// not meaningfully filled in yet. A session is a placeholder only when its
// name is empty/default ("New Session"), EVERY user-editable session-level
// field is at its fresh-session default (header fields + the S2-1 extras:
// facilitator name/comment, clog oil level, mouthpiece notes, viscosity,
// device return date, oilSmellLiking == 3, clog == false), and ALL samples
// are empty. date/timestamp are auto-populated by newSession() and so are
// excluded from the check. Plan C uses this to avoid writing empty
// "New Session" placeholders into recovery snapshots.
bool isPlaceholderSession(const DetailedSensorySession& s);

} // namespace DVE
