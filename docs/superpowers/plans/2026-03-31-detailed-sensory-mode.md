# Detailed Sensory Mode Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Detailed Sensory" mode implementing the S2-1 user test template with dual radar charts (Vapor Quality + Consistency), per-sample navigation, database storage, Excel export, and PowerPoint report generation.

**Architecture:** New `DetailedSensoryPanel` as a parallel sibling to `SensoryPanel`, with its own data structures (`DetailedSensoryData.h`), separate database table (`detailed_sensory_sessions`), and extended `RadarChartWidget` supporting configurable axes. Three-way mode switch (TPM / Sensory / Detailed Sensory) in MainWindow.

**Tech Stack:** Qt 6 (C++ Widgets), SQLite, QXlsx, custom PptxWriter

---

## File Map

| Action | File | Responsibility |
|--------|------|----------------|
| Create | `src/pipeline/DetailedSensoryData.h` | Data structs, metric definitions, normalization |
| Create | `src/ui/DetailedSensoryPanel.h` | Panel class declaration |
| Create | `src/ui/DetailedSensoryPanel.cpp` | Panel UI, session mgmt, save/load, report gen |
| Create | `resources/images/ccell_icon_black.png` | Black variant of ccell icon for Tools tab |
| Modify | `src/ui/RadarChartWidget.h` | Add configurable axes constructor/method |
| Modify | `src/ui/RadarChartWidget.cpp` | Support custom axis lists + label annotations |
| Modify | `src/database/DatabaseManager.h` | Add DetailedSensory CRUD method declarations |
| Modify | `src/database/DatabaseManager.cpp` | New table schema + CRUD implementations |
| Modify | `src/MainWindow.h` | New member variables for detailed sensory mode |
| Modify | `src/MainWindow.cpp` | Three-way toggle, ribbon wiring, init panel |
| Modify | `src/ui/DatabaseBrowserDialog.h` | New tab member variables |
| Modify | `src/ui/DatabaseBrowserDialog.cpp` | "Detailed Sensory" tab |
| Modify | `DataViewerEnterprise.pro` | Add new source/header files |

---

### Task 1: Data Structures — `DetailedSensoryData.h`

**Files:**
- Create: `src/pipeline/DetailedSensoryData.h`

- [ ] **Step 1: Create the data header file**

```cpp
#pragma once

#include <QString>
#include <QStringList>
#include <QMap>
#include <QVector>
#include <QRectF>

namespace DVE {

// ── Vapor Quality metrics (6 axes) ──────────────────────────────────────────
// Displayed on left radar chart. Stored at original scale.
static const QStringList kDetailedVaporQualityMetrics = {
    "Burn Taste",
    "Flavor Intensity",
    "Throat Irritation",
    "Nasal Irritation",
    "Vapor Quality Overall",
    "Cough"
};

// ── Consistency metrics (5 axes) ────────────────────────────────────────────
// Displayed on right radar chart. Stored at original scale.
static const QStringList kDetailedConsistencyMetrics = {
    "Volume Consistency",
    "Performance Consistency",
    "Vapor Temperature",
    "Vapor vs Oil",
    "Vapor Volume"
};

// ── All metrics in data-entry order (for tables/export) ─────────────────────
static const QStringList kDetailedAllMetrics = {
    "Burn Taste", "Flavor Intensity", "Throat Irritation",
    "Nasal Irritation", "Vapor Quality Overall", "Cough",
    "Volume Consistency", "Performance Consistency",
    "Vapor Temperature", "Vapor vs Oil", "Vapor Volume"
};

// ── Radar axis labels with scale annotations ────────────────────────────────
static const QMap<QString, QString> kDetailedAxisLabels = {
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

static const QVector<ChoiceOption> kCoughOptions = {
    {1, "1 - No"},
    {2, "2 - Yes, but doesn't bother me"},
    {3, "3 - Yes, will avoid buying next time"},
    {4, "4 - Yes, will stop immediately"}
};

static const QVector<ChoiceOption> kVaporVsOilOptions = {
    {1, "1 - Very true to original oil"},
    {2, "2 - Main features but not everything"},
    {3, "3 - Minor off taste/flavor issue"},
    {4, "4 - Major off taste/flavor issue"}
};

static const QVector<ChoiceOption> kVolumeConsistencyOptions = {
    {1, "1 - Very consistent"},
    {2, "2 - Mostly yes, increased a little"},
    {3, "3 - No, increased obviously"},
    {4, "4 - No, jumped everywhere"}
};

static const QVector<ChoiceOption> kVaporVolumeOptions = {
    {1, "1 - Too little"},
    {2, "2 - A bit little"},
    {3, "3 - Ok"},
    {4, "4 - Big"},
    {5, "5 - Very big"}
};

static const QVector<ChoiceOption> kVaporTemperatureOptions = {
    {1, "1 - Always good range"},
    {2, "2 - Got hot after puffs"},
    {3, "3 - Always too hot"},
    {4, "4 - Too cold"}
};

static const QVector<ChoiceOption> kPerformanceConsistencyOptions = {
    {1, "1 - Very consistent"},
    {2, "2 - Mostly yes, small changes"},
    {3, "3 - No, not stable"}
};

// ── Max score per metric (for normalization to 1-9) ─────────────────────────
static const QMap<QString, int> kDetailedMetricMaxScore = {
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
    QString deviceReturnDate;
    QString viscosity;

    // Images
    QStringList     imagePaths;
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
};

} // namespace DVE
```

- [ ] **Step 2: Add to `.pro` file**

In `DataViewerEnterprise.pro`, add to the HEADERS section after `src/pipeline/SensoryData.h`:

```
    src/pipeline/DetailedSensoryData.h
```

- [ ] **Step 3: Verify build**

Run: `cd build && qmake ../DataViewerEnterprise.pro && mingw32-make -j8 2>&1 | tail -5`
Expected: Build succeeds (new header is include-only, no link errors)

- [ ] **Step 4: Commit**

```bash
git add src/pipeline/DetailedSensoryData.h DataViewerEnterprise.pro
git commit -m "feat: add DetailedSensoryData.h data structures for S2-1 template"
```

---

### Task 2: Extend RadarChartWidget for Configurable Axes

**Files:**
- Modify: `src/ui/RadarChartWidget.h`
- Modify: `src/ui/RadarChartWidget.cpp`

The current RadarChartWidget is hardcoded to use `kSensoryMetricsPlot` (5 axes) and `SensorySession`/`SensorySample`. We need to add a mode where it accepts custom axis names, custom label annotations, and generic score data — without breaking the existing sensory mode usage.

- [ ] **Step 1: Add configurable axes API to header**

In `src/ui/RadarChartWidget.h`, add these includes and members:

After line 5 (`#include <QSet>`), add:
```cpp
#include <QMap>
```

After line 19 (`void setReportCropTop(int pixels) { m_reportCropTop = pixels; }`), add:
```cpp

    // ── Configurable-axes mode (for DetailedSensoryPanel) ───────────────
    // When set, overrides kSensoryMetricsPlot with custom axis names/labels.
    void setCustomAxes(const QStringList& metricKeys,
                       const QMap<QString, QString>& axisLabels);
    void clearCustomAxes();

    // Set data as raw score maps (key->value) instead of SensorySession.
    // Each entry: sample name -> metric scores map.
    struct SampleData { QString name; QMap<QString, double> scores; };
    void setCustomData(const QVector<SampleData>& samples);
```

After line 37 (`int  m_reportCropTop = 0;`), add:
```cpp

    // Custom axes (empty = use default kSensoryMetricsPlot)
    QStringList m_customMetrics;
    QMap<QString, QString> m_customLabels;  // metric key -> display label
    QVector<SampleData> m_customSamples;
    bool m_useCustomAxes = false;

    // Helpers for custom/default mode
    int axisCount() const;
    QStringList axisMetrics() const;
    QString axisLabel(int i) const;
    double sampleScore(const SampleData& sd, int axisIdx) const;
```

- [ ] **Step 2: Implement configurable axes in .cpp**

In `src/ui/RadarChartWidget.cpp`, add after the constructor (after line 30):

```cpp
void RadarChartWidget::setCustomAxes(const QStringList& metricKeys,
                                     const QMap<QString, QString>& axisLabels)
{
    m_customMetrics = metricKeys;
    m_customLabels  = axisLabels;
    m_useCustomAxes = true;
    update();
}

void RadarChartWidget::clearCustomAxes()
{
    m_customMetrics.clear();
    m_customLabels.clear();
    m_customSamples.clear();
    m_useCustomAxes = false;
    update();
}

void RadarChartWidget::setCustomData(const QVector<SampleData>& samples)
{
    m_customSamples = samples;
    // Prune hidden indices
    QSet<int> pruned;
    for (int idx : m_hiddenSamples)
        if (idx < samples.size()) pruned.insert(idx);
    m_hiddenSamples = pruned;
    update();
}

int RadarChartWidget::axisCount() const
{
    return m_useCustomAxes ? m_customMetrics.size() : kSensoryMetricsPlot.size();
}

QStringList RadarChartWidget::axisMetrics() const
{
    return m_useCustomAxes ? m_customMetrics : kSensoryMetricsPlot;
}

QString RadarChartWidget::axisLabel(int i) const
{
    QStringList metrics = axisMetrics();
    if (i < 0 || i >= metrics.size()) return {};
    if (m_useCustomAxes)
        return m_customLabels.value(metrics[i], metrics[i]);
    return metrics[i];
}

double RadarChartWidget::sampleScore(const SampleData& sd, int axisIdx) const
{
    QStringList metrics = axisMetrics();
    if (axisIdx < 0 || axisIdx >= metrics.size()) return 5.0;
    return sd.scores.value(metrics[axisIdx], 5.0);
}
```

- [ ] **Step 3: Refactor `axisPoint()` to use `axisCount()`**

Replace the `axisPoint` method (lines 49-59) with:

```cpp
QPointF RadarChartWidget::axisPoint(int axisIndex, double value,
                                    QPointF center, double radius) const
{
    int n = axisCount();
    double angleDeg = 270.0 + (360.0 / n) * axisIndex;
    double angleRad = qDegreesToRadians(angleDeg);
    double t = (value - 1.0) / 8.0;
    double r = t * radius;
    return QPointF(center.x() + r * qCos(angleRad),
                   center.y() + r * qSin(angleRad));
}
```

- [ ] **Step 4: Refactor `drawGrid()` to use `axisCount()`**

Replace `int n = kSensoryMetricsPlot.size();` at line 67 with:
```cpp
    int n = axisCount();
```

- [ ] **Step 5: Refactor `drawAxes()` to use `axisCount()` and `axisLabel()`**

Replace the `drawAxes` method (lines 95-125) with:

```cpp
void RadarChartWidget::drawAxes(QPainter& p, QPointF center, double radius) const
{
    int n = axisCount();

    QFont labelFont = p.font();
    labelFont.setPointSize(m_useCustomAxes ? 7 : 8);
    p.setFont(labelFont);

    for (int i = 0; i < n; ++i) {
        QPointF tip = axisPoint(i, 9, center, radius);

        // Spoke
        p.setPen(QPen(QColor(150, 150, 150), 1));
        p.drawLine(center, tip);

        // Label
        double angleDeg = 270.0 + (360.0 / n) * i;
        double angleRad = qDegreesToRadians(angleDeg);
        double labelDist = radius + 18;
        QPointF labelCenter(center.x() + labelDist * qCos(angleRad),
                            center.y() + labelDist * qSin(angleRad));

        QFontMetrics fm(labelFont);
        QString label = axisLabel(i);
        // For multi-line labels, use boundingRect with newlines
        QRect textRect = fm.boundingRect(QRect(0, 0, 200, 200),
                                         Qt::AlignCenter | Qt::TextWordWrap, label);
        textRect.moveCenter(labelCenter.toPoint());

        p.setPen(QColor(40, 40, 40));
        p.drawText(textRect, Qt::AlignCenter | Qt::TextWordWrap, label);
    }
}
```

- [ ] **Step 6: Refactor `drawSample()` to use `axisCount()`**

Replace the `drawSample` method (lines 127-142) with:

```cpp
void RadarChartWidget::drawSample(QPainter& p, const SensorySample& sample,
                                  QPointF center, double radius, QColor color) const
{
    int n = axisCount();
    QStringList metrics = axisMetrics();
    QPolygonF poly;
    for (int i = 0; i < n; ++i) {
        double score = sample.scores.value(metrics[i], 5.0);
        poly << axisPoint(i, score, center, radius);
    }

    QColor fillColor = color;
    fillColor.setAlphaF(0.18);
    p.setBrush(fillColor);
    p.setPen(QPen(color, 2));
    p.drawPolygon(poly);
}
```

- [ ] **Step 7: Add `drawCustomSample()` helper**

Add after the `drawSample` method:

```cpp
void RadarChartWidget::drawCustomSample(QPainter& p, const SampleData& sample,
                                        QPointF center, double radius, QColor color) const
{
    int n = axisCount();
    QPolygonF poly;
    for (int i = 0; i < n; ++i) {
        double score = sampleScore(sample, i);
        poly << axisPoint(i, score, center, radius);
    }

    QColor fillColor = color;
    fillColor.setAlphaF(0.18);
    p.setBrush(fillColor);
    p.setPen(QPen(color, 2));
    p.drawPolygon(poly);
}
```

Also add the declaration to the header, in the private section after `drawLegendReport`:
```cpp
    void drawCustomSample(QPainter& p, const SampleData& sample,
                          QPointF center, double radius, QColor color) const;
```

- [ ] **Step 8: Update `paintEvent()` to handle custom data mode**

In the `paintEvent` method, wherever samples are drawn (the nested `for` loops iterating `m_sessions`), add an alternative path for custom data. Replace the drawing loops in both report and normal modes.

In **normal UI mode** (the `else` branch starting ~line 342), replace the sample drawing loop (lines 358-365):

```cpp
        if (m_useCustomAxes && !m_customSamples.isEmpty()) {
            for (int ci = 0; ci < m_customSamples.size(); ++ci) {
                if (!m_hiddenSamples.contains(ci))
                    drawCustomSample(p, m_customSamples[ci], center, radius,
                                     kColors[ci % kColors.size()]);
            }
        } else {
            int colorIdx = 0;
            for (const SensorySession& sess : m_sessions) {
                for (const SensorySample& sample : sess.samples) {
                    if (!m_hiddenSamples.contains(colorIdx))
                        drawSample(p, sample, center, radius, kColors[colorIdx % kColors.size()]);
                    ++colorIdx;
                }
            }
        }
```

In **report mode** (the `if (m_reportMode)` branch starting ~line 310), replace the sample drawing loop (lines 328-335) with the same pattern:

```cpp
        if (m_useCustomAxes && !m_customSamples.isEmpty()) {
            for (int ci = 0; ci < m_customSamples.size(); ++ci) {
                if (!m_hiddenSamples.contains(ci))
                    drawCustomSample(p, m_customSamples[ci], center, radius,
                                     kColors[ci % kColors.size()]);
            }
        } else {
            int colorIdx = 0;
            for (const SensorySession& sess : m_sessions) {
                for (const SensorySample& sample : sess.samples) {
                    if (!m_hiddenSamples.contains(colorIdx))
                        drawSample(p, sample, center, radius, kColors[colorIdx % kColors.size()]);
                    ++colorIdx;
                }
            }
        }
```

- [ ] **Step 9: Update legend methods to handle custom data**

In `drawLegend()` (line 144), replace the entry-building loop (lines 151-160):

```cpp
    int colorIdx = 0;
    if (m_useCustomAxes && !m_customSamples.isEmpty()) {
        for (const SampleData& sd : m_customSamples) {
            entries.append({sd.name.isEmpty() ? QString("Sample %1").arg(colorIdx + 1)
                                              : sd.name,
                            kColors[colorIdx % kColors.size()],
                            colorIdx});
            ++colorIdx;
        }
    } else {
        for (const SensorySession& sess : m_sessions) {
            for (const SensorySample& sample : sess.samples) {
                entries.append({sample.name.isEmpty() ? QString("Sample %1").arg(colorIdx + 1)
                                                      : sample.name,
                                kColors[colorIdx % kColors.size()],
                                colorIdx});
                ++colorIdx;
            }
        }
    }
```

Apply the same change to `drawLegendReport()` (line 214), replacing lines 218-228 with the same pattern.

- [ ] **Step 10: Verify existing sensory mode still works**

Run: `cd build && qmake ../DataViewerEnterprise.pro && mingw32-make -j8 2>&1 | tail -10`
Expected: Build succeeds. Launch the app, toggle Sensory mode, verify radar chart renders correctly with existing data.

- [ ] **Step 11: Commit**

```bash
git add src/ui/RadarChartWidget.h src/ui/RadarChartWidget.cpp
git commit -m "feat: extend RadarChartWidget with configurable axes and custom data mode"
```

---

### Task 3: Database Schema + CRUD for Detailed Sensory

**Files:**
- Modify: `src/database/DatabaseManager.h`
- Modify: `src/database/DatabaseManager.cpp`

- [ ] **Step 1: Add record struct and method declarations to header**

In `src/database/DatabaseManager.h`, add after line 10 (`#include "../pipeline/SensoryData.h"`):
```cpp
#include "../pipeline/DetailedSensoryData.h"
```

After line 33 (closing `};` of `SensoryRecord`), add:
```cpp

struct DetailedSensoryRecord {
    int     id;
    QString sessionName;
    QString testTitle;
    QString assessorName;
    QString testerName;
    QString media;
    QString date;
    int     sampleCount;
};
```

After line 77 (`QString nextDefaultTestName() const;`), add:
```cpp

    // ── Detailed sensory sessions ───────────────────────────────────────────
    bool saveDetailedSensorySession(const DetailedSensorySession& s);
    QVector<DetailedSensorySession> loadDetailedSensorySessions() const;
    DetailedSensorySession loadDetailedSensorySession(int id) const;
    QVector<DetailedSensoryRecord> listDetailedSensoryRecords() const;
    bool removeDetailedSensorySession(int id);
```

- [ ] **Step 2: Add table creation in `initSchema()`**

In `src/database/DatabaseManager.cpp`, in the `initSchema()` method, after the `sensory_images` table creation (after line 221, `if (!ok) { ... return false; }`), add:

```cpp

    // ── detailed_sensory_sessions ───────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS detailed_sensory_sessions ("
        "  id            INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  session_name  TEXT,"
        "  tester_name   TEXT,"
        "  assessor_name TEXT,"
        "  media         TEXT,"
        "  date          TEXT,"
        "  timestamp     TEXT,"
        "  json_data     TEXT"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }

    // ── detailed_sensory_images ─────────────────────────────────────────────
    ok = q.exec(
        "CREATE TABLE IF NOT EXISTS detailed_sensory_images ("
        "  id               INTEGER PRIMARY KEY AUTOINCREMENT,"
        "  session_id       INTEGER NOT NULL REFERENCES detailed_sensory_sessions(id) ON DELETE CASCADE,"
        "  sort_order       INTEGER DEFAULT 0,"
        "  file_name        TEXT,"
        "  image_data       BLOB,"
        "  layout_x         REAL,"
        "  layout_y         REAL,"
        "  layout_w         REAL,"
        "  layout_h         REAL,"
        "  crop_x           REAL,"
        "  crop_y           REAL,"
        "  crop_w           REAL,"
        "  crop_h           REAL"
        ")"
    );
    if (!ok) { m_lastError = q.lastError().text(); return false; }
```

- [ ] **Step 3: Implement `saveDetailedSensorySession()`**

Add at the end of the file, before the closing `} // namespace DVE`:

```cpp
// ============================================================================
// Detailed Sensory Sessions
// ============================================================================

bool DatabaseManager::saveDetailedSensorySession(const DetailedSensorySession& s)
{
    if (!m_open) return false;

    QJsonObject root;
    root["session_name"]        = s.sessionName;
    root["test_title"]          = s.testTitle;
    root["assessor_name"]       = s.assessorName;
    root["tester_name"]         = s.testerName;
    root["facilitator_name"]    = s.facilitatorName;
    root["facilitator_comment"] = s.facilitatorComment;
    root["media"]               = s.media;
    root["date"]                = s.date;
    root["timestamp"]           = s.timestamp;
    root["oil_smell_liking"]    = s.oilSmellLiking;
    root["clog"]                = s.clog;
    root["clog_oil_level"]      = s.clogOilLevel;
    root["device_return_date"]  = s.deviceReturnDate;
    root["viscosity"]           = s.viscosity;

    QJsonArray samplesArr;
    for (const DetailedSensorySample& sample : s.samples) {
        QJsonObject sObj;
        sObj["name"]     = sample.name;
        sObj["comments"] = sample.comments;
        for (const QString& metric : kDetailedAllMetrics) {
            sObj[metric] = sample.scores.value(metric, 0.0);
        }
        sObj["voltage"]            = sample.voltage;
        sObj["resistance"]         = sample.resistance;
        sObj["power"]              = sample.power;
        sObj["heating_technology"] = sample.heatingTechnology;
        samplesArr.append(sObj);
    }
    root["samples"] = samplesArr;

    QString jsonStr = QString::fromUtf8(
        QJsonDocument(root).toJson(QJsonDocument::Compact));

    // Upsert: delete existing matching record
    {
        QSqlQuery findOld(m_db);
        findOld.prepare("SELECT id FROM detailed_sensory_sessions "
                        "WHERE session_name = ? AND tester_name = ? AND date = ?");
        findOld.addBindValue(s.sessionName);
        findOld.addBindValue(s.testerName);
        findOld.addBindValue(s.date);
        if (findOld.exec()) {
            while (findOld.next()) {
                int oldId = findOld.value(0).toInt();
                QSqlQuery delImg(m_db);
                delImg.prepare("DELETE FROM detailed_sensory_images WHERE session_id = ?");
                delImg.addBindValue(oldId);
                delImg.exec();
            }
        }

        QSqlQuery del(m_db);
        del.prepare("DELETE FROM detailed_sensory_sessions "
                    "WHERE session_name = ? AND tester_name = ? AND date = ?");
        del.addBindValue(s.sessionName);
        del.addBindValue(s.testerName);
        del.addBindValue(s.date);
        del.exec();
    }

    QSqlQuery q(m_db);
    q.prepare("INSERT INTO detailed_sensory_sessions "
              "(session_name, tester_name, assessor_name, media, date, timestamp, json_data) "
              "VALUES (?, ?, ?, ?, ?, ?, ?)");
    q.addBindValue(s.sessionName);
    q.addBindValue(s.testerName);
    q.addBindValue(s.assessorName);
    q.addBindValue(s.media);
    q.addBindValue(s.date);
    q.addBindValue(s.timestamp);
    q.addBindValue(jsonStr);

    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }

    // Save images
    int sessionId = q.lastInsertId().toInt();
    if (!s.imagePaths.isEmpty()) {
        QSqlQuery imgQ(m_db);
        imgQ.prepare("INSERT INTO detailed_sensory_images "
                     "(session_id, sort_order, file_name, image_data,"
                     " layout_x, layout_y, layout_w, layout_h,"
                     " crop_x, crop_y, crop_w, crop_h) "
                     "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
        for (int i = 0; i < s.imagePaths.size(); ++i) {
            QByteArray imgData;
            QFile imgFile(s.imagePaths[i]);
            if (imgFile.open(QIODevice::ReadOnly))
                imgData = imgFile.readAll();

            QRectF layout = (i < s.imageLayouts.size()) ? s.imageLayouts[i] : QRectF();
            QRectF crop   = (i < s.imageCrops.size())   ? s.imageCrops[i]   : QRectF(0,0,1,1);

            imgQ.addBindValue(sessionId);
            imgQ.addBindValue(i);
            imgQ.addBindValue(QFileInfo(s.imagePaths[i]).fileName());
            imgQ.addBindValue(imgData);
            imgQ.addBindValue(layout.x());
            imgQ.addBindValue(layout.y());
            imgQ.addBindValue(layout.width());
            imgQ.addBindValue(layout.height());
            imgQ.addBindValue(crop.x());
            imgQ.addBindValue(crop.y());
            imgQ.addBindValue(crop.width());
            imgQ.addBindValue(crop.height());
            imgQ.exec();
        }
    }

    return true;
}
```

- [ ] **Step 4: Implement `loadDetailedSensorySessions()`**

```cpp
QVector<DetailedSensorySession> DatabaseManager::loadDetailedSensorySessions() const
{
    QVector<DetailedSensorySession> result;
    if (!m_open) return result;

    QSqlQuery q("SELECT id, json_data FROM detailed_sensory_sessions ORDER BY id DESC", m_db);
    while (q.next()) {
        int sessId = q.value(0).toInt();
        QByteArray jsonBytes = q.value(1).toString().toUtf8();
        QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
        if (doc.isNull() || !doc.isObject()) continue;

        QJsonObject root = doc.object();
        DetailedSensorySession sess;
        sess.sessionName        = root["session_name"].toString();
        sess.testTitle          = root["test_title"].toString();
        sess.assessorName       = root["assessor_name"].toString();
        sess.testerName         = root["tester_name"].toString();
        sess.facilitatorName    = root["facilitator_name"].toString();
        sess.facilitatorComment = root["facilitator_comment"].toString();
        sess.media              = root["media"].toString();
        sess.date               = root["date"].toString();
        sess.timestamp          = root["timestamp"].toString();
        sess.oilSmellLiking     = root["oil_smell_liking"].toInt(3);
        sess.clog               = root["clog"].toBool(false);
        sess.clogOilLevel       = root["clog_oil_level"].toString();
        sess.deviceReturnDate   = root["device_return_date"].toString();
        sess.viscosity          = root["viscosity"].toString();

        for (const QJsonValue& sv : root["samples"].toArray()) {
            QJsonObject sObj = sv.toObject();
            DetailedSensorySample sample;
            sample.name              = sObj["name"].toString();
            sample.comments          = sObj["comments"].toString();
            sample.voltage           = sObj["voltage"].toDouble();
            sample.resistance        = sObj["resistance"].toDouble();
            sample.power             = sObj["power"].toDouble();
            sample.heatingTechnology = sObj["heating_technology"].toString();
            for (const QString& metric : kDetailedAllMetrics) {
                double maxVal = kDetailedMetricMaxScore.value(metric, 9);
                sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(1.0), maxVal);
            }
            sess.samples.append(sample);
        }

        // Load images
        QSqlQuery qi(m_db);
        qi.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
                   "crop_x, crop_y, crop_w, crop_h "
                   "FROM detailed_sensory_images WHERE session_id = ? ORDER BY sort_order");
        qi.addBindValue(sessId);
        if (qi.exec()) {
            int ii = 0;
            while (qi.next()) {
                QByteArray blob = qi.value(1).toByteArray();
                QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                             qi.value(4).toDouble(), qi.value(5).toDouble());
                QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                           qi.value(8).toDouble(), qi.value(9).toDouble());

                QString tempPath = QDir::temp().filePath(
                    QString("dve_detsensimg_%1_%2.png").arg(sessId).arg(ii++));
                QFile tmpFile(tempPath);
                if (tmpFile.open(QIODevice::WriteOnly)) {
                    tmpFile.write(blob);
                    tmpFile.close();
                }
                sess.imagePaths.append(tempPath);
                sess.imageLayouts.append(layout);
                sess.imageCrops.append(crop);
            }
        }

        result.append(sess);
    }
    return result;
}
```

- [ ] **Step 5: Implement `loadDetailedSensorySession(int id)`**

```cpp
DetailedSensorySession DatabaseManager::loadDetailedSensorySession(int id) const
{
    DetailedSensorySession sess;
    if (!m_open) return sess;

    QSqlQuery q(m_db);
    q.prepare("SELECT json_data FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec() || !q.next()) return sess;

    QByteArray jsonBytes = q.value(0).toString().toUtf8();
    QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
    if (doc.isNull() || !doc.isObject()) return sess;

    QJsonObject root = doc.object();
    sess.sessionName        = root["session_name"].toString();
    sess.testTitle          = root["test_title"].toString();
    sess.assessorName       = root["assessor_name"].toString();
    sess.testerName         = root["tester_name"].toString();
    sess.facilitatorName    = root["facilitator_name"].toString();
    sess.facilitatorComment = root["facilitator_comment"].toString();
    sess.media              = root["media"].toString();
    sess.date               = root["date"].toString();
    sess.timestamp          = root["timestamp"].toString();
    sess.oilSmellLiking     = root["oil_smell_liking"].toInt(3);
    sess.clog               = root["clog"].toBool(false);
    sess.clogOilLevel       = root["clog_oil_level"].toString();
    sess.deviceReturnDate   = root["device_return_date"].toString();
    sess.viscosity          = root["viscosity"].toString();

    for (const QJsonValue& sv : root["samples"].toArray()) {
        QJsonObject sObj = sv.toObject();
        DetailedSensorySample sample;
        sample.name              = sObj["name"].toString();
        sample.comments          = sObj["comments"].toString();
        sample.voltage           = sObj["voltage"].toDouble();
        sample.resistance        = sObj["resistance"].toDouble();
        sample.power             = sObj["power"].toDouble();
        sample.heatingTechnology = sObj["heating_technology"].toString();
        for (const QString& metric : kDetailedAllMetrics) {
            double maxVal = kDetailedMetricMaxScore.value(metric, 9);
            sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(1.0), maxVal);
        }
        sess.samples.append(sample);
    }

    // Load images
    QSqlQuery qi(m_db);
    qi.prepare("SELECT file_name, image_data, layout_x, layout_y, layout_w, layout_h, "
               "crop_x, crop_y, crop_w, crop_h "
               "FROM detailed_sensory_images WHERE session_id = ? ORDER BY sort_order");
    qi.addBindValue(id);
    if (qi.exec()) {
        int ii = 0;
        while (qi.next()) {
            QByteArray blob = qi.value(1).toByteArray();
            QRectF layout(qi.value(2).toDouble(), qi.value(3).toDouble(),
                         qi.value(4).toDouble(), qi.value(5).toDouble());
            QRectF crop(qi.value(6).toDouble(), qi.value(7).toDouble(),
                       qi.value(8).toDouble(), qi.value(9).toDouble());

            QString tempPath = QDir::temp().filePath(
                QString("dve_detsensimg_%1_%2.png").arg(id).arg(ii++));
            QFile tmpFile(tempPath);
            if (tmpFile.open(QIODevice::WriteOnly)) {
                tmpFile.write(blob);
                tmpFile.close();
            }
            sess.imagePaths.append(tempPath);
            sess.imageLayouts.append(layout);
            sess.imageCrops.append(crop);
        }
    }

    return sess;
}
```

- [ ] **Step 6: Implement `listDetailedSensoryRecords()` and `removeDetailedSensorySession()`**

```cpp
QVector<DetailedSensoryRecord> DatabaseManager::listDetailedSensoryRecords() const
{
    QVector<DetailedSensoryRecord> result;
    if (!m_open) return result;

    QSqlQuery q("SELECT id, session_name, assessor_name, media, date, json_data "
                "FROM detailed_sensory_sessions ORDER BY id DESC", m_db);
    while (q.next()) {
        DetailedSensoryRecord rec;
        rec.id           = q.value(0).toInt();
        rec.sessionName  = q.value(1).toString();
        rec.assessorName = q.value(2).toString();
        rec.media        = q.value(3).toString();
        rec.date         = q.value(4).toString();

        QByteArray jsonBytes = q.value(5).toString().toUtf8();
        QJsonDocument doc = QJsonDocument::fromJson(jsonBytes);
        if (!doc.isNull() && doc.isObject()) {
            QJsonObject root = doc.object();
            rec.testTitle    = root["test_title"].toString();
            rec.testerName   = root["tester_name"].toString();
            rec.sampleCount  = root["samples"].toArray().size();
        } else {
            rec.sampleCount = 0;
        }

        result.append(rec);
    }
    return result;
}

bool DatabaseManager::removeDetailedSensorySession(int id)
{
    if (!m_open) return false;
    QSqlQuery q(m_db);
    q.prepare("DELETE FROM detailed_sensory_sessions WHERE id = ?");
    q.addBindValue(id);
    if (!q.exec()) {
        m_lastError = q.lastError().text();
        return false;
    }
    return true;
}
```

- [ ] **Step 7: Add required includes to DatabaseManager.cpp**

At the top of `DatabaseManager.cpp`, verify these includes exist (add if missing):
```cpp
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QFileInfo>
#include <QDir>
#include <QFile>
```

- [ ] **Step 8: Verify build**

Run: `cd build && qmake ../DataViewerEnterprise.pro && mingw32-make -j8 2>&1 | tail -10`
Expected: Build succeeds.

- [ ] **Step 9: Commit**

```bash
git add src/database/DatabaseManager.h src/database/DatabaseManager.cpp
git commit -m "feat: add detailed_sensory_sessions database table and CRUD methods"
```

---

### Task 4: Create Black Icon

**Files:**
- Create: `resources/images/ccell_icon_black.png`

- [ ] **Step 1: Generate black variant of ccell icon**

Use Python/PIL to convert the existing blue ccell icon to all-black while preserving alpha:

```bash
python -c "
from PIL import Image
img = Image.open(r'resources/images/ccell_icon.png').convert('RGBA')
pixels = img.load()
for y in range(img.height):
    for x in range(img.width):
        r, g, b, a = pixels[x, y]
        if a > 0:
            pixels[x, y] = (30, 30, 30, a)
img.save(r'resources/images/ccell_icon_black.png')
print('Saved ccell_icon_black.png')
"
```

- [ ] **Step 2: Verify the icon file was created**

Run: `ls -la resources/images/ccell_icon_black.png`
Expected: File exists with reasonable size.

- [ ] **Step 3: Commit**

```bash
git add resources/images/ccell_icon_black.png
git commit -m "feat: add black ccell icon for Detailed Sensory mode button"
```

---

### Task 5: DetailedSensoryPanel — Header + Constructor + Basic UI Layout

**Files:**
- Create: `src/ui/DetailedSensoryPanel.h`
- Create: `src/ui/DetailedSensoryPanel.cpp`
- Modify: `DataViewerEnterprise.pro`

This is the largest task. The panel class mirrors `SensoryPanel` in structure but with:
- Paged sample navigation (prev/next) instead of flow-layout cards
- Dual radar charts side by side (bottom half)
- 11 questions in a compact form (top strip)

- [ ] **Step 1: Create the header file**

```cpp
#pragma once

#include <QWidget>
#include <QScrollArea>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QLineEdit>
#include <QDoubleSpinBox>
#include <QComboBox>
#include <QTextEdit>
#include <QLabel>
#include <QSplitter>
#include <QTimer>
#include <QVector>
#include <QPushButton>
#include <QStackedWidget>
#include <QTableWidget>

#include "pipeline/DetailedSensoryData.h"
#include "ui/RadarChartWidget.h"
#include "database/DatabaseManager.h"

namespace QXlsx { class Document; }

namespace DVE {

class DetailedSensoryPanel : public QWidget
{
    Q_OBJECT

public:
    explicit DetailedSensoryPanel(DatabaseManager* db, QWidget* parent = nullptr);

    // ── Session management ──────────────────────────────────────────────────
    void loadSessions(const QVector<DetailedSensorySession>& sessions);
    void selectSession(int index);
    void showAveragedChart(const QVector<int>& sessionIndices);
    void newSession();
    void closeSessions(const QVector<int>& indices);
    void renameSession(int index, const QString& newLabel);

    // ── File operations ─────────────────────────────────────────────────────
    void save();
    void loadFiles();
    void loadFromDatabase();

    // ── Report generation ───────────────────────────────────────────────────
    void generateFullReport();

    // ── Session access ──────────────────────────────────────────────────────
    QVector<DetailedSensorySession> allSessions();
    int  currentSessionIndex() const { return m_currentTesterIdx; }
    QString sessionLabel(const DetailedSensorySession& s) const;
    DetailedSensorySession* currentSession();

    // ── Averaged table overlay ──────────────────────────────────────────────
    void showAveragedTable(const QStringList& deviceNames,
                           const QVector<QMap<QString, double>>& deviceAvgs);
    void showNormalView();

    // ── Static combined PPTX ────────────────────────────────────────────────
    static bool generateCombinedPptx(const QVector<DetailedSensorySession>& sessions,
                                      const QString& filePath,
                                      QString& errorOut);

signals:
    void sessionsChanged();

private:
    // ── UI build ────────────────────────────────────────────────────────────
    void buildHeaderRow(QWidget* container);
    void buildQuestionForm();
    void buildSampleNavBar();

    // ── Sample management ───────────────────────────────────────────────────
    void displayCurrentSample();
    void saveCurrentSampleToSession();
    void onPrevSample();
    void onNextSample();
    void onAddSample();
    void onRemoveSample();
    void updateSampleNav();

    // ── Session management ──────────────────────────────────────────────────
    DetailedSensorySession buildSession() const;
    void           applySession(const DetailedSensorySession& session);
    void           saveCurrentTester();
    bool           isDefaultState() const;

    // ── Chart ───────────────────────────────────────────────────────────────
    void scheduleChartRefresh();
    void onRefreshChart();

    // ── Save helpers ────────────────────────────────────────────────────────
    void saveToExcel(const QString& path, const DetailedSensorySession& sess);

    // ── UI elements ─────────────────────────────────────────────────────────
    // Header row
    QLineEdit*        m_testTitleEdit;
    QLineEdit*        m_assessorEdit;
    QLineEdit*        m_testerEdit;
    QLineEdit*        m_mediaEdit;
    QLabel*           m_dateLabel;

    // Sample navigation
    QPushButton*      m_prevBtn;
    QPushButton*      m_nextBtn;
    QLabel*           m_sampleCountLabel;
    QPushButton*      m_addSampleBtn;
    QPushButton*      m_removeSampleBtn;

    // Question form (top strip)
    QWidget*          m_questionForm;
    QScrollArea*      m_questionScroll;
    // Scored questions (1-9 spinboxes)
    QMap<QString, QDoubleSpinBox*> m_spinBoxes;
    // Multiple-choice questions (comboboxes)
    QMap<QString, QComboBox*>      m_comboBoxes;
    QLineEdit*        m_sampleNameEdit;
    QTextEdit*        m_commentsEdit;

    // Dual radar charts (bottom half)
    RadarChartWidget* m_vaporQualityChart;
    RadarChartWidget* m_consistencyChart;

    // Averaged table overlay
    QStackedWidget*   m_topStack = nullptr;
    QTableWidget*     m_avgOverlayTable = nullptr;

    // ── Session data ────────────────────────────────────────────────────────
    QVector<DetailedSensorySession> m_sessions;
    int  m_currentTesterIdx  = -1;
    int  m_currentSampleIdx  = 0;

    // ── Debounce timer ──────────────────────────────────────────────────────
    QTimer* m_refreshTimer;

    // ── Save path ───────────────────────────────────────────────────────────
    QString m_savePath;

    // ── Database ────────────────────────────────────────────────────────────
    DatabaseManager* m_db;

    // ── Browse dir ──────────────────────────────────────────────────────────
    QString m_lastBrowseDir;
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
};

} // namespace DVE
```

- [ ] **Step 2: Create the implementation file with constructor and UI layout**

Create `src/ui/DetailedSensoryPanel.cpp`. This is a large file — the constructor builds the full layout:

```cpp
#include "DetailedSensoryPanel.h"

#include <QDate>
#include <QFileDialog>
#include <QMessageBox>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QHeaderView>
#include <QScrollBar>

#include "utils/AppTheme.h"
#include "reporting/PptxWriter.h"
#include "utils/ImageUtils.h"

// Re-use the same no-scroll spinbox from SensoryPanel
#include <QWheelEvent>
namespace {
class NoWheelDoubleSpinBox : public QDoubleSpinBox {
public:
    using QDoubleSpinBox::QDoubleSpinBox;
    void wheelEvent(QWheelEvent* e) override { e->ignore(); }
};
class NoWheelComboBox : public QComboBox {
public:
    using QComboBox::QComboBox;
    void wheelEvent(QWheelEvent* e) override { e->ignore(); }
};
} // anon

namespace DVE {

// Helper: calculate power from V, R, heating tech (same as SampleCard)
static double calcPower(double voltage, double resistance, const QString& tech)
{
    double rOffset = 0.0;
    if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
        rOffset = 0.78;
    else if (tech == "T51")
        rOffset = 0.25;
    double denom = resistance + rOffset;
    return (voltage > 0 && denom > 0) ? (voltage * voltage) / denom : 0.0;
}

DetailedSensoryPanel::DetailedSensoryPanel(DatabaseManager* db, QWidget* parent)
    : QWidget(parent), m_db(db)
{
    m_refreshTimer = new QTimer(this);
    m_refreshTimer->setSingleShot(true);
    m_refreshTimer->setInterval(150);
    connect(m_refreshTimer, &QTimer::timeout, this, &DetailedSensoryPanel::onRefreshChart);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(6, 2, 6, 2);
    mainLayout->setSpacing(4);

    // ── Header row (Test Title, Assessor, Tester, Media, Date) ──────────
    auto* headerWidget = new QWidget(this);
    buildHeaderRow(headerWidget);
    mainLayout->addWidget(headerWidget);

    // ── Sample navigation bar ───────────────────────────────────────────
    auto* navBar = new QWidget(this);
    buildSampleNavBar();
    mainLayout->addWidget(m_prevBtn->parentWidget()); // navBar is set up in buildSampleNavBar

    // ── Main splitter: top=questions, bottom=dual charts ────────────────
    auto* splitter = new QSplitter(Qt::Vertical, this);

    // -- Top: question form in scroll area --
    m_questionScroll = new QScrollArea(this);
    m_questionScroll->setWidgetResizable(true);
    m_questionScroll->setFrameShape(QFrame::NoFrame);
    buildQuestionForm();
    m_questionScroll->setWidget(m_questionForm);

    // Wrap in stacked widget for averaged table overlay
    m_topStack = new QStackedWidget(this);
    m_topStack->addWidget(m_questionScroll);   // index 0 = normal
    m_avgOverlayTable = new QTableWidget(this);
    m_avgOverlayTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    m_avgOverlayTable->setAlternatingRowColors(true);
    m_topStack->addWidget(m_avgOverlayTable);  // index 1 = averaged
    splitter->addWidget(m_topStack);

    // -- Bottom: dual radar charts side by side --
    auto* chartContainer = new QWidget(this);
    auto* chartLayout = new QHBoxLayout(chartContainer);
    chartLayout->setContentsMargins(0, 0, 0, 0);
    chartLayout->setSpacing(4);

    m_vaporQualityChart = new RadarChartWidget(this);
    m_vaporQualityChart->setCustomAxes(kDetailedVaporQualityMetrics, kDetailedAxisLabels);

    m_consistencyChart = new RadarChartWidget(this);
    m_consistencyChart->setCustomAxes(kDetailedConsistencyMetrics, kDetailedAxisLabels);

    chartLayout->addWidget(m_vaporQualityChart, 1);
    chartLayout->addWidget(m_consistencyChart, 1);
    splitter->addWidget(chartContainer);

    // 40% questions, 60% charts
    splitter->setStretchFactor(0, 2);
    splitter->setStretchFactor(1, 3);

    mainLayout->addWidget(splitter, 1);
}

void DetailedSensoryPanel::buildHeaderRow(QWidget* container)
{
    auto* hl = new QHBoxLayout(container);
    hl->setContentsMargins(0, 0, 0, 0);

    auto addField = [&](const QString& label) -> QLineEdit* {
        hl->addWidget(new QLabel(label + ":", container));
        auto* edit = new QLineEdit(container);
        edit->setMinimumWidth(100);
        hl->addWidget(edit);
        return edit;
    };

    m_testTitleEdit = addField("Test Title");
    m_assessorEdit  = addField("Assessor");
    m_testerEdit    = addField("Tester");
    m_mediaEdit     = addField("Media");

    hl->addWidget(new QLabel("Date:", container));
    m_dateLabel = new QLabel(QDate::currentDate().toString("yyyy-MM-dd"), container);
    hl->addWidget(m_dateLabel);

    hl->addStretch();

    // Connect changes to refresh
    for (auto* edit : {m_testTitleEdit, m_assessorEdit, m_testerEdit, m_mediaEdit}) {
        connect(edit, &QLineEdit::textChanged, this, &DetailedSensoryPanel::scheduleChartRefresh);
    }
}

void DetailedSensoryPanel::buildSampleNavBar()
{
    auto* navBar = new QWidget(this);
    auto* hl = new QHBoxLayout(navBar);
    hl->setContentsMargins(0, 0, 0, 0);

    m_prevBtn = new QPushButton(QStringLiteral("\u25C0"), navBar);
    m_nextBtn = new QPushButton(QStringLiteral("\u25B6"), navBar);
    m_prevBtn->setToolTip("Previous sample (Ctrl+Left)");
    m_nextBtn->setToolTip("Next sample (Ctrl+Right)");
    m_prevBtn->setFixedSize(28, 24);
    m_nextBtn->setFixedSize(28, 24);

    m_sampleCountLabel = new QLabel(QStringLiteral("\u2014"), navBar);
    m_sampleCountLabel->setAlignment(Qt::AlignCenter);
    m_sampleCountLabel->setFont(AppTheme::fontSmall());
    m_sampleCountLabel->setMinimumWidth(70);

    m_addSampleBtn = new QPushButton("+ Add Sample", navBar);
    m_removeSampleBtn = new QPushButton("Remove", navBar);

    hl->addWidget(m_prevBtn);
    hl->addWidget(m_sampleCountLabel);
    hl->addWidget(m_nextBtn);
    hl->addStretch();
    hl->addWidget(m_addSampleBtn);
    hl->addWidget(m_removeSampleBtn);

    connect(m_prevBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onPrevSample);
    connect(m_nextBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onNextSample);
    connect(m_addSampleBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onAddSample);
    connect(m_removeSampleBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onRemoveSample);

    // Insert navBar into the main layout — the caller handles this
    // We need to parent it properly
    navBar->setParent(this);
}

void DetailedSensoryPanel::buildQuestionForm()
{
    m_questionForm = new QWidget(this);
    auto* grid = new QGridLayout(m_questionForm);
    grid->setContentsMargins(4, 4, 4, 4);
    grid->setSpacing(4);

    int row = 0;

    // Sample name
    grid->addWidget(new QLabel("Sample Name:", m_questionForm), row, 0);
    m_sampleNameEdit = new QLineEdit(m_questionForm);
    grid->addWidget(m_sampleNameEdit, row, 1, 1, 3);
    connect(m_sampleNameEdit, &QLineEdit::textChanged, this, &DetailedSensoryPanel::scheduleChartRefresh);
    ++row;

    // ── Vapor Quality section header ────────────────────────────────────
    auto* vqLabel = new QLabel("<b>Vapor Quality</b>", m_questionForm);
    grid->addWidget(vqLabel, row, 0, 1, 4);
    ++row;

    // Scored questions (1-9 spinboxes)
    auto addSpinQuestion = [&](const QString& metric, double min, double max, double step, double defaultVal) {
        QString label = kDetailedAxisLabels.value(metric, metric);
        // Flatten multi-line label for form
        label.replace('\n', ' ');
        grid->addWidget(new QLabel(label + ":", m_questionForm), row, 0, 1, 2);
        auto* spin = new NoWheelDoubleSpinBox(m_questionForm);
        spin->setRange(min, max);
        spin->setSingleStep(step);
        spin->setDecimals(1);
        spin->setValue(defaultVal);
        spin->setFixedWidth(70);
        grid->addWidget(spin, row, 2);
        m_spinBoxes[metric] = spin;
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, &DetailedSensoryPanel::scheduleChartRefresh);
        ++row;
    };

    addSpinQuestion("Burn Taste", 1.0, 9.0, 0.1, 1.0);
    addSpinQuestion("Flavor Intensity", 1.0, 9.0, 0.1, 5.0);
    addSpinQuestion("Throat Irritation", 1.0, 9.0, 0.1, 1.0);
    addSpinQuestion("Nasal Irritation", 1.0, 9.0, 0.1, 1.0);
    addSpinQuestion("Vapor Quality Overall", 1.0, 9.0, 0.1, 5.0);

    // Cough — combobox
    auto addComboQuestion = [&](const QString& metric, const QVector<ChoiceOption>& options) {
        QString label = kDetailedAxisLabels.value(metric, metric);
        label.replace('\n', ' ');
        grid->addWidget(new QLabel(label + ":", m_questionForm), row, 0, 1, 2);
        auto* combo = new NoWheelComboBox(m_questionForm);
        for (const auto& opt : options) {
            combo->addItem(opt.text, opt.value);
        }
        grid->addWidget(combo, row, 2, 1, 2);
        m_comboBoxes[metric] = combo;
        connect(combo, QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, &DetailedSensoryPanel::scheduleChartRefresh);
        ++row;
    };

    addComboQuestion("Cough", kCoughOptions);

    // ── Consistency section header ──────────────────────────────────────
    auto* consLabel = new QLabel("<b>Consistency</b>", m_questionForm);
    grid->addWidget(consLabel, row, 0, 1, 4);
    ++row;

    addComboQuestion("Volume Consistency", kVolumeConsistencyOptions);
    addComboQuestion("Performance Consistency", kPerformanceConsistencyOptions);
    addComboQuestion("Vapor Temperature", kVaporTemperatureOptions);
    addComboQuestion("Vapor vs Oil", kVaporVsOilOptions);
    addComboQuestion("Vapor Volume", kVaporVolumeOptions);

    // ── Comments ────────────────────────────────────────────────────────
    grid->addWidget(new QLabel("Comments:", m_questionForm), row, 0);
    m_commentsEdit = new QTextEdit(m_questionForm);
    m_commentsEdit->setMaximumHeight(60);
    grid->addWidget(m_commentsEdit, row, 1, 1, 3);
    ++row;

    grid->setRowStretch(row, 1); // push everything up
}

// ── Sample navigation ───────────────────────────────────────────────────────

void DetailedSensoryPanel::onPrevSample()
{
    if (m_currentTesterIdx < 0) return;
    saveCurrentSampleToSession();
    if (m_currentSampleIdx > 0) {
        --m_currentSampleIdx;
        displayCurrentSample();
    }
}

void DetailedSensoryPanel::onNextSample()
{
    if (m_currentTesterIdx < 0) return;
    saveCurrentSampleToSession();
    auto* sess = currentSession();
    if (!sess) return;
    if (m_currentSampleIdx < sess->samples.size() - 1) {
        ++m_currentSampleIdx;
        displayCurrentSample();
    }
}

void DetailedSensoryPanel::onAddSample()
{
    if (m_currentTesterIdx < 0) {
        newSession();
        return;
    }
    saveCurrentSampleToSession();
    auto* sess = currentSession();
    if (!sess) return;
    sess->samples.append(DetailedSensorySample{});
    m_currentSampleIdx = sess->samples.size() - 1;
    displayCurrentSample();
    scheduleChartRefresh();
    emit sessionsChanged();
}

void DetailedSensoryPanel::onRemoveSample()
{
    auto* sess = currentSession();
    if (!sess || sess->samples.isEmpty()) return;
    if (sess->samples.size() == 1) return; // keep at least one

    sess->samples.remove(m_currentSampleIdx);
    if (m_currentSampleIdx >= sess->samples.size())
        m_currentSampleIdx = sess->samples.size() - 1;
    displayCurrentSample();
    scheduleChartRefresh();
    emit sessionsChanged();
}

void DetailedSensoryPanel::updateSampleNav()
{
    auto* sess = currentSession();
    if (!sess || sess->samples.isEmpty()) {
        m_sampleCountLabel->setText(QStringLiteral("\u2014"));
        m_prevBtn->setEnabled(false);
        m_nextBtn->setEnabled(false);
        m_removeSampleBtn->setEnabled(false);
        return;
    }
    m_sampleCountLabel->setText(QString("Sample %1 of %2")
                                    .arg(m_currentSampleIdx + 1)
                                    .arg(sess->samples.size()));
    m_prevBtn->setEnabled(m_currentSampleIdx > 0);
    m_nextBtn->setEnabled(m_currentSampleIdx < sess->samples.size() - 1);
    m_removeSampleBtn->setEnabled(sess->samples.size() > 1);
}

void DetailedSensoryPanel::saveCurrentSampleToSession()
{
    auto* sess = currentSession();
    if (!sess || m_currentSampleIdx < 0 || m_currentSampleIdx >= sess->samples.size()) return;

    auto& sample = sess->samples[m_currentSampleIdx];
    sample.name     = m_sampleNameEdit->text();
    sample.comments = m_commentsEdit->toPlainText();

    for (auto it = m_spinBoxes.begin(); it != m_spinBoxes.end(); ++it)
        sample.scores[it.key()] = it.value()->value();

    for (auto it = m_comboBoxes.begin(); it != m_comboBoxes.end(); ++it)
        sample.scores[it.key()] = it.value()->currentData().toDouble();
}

void DetailedSensoryPanel::displayCurrentSample()
{
    auto* sess = currentSession();
    if (!sess || m_currentSampleIdx < 0 || m_currentSampleIdx >= sess->samples.size()) {
        updateSampleNav();
        return;
    }

    const auto& sample = sess->samples[m_currentSampleIdx];

    // Block signals while populating
    for (auto* spin : m_spinBoxes) spin->blockSignals(true);
    for (auto* combo : m_comboBoxes) combo->blockSignals(true);
    m_sampleNameEdit->blockSignals(true);

    m_sampleNameEdit->setText(sample.name);
    m_commentsEdit->setPlainText(sample.comments);

    for (auto it = m_spinBoxes.begin(); it != m_spinBoxes.end(); ++it) {
        double val = sample.scores.value(it.key(), it.value()->minimum());
        it.value()->setValue(val);
    }

    for (auto it = m_comboBoxes.begin(); it != m_comboBoxes.end(); ++it) {
        int rawVal = static_cast<int>(sample.scores.value(it.key(), 1.0));
        int idx = it.value()->findData(rawVal);
        if (idx >= 0) it.value()->setCurrentIndex(idx);
    }

    for (auto* spin : m_spinBoxes) spin->blockSignals(false);
    for (auto* combo : m_comboBoxes) combo->blockSignals(false);
    m_sampleNameEdit->blockSignals(false);

    updateSampleNav();
}

// ── Chart refresh ───────────────────────────────────────────────────────────

void DetailedSensoryPanel::scheduleChartRefresh()
{
    m_refreshTimer->start();
}

void DetailedSensoryPanel::onRefreshChart()
{
    // Save current form state into session
    saveCurrentSampleToSession();

    // Build chart data from all sessions' samples
    QVector<RadarChartWidget::SampleData> vqData, consData;

    for (const auto& sess : m_sessions) {
        for (const auto& sample : sess.samples) {
            RadarChartWidget::SampleData vqSample, consSample;
            vqSample.name = sample.name;
            consSample.name = sample.name;

            // Vapor Quality metrics — normalize to 1-9
            for (const QString& metric : kDetailedVaporQualityMetrics) {
                double raw = sample.scores.value(metric, 1.0);
                vqSample.scores[metric] = normalizeToRadar(metric, raw);
            }
            // Consistency metrics — normalize to 1-9
            for (const QString& metric : kDetailedConsistencyMetrics) {
                double raw = sample.scores.value(metric, 1.0);
                consSample.scores[metric] = normalizeToRadar(metric, raw);
            }

            vqData.append(vqSample);
            consData.append(consSample);
        }
    }

    m_vaporQualityChart->setCustomData(vqData);
    m_consistencyChart->setCustomData(consData);
}

// ── Session management ──────────────────────────────────────────────────────

DetailedSensorySession* DetailedSensoryPanel::currentSession()
{
    if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size())
        return nullptr;
    return &m_sessions[m_currentTesterIdx];
}

void DetailedSensoryPanel::newSession()
{
    saveCurrentTester();

    DetailedSensorySession sess;
    sess.date = QDate::currentDate().toString("yyyy-MM-dd");
    sess.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    sess.samples.append(DetailedSensorySample{});

    m_sessions.append(sess);
    m_currentTesterIdx = m_sessions.size() - 1;
    m_currentSampleIdx = 0;
    applySession(sess);
    emit sessionsChanged();
}

void DetailedSensoryPanel::closeSessions(const QVector<int>& indices)
{
    QVector<int> sorted = indices;
    std::sort(sorted.begin(), sorted.end(), std::greater<int>());
    for (int idx : sorted) {
        if (idx >= 0 && idx < m_sessions.size())
            m_sessions.remove(idx);
    }
    if (m_sessions.isEmpty()) {
        m_currentTesterIdx = -1;
        m_currentSampleIdx = 0;
        // Clear UI
        m_testTitleEdit->clear();
        m_assessorEdit->clear();
        m_testerEdit->clear();
        m_mediaEdit->clear();
        m_sampleNameEdit->clear();
        m_commentsEdit->clear();
        for (auto* spin : m_spinBoxes) spin->setValue(spin->minimum());
        for (auto* combo : m_comboBoxes) combo->setCurrentIndex(0);
    } else {
        m_currentTesterIdx = qMin(m_currentTesterIdx, m_sessions.size() - 1);
        m_currentSampleIdx = 0;
        applySession(m_sessions[m_currentTesterIdx]);
    }
    onRefreshChart();
    emit sessionsChanged();
}

void DetailedSensoryPanel::loadSessions(const QVector<DetailedSensorySession>& sessions)
{
    m_sessions = sessions;
    if (!m_sessions.isEmpty()) {
        m_currentTesterIdx = 0;
        m_currentSampleIdx = 0;
        applySession(m_sessions[0]);
    }
    onRefreshChart();
    emit sessionsChanged();
}

void DetailedSensoryPanel::selectSession(int index)
{
    if (index < 0 || index >= m_sessions.size()) return;
    saveCurrentTester();
    m_currentTesterIdx = index;
    m_currentSampleIdx = 0;
    applySession(m_sessions[index]);
    onRefreshChart();
}

void DetailedSensoryPanel::renameSession(int index, const QString& newLabel)
{
    if (index < 0 || index >= m_sessions.size()) return;
    // Parse "Title - Tester" format
    int sep = newLabel.indexOf(" - ");
    if (sep >= 0) {
        m_sessions[index].testTitle  = newLabel.left(sep).trimmed();
        m_sessions[index].testerName = newLabel.mid(sep + 3).trimmed();
    } else {
        m_sessions[index].sessionName = newLabel;
    }
    if (index == m_currentTesterIdx)
        applySession(m_sessions[index]);
    emit sessionsChanged();
}

void DetailedSensoryPanel::showAveragedChart(const QVector<int>& sessionIndices)
{
    // Build averaged data across selected sessions, per device
    QMap<QString, QMap<QString, QVector<double>>> deviceScores; // device -> metric -> values
    for (int idx : sessionIndices) {
        if (idx < 0 || idx >= m_sessions.size()) continue;
        for (const auto& sample : m_sessions[idx].samples) {
            for (const QString& metric : kDetailedAllMetrics) {
                deviceScores[sample.name][metric].append(sample.scores.value(metric, 0.0));
            }
        }
    }

    // Build averaged samples for charts
    QVector<RadarChartWidget::SampleData> vqData, consData;
    for (auto it = deviceScores.begin(); it != deviceScores.end(); ++it) {
        RadarChartWidget::SampleData vqSample, consSample;
        vqSample.name = consSample.name = it.key();
        for (const QString& metric : kDetailedVaporQualityMetrics) {
            const auto& vals = it.value().value(metric);
            double avg = 0;
            for (double v : vals) avg += v;
            if (!vals.isEmpty()) avg /= vals.size();
            vqSample.scores[metric] = normalizeToRadar(metric, avg);
        }
        for (const QString& metric : kDetailedConsistencyMetrics) {
            const auto& vals = it.value().value(metric);
            double avg = 0;
            for (double v : vals) avg += v;
            if (!vals.isEmpty()) avg /= vals.size();
            consSample.scores[metric] = normalizeToRadar(metric, avg);
        }
        vqData.append(vqSample);
        consData.append(consSample);
    }

    m_vaporQualityChart->setCustomData(vqData);
    m_consistencyChart->setCustomData(consData);
}

QVector<DetailedSensorySession> DetailedSensoryPanel::allSessions()
{
    saveCurrentTester();
    return m_sessions;
}

QString DetailedSensoryPanel::sessionLabel(const DetailedSensorySession& s) const
{
    QString title  = s.testTitle.isEmpty()  ? s.sessionName : s.testTitle;
    QString tester = s.testerName;
    if (!title.isEmpty() && !tester.isEmpty())
        return title + " - " + tester;
    if (!title.isEmpty()) return title;
    if (!tester.isEmpty()) return tester;
    return s.sessionName.isEmpty() ? "(untitled)" : s.sessionName;
}

DetailedSensorySession DetailedSensoryPanel::buildSession() const
{
    DetailedSensorySession sess;
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size())
        sess = m_sessions[m_currentTesterIdx];

    sess.testTitle    = m_testTitleEdit->text();
    sess.assessorName = m_assessorEdit->text();
    sess.testerName   = m_testerEdit->text();
    sess.media        = m_mediaEdit->text();
    sess.date         = m_dateLabel->text();
    if (sess.sessionName.isEmpty())
        sess.sessionName = sess.testTitle;
    if (sess.timestamp.isEmpty())
        sess.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    return sess;
}

void DetailedSensoryPanel::applySession(const DetailedSensorySession& session)
{
    m_testTitleEdit->blockSignals(true);
    m_assessorEdit->blockSignals(true);
    m_testerEdit->blockSignals(true);
    m_mediaEdit->blockSignals(true);

    m_testTitleEdit->setText(session.testTitle);
    m_assessorEdit->setText(session.assessorName);
    m_testerEdit->setText(session.testerName);
    m_mediaEdit->setText(session.media);
    m_dateLabel->setText(session.date.isEmpty()
                             ? QDate::currentDate().toString("yyyy-MM-dd")
                             : session.date);

    m_testTitleEdit->blockSignals(false);
    m_assessorEdit->blockSignals(false);
    m_testerEdit->blockSignals(false);
    m_mediaEdit->blockSignals(false);

    m_currentSampleIdx = 0;
    displayCurrentSample();
}

void DetailedSensoryPanel::saveCurrentTester()
{
    if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size()) return;
    saveCurrentSampleToSession();
    m_sessions[m_currentTesterIdx] = buildSession();
}

bool DetailedSensoryPanel::isDefaultState() const
{
    return m_sessions.isEmpty()
        || (m_sessions.size() == 1 && m_sessions[0].samples.isEmpty());
}

void DetailedSensoryPanel::showAveragedTable(const QStringList& deviceNames,
                                              const QVector<QMap<QString, double>>& deviceAvgs)
{
    m_avgOverlayTable->clear();
    m_avgOverlayTable->setColumnCount(kDetailedAllMetrics.size() + 1);

    QStringList headers;
    headers << "Device";
    for (const QString& m : kDetailedAllMetrics)
        headers << m;
    m_avgOverlayTable->setHorizontalHeaderLabels(headers);
    m_avgOverlayTable->setRowCount(deviceNames.size());

    for (int r = 0; r < deviceNames.size(); ++r) {
        m_avgOverlayTable->setItem(r, 0, new QTableWidgetItem(deviceNames[r]));
        for (int c = 0; c < kDetailedAllMetrics.size(); ++c) {
            double val = deviceAvgs[r].value(kDetailedAllMetrics[c], 0.0);
            m_avgOverlayTable->setItem(r, c + 1,
                new QTableWidgetItem(QString::number(val, 'f', 1)));
        }
    }

    m_avgOverlayTable->resizeColumnsToContents();
    m_topStack->setCurrentIndex(1);
}

void DetailedSensoryPanel::showNormalView()
{
    m_topStack->setCurrentIndex(0);
}

// ── File I/O ────────────────────────────────────────────────────────────────

void DetailedSensoryPanel::save()
{
    saveCurrentTester();
    if (m_currentTesterIdx < 0) return;

    const auto& sess = m_sessions[m_currentTesterIdx];

    if (m_savePath.isEmpty()) {
        QString defaultName = sess.testTitle.isEmpty() ? "detailed_sensory" : sess.testTitle;
        defaultName.replace(' ', '_');
        m_savePath = QFileDialog::getSaveFileName(
            this, "Save Detailed Sensory Session",
            lastBrowseDir() + "/" + defaultName + ".xlsx",
            "Excel Files (*.xlsx)");
        if (m_savePath.isEmpty()) return;
        setLastBrowseDir(m_savePath);
    }

    saveToExcel(m_savePath, sess);

    // Also save to database
    if (m_db && m_db->isOpen()) {
        m_db->saveDetailedSensorySession(sess);
    }
}

void DetailedSensoryPanel::saveToExcel(const QString& path, const DetailedSensorySession& sess)
{
    QXlsx::Document xlsx;

    // Header row
    int col = 1;
    xlsx.write(1, col++, "Sample Name");
    for (const QString& metric : kDetailedAllMetrics)
        xlsx.write(1, col++, metric);
    xlsx.write(1, col++, "V");
    xlsx.write(1, col++, "R");
    xlsx.write(1, col++, "P");
    xlsx.write(1, col++, "HT");
    xlsx.write(1, col++, "Comments");

    // Data rows
    int row = 2;
    for (const auto& sample : sess.samples) {
        col = 1;
        xlsx.write(row, col++, sample.name);
        for (const QString& metric : kDetailedAllMetrics)
            xlsx.write(row, col++, sample.scores.value(metric, 0.0));
        xlsx.write(row, col++, sample.voltage);
        xlsx.write(row, col++, sample.resistance);
        xlsx.write(row, col++, sample.power);
        xlsx.write(row, col++, sample.heatingTechnology);
        xlsx.write(row, col++, sample.comments);
        ++row;
    }

    // Average row
    if (sess.samples.size() > 1) {
        col = 1;
        xlsx.write(row, col++, "Average");
        for (const QString& metric : kDetailedAllMetrics) {
            double sum = 0;
            for (const auto& s : sess.samples)
                sum += s.scores.value(metric, 0.0);
            xlsx.write(row, col++, sum / sess.samples.size());
        }
        ++row;
    }

    // Footer metadata
    ++row;
    xlsx.write(row, 1, "Test Title"); xlsx.write(row, 2, sess.testTitle); ++row;
    xlsx.write(row, 1, "Tester");     xlsx.write(row, 2, sess.testerName); ++row;
    xlsx.write(row, 1, "Assessor");   xlsx.write(row, 2, sess.assessorName); ++row;
    xlsx.write(row, 1, "Media");      xlsx.write(row, 2, sess.media); ++row;
    xlsx.write(row, 1, "Date");       xlsx.write(row, 2, sess.date); ++row;
    xlsx.write(row, 1, "Facilitator"); xlsx.write(row, 2, sess.facilitatorName); ++row;
    xlsx.write(row, 1, "Oil Smell Liking"); xlsx.write(row, 2, sess.oilSmellLiking); ++row;
    xlsx.write(row, 1, "Viscosity");  xlsx.write(row, 2, sess.viscosity); ++row;

    xlsx.saveAs(path);
}

void DetailedSensoryPanel::loadFiles()
{
    QStringList files = QFileDialog::getOpenFileNames(
        this, "Load Detailed Sensory Data",
        lastBrowseDir(),
        "Excel Files (*.xlsx);;All Files (*)");
    if (files.isEmpty()) return;
    setLastBrowseDir(files.first());

    // TODO: implement Excel loading (reverse of saveToExcel)
    // For now, just load from database
    loadFromDatabase();
}

void DetailedSensoryPanel::loadFromDatabase()
{
    if (!m_db || !m_db->isOpen()) return;
    auto sessions = m_db->loadDetailedSensorySessions();
    if (!sessions.isEmpty())
        loadSessions(sessions);
}

// ── Browse dir helpers ──────────────────────────────────────────────────────

QString DetailedSensoryPanel::lastBrowseDir() const
{
    if (!m_lastBrowseDir.isEmpty()) return m_lastBrowseDir;
    if (m_db) return m_db->getSetting("last_browse_dir");
    return QDir::homePath();
}

void DetailedSensoryPanel::setLastBrowseDir(const QString& filePath)
{
    m_lastBrowseDir = QFileInfo(filePath).absolutePath();
}

} // namespace DVE
```

- [ ] **Step 3: Update `.pro` file**

In `DataViewerEnterprise.pro`, add to SOURCES after `src/ui/SensoryPanel.cpp \`:
```
    src/ui/DetailedSensoryPanel.cpp \
```

Add to HEADERS after `src/ui/SensoryPanel.h \`:
```
    src/ui/DetailedSensoryPanel.h \
```

- [ ] **Step 4: Verify build**

Run: `cd build && qmake ../DataViewerEnterprise.pro && mingw32-make -j8 2>&1 | tail -10`
Expected: Build succeeds.

- [ ] **Step 5: Commit**

```bash
git add src/ui/DetailedSensoryPanel.h src/ui/DetailedSensoryPanel.cpp DataViewerEnterprise.pro
git commit -m "feat: add DetailedSensoryPanel with question form, dual radar charts, sample navigation"
```

---

### Task 6: MainWindow — Three-Way Mode Toggle + Ribbon Wiring

**Files:**
- Modify: `src/MainWindow.h`
- Modify: `src/MainWindow.cpp`

- [ ] **Step 1: Add member variables to MainWindow.h**

After line 42 (`class SensoryPanel;`), add:
```cpp
class DetailedSensoryPanel;
```

After line 184 (`QToolButton*   m_sensoryBtn = nullptr;`), add:
```cpp
    QToolButton*   m_detailedSensoryBtn = nullptr;  // checkable toggle
```

After line 177 (`QToolButton*   m_homeSensCloseBtn  = nullptr;`), add:
```cpp
    // Home tab — Detailed Sensory buttons (hidden in TPM/Sensory mode)
    QToolButton*   m_homeDetSensNewBtn    = nullptr;
    QToolButton*   m_homeDetSensSaveBtn   = nullptr;
    QToolButton*   m_homeDetSensLoadXlBtn = nullptr;
    QToolButton*   m_homeDetSensCloseBtn  = nullptr;
```

After line 291 (`SensoryPanel*   m_sensoryPanel = nullptr;`), add:
```cpp
    bool                    m_detailedSensoryMode = false;
    bool                    m_detailedSensorySessionsDirty = false;
    DetailedSensoryPanel*   m_detailedSensoryPanel = nullptr;
    QListWidget*            m_detailedSensoryNav = nullptr;
```

After line 301 (`void updateSensoryProperties();`), add:
```cpp
    void initDetailedSensoryPanel();
    void toggleDetailedSensoryMode(bool checked);
    void refreshDetailedSensoryNavigator();
    void updateDetailedSensoryProperties();
```

At the top of MainWindow.h, after line 36 (`#include <QListWidget>`), add:
```cpp
#include "ui/DetailedSensoryPanel.h"
```

- [ ] **Step 2: Add Detailed Sensory button to `buildToolsTab()`**

In `MainWindow.cpp`, in `buildToolsTab()` (after the `m_sensoryBtn` setup, after line 221), add:

```cpp
    m_detailedSensoryBtn = grp->addLargeButton("Detailed\nSensory",
        QIcon(resourcePath() + "/images/ccell_icon_black.png"),
        "Toggle detailed sensory evaluation mode (S2-1)");
    m_detailedSensoryBtn->setCheckable(true);
    connect(m_detailedSensoryBtn, &QToolButton::toggled,
            this, &MainWindow::toggleDetailedSensoryMode);
```

- [ ] **Step 3: Add Detailed Sensory buttons to `buildHomeTab()`**

In `MainWindow.cpp`, after the sensory buttons setup (after the `m_homeSensCloseBtn` connection block, around line 156), add:

```cpp
    // ── Detailed Sensory buttons (hidden by default) ────────────────────
    m_homeDetSensNewBtn   = fileGrp->addLargeButton("New\nSession",
        style()->standardIcon(QStyle::SP_FileIcon), "Create a new detailed sensory session");
    m_homeDetSensSaveBtn  = fileGrp->addLargeButton("Save",
        style()->standardIcon(QStyle::SP_DialogSaveButton), "Save session (Ctrl+S)");
    m_homeDetSensLoadXlBtn = fileGrp->addLargeButton("Load\nExcel",
        style()->standardIcon(QStyle::SP_DialogOpenButton), "Load detailed sensory data from Excel");
    m_homeDetSensCloseBtn  = fileGrp->addLargeButton("Close",
        style()->standardIcon(QStyle::SP_DialogCloseButton), "Close selected session(s)");

    m_homeDetSensNewBtn->setVisible(false);
    m_homeDetSensSaveBtn->setVisible(false);
    m_homeDetSensLoadXlBtn->setVisible(false);
    m_homeDetSensCloseBtn->setVisible(false);

    connect(m_homeDetSensNewBtn, &QToolButton::clicked, this, [this]() {
        if (m_detailedSensoryPanel) m_detailedSensoryPanel->newSession();
    });
    connect(m_homeDetSensSaveBtn, &QToolButton::clicked, this, [this]() {
        if (m_detailedSensoryPanel) m_detailedSensoryPanel->save();
    });
    connect(m_homeDetSensLoadXlBtn, &QToolButton::clicked, this, [this]() {
        if (m_detailedSensoryPanel) m_detailedSensoryPanel->loadFiles();
    });
    connect(m_homeDetSensCloseBtn, &QToolButton::clicked, this, [this]() {
        if (!m_detailedSensoryPanel) return;
        QVector<int> indices;
        for (auto* item : m_detailedSensoryNav->selectedItems())
            indices.append(m_detailedSensoryNav->row(item));
        if (indices.isEmpty() && m_detailedSensoryPanel->currentSessionIndex() >= 0)
            indices.append(m_detailedSensoryPanel->currentSessionIndex());
        if (indices.isEmpty()) return;
        m_detailedSensoryPanel->closeSessions(indices);
        updateImageButton();
    });
```

- [ ] **Step 4: Implement `toggleDetailedSensoryMode()`**

Add to `MainWindow.cpp`:

```cpp
void MainWindow::toggleDetailedSensoryMode(bool checked)
{
    m_detailedSensoryMode = checked;

    if (checked) {
        // Uncheck sensory mode if active
        if (m_sensoryMode) {
            m_sensoryBtn->blockSignals(true);
            m_sensoryBtn->setChecked(false);
            m_sensoryBtn->blockSignals(false);
            m_sensoryMode = false;
        }

        if (!m_detailedSensoryPanel) {
            initDetailedSensoryPanel();
        }
        m_centralStack->setCurrentWidget(m_detailedSensoryPanel);
        m_navStack->setCurrentIndex(2);  // detailed sensory navigator
        m_navLabel->setText("Sessions:  <span style='color:gray; font-size:11px;'>select multiple to show average score</span>");
        refreshDetailedSensoryNavigator();
        if (m_testAvgPanel) m_testAvgPanel->setVisible(true);
        updateDetailedSensoryProperties();
    } else {
        m_centralStack->setCurrentIndex(0);   // TPM splitter
        m_navStack->setCurrentIndex(0);        // file tree
        m_navLabel->setText("Loaded Files:");
        if (m_testAvgPanel) m_testAvgPanel->setVisible(false);
        if (currentSheet() && m_currentSampleIndex >= 0
            && m_currentSampleIndex < currentSheet()->samples.size()) {
            updateProperties(currentSheet()->samples[m_currentSampleIndex]);
        } else {
            m_propTable->setRowCount(0);
        }
    }

    updateRibbonForMode();
    updateImageButton();
}
```

- [ ] **Step 5: Implement `initDetailedSensoryPanel()`**

```cpp
void MainWindow::initDetailedSensoryPanel()
{
    m_detailedSensoryPanel = new DetailedSensoryPanel(m_db, this);
    m_centralStack->addWidget(m_detailedSensoryPanel);

    // Add navigator list for detailed sensory sessions
    m_detailedSensoryNav = new QListWidget(this);
    m_detailedSensoryNav->setSelectionMode(QAbstractItemView::ExtendedSelection);
    m_navStack->addWidget(m_detailedSensoryNav);  // index 2

    connect(m_detailedSensoryNav, &QListWidget::currentRowChanged, this, [this](int row) {
        if (m_detailedSensoryPanel && row >= 0) {
            m_detailedSensoryPanel->selectSession(row);
            updateDetailedSensoryProperties();
        }
    });

    connect(m_detailedSensoryNav, &QListWidget::itemSelectionChanged, this, [this]() {
        auto items = m_detailedSensoryNav->selectedItems();
        if (items.size() > 1) {
            QVector<int> indices;
            for (auto* item : items)
                indices.append(m_detailedSensoryNav->row(item));
            m_detailedSensoryPanel->showAveragedChart(indices);
        }
    });

    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, &MainWindow::refreshDetailedSensoryNavigator);
    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, &MainWindow::updateDetailedSensoryProperties);
    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, &MainWindow::updateImageButton);
    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, [this]() {
        m_detailedSensorySessionsDirty = true;
        updateDbSyncIndicator();
    });
}
```

- [ ] **Step 6: Implement `refreshDetailedSensoryNavigator()` and `updateDetailedSensoryProperties()`**

```cpp
void MainWindow::refreshDetailedSensoryNavigator()
{
    if (!m_detailedSensoryNav || !m_detailedSensoryPanel) return;
    m_detailedSensoryNav->blockSignals(true);
    m_detailedSensoryNav->clear();
    auto sessions = m_detailedSensoryPanel->allSessions();
    for (const auto& s : sessions)
        m_detailedSensoryNav->addItem(m_detailedSensoryPanel->sessionLabel(s));
    int idx = m_detailedSensoryPanel->currentSessionIndex();
    if (idx >= 0 && idx < m_detailedSensoryNav->count())
        m_detailedSensoryNav->setCurrentRow(idx);
    m_detailedSensoryNav->blockSignals(false);
}

void MainWindow::updateDetailedSensoryProperties()
{
    if (!m_detailedSensoryPanel) return;
    auto* sess = m_detailedSensoryPanel->currentSession();
    if (!sess) { m_propTable->setRowCount(0); return; }

    m_propTable->blockSignals(true);
    m_propTable->setRowCount(0);

    auto addRow = [&](const QString& prop, const QString& val) {
        int r = m_propTable->rowCount();
        m_propTable->insertRow(r);
        auto* pItem = new QTableWidgetItem(prop);
        pItem->setFlags(pItem->flags() & ~Qt::ItemIsEditable);
        m_propTable->setItem(r, 0, pItem);
        m_propTable->setItem(r, 1, new QTableWidgetItem(val));
    };

    // Session Info header
    int r = m_propTable->rowCount();
    m_propTable->insertRow(r);
    auto* hdr = new QTableWidgetItem("Session Info");
    hdr->setFlags(hdr->flags() & ~Qt::ItemIsEditable);
    hdr->setForeground(QColor(0, 120, 215));
    hdr->setFont(AppTheme::fontBold());
    m_propTable->setItem(r, 0, hdr);
    m_propTable->setItem(r, 1, new QTableWidgetItem());

    addRow("Test Title", sess->testTitle);
    addRow("Assessor", sess->assessorName);
    addRow("Tester", sess->testerName);
    addRow("Media", sess->media);
    addRow("Date", sess->date);
    addRow("Samples", QString::number(sess->samples.size()));
    addRow("Viscosity", sess->viscosity);
    addRow("Oil Smell Liking", QString::number(sess->oilSmellLiking));
    addRow("Clog", sess->clog ? "Yes" : "No");
    if (sess->clog) addRow("Clog Oil Level", sess->clogOilLevel);
    addRow("Device Return Date", sess->deviceReturnDate);
    addRow("Facilitator", sess->facilitatorName);
    addRow("Facilitator Comment", sess->facilitatorComment);

    m_propTable->blockSignals(false);
}
```

- [ ] **Step 7: Update `toggleSensoryMode()` for mutual exclusion**

In the existing `toggleSensoryMode()` method, at the beginning of the `if (checked)` block (after `m_sensoryMode = checked;`), add:

```cpp
    if (checked) {
        // Uncheck detailed sensory mode if active
        if (m_detailedSensoryMode && m_detailedSensoryBtn) {
            m_detailedSensoryBtn->blockSignals(true);
            m_detailedSensoryBtn->setChecked(false);
            m_detailedSensoryBtn->blockSignals(false);
            m_detailedSensoryMode = false;
        }
```

- [ ] **Step 8: Update `updateRibbonForMode()` for three-way switching**

Replace the `updateRibbonForMode()` method with:

```cpp
void MainWindow::updateRibbonForMode()
{
    bool sensory = m_sensoryMode;
    bool detailedSensory = m_detailedSensoryMode;
    bool tpm = !sensory && !detailedSensory;

    // Home tab: show/hide TPM vs sensory vs detailed sensory buttons
    m_homeNewBtn->setVisible(tpm);
    m_homeLoadBtn->setVisible(tpm);
    m_homeCloseBtn->setVisible(tpm);

    m_homeSensNewBtn->setVisible(sensory);
    m_homeSensSaveBtn->setVisible(sensory);
    m_homeSensLoadXlBtn->setVisible(sensory);
    m_homeSensCloseBtn->setVisible(sensory);

    m_homeDetSensNewBtn->setVisible(detailedSensory);
    m_homeDetSensSaveBtn->setVisible(detailedSensory);
    m_homeDetSensLoadXlBtn->setVisible(detailedSensory);
    m_homeDetSensCloseBtn->setVisible(detailedSensory);

    // Reports tab: swap labels and connections
    disconnect(m_reportBtn1, &QToolButton::clicked, nullptr, nullptr);
    disconnect(m_reportBtn2, &QToolButton::clicked, nullptr, nullptr);

    if (sensory) {
        m_reportBtn1->setText("Sensory\nReport");
        m_reportBtn1->setIcon(QIcon(resourcePath() + "/images/ccell_icon.png"));
        m_reportBtn1->setToolTip("Generate PPTX report for selected sensory sessions");
        m_reportBtn2->setVisible(false);
        connect(m_reportBtn1, &QToolButton::clicked, this, [this]() {
            if (m_sensoryPanel) m_sensoryPanel->generateFullReport();
        });
        if (m_cleanupGroup) m_cleanupGroup->setVisible(false);

    } else if (detailedSensory) {
        m_reportBtn1->setText("Detailed\nSensory Report");
        m_reportBtn1->setIcon(QIcon(resourcePath() + "/images/ccell_icon_black.png"));
        m_reportBtn1->setToolTip("Generate PPTX report for selected detailed sensory sessions");
        m_reportBtn2->setVisible(false);
        connect(m_reportBtn1, &QToolButton::clicked, this, [this]() {
            if (m_detailedSensoryPanel) m_detailedSensoryPanel->generateFullReport();
        });
        if (m_cleanupGroup) m_cleanupGroup->setVisible(false);

    } else {
        m_reportBtn1->setText("Test Report");
        m_reportBtn1->setIcon(style()->standardIcon(QStyle::SP_FileDialogDetailedView));
        m_reportBtn1->setToolTip("Generate a PPTX report for the current sheet");
        m_reportBtn2->setVisible(true);
        m_reportBtn2->setText("Full Report");
        m_reportBtn2->setIcon(style()->standardIcon(QStyle::SP_FileDialogListView));
        m_reportBtn2->setToolTip("Generate a PPTX report for all sheets");
        connect(m_reportBtn1, &QToolButton::clicked, this, &MainWindow::onGenerateTestReport);
        connect(m_reportBtn2, &QToolButton::clicked, this, &MainWindow::onGenerateFullReport);
        if (m_cleanupGroup) m_cleanupGroup->setVisible(true);
    }
}
```

- [ ] **Step 9: Verify build and test mode switching**

Run: `cd build && qmake ../DataViewerEnterprise.pro && mingw32-make -j8 2>&1 | tail -10`
Expected: Build succeeds. Launch app, verify:
- Detailed Sensory button appears in Tools tab with black icon
- Clicking it switches to detailed sensory panel
- Clicking Sensory unchecks Detailed Sensory and vice versa
- Home tab buttons swap correctly
- Reports tab shows "Detailed Sensory Report" in detailed mode

- [ ] **Step 10: Commit**

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat: three-way mode toggle (TPM/Sensory/Detailed Sensory) with ribbon wiring"
```

---

### Task 7: Database Browser — Detailed Sensory Tab

**Files:**
- Modify: `src/ui/DatabaseBrowserDialog.h`
- Modify: `src/ui/DatabaseBrowserDialog.cpp`

- [ ] **Step 1: Add member variables to header**

In `DatabaseBrowserDialog.h`, add the necessary members for a third tab. After the existing sensory tab members, add:

```cpp
    // Detailed Sensory tab
    QTreeWidget*   m_detSensTree   = nullptr;
    QLabel*        m_detSensStatus = nullptr;
    QVector<DetailedSensoryRecord> m_detSensRecords;

    void populateDetailedSensoryTree(const QString& filter = {});
    void onDetailedSensoryDelete();
    void onDetailedSensoryGenerateReport();
```

Also add the include at the top:
```cpp
#include "pipeline/DetailedSensoryData.h"
```

- [ ] **Step 2: Build the "Detailed Sensory" tab in the constructor**

In `DatabaseBrowserDialog.cpp`, after the sensory tab setup, add a third tab following the exact same pattern. Create a "Detailed Sensory Data" tab with:
- Filter bar + refresh button
- Tree widget (5 columns: Test Title / Tester, Assessor, Media, Date, Samples)
- Status label
- Open Selected, Delete Selected, Generate Combined Report buttons
- Same button connections pattern as the sensory tab but calling `m_db->listDetailedSensoryRecords()`, `m_db->removeDetailedSensorySession()`, and `DetailedSensoryPanel::generateCombinedPptx()`

- [ ] **Step 3: Implement `populateDetailedSensoryTree()`**

Same pattern as `populateSensoryTree()` but using `m_detSensRecords` (which are `DetailedSensoryRecord` structs).

- [ ] **Step 4: Implement delete and report generation**

Same patterns as existing sensory tab handlers but calling the detailed sensory database methods.

- [ ] **Step 5: Verify build and test**

Run: `cd build && qmake ../DataViewerEnterprise.pro && mingw32-make -j8 2>&1 | tail -10`
Expected: Build succeeds. Database browser shows three tabs.

- [ ] **Step 6: Commit**

```bash
git add src/ui/DatabaseBrowserDialog.h src/ui/DatabaseBrowserDialog.cpp
git commit -m "feat: add Detailed Sensory tab to Database Browser dialog"
```

---

### Task 8: PowerPoint Report Generation

**Files:**
- Modify: `src/ui/DetailedSensoryPanel.cpp`

- [ ] **Step 1: Implement `generateFullReport()`**

Add the tester selection dialog (same pattern as `SensoryPanel::generateFullReport()`) and call `generateCombinedPptx()`:

```cpp
void DetailedSensoryPanel::generateFullReport()
{
    saveCurrentTester();
    if (m_sessions.isEmpty()) return;

    // Tester selection dialog
    QDialog dlg(this);
    dlg.setWindowTitle("Select Sessions for Report");
    auto* vl = new QVBoxLayout(&dlg);
    auto* listWidget = new QListWidget(&dlg);
    listWidget->setSelectionMode(QAbstractItemView::ExtendedSelection);
    for (const auto& s : m_sessions)
        listWidget->addItem(sessionLabel(s));
    listWidget->selectAll();
    vl->addWidget(listWidget);

    auto* btnBox = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
    connect(btnBox, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
    connect(btnBox, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);
    vl->addWidget(btnBox);

    if (dlg.exec() != QDialog::Accepted) return;

    QVector<DetailedSensorySession> selected;
    for (auto* item : listWidget->selectedItems()) {
        int idx = listWidget->row(item);
        if (idx >= 0 && idx < m_sessions.size())
            selected.append(m_sessions[idx]);
    }
    if (selected.isEmpty()) return;

    QString defaultName = selected[0].testTitle.isEmpty() ? "detailed_sensory_report" : selected[0].testTitle;
    defaultName.replace(' ', '_');
    QString path = QFileDialog::getSaveFileName(
        this, "Save Report",
        lastBrowseDir() + "/" + defaultName + ".pptx",
        "PowerPoint Files (*.pptx)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    QString errorOut;
    if (generateCombinedPptx(selected, path, errorOut)) {
        QMessageBox::information(this, "Report Saved", "Report saved to:\n" + path);
    } else {
        QMessageBox::warning(this, "Report Error", "Failed to generate report:\n" + errorOut);
    }
}
```

- [ ] **Step 2: Implement `generateCombinedPptx()`**

This is the core report method. It creates:
1. Title slide
2. Per sample: table slide (testers as rows, questions as columns, average row) + plot slide (dual radar charts side by side)

```cpp
bool DetailedSensoryPanel::generateCombinedPptx(
    const QVector<DetailedSensorySession>& sessions,
    const QString& filePath,
    QString& errorOut)
{
    if (sessions.isEmpty()) {
        errorOut = "No sessions provided";
        return false;
    }

    PptxWriter pptx;

    // ── Title slide ─────────────────────────────────────────────────────
    QString coverTitle = sessions[0].testTitle.isEmpty()
                             ? "Detailed Sensory Evaluation"
                             : sessions[0].testTitle;
    pptx.addCoverSlide(coverTitle, sessions[0].date);

    // ── Gather all unique sample names across sessions ──────────────────
    QStringList sampleNames;
    for (const auto& sess : sessions) {
        for (const auto& sample : sess.samples) {
            if (!sampleNames.contains(sample.name))
                sampleNames.append(sample.name);
        }
    }

    // ── Per sample: table slide + plot slide ────────────────────────────
    for (const QString& sampleName : sampleNames) {
        // Collect this sample's data across all testers
        struct TesterRow {
            QString testerName;
            QMap<QString, double> scores;
        };
        QVector<TesterRow> rows;

        for (const auto& sess : sessions) {
            for (const auto& sample : sess.samples) {
                if (sample.name == sampleName) {
                    TesterRow row;
                    row.testerName = sess.testerName.isEmpty() ? sess.assessorName : sess.testerName;
                    row.scores = sample.scores;
                    rows.append(row);
                }
            }
        }

        if (rows.isEmpty()) continue;

        // ── TABLE SLIDE ─────────────────────────────────────────────────
        // Headers: Tester | all 11 metrics
        QStringList headers;
        headers << "Tester";
        for (const QString& m : kDetailedAllMetrics)
            headers << m;

        QVector<QStringList> tableRows;
        QMap<QString, double> sums;
        for (const auto& r : rows) {
            QStringList cells;
            cells << r.testerName;
            for (const QString& m : kDetailedAllMetrics) {
                double val = r.scores.value(m, 0.0);
                cells << QString::number(val, 'f', 1);
                sums[m] += val;
            }
            tableRows.append(cells);
        }

        // Average row
        QStringList avgRow;
        avgRow << "Average";
        for (const QString& m : kDetailedAllMetrics)
            avgRow << QString::number(sums[m] / rows.size(), 'f', 1);
        tableRows.append(avgRow);

        // Build column widths (proportional)
        QVector<double> colWidths;
        colWidths << 0.12;  // Tester column
        double metricW = (1.0 - 0.12) / kDetailedAllMetrics.size();
        for (int i = 0; i < kDetailedAllMetrics.size(); ++i)
            colWidths << metricW;

        SlideTable table;
        table.headers = headers;
        table.rows = tableRows;
        table.colWidthFractions = colWidths;
        table.x = 0.3;
        table.y = 0.64;
        table.w = 12.7;
        pptx.addContentSlide(sampleName, table, {});

        // ── PLOT SLIDE ──────────────────────────────────────────────────
        // Render two radar charts side by side

        // Build chart data for this sample across testers
        QVector<RadarChartWidget::SampleData> vqData, consData;
        for (const auto& r : rows) {
            RadarChartWidget::SampleData vqSample, consSample;
            vqSample.name = consSample.name = r.testerName;
            for (const QString& m : kDetailedVaporQualityMetrics)
                vqSample.scores[m] = normalizeToRadar(m, r.scores.value(m, 1.0));
            for (const QString& m : kDetailedConsistencyMetrics)
                consSample.scores[m] = normalizeToRadar(m, r.scores.value(m, 1.0));
            vqData.append(vqSample);
            consData.append(consSample);
        }

        // Render left chart (Vapor Quality)
        RadarChartWidget vqChart;
        vqChart.setCustomAxes(kDetailedVaporQualityMetrics, kDetailedAxisLabels);
        vqChart.setCustomData(vqData);
        vqChart.setReportMode(true);
        vqChart.setReportCropTop(70);
        vqChart.resize(580, 858);
        QPixmap vqPix = vqChart.grab();

        // Render right chart (Consistency)
        RadarChartWidget consChart;
        consChart.setCustomAxes(kDetailedConsistencyMetrics, kDetailedAxisLabels);
        consChart.setCustomData(consData);
        consChart.setReportMode(true);
        consChart.setReportCropTop(70);
        consChart.resize(580, 858);
        QPixmap consPix = consChart.grab();

        // Save to temp files
        QString vqPath  = QDir::temp().filePath("dve_vq_report.png");
        QString consPath = QDir::temp().filePath("dve_cons_report.png");
        vqPix.save(vqPath);
        consPix.save(consPath);

        // Add dual-image slide using addContentSlide with two SlideImages
        QByteArray vqData, consData2;
        { QFile f(vqPath); if (f.open(QIODevice::ReadOnly)) vqData = f.readAll(); }
        { QFile f(consPath); if (f.open(QIODevice::ReadOnly)) consData2 = f.readAll(); }

        SlideImage vqImg;
        vqImg.pngData = vqData;
        vqImg.x = 0.3; vqImg.y = 0.8; vqImg.w = 6.0; vqImg.h = 5.5;

        SlideImage consImg;
        consImg.pngData = consData2;
        consImg.x = 6.7; consImg.y = 0.8; consImg.w = 6.0; consImg.h = 5.5;

        SlideTable emptyTable;
        emptyTable.headers = {};
        emptyTable.rows = {};
        pptx.addContentSlide(sampleName + " - Charts", emptyTable,
                             {vqImg, consImg});

        QFile::remove(vqPath);
        QFile::remove(consPath);
    }

    if (!pptx.save(filePath)) {
        errorOut = "Failed to write file: " + filePath;
        return false;
    }
    return true;
}
```

**PptxWriter API:** Use the existing `addContentSlide(title, SlideTable, QVector<SlideImage>)` method.
- For the table slide: pass the `SlideTable` with an empty `plots` vector.
- For the plot slide: pass an empty `SlideTable` (or minimal) with two `SlideImage` entries positioned side by side (left half: x=0.3, w=6.0; right half: x=6.7, w=6.0).
- The `addCoverSlide()` method is already available for the title slide.
- Remember to call `pptx.setResourcePath(resourcePath())` before generating if resource images are needed for branding.

- [ ] **Step 3: Verify build**

Run: `cd build && qmake ../DataViewerEnterprise.pro && mingw32-make -j8 2>&1 | tail -10`
Expected: Build succeeds.

- [ ] **Step 4: Commit**

```bash
git add src/ui/DetailedSensoryPanel.cpp
git commit -m "feat: add PowerPoint report generation for detailed sensory mode"
```

---

### Task 9: Integration Testing and Polish

- [ ] **Step 1: End-to-end test — Create session**

Launch the app. Click "Detailed Sensory" in Tools tab. Click "New Session" in Home tab. Fill in:
- Test Title, Assessor, Tester, Media
- Fill in all 11 questions for Sample 1
- Click "+ Add Sample", fill in Sample 2
- Navigate with prev/next arrows
- Verify both radar charts update live

- [ ] **Step 2: Test save and reload**

Click "Save" — save as .xlsx. Close the session. Open Database Browser, verify "Detailed Sensory" tab shows the session. Click "Open Selected" to reload.

- [ ] **Step 3: Test report generation**

Click "Detailed Sensory Report" in Reports tab. Select sessions. Save as .pptx. Open the file and verify:
- Title slide
- Per-sample table slide with testers as rows and average row
- Per-sample plot slide with dual radar charts

- [ ] **Step 4: Test mode switching**

Verify seamless transitions:
- TPM -> Detailed Sensory -> TPM (no panel size changes)
- Sensory -> Detailed Sensory (mutual exclusion works)
- Detailed Sensory -> Sensory (mutual exclusion works)

- [ ] **Step 5: Test multi-tester averaging**

Load multiple sessions with different testers. Ctrl+click multiple in navigator. Verify averaged charts display correctly.

- [ ] **Step 6: Final commit**

```bash
git add -A
git commit -m "feat: detailed sensory mode — integration testing and polish"
```

---

## Dependency Graph

```
Task 1 (Data Structs)
  └─> Task 2 (RadarChart Extension)
  └─> Task 3 (Database CRUD)
  └─> Task 4 (Black Icon)

Task 1 + Task 2 + Task 3
  └─> Task 5 (DetailedSensoryPanel)
      └─> Task 6 (MainWindow Wiring)
          └─> Task 7 (Database Browser Tab)
          └─> Task 8 (Report Generation)
              └─> Task 9 (Integration Testing)
```

Tasks 1, 2, 3, 4 can be done in parallel. Task 5 depends on 1+2+3. Tasks 6, 7, 8 depend on 5. Task 9 is final.
