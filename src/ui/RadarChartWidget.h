#pragma once

#include <QWidget>
#include <QColor>
#include <QVector>
#include <QSet>
#include <QMap>
#include "pipeline/SensoryData.h"

namespace DVE {

class RadarChartWidget : public QWidget
{
    Q_OBJECT

public:
    explicit RadarChartWidget(QWidget* parent = nullptr);

    void setSessions(const QVector<SensorySession>& sessions);
    void setReportMode(bool reportMode) { m_reportMode = reportMode; }
    void setReportCropTop(int pixels) { m_reportCropTop = pixels; }
    void setReportCropBottom(int pixels) { m_reportCropBottom = pixels; }

    // ── Configurable-axes mode (for DetailedSensoryPanel) ───────────────
    void setCustomAxes(const QStringList& metricKeys,
                       const QMap<QString, QString>& axisLabels);
    void clearCustomAxes();

    struct SampleData { QString name; QMap<QString, double> scores; };
    void setCustomData(const QVector<SampleData>& samples);

protected:
    void paintEvent(QPaintEvent* event) override;
    void resizeEvent(QResizeEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;

private:
    QVector<SensorySession> m_sessions;
    QSet<int> m_hiddenSamples;  // global sample indices that are toggled off

    // Legend hit-testing
    struct LegendItem { QRect rect; int globalIdx; };
    QVector<LegendItem> m_legendItems;

    bool m_reportMode = false;
    int  m_reportCropTop = 0;
    int  m_reportCropBottom = 0;

    // Custom axes (empty = use default kSensoryMetricsPlot)
    QStringList m_customMetrics;
    QMap<QString, QString> m_customLabels;
    QVector<SampleData> m_customSamples;
    bool m_useCustomAxes = false;

    int axisCount() const;
    QStringList axisMetrics() const;
    QString axisLabel(int i) const;
    double sampleScore(const SampleData& sd, int axisIdx) const;

    QPointF axisPoint(int axisIndex, double value, QPointF center, double radius) const;
    void drawGrid(QPainter& p, QPointF center, double radius) const;
    void drawAxes(QPainter& p, QPointF center, double radius) const;
    // Drawn AFTER samples so the polygon outlines never cover the axis labels.
    void drawAxisLabels(QPainter& p, QPointF center, double radius) const;
    void drawSample(QPainter& p, const SensorySample& sample,
                    QPointF center, double radius, QColor color) const;
    void drawLegend(QPainter& p, const QRectF& legendRect);
    void drawLegendReport(QPainter& p, const QRectF& legendRect);
    void drawCustomSample(QPainter& p, const SampleData& sample,
                          QPointF center, double radius, QColor color) const;
};

} // namespace DVE
