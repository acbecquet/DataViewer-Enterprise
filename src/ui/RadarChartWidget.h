#pragma once

#include <QWidget>
#include <QVector>
#include <QSet>
#include "pipeline/SensoryData.h"

namespace DVE {

class RadarChartWidget : public QWidget
{
    Q_OBJECT

public:
    explicit RadarChartWidget(QWidget* parent = nullptr);

    void setSessions(const QVector<SensorySession>& sessions);

protected:
    void paintEvent(QPaintEvent* event) override;
    void resizeEvent(QResizeEvent* event) override;
    void mousePressEvent(QMouseEvent* event) override;

private:
    QVector<SensorySession> m_sessions;
    QSet<int> m_hiddenSamples;  // global sample indices that are toggled off

    static const QList<QColor> kColors;

    // Legend hit-testing
    struct LegendItem { QRect rect; int globalIdx; };
    QVector<LegendItem> m_legendItems;

    QPointF axisPoint(int axisIndex, double value, QPointF center, double radius) const;
    void drawGrid(QPainter& p, QPointF center, double radius) const;
    void drawAxes(QPainter& p, QPointF center, double radius) const;
    void drawSample(QPainter& p, const SensorySample& sample,
                    QPointF center, double radius, QColor color) const;
    void drawLegend(QPainter& p, const QRectF& legendRect);
};

} // namespace DVE
