#include "RadarChartWidget.h"

#include <QPainter>
#include <QPainterPath>
#include <QResizeEvent>
#include <QMouseEvent>
#include <QtMath>

namespace DVE {

// 10 preset fill colors
const QList<QColor> RadarChartWidget::kColors = {
    QColor(0,   114, 189),
    QColor(217,  83,  25),
    QColor(237, 177,  32),
    QColor(126,  47, 142),
    QColor(119, 172,  48),
    QColor( 77, 190, 238),
    QColor(162,  20,  47),
    QColor( 76,  76,  76),
    QColor(153, 153,   0),
    QColor(  0, 128, 128),
};

RadarChartWidget::RadarChartWidget(QWidget* parent)
    : QWidget(parent)
{
    setMinimumSize(200, 200);
    setCursor(Qt::PointingHandCursor);
}

void RadarChartWidget::setSessions(const QVector<SensorySession>& sessions)
{
    m_sessions = sessions;
    // Prune hidden indices that no longer exist
    int total = 0;
    for (const auto& s : sessions) total += s.samples.size();
    QSet<int> pruned;
    for (int idx : m_hiddenSamples)
        if (idx < total) pruned.insert(idx);
    m_hiddenSamples = pruned;
    update();
}

// ─────────────────────────────────────────────────────────────────────────────
// Geometry helpers
// ─────────────────────────────────────────────────────────────────────────────

QPointF RadarChartWidget::axisPoint(int axisIndex, double value,
                                    QPointF center, double radius) const
{
    int n = kSensoryMetricsPlot.size();
    double angleDeg = 270.0 + (360.0 / n) * axisIndex;
    double angleRad = qDegreesToRadians(angleDeg);
    double t = (value - 1.0) / 8.0;
    double r = t * radius;
    return QPointF(center.x() + r * qCos(angleRad),
                   center.y() + r * qSin(angleRad));
}

// ─────────────────────────────────────────────────────────────────────────────
// Drawing
// ─────────────────────────────────────────────────────────────────────────────

void RadarChartWidget::drawGrid(QPainter& p, QPointF center, double radius) const
{
    int n = kSensoryMetricsPlot.size();
    const QList<int> ringScores = {3, 5, 7, 9};
    QPen ringPen(QColor(200, 200, 200), 1, Qt::SolidLine);
    p.setPen(ringPen);
    p.setBrush(Qt::NoBrush);

    for (int score : ringScores) {
        QPolygonF poly;
        for (int i = 0; i < n; ++i)
            poly << axisPoint(i, score, center, radius);
        p.drawPolygon(poly);
    }
}

void RadarChartWidget::drawAxes(QPainter& p, QPointF center, double radius) const
{
    int n = kSensoryMetricsPlot.size();

    QFont labelFont = p.font();
    labelFont.setPointSize(8);
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
        QString label = kSensoryMetricsPlot[i];
        QRect textRect = fm.boundingRect(label);
        textRect.moveCenter(labelCenter.toPoint());

        p.setPen(QColor(40, 40, 40));
        p.drawText(textRect, Qt::AlignCenter, label);
    }
}

void RadarChartWidget::drawSample(QPainter& p, const SensorySample& sample,
                                  QPointF center, double radius, QColor color) const
{
    int n = kSensoryMetricsPlot.size();
    QPolygonF poly;
    for (int i = 0; i < n; ++i) {
        int score = sample.scores.value(kSensoryMetricsPlot[i], 5);
        poly << axisPoint(i, score, center, radius);
    }

    QColor fillColor = color;
    fillColor.setAlphaF(0.18);
    p.setBrush(fillColor);
    p.setPen(QPen(color, 2));
    p.drawPolygon(poly);
}

void RadarChartWidget::drawLegend(QPainter& p, const QRectF& legendRect)
{
    m_legendItems.clear();

    struct Entry { QString name; QColor color; int globalIdx; };
    QList<Entry> entries;

    int colorIdx = 0;
    for (const SensorySession& sess : m_sessions) {
        for (const SensorySample& sample : sess.samples) {
            entries.append({sample.name.isEmpty() ? QString("Sample %1").arg(colorIdx + 1)
                                                  : sample.name,
                            kColors[colorIdx % kColors.size()],
                            colorIdx});
            ++colorIdx;
        }
    }
    if (entries.isEmpty()) return;

    const int swatchSize = 12;
    const int spacing    = 6;
    const int entryWidth = 200;

    QFont f = p.font();
    f.setPointSize(8);
    p.setFont(f);

    int x = static_cast<int>(legendRect.left()) + 4;
    int y = static_cast<int>(legendRect.top()) + 4;

    for (const Entry& e : entries) {
        bool hidden = m_hiddenSamples.contains(e.globalIdx);

        QColor swatchColor = hidden ? QColor(200, 200, 200) : e.color;
        QColor textColor   = hidden ? QColor(170, 170, 170) : QColor(40, 40, 40);

        // Swatch
        p.fillRect(x, y, swatchSize, swatchSize, swatchColor);
        p.setPen(swatchColor.darker(150));
        p.drawRect(x, y, swatchSize, swatchSize);

        // Strikethrough indicator for hidden
        if (hidden) {
            p.setPen(QPen(QColor(150, 150, 150), 1));
            p.drawLine(x, y + swatchSize / 2, x + swatchSize, y + swatchSize / 2);
        }

        // Label
        p.setPen(textColor);
        if (hidden) {
            QFont strikeFont = f;
            strikeFont.setStrikeOut(true);
            p.setFont(strikeFont);
        }
        p.drawText(x + swatchSize + 3, y, entryWidth - swatchSize - 3, swatchSize + 4,
                   Qt::AlignVCenter | Qt::AlignLeft, e.name);
        if (hidden) p.setFont(f);

        // Store hit rect for click detection
        QRect hitRect(x, y, entryWidth, swatchSize + 4);
        m_legendItems.append({hitRect, e.globalIdx});

        x += entryWidth + spacing;
        if (x + entryWidth > legendRect.right()) {
            x = static_cast<int>(legendRect.left()) + 4;
            y += swatchSize + spacing;
        }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Events
// ─────────────────────────────────────────────────────────────────────────────

void RadarChartWidget::resizeEvent(QResizeEvent* event)
{
    QWidget::resizeEvent(event);
    update();
}

void RadarChartWidget::mousePressEvent(QMouseEvent* event)
{
    QPoint pos = event->pos();
    for (const LegendItem& item : m_legendItems) {
        if (item.rect.contains(pos)) {
            if (m_hiddenSamples.contains(item.globalIdx))
                m_hiddenSamples.remove(item.globalIdx);
            else
                m_hiddenSamples.insert(item.globalIdx);
            update();
            return;
        }
    }
    QWidget::mousePressEvent(event);
}

void RadarChartWidget::paintEvent(QPaintEvent* /*event*/)
{
    QPainter p(this);
    p.setRenderHint(QPainter::Antialiasing);

    p.fillRect(rect(), QColor(248, 248, 248));

    const int legendH = 40;
    QRectF chartArea(0, 0, width(), height() - legendH);
    QRectF legendArea(0, height() - legendH, width(), legendH);

    double margin = 48.0;
    double radius = (qMin(chartArea.width(), chartArea.height()) / 2.0) - margin;
    if (radius < 20) return;

    QPointF center(chartArea.left() + chartArea.width() / 2.0,
                   chartArea.top()  + chartArea.height() / 2.0);

    drawGrid(p, center, radius);
    drawAxes(p, center, radius);

    // Draw visible samples
    int colorIdx = 0;
    for (const SensorySession& sess : m_sessions) {
        for (const SensorySample& sample : sess.samples) {
            if (!m_hiddenSamples.contains(colorIdx)) {
                QColor c = kColors[colorIdx % kColors.size()];
                drawSample(p, sample, center, radius, c);
            }
            ++colorIdx;
        }
    }

    // Legend separator
    p.setPen(QColor(200, 200, 200));
    p.drawLine(QPointF(0, legendArea.top()), QPointF(width(), legendArea.top()));

    drawLegend(p, legendArea);
}

} // namespace DVE
