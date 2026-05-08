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

// ─────────────────────────────────────────────────────────────────────────────
// Geometry helpers
// ─────────────────────────────────────────────────────────────────────────────

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

// ─────────────────────────────────────────────────────────────────────────────
// Drawing
// ─────────────────────────────────────────────────────────────────────────────

void RadarChartWidget::drawGrid(QPainter& p, QPointF center, double radius) const
{
    int n = axisCount();
    // Fine-grained gridlines at every integer 1-9 give the chart the same
    // visual density as the production target reference. Outermost ring (9)
    // is drawn slightly heavier as the chart boundary.
    const QList<int> ringScores = {1, 2, 3, 4, 5, 6, 7, 8, 9};
    QPen ringPen(QColor(200, 200, 200), 1, Qt::SolidLine);
    p.setBrush(Qt::NoBrush);

    for (int score : ringScores) {
        ringPen.setWidthF(score == 9 ? 1.2 : 1.0);
        p.setPen(ringPen);
        QPolygonF poly;
        for (int i = 0; i < n; ++i)
            poly << axisPoint(i, score, center, radius);
        p.drawPolygon(poly);
    }

    // Scale labels along the first spoke (Overall Liking, 12-o'clock).
    // Bigger, bolder font in report mode so the numbers read clearly in PPTX.
    QFont scaleFont = p.font();
    scaleFont.setPointSize(m_reportMode ? 14 : 10);
    scaleFont.setBold(true);
    p.setFont(scaleFont);
    p.setPen(QColor(80, 80, 80));

    for (int score : {1, 2, 3, 4, 5, 6, 7, 8, 9}) {
        QPointF pt = axisPoint(0, score, center, radius);
        const int boxW = m_reportMode ? 30 : 22;
        const int boxH = m_reportMode ? 22 : 16;
        p.drawText(QRectF(pt.x() - boxW - 2, pt.y() - boxH/2.0, boxW, boxH),
                   Qt::AlignRight | Qt::AlignVCenter,
                   QString::number(score));
    }
}

void RadarChartWidget::drawAxes(QPainter& p, QPointF center, double radius) const
{
    int n = axisCount();

    // Big bold axis labels in both modes; report mode gets the larger size to
    // match the production target reference. Custom-axis mode uses a smaller
    // size since it can carry longer labels.
    QFont labelFont = p.font();
    labelFont.setPointSize(m_reportMode ? 20 : (m_useCustomAxes ? 10 : 12));
    labelFont.setBold(true);
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
        QRect textRect = fm.boundingRect(QRect(0, 0, 200, 200),
                                         Qt::AlignCenter | Qt::TextWordWrap, label);
        textRect.moveCenter(labelCenter.toPoint());

        p.setPen(QColor(40, 40, 40));
        p.drawText(textRect, Qt::AlignCenter | Qt::TextWordWrap, label);
    }
}

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

void RadarChartWidget::drawLegend(QPainter& p, const QRectF& legendRect)
{
    m_legendItems.clear();

    struct Entry { QString name; QColor color; int globalIdx; };
    QList<Entry> entries;

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

void RadarChartWidget::drawLegendReport(QPainter& p, const QRectF& legendRect)
{
    struct Entry { QString name; QColor color; int globalIdx; };
    QList<Entry> entries;
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
                entries.append({sample.name.isEmpty()
                                    ? QString("Sample %1").arg(colorIdx + 1)
                                    : sample.name,
                                kColors[colorIdx % kColors.size()],
                                colorIdx});
                ++colorIdx;
            }
        }
    }
    if (entries.isEmpty()) return;

    const int swatchSize = 20;
    const int rowH       = 34;
    const int margin     = 10;

    // Title
    QFont titleFont = p.font();
    titleFont.setBold(true);
    titleFont.setPointSize(20);
    p.setFont(titleFont);
    p.setPen(QColor(40, 40, 40));
    int titleH = 36;
    p.drawText(QRectF(legendRect.left() + margin, legendRect.top() + margin,
                      legendRect.width() - 2 * margin, titleH),
               Qt::AlignLeft | Qt::AlignVCenter, "Samples");

    QFont f = p.font();
    f.setBold(false);
    f.setPointSize(18);
    p.setFont(f);

    int y = static_cast<int>(legendRect.top()) + margin + titleH + 6;

    for (const Entry& e : entries) {
        bool hidden = m_hiddenSamples.contains(e.globalIdx);
        QColor swatchColor = hidden ? QColor(200, 200, 200) : e.color;
        QColor textColor   = hidden ? QColor(170, 170, 170) : QColor(40, 40, 40);

        // Swatch
        p.fillRect(static_cast<int>(legendRect.left()) + margin, y,
                   swatchSize, swatchSize, swatchColor);
        p.setPen(swatchColor.darker(150));
        p.drawRect(static_cast<int>(legendRect.left()) + margin, y, swatchSize, swatchSize);

        // Label
        p.setPen(textColor);
        p.drawText(static_cast<int>(legendRect.left()) + margin + swatchSize + 6,
                   y,
                   static_cast<int>(legendRect.width()) - margin - swatchSize - 10,
                   rowH,
                   Qt::AlignVCenter | Qt::AlignLeft,
                   e.name);

        y += rowH;
        if (y + rowH > legendRect.bottom()) break;
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

    if (m_reportMode) {
        // Report mode: legend in top-right panel, chart nearly edge-to-edge
        const int legendW = qMax(220, width() / 3);
        QRectF chartArea(0, 0, width() - legendW, height());
        // Offset legend top by crop amount so it survives top cropping
        QRectF legendArea(width() - legendW, m_reportCropTop, legendW, height() - m_reportCropTop);

        double margin = 14.0;
        double labelSpace = 46.0;  // room for axis labels outside the grid
        double radius = (qMin(chartArea.width(), chartArea.height()) / 2.0) - margin - labelSpace;
        if (radius < 20) return;

        QPointF center(chartArea.left() + chartArea.width() / 2.0,
                       chartArea.top()  + chartArea.height() / 2.0);

        drawGrid(p, center, radius);
        drawAxes(p, center, radius);

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

        // Vertical separator between chart and legend
        p.setPen(QColor(200, 200, 200));
        p.drawLine(QPointF(legendArea.left(), 0), QPointF(legendArea.left(), height()));

        drawLegendReport(p, legendArea);
    } else {
        // Normal UI mode: legend at bottom (unchanged)
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

        p.setPen(QColor(200, 200, 200));
        p.drawLine(QPointF(0, legendArea.top()), QPointF(width(), legendArea.top()));

        drawLegend(p, legendArea);
    }
}

} // namespace DVE
