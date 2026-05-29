#include "RadarChartWidget.h"

#include <QPainter>
#include <QPainterPath>
#include <QResizeEvent>
#include <QMouseEvent>
#include <QtMath>

#include "../utils/AppTheme.h"

namespace DVE {

RadarChartWidget::RadarChartWidget(QWidget* parent)
    : QWidget(parent)
{
    setMinimumSize(200, 200);
    setCursor(Qt::PointingHandCursor);
    // White background is painted explicitly in paintEvent; no palette setup
    // needed (palette-based fill conflicts with off-screen render(QPainter)).
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
    // Lighter + dashed inner rings — keeps the grid visible without competing
    // with the sample polygons. Score-9 ring stays solid as the chart
    // boundary so the outer edge reads as a frame, not as just another
    // grid step.
    QPen ringPen(QColor(225, 225, 225), 1, Qt::DashLine);
    p.setBrush(Qt::NoBrush);

    for (int score : ringScores) {
        ringPen.setWidthF(score == 9 ? 1.2 : 1.0);
        ringPen.setStyle(score == 9 ? Qt::SolidLine : Qt::DashLine);
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
    // Spokes only — labels are drawn AFTER samples by drawAxisLabels so they
    // never get covered by polygon outlines that reach score 8 or 9.
    const int n = axisCount();
    for (int i = 0; i < n; ++i) {
        const QPointF tip = axisPoint(i, 9, center, radius);
        p.setPen(QPen(QColor(150, 150, 150), 1));
        p.drawLine(center, tip);
    }
}

void RadarChartWidget::drawAxisLabels(QPainter& p, QPointF center, double radius) const
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
        const double angleDeg = 270.0 + (360.0 / n) * i;
        const double angleRad = qDegreesToRadians(angleDeg);
        const double cosA = qCos(angleRad);
        const double sinA = qSin(angleRad);

        // Burnt Taste (i=1) and Smoothness (i=4): place the label centre at
        // fraction t along the upper polygon edge from this vertex toward
        // Overall Liking, then offset perpendicular-outward from the edge by
        // `perp` pixels. The user marked the desired position with a green
        // box in v1.3.9 feedback — close to the polygon vertex (NOT extended
        // out to the chart edge), slightly above and along the upper edge.
        // t=0.08 and perp=25 hit the target position for the chart in the
        // screenshot.
        QPointF anchor;
        if (i == 1 || i == 4) {
            const double t = 0.08;
            const double perp = 12.5;
            const double thisVertexDeg = (i == 1) ? 342.0 : 198.0;
            const double olVertexDeg = 270.0;
            const QPointF V(radius * qCos(qDegreesToRadians(thisVertexDeg)),
                            radius * qSin(qDegreesToRadians(thisVertexDeg)));
            const QPointF OL(radius * qCos(qDegreesToRadians(olVertexDeg)),
                             radius * qSin(qDegreesToRadians(olVertexDeg)));
            const QPointF edgeAtT = V + t * (OL - V);
            const double normalDeg = (i == 1) ? 306.0 : 234.0;
            const QPointF normal(qCos(qDegreesToRadians(normalDeg)),
                                 qSin(qDegreesToRadians(normalDeg)));
            const QPointF labelCenter = edgeAtT + perp * normal;

            QFontMetrics fmTmp(labelFont);
            const QRect probe = fmTmp.boundingRect(
                QRect(0, 0, 220, 80),
                Qt::AlignCenter | Qt::TextWordWrap,
                axisLabel(i));
            const double labelHalfH = probe.height() / 2.0;
            anchor = QPointF(
                center.x() + labelCenter.x(),
                center.y() + labelCenter.y() + labelHalfH + 2.0);
        } else {
            anchor = QPointF(center.x() + (radius + 8) * cosA,
                             center.y() + (radius + 8) * sinA);
        }

        QFontMetrics fm(labelFont);
        QString label = axisLabel(i);
        QRect textRect = fm.boundingRect(QRect(0, 0, 220, 80),
                                         Qt::AlignCenter | Qt::TextWordWrap, label);

        // Position the label so it sits OUTSIDE the polygon vertically:
        //  - Upper-half axes (Overall Liking, Burnt Taste, Smoothness): bottom
        //    edge of the label sits just above the anchor.
        //  - Lower-half axes (Vapor Volume, Overall Flavor): top edge of the
        //    label sits just below the anchor.
        // Then clamp to widget bounds so no text is cut off the edges.
        QPoint topLeft;
        topLeft.setX(int(anchor.x() - textRect.width() / 2.0));
        if (sinA < -0.1) {
            topLeft.setY(int(anchor.y() - textRect.height() - 2));
        } else if (sinA > 0.1) {
            topLeft.setY(int(anchor.y() + 2));
        } else {
            topLeft.setY(int(anchor.y() - textRect.height() / 2.0));
        }
        constexpr int kEdgeMargin = 4;
        // In report mode the rendered pixmap is cropped from the top and
        // bottom — clamp labels to stay within the visible region so they
        // don't get sliced off.
        const int yLo = m_reportMode ? (m_reportCropTop + kEdgeMargin)
                                     : kEdgeMargin;
        const int yHi = m_reportMode
            ? (height() - m_reportCropBottom - textRect.height() - kEdgeMargin)
            : (height() - textRect.height() - kEdgeMargin);
        topLeft.setX(qBound(kEdgeMargin,
                             topLeft.x(),
                             width() - textRect.width() - kEdgeMargin));
        topLeft.setY(qBound(yLo, topLeft.y(), yHi));
        textRect.moveTopLeft(topLeft);

        // Burnt Taste (i=1, upper-right) and Smoothness (i=4, upper-left)
        // overlap with the polygon edge that meets their vertex when drawn
        // horizontally. Rotate each so the text runs parallel to that edge.
        // Geometry: the pentagon edge adjacent to a vertex makes a 36° angle
        // with horizontal (since interior pentagon angles are 108° and the
        // radial axis bisects the vertex, so the edge tilts 90 - 54 = 36°
        // from horizontal). Earlier 72° was a full axis-step, way too steep.
        //   * Smoothness: 36° CCW (negative QPainter rotation)
        //   * Burnt Taste: 36° CW (positive)
        const double rotDeg = (i == 1) ?  36.0
                            : (i == 4) ? -36.0
                                       :   0.0;

        if (rotDeg != 0.0) {
            // Rotate around the label's center so the anchor stays fixed.
            p.save();
            p.translate(textRect.center());
            p.rotate(rotDeg);
            QRectF localRect(-textRect.width() / 2.0,
                             -textRect.height() / 2.0,
                              textRect.width(),
                              textRect.height());
            p.setPen(Qt::NoPen);
            p.setBrush(Qt::white);
            p.drawRect(localRect.adjusted(-2, -1, 2, 1));
            p.setPen(QColor(40, 40, 40));
            p.drawText(localRect, Qt::AlignCenter | Qt::TextWordWrap, label);
            p.restore();
        } else {
            // Knock out a small white halo behind the text so the polygon
            // outline doesn't bleed through the descenders of the bold font.
            p.setPen(Qt::NoPen);
            p.setBrush(Qt::white);
            p.drawRect(textRect.adjusted(-2, -1, 2, 1));

            p.setPen(QColor(40, 40, 40));
            p.drawText(textRect, Qt::AlignCenter | Qt::TextWordWrap, label);
        }
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

    // Outline-only — matches the production-target chart style. Legend
    // colors disambiguate samples; the inner shading was visually noisy when
    // multiple samples overlapped.
    p.setBrush(Qt::NoBrush);
    p.setPen(QPen(color, 2.5));
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

    // Outline-only (see drawSample for rationale).
    p.setBrush(Qt::NoBrush);
    p.setPen(QPen(color, 2.5));
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
                            AppTheme::seriesColor(colorIdx),
                            colorIdx});
            ++colorIdx;
        }
    } else {
        for (const SensorySession& sess : m_sessions) {
            for (const SensorySample& sample : sess.samples) {
                entries.append({sample.name.isEmpty() ? QString("Sample %1").arg(colorIdx + 1)
                                                      : sample.name,
                                AppTheme::seriesColor(colorIdx),
                                colorIdx});
                ++colorIdx;
            }
        }
    }
    if (entries.isEmpty()) return;

    const int swatchSize = 14;
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

        // Swatch — solid colour fill, no outline (the outline previously drawn
        // with swatchColor.darker(150) on a 12 px square competed with the
        // fill at on-screen scales and read as outline-only).
        p.fillRect(x, y, swatchSize, swatchSize, swatchColor);

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
                            AppTheme::seriesColor(colorIdx),
                            colorIdx});
            ++colorIdx;
        }
    } else {
        for (const SensorySession& sess : m_sessions) {
            for (const SensorySample& sample : sess.samples) {
                entries.append({sample.name.isEmpty()
                                    ? QString("Sample %1").arg(colorIdx + 1)
                                    : sample.name,
                                AppTheme::seriesColor(colorIdx),
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

        // Swatch — solid colour fill, no outline (matches drawLegend's bottom
        // swatch styling so the two legend paths look consistent).
        p.fillRect(static_cast<int>(legendRect.left()) + margin, y,
                   swatchSize, swatchSize, swatchColor);

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
    // White background — flush with the rest of the sensory panel so the
    // chart doesn't look like it's sitting on a different-colored surface.
    p.fillRect(rect(), Qt::white);

    if (m_reportMode) {
        // Report mode: legend in top-right panel, chart nearly edge-to-edge.
        // The chart area excludes the crop region so labels and polygon both
        // sit within the visible portion of the rendered pixmap. Without this,
        // bottom labels (Vapor Volume, Overall Flavor) and the top label
        // (Overall Liking) end up positioned in the cropped-away region.
        const int legendW = qMax(220, width() / 3);
        const int chartH  = qMax(0, height() - m_reportCropTop - m_reportCropBottom);
        QRectF chartArea(0, m_reportCropTop, width() - legendW, chartH);
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
                                     AppTheme::seriesColor(ci));
            }
        } else {
            int colorIdx = 0;
            for (const SensorySession& sess : m_sessions) {
                for (const SensorySample& sample : sess.samples) {
                    if (!m_hiddenSamples.contains(colorIdx))
                        drawSample(p, sample, center, radius, AppTheme::seriesColor(colorIdx));
                    ++colorIdx;
                }
            }
        }

        // Labels on top so the polygon outlines never cover them.
        drawAxisLabels(p, center, radius);

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
                                     AppTheme::seriesColor(ci));
            }
        } else {
            int colorIdx = 0;
            for (const SensorySession& sess : m_sessions) {
                for (const SensorySample& sample : sess.samples) {
                    if (!m_hiddenSamples.contains(colorIdx))
                        drawSample(p, sample, center, radius, AppTheme::seriesColor(colorIdx));
                    ++colorIdx;
                }
            }
        }

        // Labels on top so the polygon outlines never cover them.
        drawAxisLabels(p, center, radius);

        p.setPen(QColor(200, 200, 200));
        p.drawLine(QPointF(0, legendArea.top()), QPointF(width(), legendArea.top()));

        drawLegend(p, legendArea);
    }
}

} // namespace DVE
