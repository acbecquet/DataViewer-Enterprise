#include "PlotEngine.h"

#include "../utils/AppTheme.h"
#include <QPainter>
#include <QPen>
#include <QBrush>
#include <QBuffer>
#include <QImageWriter>
#include <QImage>
#include <QFontMetrics>
#include <QRectF>
#include <QtMath>
#include <algorithm>
#include <cmath>
#include <limits>

namespace DVE {

// ═══════════════════════════════════════════════════════════════════════════════
// Internal helpers
// ═══════════════════════════════════════════════════════════════════════════════

QPointF PlotEngine::dataToPixel(double x, double y,
                                double xMin, double xMax,
                                double yMin, double yMax,
                                int pxLeft, int pxRight,
                                int pxTop,  int pxBottom)
{
    if (qFuzzyCompare(xMax, xMin)) xMax = xMin + 1.0;
    if (qFuzzyCompare(yMax, yMin)) yMax = yMin + 1.0;

    double px = pxLeft  + (x - xMin) / (xMax - xMin) * (pxRight  - pxLeft);
    double py = pxBottom - (y - yMin) / (yMax - yMin) * (pxBottom - pxTop);
    return QPointF(px, py);
}

double PlotEngine::niceStep(double range, int targetTicks)
{
    if (range <= 0.0) return 1.0;
    double rawStep = range / targetTicks;
    double mag     = std::pow(10.0, std::floor(std::log10(rawStep)));
    double norm    = rawStep / mag;  // in [1, 10)

    double niceMult;
    if      (norm < 1.5) niceMult = 1.0;
    else if (norm < 3.0) niceMult = 2.0;
    else if (norm < 7.0) niceMult = 5.0;
    else                 niceMult = 10.0;

    return niceMult * mag;
}

void PlotEngine::autoRange(const QVector<PlotSeries>& series,
                           double& xMin, double& xMax,
                           double& yMin, double& yMax)
{
    xMin = yMin =  std::numeric_limits<double>::max();
    xMax = yMax = -std::numeric_limits<double>::max();

    for (const auto& s : series) {
        for (double v : s.x) { xMin = qMin(xMin, v); xMax = qMax(xMax, v); }
        for (double v : s.y) { yMin = qMin(yMin, v); yMax = qMax(yMax, v); }
    }

    if (xMin > xMax) { xMin = 0; xMax = 0; }
    if (yMin > yMax) { yMin = 0; yMax = 0; }

    // Standing axis rules for every x-y plot: both axes anchored at 0 and
    // topping out at the data maximum + 1, so no point sits on an edge.
    xMin = 0.0;
    xMax = xMax + 1.0;
    yMin = 0.0;
    yMax = yMax + 1.0;
}

void PlotEngine::applyDataXRange(PlotConfig& cfg,
                                 const QVector<PlotSeries>& series)
{
    double hi = 0.0;
    for (const PlotSeries& s : series)
        for (double v : s.x) hi = qMax(hi, v);

    cfg.xMin = 0.0;
    cfg.xMax = hi + 1.0;   // +1 so the last point never sits on the edge
}

void PlotEngine::applyAnchoredYRange(PlotConfig& cfg,
                                     const QVector<PlotSeries>& series)
{
    double hi = 0.0;
    for (const PlotSeries& s : series)
        for (double v : s.y) hi = qMax(hi, v);

    cfg.yMin = 0.0;
    cfg.yMax = hi + 1.0;
}

QString PlotEngine::formatTickLabel(double v)
{
    // Use integer format if the value is close to a whole number, otherwise
    // up to 3 significant decimal digits.
    if (v == 0.0) return "0";
    double absV = std::abs(v);

    if (absV >= 1000.0)
        return QString::number(v, 'f', 0);
    if (absV >= 100.0)
        return QString::number(v, 'f', 1);
    if (absV >= 1.0)
        return QString::number(v, 'f', (std::fmod(v, 1.0) < 1e-9) ? 0 : 2);
    // Small values
    return QString::number(v, 'g', 3);
}

// ── Grid ──────────────────────────────────────────────────────────────────────

void PlotEngine::drawGrid(QPainter& p, const PlotConfig& cfg,
                          int pxLeft, int pxRight, int pxTop, int pxBottom,
                          double xMin, double xMax, double yMin, double yMax)
{
    p.save();
    QPen gridPen(cfg.gridColor, 1, Qt::SolidLine);
    p.setPen(gridPen);

    // Horizontal grid lines (Y ticks)
    double yStep = niceStep(yMax - yMin);
    for (double yv = std::ceil(yMin / yStep) * yStep; yv <= yMax + 1e-9; yv += yStep) {
        QPointF pt = dataToPixel(xMin, yv, xMin, xMax, yMin, yMax,
                                 pxLeft, pxRight, pxTop, pxBottom);
        p.drawLine(QPointF(pxLeft, pt.y()), QPointF(pxRight, pt.y()));
    }

    // Vertical grid lines (X ticks)
    double xStep = niceStep(xMax - xMin);
    for (double xv = std::ceil(xMin / xStep) * xStep; xv <= xMax + 1e-9; xv += xStep) {
        QPointF pt = dataToPixel(xv, yMin, xMin, xMax, yMin, yMax,
                                 pxLeft, pxRight, pxTop, pxBottom);
        p.drawLine(QPointF(pt.x(), pxTop), QPointF(pt.x(), pxBottom));
    }

    p.restore();
}

// ── Axes ──────────────────────────────────────────────────────────────────────

void PlotEngine::drawAxes(QPainter& p, const PlotConfig& cfg,
                          int pxLeft, int pxRight, int pxTop, int pxBottom,
                          double xMin, double xMax, double yMin, double yMax)
{
    p.save();

    QPen axisPen(cfg.axisColor, 2, Qt::SolidLine);
    p.setPen(axisPen);

    // Bottom (X) axis
    p.drawLine(QPointF(pxLeft, pxBottom), QPointF(pxRight, pxBottom));
    // Left (Y) axis
    p.drawLine(QPointF(pxLeft, pxTop),    QPointF(pxLeft,  pxBottom));

    // ── Y tick marks and labels ───────────────────────────────────────────────
    p.setFont(cfg.labelFont);
    QFontMetrics fm(cfg.labelFont);

    double yStep = niceStep(yMax - yMin);
    QPen tickPen(cfg.axisColor, 1);
    p.setPen(tickPen);

    for (double yv = std::ceil(yMin / yStep) * yStep; yv <= yMax + 1e-9; yv += yStep) {
        QPointF pt = dataToPixel(xMin, yv, xMin, xMax, yMin, yMax,
                                 pxLeft, pxRight, pxTop, pxBottom);
        int py = qRound(pt.y());
        // Tick mark
        p.drawLine(pxLeft - 5, py, pxLeft, py);
        // Label
        QString lbl = formatTickLabel(yv);
        int tw = fm.horizontalAdvance(lbl);
        int th = fm.height();
        p.setPen(cfg.axisColor);
        p.drawText(QRect(pxLeft - 8 - tw, py - th / 2, tw, th),
                   Qt::AlignRight | Qt::AlignVCenter, lbl);
        p.setPen(tickPen);
    }

    // ── X tick marks and labels ───────────────────────────────────────────────
    double xStep = niceStep(xMax - xMin);
    for (double xv = std::ceil(xMin / xStep) * xStep; xv <= xMax + 1e-9; xv += xStep) {
        QPointF pt = dataToPixel(xv, yMin, xMin, xMax, yMin, yMax,
                                 pxLeft, pxRight, pxTop, pxBottom);
        int px = qRound(pt.x());
        // Tick mark
        p.drawLine(px, pxBottom, px, pxBottom + 5);
        // Label
        QString lbl = formatTickLabel(xv);
        int tw = fm.horizontalAdvance(lbl);
        int th = fm.height();
        p.setPen(cfg.axisColor);
        p.drawText(QRect(px - tw / 2, pxBottom + 7, tw, th),
                   Qt::AlignHCenter | Qt::AlignTop, lbl);
        p.setPen(tickPen);
    }

    // ── Axis labels ───────────────────────────────────────────────────────────
    p.setFont(cfg.axisFont);
    QFontMetrics afm(cfg.axisFont);

    if (!cfg.xLabel.isEmpty()) {
        int lw = afm.horizontalAdvance(cfg.xLabel);
        int midX = (pxLeft + pxRight) / 2;
        p.setPen(cfg.axisColor);
        // Position below the X tick labels (which sit at pxBottom+7 with height fm.height())
        int xLabelY = pxBottom + 7 + fm.height() + 6;
        p.drawText(QRect(midX - lw / 2, xLabelY, lw, afm.height() + 2),
                   Qt::AlignHCenter | Qt::AlignTop, cfg.xLabel);
    }

    if (!cfg.yLabel.isEmpty()) {
        int midY = (pxTop + pxBottom) / 2;
        p.save();
        // Center the rotated label so its full height fits within the left margin.
        int yLabelX = qMax(14, afm.height() / 2 + 6);
        p.translate(yLabelX, midY);
        p.rotate(-90);
        int lw = afm.horizontalAdvance(cfg.yLabel);
        p.setPen(cfg.axisColor);
        p.drawText(QRect(-lw / 2, -afm.height(), lw, afm.height() + 2),
                   Qt::AlignHCenter | Qt::AlignBottom, cfg.yLabel);
        p.restore();
    }

    p.restore();
}

// ── Legend ────────────────────────────────────────────────────────────────────

void PlotEngine::drawLegend(QPainter& p, const QVector<PlotSeries>& series,
                            const PlotConfig& cfg,
                            int pxRight, int /*pxTop*/)
{
    if (series.isEmpty()) return;

    p.save();
    p.setFont(cfg.labelFont);
    QFontMetrics fm(cfg.labelFont);

    const int swatchW  = 20;
    const int padX     = 8;
    const int padY     = 4;
    const int rowH     = fm.height() + 4;
    const int itemGap  = 16;

    int availW   = pxRight - cfg.marginLeft;

    // Measure items and wrap into rows
    struct LegItem { int x; int y; int w; int seriesIdx; };
    QVector<LegItem> items;
    int lx = 0, ly = 0;
    for (int i = 0; i < series.size(); ++i) {
        int itemW = swatchW + 6 + fm.horizontalAdvance(series[i].label) + itemGap;
        if (lx + itemW > availW && lx > 0) {
            lx = 0;
            ly += rowH;
        }
        items.append({ lx, ly, itemW, i });
        lx += itemW;
    }

    int legendH = ly + rowH;
    int bx = cfg.marginLeft;
    // Anchor legend to the bottom of the canvas so it sits below the X-axis labels.
    int by = cfg.height - 8 - legendH;

    // Semi-transparent background
    p.setBrush(QColor(255, 255, 255, 200));
    p.setPen(QColor(0xBC, 0xBC, 0xBC));
    p.drawRoundedRect(QRect(bx - padX, by - padY,
                             availW + padX * 2, legendH + padY * 2), 3, 3);

    for (const auto& item : items) {
        const auto& s = series[item.seriesIdx];
        int ix = bx + item.x;
        int iy = by + item.y;
        int swatchY = iy + rowH / 2;

        // Color swatch (line)
        QPen swatchPen(s.color, 3, s.dashed ? Qt::DashLine : Qt::SolidLine);
        p.setPen(swatchPen);
        p.drawLine(ix, swatchY, ix + swatchW, swatchY);

        // Dot
        if (s.drawDots) {
            p.setBrush(s.color);
            p.setPen(Qt::NoPen);
            p.drawEllipse(QPoint(ix + swatchW / 2, swatchY), 3, 3);
        }

        // Label
        p.setPen(cfg.axisColor);
        p.setBrush(Qt::NoBrush);
        p.setFont(cfg.labelFont);
        p.drawText(ix + swatchW + 6, iy,
                   fm.horizontalAdvance(s.label), rowH,
                   Qt::AlignLeft | Qt::AlignVCenter, s.label);
    }

    p.restore();
}


// ═══════════════════════════════════════════════════════════════════════════════
// renderLinePlot
// ═══════════════════════════════════════════════════════════════════════════════

QPixmap PlotEngine::renderLinePlot(const QVector<PlotSeries>& series,
                                   const PlotConfig&           config,
                                   PlotTransform*               outTransform)
{
    QPixmap pm(config.width, config.height);
    pm.fill(config.bgColor);

    if (series.isEmpty())
        return pm;

    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing,    true);
    p.setRenderHint(QPainter::TextAntialiasing, true);

    // ── Plot area pixel bounds ─────────────────────────────────────────────────
    // Reserve extra bottom margin for legend when shown
    int legendExtraH = 0;
    if (config.showLegend && !series.isEmpty()) {
        QFontMetrics lfm(config.labelFont);
        int rowH = lfm.height() + 4;
        int availW = config.width - config.marginLeft - config.marginRight;
        int lx = 0, rows = 1;
        for (const auto& s : series) {
            int itemW = 20 + 6 + lfm.horizontalAdvance(s.label) + 16;
            if (lx + itemW > availW && lx > 0) { lx = 0; ++rows; }
            lx += itemW;
        }
        legendExtraH = rows * rowH + 12;
    }

    // Bottom area needs to fit X tick labels and (optionally) the X-axis title.
    QFontMetrics tickFm(config.labelFont);
    QFontMetrics axisFm(config.axisFont);
    int xAxisLabelsH = 7 + tickFm.height()
                     + (config.xLabel.isEmpty() ? 0 : (6 + axisFm.height()))
                     + 6;

    int pxLeft   = config.marginLeft;
    int pxRight  = config.width  - config.marginRight;
    int pxTop    = config.marginTop;
    int pxBottom = config.height
                 - qMax(config.marginBottom, xAxisLabelsH)
                 - legendExtraH;

    // ── Data ranges ───────────────────────────────────────────────────────────
    double xMin, xMax, yMin, yMax;
    if (config.autoScale) {
        autoRange(series, xMin, xMax, yMin, yMax);
    } else {
        xMin = config.xMin; xMax = config.xMax;
        yMin = config.yMin; yMax = config.yMax;
    }

    // ── Title ─────────────────────────────────────────────────────────────────
    if (!config.title.isEmpty()) {
        p.setFont(config.titleFont);
        QFontMetrics tfm(config.titleFont);
        p.setPen(config.titleColor);
        p.drawText(QRect(pxLeft, 8, pxRight - pxLeft, pxTop - 8),
                   Qt::AlignHCenter | Qt::AlignVCenter, config.title);
    }

    // ── Plot background ───────────────────────────────────────────────────────
    p.fillRect(QRect(pxLeft, pxTop, pxRight - pxLeft, pxBottom - pxTop),
               config.bgColor);

    // ── Grid ──────────────────────────────────────────────────────────────────
    if (config.showGrid)
        drawGrid(p, config, pxLeft, pxRight, pxTop, pxBottom,
                 xMin, xMax, yMin, yMax);

    // ── Clip to plot area for drawing series ──────────────────────────────────
    p.setClipRect(QRect(pxLeft, pxTop, pxRight - pxLeft + 1, pxBottom - pxTop + 1));

    // ── Series ────────────────────────────────────────────────────────────────
    for (const auto& s : series) {
        if (s.x.isEmpty() || s.y.isEmpty()) continue;

        int n = qMin(s.x.size(), s.y.size());

        // Build pixel path
        QVector<QPointF> pts;
        pts.reserve(n);
        for (int i = 0; i < n; ++i)
            pts.append(dataToPixel(s.x[i], s.y[i],
                                   xMin, xMax, yMin, yMax,
                                   pxLeft, pxRight, pxTop, pxBottom));

        // Draw line
        if (s.drawLine && n >= 2) {
            Qt::PenStyle style = s.dashed ? Qt::DashLine : Qt::SolidLine;
            QPen linePen(s.color, s.lineWidth, style, Qt::RoundCap, Qt::RoundJoin);
            p.setPen(linePen);
            p.setBrush(Qt::NoBrush);
            for (int i = 0; i < n - 1; ++i)
                p.drawLine(pts[i], pts[i + 1]);
        }

        // Draw dots
        if (s.drawDots) {
            p.setPen(QPen(s.color.darker(130), 1));
            p.setBrush(s.color);
            int r = s.dotRadius;
            for (const auto& pt : pts)
                p.drawEllipse(pt, r, r);
        }

        // ── Note→plot linking (DATAVIEWER-6 v1) ────────────────────────────────
        // Amber ring on every note-bearing point, and a larger ring + a dashed
        // vertical guide for the emphasized (clicked) point.
        const QColor amber(0xBA, 0x75, 0x17);
        const int ringR = s.dotRadius + 5;
        for (int idx : s.ringed) {
            if (idx < 0 || idx >= pts.size()) continue;
            p.setPen(QPen(amber, 2));
            p.setBrush(Qt::NoBrush);
            const QPointF pt(qRound(pts[idx].x()), qRound(pts[idx].y()));
            p.drawEllipse(pt, ringR, ringR);
        }
        if (s.emphasized >= 0 && s.emphasized < pts.size()) {
            const QPointF pt(qRound(pts[s.emphasized].x()),
                             qRound(pts[s.emphasized].y()));
            // Dashed guide down to the x-axis baseline.
            p.setPen(QPen(amber, 1, Qt::DashLine));
            p.drawLine(QPointF(pt.x(), pt.y()), QPointF(pt.x(), pxBottom));
            // Larger solid ring on top.
            p.setPen(QPen(amber, 2));
            p.setBrush(Qt::NoBrush);
            p.drawEllipse(pt, ringR + 3, ringR + 3);
        }
    }

    p.setClipping(false);

    // ── Axes (drawn on top of everything) ─────────────────────────────────────
    drawAxes(p, config, pxLeft, pxRight, pxTop, pxBottom,
             xMin, xMax, yMin, yMax);

    // ── Legend ────────────────────────────────────────────────────────────────
    if (config.showLegend && series.size() > 1) {
        // Only show legend when there are multiple series
        drawLegend(p, series, config, pxRight, pxTop);
    } else if (config.showLegend && !series.isEmpty() && !series[0].label.isEmpty()) {
        drawLegend(p, series, config, pxRight, pxTop);
    }

    p.end();

    if (outTransform) {
        outTransform->plotRect = QRectF(QPointF(pxLeft, pxTop), QPointF(pxRight, pxBottom));
        outTransform->xMin = xMin; outTransform->xMax = xMax;
        outTransform->yMin = yMin; outTransform->yMax = yMax;
        outTransform->valid = true;
    }

    return pm;
}


// ═══════════════════════════════════════════════════════════════════════════════
// renderBarChart
// ═══════════════════════════════════════════════════════════════════════════════

QPixmap PlotEngine::renderBarChart(const QVector<QString>& labels,
                                   const QVector<double>&  values,
                                   const PlotConfig&       config,
                                   const QVector<QColor>&  colors,
                                   const QVector<double>&  stdDevValues)
{
    QPixmap pm(config.width, config.height);
    pm.fill(config.bgColor);

    if (labels.isEmpty() || values.isEmpty())
        return pm;

    int n = qMin(labels.size(), values.size());

    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing,    true);
    p.setRenderHint(QPainter::TextAntialiasing, true);

    // ── Legend sizing (computed first to reserve space below the plot) ────────
    p.setFont(config.labelFont);
    QFontMetrics fm(config.labelFont);

    // Build resolved color list for legend. Caller-supplied colors win; any
    // remainder comes from the shared distinct palette (no reuse, no yellow).
    const QVector<QColor> fallback = AppTheme::seriesColors(n);
    QVector<QColor> resolvedColors(n);
    for (int i = 0; i < n; ++i)
        resolvedColors[i] = (i < colors.size()) ? colors[i] : fallback[i];

    // Measure legend to determine how much bottom space to reserve.
    // Legend is drawn as horizontal rows of (swatch + label) pairs that wrap.
    const int swatchSz = 10;
    const int swatchGap = 4;
    const int itemPadR  = 14;
    int availLegendW = config.width - config.marginLeft - config.marginRight;

    struct LegendItem { int x; int y; int w; };
    QVector<LegendItem> legendItems(n);
    int legX = 0, legY = 0, legRowH = fm.height() + 4;
    int legendTotalH = legRowH;
    for (int i = 0; i < n; ++i) {
        int itemW = swatchSz + swatchGap + fm.horizontalAdvance(labels[i]) + itemPadR;
        if (legX + itemW > availLegendW && legX > 0) {
            legX = 0;
            legY += legRowH;
            legendTotalH += legRowH;
        }
        legendItems[i] = { legX, legY, itemW };
        legX += itemW;
    }

    int legendAreaH = legendTotalH + 8;  // padding above/below

    // ── Plot area ─────────────────────────────────────────────────────────────
    int pxLeft   = config.marginLeft;
    int pxRight  = config.width  - config.marginRight;
    int pxTop    = config.marginTop;
    int pxBottom = config.height - config.marginBottom - legendAreaH;

    // ── Title ─────────────────────────────────────────────────────────────────
    if (!config.title.isEmpty()) {
        p.setFont(config.titleFont);
        p.setPen(config.titleColor);
        p.drawText(QRect(pxLeft, 8, pxRight - pxLeft, pxTop - 8),
                   Qt::AlignHCenter | Qt::AlignVCenter, config.title);
    }

    // ── Y range (always starts at 0) ──────────────────────────────────────────
    double yMax = 0;
    for (int i = 0; i < n; ++i) {
        double top = values[i];
        if (!stdDevValues.isEmpty() && i < stdDevValues.size())
            top += stdDevValues[i];
        yMax = qMax(yMax, top);
    }
    // Standing axis rules: y anchored at 0, max visible value (bar top or
    // error-bar top) + 1.
    yMax += 1.0;
    double yStep = niceStep(yMax);
    double yMin = 0.0;

    // ── Plot background ───────────────────────────────────────────────────────
    p.fillRect(QRect(pxLeft, pxTop, pxRight - pxLeft, pxBottom - pxTop),
               config.bgColor);

    // ── Grid ──────────────────────────────────────────────────────────────────
    if (config.showGrid) {
        QPen gridPen(config.gridColor, 1);
        p.setPen(gridPen);
        for (double yv = 0; yv <= yMax + 1e-9; yv += yStep) {
            QPointF pt = dataToPixel(0, yv, 0, 1, yMin, yMax,
                                     pxLeft, pxRight, pxTop, pxBottom);
            p.drawLine(QPointF(pxLeft, pt.y()), QPointF(pxRight, pt.y()));
        }
    }

    // ── Bars ──────────────────────────────────────────────────────────────────
    double totalW = static_cast<double>(pxRight - pxLeft);
    double barW   = totalW / n * 0.65;

    p.setFont(config.labelFont);

    for (int i = 0; i < n; ++i) {
        double barCenterX = pxLeft + (i + 0.5) * (totalW / n);
        double barLeft    = barCenterX - barW / 2.0;

        QPointF ptBase = dataToPixel(0, 0,        0, 1, yMin, yMax,
                                     pxLeft, pxRight, pxTop, pxBottom);
        QPointF ptTop  = dataToPixel(0, values[i], 0, 1, yMin, yMax,
                                     pxLeft, pxRight, pxTop, pxBottom);

        QRectF barRect(barLeft, ptTop.y(),
                       barW,    ptBase.y() - ptTop.y());

        QColor barColor = resolvedColors[i];

        // Bar fill with a subtle gradient
        QLinearGradient grad(barRect.topLeft(), barRect.topRight());
        grad.setColorAt(0.0, barColor.lighter(115));
        grad.setColorAt(1.0, barColor);

        p.setBrush(grad);
        p.setPen(QPen(barColor.darker(130), 1));
        p.drawRect(barRect);

        // ── Error bar ─────────────────────────────────────────────────────
        bool hasError = (!stdDevValues.isEmpty() && i < stdDevValues.size()
                         && stdDevValues[i] > 0);
        if (hasError) {
            double sd   = stdDevValues[i];
            QPointF ptU = dataToPixel(0, values[i] + sd, 0, 1, yMin, yMax,
                                      pxLeft, pxRight, pxTop, pxBottom);
            QPointF ptL = dataToPixel(0, values[i] - sd, 0, 1, yMin, yMax,
                                      pxLeft, pxRight, pxTop, pxBottom);
            double cx   = barCenterX;
            double capW = barW * 0.25;

            p.setPen(QPen(Qt::black, 1.5));
            p.drawLine(QPointF(cx, ptU.y()), QPointF(cx, ptL.y()));
            p.drawLine(QPointF(cx - capW, ptU.y()), QPointF(cx + capW, ptU.y()));
            p.drawLine(QPointF(cx - capW, ptL.y()), QPointF(cx + capW, ptL.y()));
        }

        // ── Value label (offset left of error bar to stay readable) ──────
        QString valLabel = formatTickLabel(values[i]);
        int     vlW      = fm.horizontalAdvance(valLabel);
        int     vlH      = fm.height();
        p.setPen(config.axisColor);

        if (hasError) {
            // Place label to the left of the bar center so error bar doesn't block it
            double capW  = barW * 0.25;
            int labelX   = qRound(barCenterX - capW) - vlW - 2;
            // Clamp so it doesn't go past the left edge of the bar
            labelX = qMax(labelX, qRound(barLeft));
            p.drawText(QRect(labelX,
                             qRound(ptTop.y()) - vlH - 3,
                             vlW, vlH),
                       Qt::AlignRight | Qt::AlignBottom, valLabel);
        } else {
            // No error bar — center the label above the bar
            p.drawText(QRect(qRound(barCenterX) - vlW / 2,
                             qRound(ptTop.y()) - vlH - 3,
                             vlW, vlH),
                       Qt::AlignHCenter | Qt::AlignBottom, valLabel);
        }
    }

    // ── Y axis ────────────────────────────────────────────────────────────────
    p.setFont(config.labelFont);
    QPen axisPen(config.axisColor, 2);
    p.setPen(axisPen);
    p.drawLine(pxLeft, pxTop,    pxLeft, pxBottom);
    p.drawLine(pxLeft, pxBottom, pxRight, pxBottom);

    // Y tick marks and labels
    for (double yv = 0; yv <= yMax + 1e-9; yv += yStep) {
        QPointF pt = dataToPixel(0, yv, 0, 1, yMin, yMax,
                                 pxLeft, pxRight, pxTop, pxBottom);
        int py = qRound(pt.y());
        p.setPen(QPen(config.axisColor, 1));
        p.drawLine(pxLeft - 5, py, pxLeft, py);

        QString lbl = formatTickLabel(yv);
        int tw = fm.horizontalAdvance(lbl);
        p.setPen(config.axisColor);
        p.drawText(QRect(pxLeft - 8 - tw, py - fm.height() / 2, tw, fm.height()),
                   Qt::AlignRight | Qt::AlignVCenter, lbl);
    }

    // Y-axis label
    if (!config.yLabel.isEmpty()) {
        p.setFont(config.axisFont);
        QFontMetrics afm(config.axisFont);
        int midY = (pxTop + pxBottom) / 2;
        p.save();
        int yLabelX = qMax(14, afm.height() / 2 + 6);
        p.translate(yLabelX, midY);
        p.rotate(-90);
        int lw = afm.horizontalAdvance(config.yLabel);
        p.setPen(config.axisColor);
        p.drawText(QRect(-lw / 2, -afm.height(), lw, afm.height() + 2),
                   Qt::AlignHCenter | Qt::AlignBottom, config.yLabel);
        p.restore();
    }

    // ── Legend (below the plot, color-coded sample IDs) ───────────────────────
    {
        p.setFont(config.labelFont);
        int legendBaseY = pxBottom + 6;
        int legendBaseX = pxLeft;

        for (int i = 0; i < n; ++i) {
            int ix = legendBaseX + legendItems[i].x;
            int iy = legendBaseY + legendItems[i].y;

            // Color swatch (filled square)
            p.setBrush(resolvedColors[i]);
            p.setPen(QPen(resolvedColors[i].darker(130), 1));
            p.drawRect(ix, iy + (legRowH - swatchSz) / 2, swatchSz, swatchSz);

            // Label
            p.setPen(config.axisColor);
            p.drawText(ix + swatchSz + swatchGap, iy,
                       fm.horizontalAdvance(labels[i]), legRowH,
                       Qt::AlignLeft | Qt::AlignVCenter, labels[i]);
        }
    }

    // Optional manual legend (top-right)
    if (!config.legendEntries.isEmpty()) {
        const int swatchW = 24;
        const int swatchH = 14;
        QFontMetrics fm2(config.labelFont);
        const int rowH    = qMax(swatchH + 6, fm2.height() + 4);
        const int padding = 12;

        int maxLabelW = 0;
        for (const auto& kv : config.legendEntries)
            maxLabelW = qMax(maxLabelW, fm2.horizontalAdvance(kv.first));
        const int boxW = swatchW + 8 + maxLabelW + padding * 2;
        const int boxH = config.legendEntries.size() * rowH + padding;

        const int boxX = pxRight - boxW - 8;
        const int boxY = pxTop + 8;

        p.setPen(QPen(config.axisColor, 1));
        p.setBrush(QBrush(QColor(255, 255, 255, 230)));
        p.drawRect(boxX, boxY, boxW, boxH);

        p.setFont(config.labelFont);
        for (int i = 0; i < config.legendEntries.size(); ++i) {
            const auto& kv = config.legendEntries[i];
            const int rowY = boxY + padding / 2 + i * rowH;
            p.setPen(Qt::NoPen);
            p.setBrush(kv.second);
            p.drawRect(boxX + padding, rowY + (rowH - swatchH) / 2, swatchW, swatchH);
            p.setPen(config.axisColor);
            p.drawText(boxX + padding + swatchW + 8,
                       rowY + (rowH + fm2.ascent()) / 2 - 2,
                       kv.first);
        }
    }

    p.end();
    return pm;
}


// ═══════════════════════════════════════════════════════════════════════════════
// renderLinePlotDualAxis
// ═══════════════════════════════════════════════════════════════════════════════

QPixmap PlotEngine::renderLinePlotDualAxis(const QVector<PlotSeries>& primarySeries,
                                            const QVector<PlotSeries>& secondarySeries,
                                            const PlotConfig&           config,
                                            PlotTransform*               outTransform)
{
    QPixmap pm(config.width, config.height);
    pm.fill(config.bgColor);

    QPainter p(&pm);
    p.setRenderHint(QPainter::Antialiasing,    true);
    p.setRenderHint(QPainter::TextAntialiasing, true);

    // Right margin wide enough for right-axis tick labels + label
    const int rMargin = 75;

    // Reserve extra bottom margin for legend when shown
    int legendExtraH = 0;
    if (config.showLegend) {
        QVector<PlotSeries> allS = primarySeries;
        allS.append(secondarySeries);
        if (!allS.isEmpty()) {
            QFontMetrics lfm(config.labelFont);
            int rowH = lfm.height() + 4;
            int availW = config.width - config.marginLeft - rMargin;
            int lx = 0, rows = 1;
            for (const auto& s : allS) {
                int itemW = 20 + 6 + lfm.horizontalAdvance(s.label) + 16;
                if (lx + itemW > availW && lx > 0) { lx = 0; ++rows; }
                lx += itemW;
            }
            legendExtraH = rows * rowH + 12;
        }
    }

    QFontMetrics tickFm2(config.labelFont);
    QFontMetrics axisFm2(config.axisFont);
    int xAxisLabelsH2 = 7 + tickFm2.height()
                      + (config.xLabel.isEmpty() ? 0 : (6 + axisFm2.height()))
                      + 6;

    int pxLeft   = config.marginLeft;
    int pxRight  = config.width  - rMargin;
    int pxTop    = config.marginTop;
    int pxBottom = config.height
                 - qMax(config.marginBottom, xAxisLabelsH2)
                 - legendExtraH;

    // ── Title ─────────────────────────────────────────────────────────────────
    if (!config.title.isEmpty()) {
        p.setFont(config.titleFont);
        p.setPen(config.titleColor);
        p.drawText(QRect(pxLeft, 8, pxRight - pxLeft, pxTop - 8),
                   Qt::AlignHCenter | Qt::AlignVCenter, config.title);
    }

    // ── Auto-range for left Y (primary) and X ─────────────────────────────────
    double xMin, xMax, yMin, yMax;
    if (!primarySeries.isEmpty()) {
        autoRange(primarySeries, xMin, xMax, yMin, yMax);
        // Also expand X range to include secondary series X values
        // (x stays anchored at 0, last point + 1, per the standing rules).
        if (!secondarySeries.isEmpty()) {
            for (const auto& s : secondarySeries)
                for (double v : s.x) xMax = qMax(xMax, v + 1.0);
        }
    } else {
        xMin = 0; xMax = 1; yMin = 0; yMax = 1;
    }

    // ── Auto-range for right Y (secondary) ────────────────────────────────────
    double y2Min = 0.0, y2Max = 1.0;
    if (!secondarySeries.isEmpty()) {
        y2Min =  std::numeric_limits<double>::max();
        y2Max = -std::numeric_limits<double>::max();
        for (const auto& s : secondarySeries)
            for (double v : s.y) { y2Min = qMin(y2Min, v); y2Max = qMax(y2Max, v); }
        if (y2Min > y2Max) { y2Min = 0; y2Max = 1; }
        // Standing axis rules: right (oil) axis anchored at 0, max + 1 on top.
        y2Min = 0.0;
        y2Max = y2Max + 1.0;
    }

    // ── Plot background ───────────────────────────────────────────────────────
    p.fillRect(QRect(pxLeft, pxTop, pxRight - pxLeft, pxBottom - pxTop), config.bgColor);

    // ── Grid (based on left Y) ────────────────────────────────────────────────
    if (config.showGrid)
        drawGrid(p, config, pxLeft, pxRight, pxTop, pxBottom,
                 xMin, xMax, yMin, yMax);

    // ── Clip to plot area ─────────────────────────────────────────────────────
    p.setClipRect(QRect(pxLeft, pxTop, pxRight - pxLeft + 1, pxBottom - pxTop + 1));

    // ── Draw primary series (left Y) ──────────────────────────────────────────
    for (const auto& s : primarySeries) {
        if (s.x.isEmpty() || s.y.isEmpty()) continue;
        int n = qMin(s.x.size(), s.y.size());
        QVector<QPointF> pts;
        pts.reserve(n);
        for (int i = 0; i < n; ++i)
            pts.append(dataToPixel(s.x[i], s.y[i], xMin, xMax, yMin, yMax,
                                   pxLeft, pxRight, pxTop, pxBottom));
        if (s.drawLine && n >= 2) {
            QPen lp(s.color, s.lineWidth, Qt::SolidLine, Qt::RoundCap, Qt::RoundJoin);
            p.setPen(lp); p.setBrush(Qt::NoBrush);
            for (int i = 0; i < n - 1; ++i) p.drawLine(pts[i], pts[i+1]);
        }
        if (s.drawDots) {
            p.setPen(QPen(s.color.darker(130), 1));
            p.setBrush(s.color);
            for (const auto& pt : pts) p.drawEllipse(pt, s.dotRadius, s.dotRadius);
        }
    }

    // ── Draw secondary series (right Y) ──────────────────────────────────────
    for (const auto& s : secondarySeries) {
        if (s.x.isEmpty() || s.y.isEmpty()) continue;
        int n = qMin(s.x.size(), s.y.size());
        QVector<QPointF> pts;
        pts.reserve(n);
        for (int i = 0; i < n; ++i)
            pts.append(dataToPixel(s.x[i], s.y[i], xMin, xMax, y2Min, y2Max,
                                   pxLeft, pxRight, pxTop, pxBottom));
        if (s.drawLine && n >= 2) {
            QPen lp(s.color, s.lineWidth, Qt::DashLine, Qt::RoundCap, Qt::RoundJoin);
            p.setPen(lp); p.setBrush(Qt::NoBrush);
            for (int i = 0; i < n - 1; ++i) p.drawLine(pts[i], pts[i+1]);
        }
        if (s.drawDots) {
            p.setPen(QPen(s.color.darker(130), 1));
            p.setBrush(s.color);
            for (const auto& pt : pts) p.drawEllipse(pt, s.dotRadius, s.dotRadius);
        }
    }

    p.setClipping(false);

    // ── Left axes (primary Y) ─────────────────────────────────────────────────
    drawAxes(p, config, pxLeft, pxRight, pxTop, pxBottom, xMin, xMax, yMin, yMax);

    // ── Right axis (secondary Y) ──────────────────────────────────────────────
    if (!secondarySeries.isEmpty()) {
        p.save();
        p.setFont(config.labelFont);
        QFontMetrics fm(config.labelFont);

        QPen axisPen(config.axisColor, 2);
        p.setPen(axisPen);
        p.drawLine(QPointF(pxRight, pxTop), QPointF(pxRight, pxBottom));

        double yStep2 = niceStep(y2Max - y2Min);
        QPen tickPen(config.axisColor, 1);

        for (double yv = std::ceil(y2Min / yStep2) * yStep2; yv <= y2Max + 1e-9; yv += yStep2) {
            QPointF pt = dataToPixel(xMax, yv, xMin, xMax, y2Min, y2Max,
                                     pxLeft, pxRight, pxTop, pxBottom);
            int py = qRound(pt.y());
            p.setPen(tickPen);
            p.drawLine(pxRight, py, pxRight + 5, py);
            QString lbl = formatTickLabel(yv);
            int tw = fm.horizontalAdvance(lbl);
            int th = fm.height();
            p.setPen(config.axisColor);
            p.drawText(QRect(pxRight + 8, py - th / 2, tw, th),
                       Qt::AlignLeft | Qt::AlignVCenter, lbl);
        }

        // Right axis label (rotated, on the far right edge)
        if (!config.y2Label.isEmpty()) {
            p.setFont(config.axisFont);
            QFontMetrics afm(config.axisFont);
            int midY = (pxTop + pxBottom) / 2;
            p.save();
            p.translate(config.width - 14, midY);
            p.rotate(90);
            int lw = afm.horizontalAdvance(config.y2Label);
            p.setPen(config.axisColor);
            p.drawText(QRect(-lw / 2, -afm.height(), lw, afm.height() + 2),
                       Qt::AlignHCenter | Qt::AlignBottom, config.y2Label);
            p.restore();
        }

        p.restore();
    }

    // ── Combined legend ───────────────────────────────────────────────────────
    if (config.showLegend) {
        QVector<PlotSeries> allSeries = primarySeries;
        for (auto s : secondarySeries) { s.dashed = true; allSeries.append(s); }
        if (!allSeries.isEmpty())
            drawLegend(p, allSeries, config, pxRight, pxTop);
    }

    p.end();

    // Populated from the PRIMARY (left Y) axis only; the secondary (right Y /
    // oil-overlay) axis is not exposed via PlotTransform.
    if (outTransform) {
        outTransform->plotRect = QRectF(QPointF(pxLeft, pxTop), QPointF(pxRight, pxBottom));
        outTransform->xMin = xMin; outTransform->xMax = xMax;
        outTransform->yMin = yMin; outTransform->yMax = yMax;
        outTransform->valid = true;
    }

    return pm;
}


// ═══════════════════════════════════════════════════════════════════════════════
// toPng
// ═══════════════════════════════════════════════════════════════════════════════

QByteArray PlotEngine::toPng(const QPixmap& pm, int dpi)
{
    QImage img = pm.toImage();

    // Set DPI metadata (dots per meter = dpi / 0.0254)
    int dpm = qRound(dpi / 0.0254);
    img.setDotsPerMeterX(dpm);
    img.setDotsPerMeterY(dpm);

    QByteArray  bytes;
    QBuffer     buf(&bytes);
    buf.open(QIODevice::WriteOnly);

    QImageWriter writer(&buf, "PNG");
    writer.setQuality(95);
    writer.write(img);

    buf.close();
    return bytes;
}


// ═══════════════════════════════════════════════════════════════════════════════
// composeAnnotatedExport (DV-22)
// ═══════════════════════════════════════════════════════════════════════════════

QImage PlotEngine::composeAnnotatedExport(const QPixmap&                 basePlot,
                                          const PlotTransform&           tf,
                                          const QVector<PlotAnnotation>& notes)
{
    // Nothing to annotate (or no transform to place arrows with) -> plain export.
    if (notes.isEmpty() || !tf.valid)
        return basePlot.toImage();

    const QColor amber(0xBA, 0x75, 0x17);

    // Textbox geometry: 6pt Segoe UI, word-wrapped. Width adapts to the room to
    // the LEFT of the note's point so the box never extends past that point - its
    // right edge stays exactly on the arrow (the spec's bottom-right-corner
    // anchor) and, crucially, a box can never slide under a neighbour's arrow line.
    QFont noteFont("Segoe UI", 6);
    QFontMetrics fm(noteFont);
    const int pad      = 4;                       // inner padding inside each box
    const int minBoxW  = 90;                      // narrowest readable box
    const int maxBoxW  = 200;                     // hard width cap (spec)
    const int lineH    = fm.lineSpacing();
    const int maxTextH = 10 * lineH;              // hard row cap (spec): 10 rows
    const int gap      = 4;                       // vertical gap between stacked boxes

    const int totalW = basePlot.width();

    // Resolve each note to a placed textbox. Non-spill boxes live in the
    // whitespace band above the plot, their position measured as a distance ABOVE
    // the plot's top edge (so the band height falls out as the tallest stack -
    // need-based). Boxes are laid out left-to-right by data x and greedily lifted
    // above any already-placed box they overlap, so clustered notes stack instead
    // of colliding; together with the never-past-the-point width above, that
    // leaves zero box/box and box/arrow overlaps. The one case that can't be
    // placed cleanly - a point so close to the left edge that even a min-width box
    // won't fit to its left - "spills": drawn 50% transparent, anchored at the
    // plot top and extending DOWN into the plot (rare; the only time a box enters
    // the plot).
    struct Placed {
        int     left, right, w, h;   // box rect (x on-image; y resolved after banding)
        int     bottomAbove;         // box bottom, measured upward from the plot top (0 if spill)
        bool    spill;               // couldn't fit above the plot -> transparent, into the plot
        double  xPix;                // arrow x == the data point's x
        double  yPointBase;          // the data point's y, in base-pixmap px
        QString text;
    };

    QVector<int> order(notes.size());
    for (int i = 0; i < notes.size(); ++i) order[i] = i;
    std::sort(order.begin(), order.end(), [&](int a, int b) {
        return tf.dataToPixel(notes[a].puff, notes[a].tpm).x()
             < tf.dataToPixel(notes[b].puff, notes[b].tpm).x();
    });

    QVector<Placed> placed;
    placed.reserve(notes.size());
    int maxTopAbove = 0;
    for (int oi : order) {
        const PlotAnnotation& n  = notes[oi];
        const QPointF         pt = tf.dataToPixel(n.puff, n.tpm);
        const int             px = int(qRound(pt.x()));

        // Width = as wide as fits to the left of the point, within [min, max];
        // then shrink to the wrapped text.
        const int   boxW  = qBound(minBoxW, px, maxBoxW);
        const int   textW = boxW - 2 * pad;
        const QRect br    = fm.boundingRect(QRect(0, 0, textW, maxTextH),
                                            Qt::TextWordWrap, n.text);
        const int   w = qBound(1,     br.width(),  textW)    + 2 * pad;
        const int   h = qBound(lineH, br.height(), maxTextH) + 2 * pad;

        int  right = px;
        int  left  = right - w;
        bool spill = false;
        if (left < 0) {                     // point too near the left edge to fit a box
            left = 0; right = w; spill = true;
        }

        int bottomAbove = 0;
        if (!spill) {
            // Greedy stack: anchor to the plot top, then lift above any placed box
            // we still overlap (re-scan to a fixpoint - one lift can expose another).
            bool moved = true;
            while (moved) {
                moved = false;
                for (const Placed& q : placed) {
                    if (q.spill) continue;
                    const bool hOverlap = left < q.right && q.left < right;
                    const bool vOverlap = bottomAbove < q.bottomAbove + q.h
                                       && q.bottomAbove < bottomAbove + h;
                    if (hOverlap && vOverlap) {
                        bottomAbove = q.bottomAbove + q.h + gap;
                        moved = true;
                    }
                }
            }
            maxTopAbove = qMax(maxTopAbove, bottomAbove + h);
        }
        placed.append({ left, right, w, h, bottomAbove, spill, pt.x(), pt.y(), n.text });
    }

    const int bandTopPad = 6;                        // breathing room above the tallest box
    const int bandH      = maxTopAbove + bandTopPad;  // whitespace band height (need-based)
    const int totalH     = bandH + basePlot.height();

    QImage img(totalW, totalH, QImage::Format_ARGB32_Premultiplied);
    img.fill(Qt::white);

    QPainter p(&img);
    p.setRenderHint(QPainter::Antialiasing, true);
    p.setRenderHint(QPainter::TextAntialiasing, true);

    // Base plot, unmodified, below the whitespace band.
    p.drawPixmap(0, bandH, basePlot);

    // Pass 1 — textboxes. Non-spill: opaque, in the band. Spill: 50% transparent,
    // top on the plot's edge, extending down into the plot (the plot shows
    // through). Drawn before the arrows so no arrow is hidden behind a later box.
    p.setFont(noteFont);
    for (const Placed& b : placed) {
        const int   boxTopY = b.spill ? bandH : (bandH - b.bottomAbove - b.h);
        const QRect boxRect(b.left, boxTopY, b.w, b.h);
        p.setOpacity(b.spill ? 0.5 : 1.0);
        p.setPen(QPen(amber, 1.0));
        p.setBrush(Qt::white);
        p.drawRect(boxRect);
        p.setPen(QColor(0x33, 0x33, 0x33));
        p.drawText(boxRect.adjusted(pad, pad, -pad, -pad),
                   Qt::AlignLeft | Qt::AlignTop | Qt::TextWordWrap, b.text);
    }
    p.setOpacity(1.0);

    // Pass 2 — vertical amber arrows on top: from each box's bottom-right corner
    // straight down to 8px above its data point, so the line never crosses into
    // the plotted data below the point. (A spill box already sits over its point,
    // so its arrow is drawn only where the point is still below the box.)
    for (const Placed& b : placed) {
        const double xPix = b.xPix;
        const double yTop = b.spill ? double(bandH + b.h) : double(bandH - b.bottomAbove);
        const double yTip = bandH + b.yPointBase - 8.0;
        if (yTip <= yTop) continue;
        p.setOpacity(b.spill ? 0.5 : 1.0);
        p.setPen(QPen(amber, 1.4, Qt::SolidLine, Qt::RoundCap));
        p.setBrush(Qt::NoBrush);
        p.drawLine(QPointF(xPix, yTop), QPointF(xPix, yTip));

        // Filled arrowhead pointing down at the tip.
        p.setBrush(amber);
        p.setPen(Qt::NoPen);
        const double aw = 3.5;   // arrowhead half-width
        const double ah = 6.0;   // arrowhead height
        QPolygonF head;
        head << QPointF(xPix - aw, yTip - ah)
             << QPointF(xPix + aw, yTip - ah)
             << QPointF(xPix,      yTip);
        p.drawPolygon(head);
    }
    p.setOpacity(1.0);

    p.end();
    return img;
}


// ═══════════════════════════════════════════════════════════════════════════════
// Convenience helpers
// ═══════════════════════════════════════════════════════════════════════════════

QPixmap PlotEngine::renderTPMTrend(const QVector<double>& puffCounts,
                                   const QVector<double>& tpmValues,
                                   const QString&         title)
{
    PlotSeries s;
    s.label     = "TPM (mg/puff)";
    s.x         = puffCounts;
    s.y         = tpmValues;
    s.color     = QColor(0x00, 0x66, 0xCC);
    s.drawLine  = true;
    s.drawDots  = true;
    s.lineWidth = 2;
    s.dotRadius = 4;

    PlotConfig cfg;
    cfg.title     = title;
    cfg.xLabel    = "Puff Count";
    cfg.yLabel    = "TPM (mg/puff)";
    cfg.width     = 800;
    cfg.height    = 480;
    cfg.autoScale = true;
    cfg.showGrid  = true;
    cfg.showLegend = false;

    return renderLinePlot({s}, cfg);
}

QPixmap PlotEngine::renderTPMBarChart(const QVector<QString>& sampleNames,
                                      const QVector<double>&  avgTPM,
                                      const QVector<double>&  stdDevTPM,
                                      const QString&          title)
{
    PlotConfig cfg;
    cfg.title     = title;
    cfg.xLabel    = "Sample";
    cfg.yLabel    = "Avg TPM (mg/puff)";
    cfg.width     = 800;
    cfg.height    = 480;
    cfg.autoScale = false;   // renderBarChart computes its own Y range
    cfg.showGrid  = true;
    cfg.showLegend = false;

    // One distinct color per bar, by sample position in AppTheme::seriesColors
    // (the shared plot palette).
    const QVector<QColor> colors = AppTheme::seriesColors(sampleNames.size());

    return renderBarChart(sampleNames, avgTPM, cfg, colors, stdDevTPM);
}

} // namespace DVE
