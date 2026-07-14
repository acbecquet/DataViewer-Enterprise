#pragma once
#include <QString>
#include <QVector>
#include <QColor>
#include <QPixmap>
#include <QFont>
#include <QPainter>
#include <QPair>
#include <QRectF>

namespace DVE {

// ─── PlotSeries ───────────────────────────────────────────────────────────────
struct PlotSeries {
    QString         label;
    QVector<double> x;
    QVector<double> y;
    QColor          color     = QColor(0x00, 0x66, 0xCC);
    bool            drawLine  = true;
    bool            drawDots  = true;
    bool            dashed    = false;  // dashed line (for overlay series)
    int             lineWidth = 2;
    int             dotRadius = 4;
    // Note→plot linking (DATAVIEWER-6 v1): indices into x/y that carry a note
    // get an amber ring; `emphasized` (an index into x/y, or -1) gets a larger
    // ring plus a dashed guide line down to the x-axis.
    QVector<int>    ringed;
    int             emphasized = -1;
};

// ─── PlotConfig ───────────────────────────────────────────────────────────────
struct PlotConfig {
    QString title;
    QString xLabel;
    QString yLabel;

    int  width  = 800;
    int  height = 500;

    bool showGrid   = true;
    bool showLegend = true;
    bool autoScale  = true;

    double yMin = 0.0, yMax = 1.0;
    double xMin = 0.0, xMax = 1.0;

    // Right Y axis (dual-axis plots)
    QString y2Label;    // empty = no right axis

    QColor bgColor   = Qt::white;
    QColor gridColor = QColor(0xDD, 0xDD, 0xDD);
    QColor axisColor = QColor(0x33, 0x33, 0x33);
    QColor titleColor = QColor(0x1A, 0x1A, 0x1A);

    QFont titleFont = QFont("Segoe UI", 11, QFont::Bold);
    QFont axisFont  = QFont("Segoe UI",  9);
    QFont labelFont = QFont("Segoe UI",  8);

    int marginLeft   = 70;
    int marginRight  = 30;
    int marginTop    = 50;
    int marginBottom = 50;

    // Optional manual legend entries — used by bar charts (which have no
    // series to derive a legend from) and any custom plot. Empty = no legend.
    QVector<QPair<QString, QColor>> legendEntries;
};

// ─── PlotTransform ─────────────────────────────────────────────────────────────
// Axis-to-pixel transform for a rendered line plot. Emitted by renderLinePlot/
// renderLinePlotDualAxis so callers (hit-testing, annotated export) can map
// between data coordinates and the pixels actually drawn, without re-deriving
// the margin/legend math. y grows upward in data space, downward in pixels.
struct PlotTransform {
    QRectF plotRect;                 // the drawable data area, in pixmap pixels
    double xMin = 0.0, xMax = 1.0;
    double yMin = 0.0, yMax = 1.0;
    bool   valid = false;
    QPointF dataToPixel(double x, double y) const {
        const double fx = (xMax > xMin) ? (x - xMin) / (xMax - xMin) : 0.0;
        const double fy = (yMax > yMin) ? (y - yMin) / (yMax - yMin) : 0.0;
        return QPointF(plotRect.left()   + fx * plotRect.width(),
                       plotRect.bottom() - fy * plotRect.height());
    }
    QPointF pixelToData(const QPointF& px) const {
        const double fx = (plotRect.width()  > 0) ? (px.x() - plotRect.left())   / plotRect.width()  : 0.0;
        const double fy = (plotRect.height() > 0) ? (plotRect.bottom() - px.y()) / plotRect.height() : 0.0;
        return QPointF(xMin + fx * (xMax - xMin), yMin + fy * (yMax - yMin));
    }
};

// ─── PlotAnnotation ─────────────────────────────────────────────────────────────
// DV-22: one note to render onto the annotated Save-plot export — a textbox in
// the whitespace band above the plot, tied by a vertical amber arrow to the data
// point at (puff, tpm). `text` is drawn word-wrapped inside the box.
struct PlotAnnotation {
    double  puff = 0.0;   // data x of the noted point
    double  tpm  = 0.0;   // data y of the noted point (the arrow's target)
    QString text;         // textbox contents (word-wrapped; capped at 200px / 10 rows)
};

// ─── PlotEngine ───────────────────────────────────────────────────────────────
class PlotEngine {
public:
    // Line/scatter plot – one or more series on the same axes.
    // outTransform (optional): if non-null, populated with the axis-to-pixel
    // transform actually used to draw this render, for hit-testing/export.
    static QPixmap renderLinePlot(const QVector<PlotSeries>& series,
                                  const PlotConfig&           config,
                                  PlotTransform*               outTransform = nullptr);

    // Bar chart – one bar per label.
    // stdDevValues may be empty (no error bars will be drawn).
    static QPixmap renderBarChart(const QVector<QString>&  labels,
                                  const QVector<double>&   values,
                                  const PlotConfig&        config,
                                  const QVector<QColor>&   colors     = {},
                                  const QVector<double>&   stdDevValues = {});

    // Export a QPixmap to in-memory PNG bytes (suitable for embedding in PPTX).
    static QByteArray toPng(const QPixmap& pm, int dpi = 150);

    // Set cfg.xMin/xMax so the x axis runs from 0 (always — starting at the
    // first data point pushes that point onto the y axis where its dot looks
    // cut off) to the last data point. Use with autoScale=false; otherwise
    // the config's 0..1 defaults silently become the x axis.
    static void applyDataXRange(PlotConfig& cfg,
                                const QVector<PlotSeries>& series);

    // Set cfg.yMin/yMax per the standing rule for EVERY x-y plot: the y axis
    // is anchored at 0 and tops out at the maximum data value + 1. Always.
    static void applyAnchoredYRange(PlotConfig& cfg,
                                    const QVector<PlotSeries>& series);

    // Test-only: returns the auto-scaled yMin for the given series (always 0).
    static double autoRangeYMinForTesting(const QVector<PlotSeries>& series) {
        double xMin, xMax, yMin, yMax;
        autoRange(series, xMin, xMax, yMin, yMax);
        return yMin;
    }

    // Dual-axis line plot: primarySeries on left Y, secondarySeries on right Y.
    // config.y2Label is used for the right axis label.
    // outTransform (optional): populated from the PRIMARY (left Y) axis only;
    // the secondary (right Y / oil-overlay) axis is not exposed.
    static QPixmap renderLinePlotDualAxis(const QVector<PlotSeries>& primarySeries,
                                          const QVector<PlotSeries>& secondarySeries,
                                          const PlotConfig&           config,
                                          PlotTransform*               outTransform = nullptr);

    // Convenience: TPM vs puff-count trend line.
    static QPixmap renderTPMTrend(const QVector<double>& puffCounts,
                                  const QVector<double>& tpmValues,
                                  const QString&         title);

    // Convenience: average TPM bar chart, one bar per sample, with error bars.
    static QPixmap renderTPMBarChart(const QVector<QString>& sampleNames,
                                     const QVector<double>&  avgTPM,
                                     const QVector<double>&  stdDevTPM,
                                     const QString&          title);

    // DV-22: build the annotated Save-plot export image. Returns a QImage taller
    // than basePlot by a whitespace band that holds one textbox per note (word-
    // wrapped, capped at 200px wide / 10 rows, 6pt). Each box is anchored with
    // its bottom on the plot's top edge and its bottom-right corner at the note's
    // x; a thin amber vertical arrow drops from there to 8px above that note's
    // data point (mapped via tf.dataToPixel(puff, tpm)) so it never crosses into
    // the plotted data. The band height is need-based (the tallest box). notes
    // empty or tf invalid -> basePlot returned unchanged. Pure — does not mutate
    // basePlot.
    // If `title` is non-empty it is re-drawn on the TOP layer (over the arrows),
    // positioned exactly where renderLinePlot drew it (from tf.plotRect), with a
    // translucent halo so it stays readable where an arrow passes behind it.
    static QImage composeAnnotatedExport(const QPixmap&                 basePlot,
                                         const PlotTransform&           tf,
                                         const QVector<PlotAnnotation>& notes,
                                         const QString&                 title      = QString(),
                                         const QFont&                   titleFont  = QFont(QStringLiteral("Segoe UI"), 11, QFont::Bold),
                                         const QColor&                  titleColor = QColor(0x1A, 0x1A, 0x1A));

private:
    // Map a data-space point to pixel coordinates.
    static QPointF dataToPixel(double x, double y,
                               double xMin, double xMax,
                               double yMin, double yMax,
                               int pxLeft, int pxRight,
                               int pxTop,  int pxBottom);

    // Determine a "nice" tick step for the given data range and approximate
    // number of desired ticks.
    static double niceStep(double range, int targetTicks = 6);

    // Compute auto-scale ranges with 5 % padding, rounded to nice numbers.
    static void autoRange(const QVector<PlotSeries>& series,
                          double& xMin, double& xMax,
                          double& yMin, double& yMax);

    // Draw grid lines aligned to tick positions.
    static void drawGrid(QPainter& p, const PlotConfig& cfg,
                         int pxLeft, int pxRight, int pxTop, int pxBottom,
                         double xMin, double xMax, double yMin, double yMax);

    // Draw the axis lines, tick marks, and tick labels.
    static void drawAxes(QPainter& p, const PlotConfig& cfg,
                         int pxLeft, int pxRight, int pxTop, int pxBottom,
                         double xMin, double xMax, double yMin, double yMax);

    // Draw the legend box (top-right of the plot area).
    static void drawLegend(QPainter& p, const QVector<PlotSeries>& series,
                           const PlotConfig& cfg,
                           int pxRight, int pxTop);

    // Format a tick-label number concisely (remove trailing zeros).
    static QString formatTickLabel(double v);
};

} // namespace DVE
