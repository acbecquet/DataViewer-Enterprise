#pragma once
#include <QString>
#include <QVector>
#include <QColor>
#include <QPixmap>
#include <QFont>
#include <QPainter>
#include <QPair>

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

// ─── PlotEngine ───────────────────────────────────────────────────────────────
class PlotEngine {
public:
    // Line/scatter plot – one or more series on the same axes.
    static QPixmap renderLinePlot(const QVector<PlotSeries>& series,
                                  const PlotConfig&           config);

    // Bar chart – one bar per label.
    // stdDevValues may be empty (no error bars will be drawn).
    static QPixmap renderBarChart(const QVector<QString>&  labels,
                                  const QVector<double>&   values,
                                  const PlotConfig&        config,
                                  const QVector<QColor>&   colors     = {},
                                  const QVector<double>&   stdDevValues = {});

    // Export a QPixmap to in-memory PNG bytes (suitable for embedding in PPTX).
    static QByteArray toPng(const QPixmap& pm, int dpi = 150);

    // Test-only: returns the auto-scaled yMin for the given series (always 0).
    static double autoRangeYMinForTesting(const QVector<PlotSeries>& series) {
        double xMin, xMax, yMin, yMax;
        autoRange(series, xMin, xMax, yMin, yMax);
        return yMin;
    }

    // Dual-axis line plot: primarySeries on left Y, secondarySeries on right Y.
    // config.y2Label is used for the right axis label.
    static QPixmap renderLinePlotDualAxis(const QVector<PlotSeries>& primarySeries,
                                          const QVector<PlotSeries>& secondarySeries,
                                          const PlotConfig&           config);

    // Convenience: TPM vs puff-count trend line.
    static QPixmap renderTPMTrend(const QVector<double>& puffCounts,
                                  const QVector<double>& tpmValues,
                                  const QString&         title);

    // Convenience: average TPM bar chart, one bar per sample, with error bars.
    static QPixmap renderTPMBarChart(const QVector<QString>& sampleNames,
                                     const QVector<double>&  avgTPM,
                                     const QVector<double>&  stdDevTPM,
                                     const QString&          title);

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

// Computes the y-axis upper bound for the draw-pressure chart.
// Returns max(2.0, ceil(seriesMax)) so the axis always shows at least
// 0-2 Pa range, expanding to the next integer when data exceeds 2 Pa.
double drawPressureYMax(double seriesMax);

} // namespace DVE
