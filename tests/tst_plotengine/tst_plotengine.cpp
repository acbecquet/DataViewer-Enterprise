#include <QtTest>
#include <QPixmap>
#include <QImage>
#include <QColor>
#include <QApplication>
#include "PlotEngine.h"
#include "TestHelpers.h"

class TestPlotEngine : public QObject
{
    Q_OBJECT

private slots:
    void testRenderLinePlot();
    void testRenderBarChart();
    void testRenderBarChartWithErrorBars();
    void testToPng();
    void testDualAxisPlot();
    void testTPMTrend();
    void testTPMBarChart();
    void testEmptyData();
    void testPlotDimensions();
    void barChart_drawsLegendEntries();
    void applyDataXRange_spansData();
    void applyAnchoredYRange_zeroToMaxPlusOne();
    void autoRange_anchorsYAtZero();
    void ringedPoint_changesRenderedPixels();
    void transform_roundTripsDataAndPixel();
    void annotatedExport_growsHeightAndDrawsAmberLines();
};

void TestPlotEngine::testRenderLinePlot()
{
    DVE::PlotSeries series;
    series.label = "Test";
    series.x = {1.0, 2.0, 3.0, 4.0, 5.0};
    series.y = {10.0, 20.0, 15.0, 25.0, 30.0};
    series.color = Qt::blue;
    series.drawLine = true;
    series.drawDots = true;

    DVE::PlotConfig config;
    config.title = "Line Plot Test";
    config.xLabel = "X";
    config.yLabel = "Y";
    config.width = 800;
    config.height = 600;

    QPixmap result = DVE::PlotEngine::renderLinePlot({series}, config);
    QVERIFY(!result.isNull());
    QCOMPARE(result.width(), 800);
    QCOMPARE(result.height(), 600);
}

void TestPlotEngine::testRenderBarChart()
{
    QStringList labels = {"A", "B", "C"};
    QVector<double> values = {10.0, 20.0, 15.0};

    DVE::PlotConfig config;
    config.title = "Bar Chart Test";
    config.xLabel = "Category";
    config.yLabel = "Value";
    config.width = 800;
    config.height = 600;

    QPixmap result = DVE::PlotEngine::renderBarChart(labels, values, config);
    QVERIFY(!result.isNull());
}

void TestPlotEngine::testRenderBarChartWithErrorBars()
{
    QStringList labels = {"A", "B", "C"};
    QVector<double> values = {10.0, 20.0, 15.0};
    QVector<double> stddev = {1.0, 2.0, 1.5};

    DVE::PlotConfig config;
    config.title = "Bar Chart with Error Bars";
    config.xLabel = "Category";
    config.yLabel = "Value";
    config.width = 800;
    config.height = 600;

    QPixmap result = DVE::PlotEngine::renderBarChart(labels, values, config, {}, stddev);
    QVERIFY(!result.isNull());
}

void TestPlotEngine::testToPng()
{
    DVE::PlotSeries series;
    series.label = "PNG Test";
    series.x = {1.0, 2.0, 3.0};
    series.y = {10.0, 20.0, 30.0};
    series.color = Qt::red;
    series.drawLine = true;
    series.drawDots = false;

    DVE::PlotConfig config;
    config.title = "PNG Export";
    config.xLabel = "X";
    config.yLabel = "Y";
    config.width = 400;
    config.height = 300;

    QPixmap pm = DVE::PlotEngine::renderLinePlot({series}, config);
    QVERIFY(!pm.isNull());

    QByteArray png = DVE::PlotEngine::toPng(pm);
    QVERIFY(!png.isEmpty());
    // PNG magic bytes: 0x89 P N G
    QCOMPARE(png[0], '\x89');
    QCOMPARE(png[1], 'P');
    QCOMPARE(png[2], 'N');
    QCOMPARE(png[3], 'G');
}

void TestPlotEngine::testDualAxisPlot()
{
    DVE::PlotSeries primary;
    primary.label = "Primary";
    primary.x = {1.0, 2.0, 3.0};
    primary.y = {10.0, 20.0, 30.0};
    primary.color = Qt::blue;
    primary.drawLine = true;
    primary.drawDots = true;

    DVE::PlotSeries secondary;
    secondary.label = "Secondary";
    secondary.x = {1.0, 2.0, 3.0};
    secondary.y = {100.0, 200.0, 300.0};
    secondary.color = Qt::red;
    secondary.drawLine = true;
    secondary.drawDots = false;

    DVE::PlotConfig config;
    config.title = "Dual Axis";
    config.xLabel = "X";
    config.yLabel = "Primary Y";
    config.width = 800;
    config.height = 600;

    QPixmap result = DVE::PlotEngine::renderLinePlotDualAxis({primary}, {secondary}, config);
    QVERIFY(!result.isNull());
}

void TestPlotEngine::testTPMTrend()
{
    QVector<double> puffs = {10, 20, 30, 40, 50};
    QVector<double> tpm   = {3.5, 3.4, 3.6, 3.5, 3.3};

    QPixmap result = DVE::PlotEngine::renderTPMTrend(puffs, tpm, "TPM Trend");
    QVERIFY(!result.isNull());
}

void TestPlotEngine::testTPMBarChart()
{
    QStringList names = {"Sample A", "Sample B", "Sample C"};
    QVector<double> avg = {3.5, 4.2, 3.8};
    QVector<double> sd  = {0.3, 0.5, 0.2};

    QPixmap result = DVE::PlotEngine::renderTPMBarChart(names, avg, sd, "TPM Bar");
    QVERIFY(!result.isNull());
}

void TestPlotEngine::testEmptyData()
{
    DVE::PlotConfig config;
    config.title = "Empty";
    config.xLabel = "X";
    config.yLabel = "Y";
    config.width = 400;
    config.height = 300;

    // Empty series list — should not crash
    QPixmap result = DVE::PlotEngine::renderLinePlot({}, config);
    // May be null or a blank pixmap — just ensure no crash
    Q_UNUSED(result);
}

void TestPlotEngine::testPlotDimensions()
{
    DVE::PlotSeries series;
    series.label = "Dim";
    series.x = {1.0, 2.0};
    series.y = {5.0, 10.0};
    series.color = Qt::green;
    series.drawLine = true;
    series.drawDots = false;

    DVE::PlotConfig config;
    config.title = "Dimensions";
    config.xLabel = "X";
    config.yLabel = "Y";
    config.width = 400;
    config.height = 300;

    QPixmap result = DVE::PlotEngine::renderLinePlot({series}, config);
    QVERIFY(!result.isNull());
    QCOMPARE(result.width(), 400);
    QCOMPARE(result.height(), 300);
}

void TestPlotEngine::barChart_drawsLegendEntries()
{
    DVE::PlotConfig cfg;
    cfg.title = "Test"; cfg.width = 600; cfg.height = 400;
    cfg.legendEntries = {
        {"File A", QColor(255, 0, 0)},
        {"File B", QColor(0, 0, 255)},
    };
    QVector<QString> labels = {"s1", "s2"};
    QVector<double>  vals   = {3.0, 4.0};
    QPixmap pm = DVE::PlotEngine::renderBarChart(labels, vals, cfg);
    QImage img = pm.toImage();

    bool sawRed = false, sawBlue = false;
    for (int y = 0; y < img.height() / 4; ++y) {
        for (int x = img.width() / 2; x < img.width(); ++x) {
            QColor c = img.pixelColor(x, y);
            if (c.red()   > 200 && c.green() < 50  && c.blue() < 50) sawRed  = true;
            if (c.blue()  > 200 && c.red()   < 50  && c.green() < 50) sawBlue = true;
            if (sawRed && sawBlue) break;
        }
        if (sawRed && sawBlue) break;
    }
    QVERIFY(sawRed);
    QVERIFY(sawBlue);
}

void TestPlotEngine::applyDataXRange_spansData()
{
    // Standing rule: the x axis always runs 0 .. last data point + 1 so
    // neither the first nor the last point sits cut off on an axis edge.
    DVE::PlotSeries a; a.x = {5.0, 10.0, 20.0}; a.y = {1.0, 1.1, 1.2};
    DVE::PlotSeries b; b.x = {10.0, 30.0};      b.y = {1.3, 1.4};

    DVE::PlotConfig cfg;   // defaults are xMin=0, xMax=1
    DVE::PlotEngine::applyDataXRange(cfg, {a, b});
    QCOMPARE(cfg.xMin, 0.0);
    QCOMPARE(cfg.xMax, 31.0);

    // No data: keep a sane 0..1 axis.
    DVE::PlotConfig empty;
    DVE::PlotEngine::applyDataXRange(empty, {});
    QCOMPARE(empty.xMin, 0.0);
    QCOMPARE(empty.xMax, 1.0);

    // Single point still yields a non-zero span from 0.
    DVE::PlotSeries one; one.x = {10.0}; one.y = {1.5};
    DVE::PlotConfig single;
    DVE::PlotEngine::applyDataXRange(single, {one});
    QCOMPARE(single.xMin, 0.0);
    QCOMPARE(single.xMax, 11.0);
}

void TestPlotEngine::applyAnchoredYRange_zeroToMaxPlusOne()
{
    // Standing rule for EVERY x-y plot: y anchored at 0, top = data max + 1.
    DVE::PlotSeries a; a.x = {1.0, 2.0}; a.y = {1.52, 1.74};
    DVE::PlotSeries b; b.x = {1.0, 2.0}; b.y = {1.60, 1.69};

    DVE::PlotConfig cfg;
    DVE::PlotEngine::applyAnchoredYRange(cfg, {a, b});
    QCOMPARE(cfg.yMin, 0.0);
    QCOMPARE(cfg.yMax, 2.74);

    // High-TPM data gets the same rule (no 7-unit floor, no 25 clamp).
    DVE::PlotSeries t; t.x = {1.0, 2.0, 3.0}; t.y = {5.0, 12.0, 8.0};
    DVE::PlotConfig tpm;
    DVE::PlotEngine::applyAnchoredYRange(tpm, {t});
    QCOMPARE(tpm.yMin, 0.0);
    QCOMPARE(tpm.yMax, 13.0);

    // No data: 0..1.
    DVE::PlotConfig empty;
    DVE::PlotEngine::applyAnchoredYRange(empty, {});
    QCOMPARE(empty.yMin, 0.0);
    QCOMPARE(empty.yMax, 1.0);
}

void TestPlotEngine::autoRange_anchorsYAtZero()
{
    // Data well above zero must still scale from a y-axis pinned at 0,
    // so the curve is never visually "floated" off the baseline.
    DVE::PlotSeries series;
    series.x = {1.0, 2.0, 3.0, 4.0, 5.0};
    series.y = {120.0, 135.0, 128.0, 142.0, 150.0};

    double yMin = DVE::PlotEngine::autoRangeYMinForTesting({series});
    QCOMPARE(yMin, 0.0);
}

void TestPlotEngine::ringedPoint_changesRenderedPixels()
{
    // A series rendered WITH a per-point amber ring must differ from the same
    // series WITHOUT rings, in the neighbourhood of the ringed point. The ring
    // is drawn at (x[2], y[2]).
    DVE::PlotSeries base;
    base.x = {1.0, 2.0, 3.0, 4.0, 5.0};
    base.y = {10.0, 20.0, 15.0, 25.0, 30.0};
    base.color = Qt::blue;
    base.drawLine = true;
    base.drawDots = true;

    DVE::PlotConfig cfg;
    cfg.title = "Ring Test";
    cfg.xLabel = "X";
    cfg.yLabel = "Y";
    cfg.width = 800;
    cfg.height = 600;
    cfg.showLegend = false;

    const QImage plain = DVE::PlotEngine::renderLinePlot({base}, cfg).toImage();

    DVE::PlotSeries ringedSeries = base;
    ringedSeries.ringed = {2};
    const QImage ringed = DVE::PlotEngine::renderLinePlot({ringedSeries}, cfg).toImage();

    QVERIFY(!plain.isNull());
    QVERIFY(!ringed.isNull());
    QCOMPARE(plain.size(), ringed.size());

    // Primary signal (per the design): the ringed render must differ from the
    // plain one — the amber ring adds pixels the plain render does not have.
    int differingPixels = 0;
    for (int y = 0; y < plain.height(); ++y)
        for (int x = 0; x < plain.width(); ++x)
            if (plain.pixel(x, y) != ringed.pixel(x, y)) ++differingPixels;
    QVERIFY2(differingPixels > 0, "ringed render is pixel-identical to the plain render");

    // The amber ring colour (0xBA7517) must appear in the RINGED render. Use a
    // tight tolerance so it cannot be confused with the blue series (0x0066CC)
    // or its darkened dot border — amber has red>>blue, blue has blue>>red.
    auto countAmber = [](const QImage& img) {
        int n = 0;
        for (int y = 0; y < img.height(); ++y)
            for (int x = 0; x < img.width(); ++x) {
                const QColor c = img.pixelColor(x, y);
                if (qAbs(c.red() - 0xBA) <= 12 && qAbs(c.green() - 0x75) <= 12
                    && qAbs(c.blue() - 0x17) <= 12)
                    ++n;
            }
        return n;
    };
    QVERIFY2(countAmber(ringed) > 0, "amber ring colour not found in the ringed render");
}

void TestPlotEngine::transform_roundTripsDataAndPixel()
{
    DVE::PlotSeries s;
    s.label = "S1";
    s.x = {0, 1, 2, 3, 4};
    s.y = {0.0, 1.0, 2.0, 3.0, 4.0};
    DVE::PlotConfig cfg;
    cfg.width = 800; cfg.height = 500; cfg.autoScale = true;
    DVE::PlotTransform tf;
    QPixmap pm = DVE::PlotEngine::renderLinePlot({s}, cfg, &tf);
    QVERIFY(!pm.isNull());
    QVERIFY(tf.valid);
    QVERIFY(tf.plotRect.width() > 0 && tf.plotRect.height() > 0);
    // A data point maps into the plot rect, and pixel->data->pixel round-trips.
    const QPointF px = tf.dataToPixel(2, 2.0);
    QVERIFY(tf.plotRect.contains(px));
    const QPointF back = tf.dataToPixel(tf.pixelToData(px).x(), tf.pixelToData(px).y());
    QVERIFY(qAbs(back.x() - px.x()) < 0.5);
    QVERIFY(qAbs(back.y() - px.y()) < 0.5);
    // x grows rightward, y grows upward.
    QVERIFY(tf.dataToPixel(3, 2.0).x() > tf.dataToPixel(1, 2.0).x());
    QVERIFY(tf.dataToPixel(2, 3.0).y() < tf.dataToPixel(2, 1.0).y());
}

void TestPlotEngine::annotatedExport_growsHeightAndDrawsAmberLines()
{
    QPixmap base(800, 500); base.fill(Qt::white);
    DVE::PlotTransform tf;
    tf.plotRect = QRectF(70, 50, 700, 400);
    tf.xMin = 0; tf.xMax = 10; tf.yMin = 0; tf.yMax = 5; tf.valid = true;
    QVector<QString> rows = {"Sample A - puff 1 - TPM 1.0 (avg 1.4) - clog",
                             "Sample A - puff 6 - TPM 3.0 (avg 1.4) - harsh"};
    QVector<double> puffs = {1, 6};
    QImage img = DVE::PlotEngine::composeAnnotatedExport(base, tf, rows, puffs);
    QVERIFY(img.height() > base.height());          // header band added
    QVERIFY(img.width()  == base.width());
    // Amber pixels (0xBA,0x75,0x17) appear (the vertical note-lines).
    const QRgb amber = qRgb(0xBA, 0x75, 0x17);
    int amberCount = 0;
    for (int y = 0; y < img.height(); ++y)
        for (int x = 0; x < img.width(); ++x)
            if (img.pixel(x, y) == amber) { ++amberCount; break; }
    QVERIFY2(amberCount >= 2, "expected amber vertical lines for 2 noted puffs");
}

QTEST_MAIN(TestPlotEngine)
#include "tst_plotengine.moc"
