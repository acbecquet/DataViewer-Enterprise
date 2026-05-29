#include <QtTest>
#include <QColor>
#include <QSet>
#include "AppTheme.h"

class TestAppTheme : public QObject
{
    Q_OBJECT
private slots:
    void testCountAndEmpty();
    void testAllUnique();
    void testSeriesColorMatchesVector();
    void testNoYellow();
    void testShadeDistinctAndBanded();
    void testShadeGrayStaysNeutral();
    void testGroupedComparisonUnique();
};

static int rgbKey(const QColor& c) { return (c.red() << 16) | (c.green() << 8) | c.blue(); }

void TestAppTheme::testCountAndEmpty()
{
    QVERIFY(AppTheme::seriesColors(0).isEmpty());
    QCOMPARE(AppTheme::seriesColors(5).size(), 5);
    QCOMPARE(AppTheme::seriesColors(25).size(), 25);
    QCOMPARE(AppTheme::seriesColor(-1), AppTheme::seriesColor(0));     // negative clamps to 0
    QCOMPARE(AppTheme::seriesColor(-100), AppTheme::seriesColor(0));
}

void TestAppTheme::testAllUnique()
{
    for (int n : {1, 5, 12, 20, 21, 25, 40, 60}) {
        QSet<int> seen;
        const QVector<QColor> pal = AppTheme::seriesColors(n);
        for (const QColor& c : pal) seen.insert(rgbKey(c));
        QCOMPARE(seen.size(), n);   // no color repeats
    }
}

void TestAppTheme::testSeriesColorMatchesVector()
{
    const QVector<QColor> pal = AppTheme::seriesColors(30);
    for (int i = 0; i < pal.size(); ++i)
        QCOMPARE(AppTheme::seriesColor(i), pal[i]);
}

void TestAppTheme::testNoYellow()
{
    for (int i = 0; i < 60; ++i) {
        const int h = AppTheme::seriesColor(i).hue();   // -1 if achromatic
        if (h >= 0)
            QVERIFY2(!(h >= 45 && h <= 70),
                     qPrintable(QString("color %1 has yellow hue %2").arg(i).arg(h)));
    }
}

void TestAppTheme::testShadeDistinctAndBanded()
{
    const QColor base = AppTheme::seriesColor(0);   // chromatic (blue)
    const int count = 6;
    QSet<int> seen;
    int minV = 255, maxV = 0;
    for (int i = 0; i < count; ++i) {
        const QColor s = AppTheme::shade(base, i, count);
        seen.insert(rgbKey(s));
        QCOMPARE(s.hue(), base.hue());      // hue preserved -> grouping intact
        minV = qMin(minV, s.value());
        maxV = qMax(maxV, s.value());
    }
    QCOMPARE(seen.size(), count);            // shades all distinct
    QVERIFY(maxV <= 217);                    // light end capped (projector-safe)
    QVERIFY(minV >= 140);                    // dark end floored
    QCOMPARE(AppTheme::shade(base, 0, 1), base);  // count<=1 returns base unchanged
}

void TestAppTheme::testShadeGrayStaysNeutral()
{
    const QColor gray(127, 127, 127);
    QSet<int> seen;
    for (int i = 0; i < 4; ++i) {
        const QColor s = AppTheme::shade(gray, i, 4);
        QCOMPARE(s.saturation(), 0);         // stays neutral, no tint
        seen.insert(rgbKey(s));
    }
    QCOMPARE(seen.size(), 4);                // distinct grays
}

void TestAppTheme::testGroupedComparisonUnique()
{
    // The Lifetime comparison guarantee: distinct file hues + per-sample shades
    // => no two bars across the whole chart share a color.
    const int files = 4, samplesPerFile = 6;
    const QVector<QColor> fileHues = AppTheme::seriesColors(files);
    QSet<int> seen;
    for (int f = 0; f < files; ++f)
        for (int s = 0; s < samplesPerFile; ++s)
            seen.insert(rgbKey(AppTheme::shade(fileHues[f], s, samplesPerFile)));
    QCOMPARE(seen.size(), files * samplesPerFile);
}

QTEST_APPLESS_MAIN(TestAppTheme)
#include "tst_apptheme.moc"
