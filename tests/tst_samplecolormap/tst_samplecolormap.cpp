#include <QtTest>
#include "utils/SampleColorMap.h"
#include "utils/AppTheme.h"

using namespace DVE;

class TestSampleColorMap : public QObject
{
    Q_OBJECT
private slots:
    void init() { SampleColorMap::instance().clear(); }

    void indexFor_stableAndDedup()
    {
        auto& m = SampleColorMap::instance();
        const int a = m.indexFor("Alpha");
        const int b = m.indexFor("Beta");
        QVERIFY(a != b);
        QCOMPARE(m.indexFor("Alpha"), a);          // stable
        QCOMPARE(m.indexFor("  Alpha  "), a);      // trimmed key
        QCOMPARE(m.indexFor(""), -1);              // blank -> sentinel
        QCOMPARE(m.indexFor("   "), -1);
    }

    void colorsForPlot_pinsNamesAcrossPlots()
    {
        auto& m = SampleColorMap::instance();
        const QVector<QColor> c1 = m.colorsForPlot({ "Alpha", "Beta" });
        const QVector<QColor> c2 = m.colorsForPlot({ "Beta", "Alpha" });   // reordered
        QCOMPARE(c1.size(), 2);
        QVERIFY(c1[0] != c1[1]);                   // distinct in-plot
        QCOMPARE(c2[0], c1[1]);                    // Beta same color both plots
        QCOMPARE(c2[1], c1[0]);                    // Alpha same color both plots
    }

    void colorsForPlot_blanksDistinctAndUnused()
    {
        auto& m = SampleColorMap::instance();
        m.indexFor("Alpha");                       // -> idx 0
        // Plot: [Alpha, "", ""]. Alpha keeps idx 0; blanks take lowest unused (1,2).
        const QVector<QColor> c = m.colorsForPlot({ "Alpha", "", "" });
        QCOMPARE(c[0], AppTheme::seriesColor(0));
        QVERIFY(c[1] != c[0] && c[2] != c[0] && c[1] != c[2]);   // all distinct
    }

    void clear_resets()
    {
        auto& m = SampleColorMap::instance();
        m.indexFor("Alpha"); m.indexFor("Beta");
        QVERIFY(m.size() >= 2);
        m.clear();
        QCOMPARE(m.size(), 0);
        QCOMPARE(m.indexFor("Zeta"), 0);           // numbering restarts
    }
};

QTEST_MAIN(TestSampleColorMap)
#include "tst_samplecolormap.moc"
