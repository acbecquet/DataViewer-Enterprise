#include <QtTest>
#include "../../src/pipeline/RegimeUtils.h"
#include "../../src/pipeline/ReportData.h"

using namespace DVE;

class TstRegimeUtils : public QObject {
    Q_OBJECT
private:
    static DataRow row(double puffs, double before, double after, const QString& regime) {
        DataRow r; r.puffs = puffs; r.beforeWeight = before; r.afterWeight = after;
        r.puffingRegime = regime; return r;
    }
    static SheetResult twoRegimeSheet() {
        SheetResult s; s.sheetName = "Lifetime Test"; s.hasPerRowRegime = true;
        SampleResult sm; sm.sampleName = "S1"; sm.power = 4.0;
        sm.rows << row(10, 25.10, 25.06, "60mL/3s/30s")
                << row(20, 25.06, 25.02, "60mL/3s/30s")
                << row(30, 25.02, 24.97, "200mL/9s/300s")
                << row(40, 24.97, 24.93, "200mL/9s/300s");
        s.samples << sm;
        return s;
    }
private slots:
    void isRegimeHeader_matches() {
        QVERIFY(RegimeUtils::isRegimeHeader("Puffing Regime"));
        QVERIFY(RegimeUtils::isRegimeHeader("puffing regime"));
        QVERIFY(!RegimeUtils::isRegimeHeader("Resistance (\xce\xa9)"));
        QVERIFY(!RegimeUtils::isRegimeHeader("Resistance"));
    }
    void regimeKey_blankBecomesUnspecified() {
        QCOMPARE(RegimeUtils::regimeKey(row(1,0,0,"  ")), RegimeUtils::unspecifiedLabel());
        QCOMPARE(RegimeUtils::regimeKey(row(1,0,0,"60mL/3s/30s")), QString("60mL/3s/30s"));
    }
    void uniqueRegimes_firstSeenOrder() {
        const QStringList u = RegimeUtils::uniqueRegimes(twoRegimeSheet());
        QCOMPARE(u, (QStringList{"60mL/3s/30s", "200mL/9s/300s"}));
    }
    void uniqueRegimeKeys_includesUnspecifiedForBlanks() {
        SheetResult s; SampleResult sm;
        sm.rows << row(10,25.10,25.06,"60mL/3s/30s") << row(20,25.06,25.02,"");
        s.samples << sm;
        QCOMPARE(RegimeUtils::uniqueRegimeKeys(s),
                 (QStringList{"60mL/3s/30s", RegimeUtils::unspecifiedLabel()}));
        QVERIFY(RegimeUtils::uniqueRegimes(s) == QStringList{"60mL/3s/30s"});
    }
    void uniqueRegimes_file_dedupAcrossSheets() {
        FileResult f;
        SheetResult a; { SampleResult s;
            s.rows << row(10,25.10,25.06,"60mL/3s/30s") << row(20,25.06,25.02,"200mL/9s/300s");
            a.samples << s; }
        SheetResult b; { SampleResult s;
            s.rows << row(10,25.10,25.06,"200mL/9s/300s")    // repeat of sheet A
                   << row(20,25.06,25.02,"100mL/2.5s/15s");  // new
            b.samples << s; }
        f.sheets << a << b;
        QCOMPARE(RegimeUtils::uniqueRegimes(f),
                 (QStringList{"60mL/3s/30s", "200mL/9s/300s", "100mL/2.5s/15s"}));
    }
    void sheetHasRegimeData_true() {
        QVERIFY(RegimeUtils::sheetHasRegimeData(twoRegimeSheet()));
        SheetResult plain; SampleResult sm; sm.rows << row(10,25.1,25.06,"");
        plain.samples << sm;
        QVERIFY(!RegimeUtils::sheetHasRegimeData(plain));
    }
    void filterByRegime_keepsOnlyMatchingRowsAndRecomputes() {
        const SheetResult f = RegimeUtils::filterByRegime(twoRegimeSheet(), "200mL/9s/300s");
        QCOMPARE(f.samples.size(), 1);
        QCOMPARE(f.samples[0].rows.size(), 2);
        for (const DataRow& r : f.samples[0].rows)
            QCOMPARE(r.puffingRegime, QString("200mL/9s/300s"));
        QVERIFY(f.samples[0].averageTPM > 0.0);
    }
    void filterByRegime_dropsEmptySamples() {
        const SheetResult f = RegimeUtils::filterByRegime(twoRegimeSheet(), "nonexistent");
        QCOMPARE(f.samples.size(), 0);
    }
};
QTEST_MAIN(TstRegimeUtils)
#include "tst_regimeutils.moc"
