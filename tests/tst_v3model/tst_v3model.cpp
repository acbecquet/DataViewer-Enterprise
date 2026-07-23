#include <QtTest>
#include "model/MetricSample.h"

using namespace DVE::model;

class TestV3Model : public QObject {
    Q_OBJECT
private slots:
    void seriesLookup();
    void rowCountIsLongestSeries();
};

void TestV3Model::seriesLookup()
{
    Sample s;
    s.data.append(MetricSeries{QStringLiteral("puffs"), {10, 20}});
    s.data.append(MetricSeries{QStringLiteral("tpm"),   {3.5, 3.5}});
    QVERIFY(s.series("tpm") != nullptr);
    QCOMPARE(s.series("tpm")->values.size(), 2);
    QVERIFY(s.series("nope") == nullptr);
}

void TestV3Model::rowCountIsLongestSeries()
{
    Sample s;
    s.data.append(MetricSeries{QStringLiteral("puffs"), {10, 20, 30}});
    s.data.append(MetricSeries{QStringLiteral("notes"), {QStringLiteral("a")}});
    QCOMPARE(s.rowCount(), 3);
}

QTEST_MAIN(TestV3Model)
#include "tst_v3model.moc"
