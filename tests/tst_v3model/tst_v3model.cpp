#include <QtTest>
#include "model/MetricSample.h"
#include "model/StandardSchema.h"
#include <algorithm>

using namespace DVE::model;

class TestV3Model : public QObject {
    Q_OBJECT
private slots:
    void seriesLookup();
    void rowCountIsLongestSeries();
    void standardV1ColumnOrderMatchesCols();
    void standardV1RegimeVariant();
    void standardV1HeaderCells();
    void standardV1HeaderAliasesMatchRealTemplate();
    void standardV1AggregatesLocked();
    void missingKeysReturnSentinels();
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

void TestV3Model::standardV1ColumnOrderMatchesCols()
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/false);
    const QStringList expected{"puffs","before_weight","after_weight","draw_pressure",
        "resistance","smell","clog","notes","tpm","tpm_power_density",
        "variation_tpm","oil_consumed"};
    QCOMPARE(s.columns.size(), 12);
    QCOMPARE(s.blockCols, 12);
    for (int i = 0; i < expected.size(); ++i)
        QCOMPARE(s.columns[i].key, expected[i]);
    QCOMPARE(s.columnPos("tpm"), 8);           // == DVE::Cols::TPM
    QVERIFY(s.column("tpm")->role == Role::Derived);
    QCOMPARE(s.column("tpm")->calculator, QStringLiteral("tpm_v1"));
}

void TestV3Model::standardV1RegimeVariant()
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/true);
    QCOMPARE(s.columns[4].key, QStringLiteral("puffing_regime"));
    QVERIFY(s.columns[4].type == ValueType::Text);
    QVERIFY(s.columns[4].role == Role::Qualitative);
}

void TestV3Model::standardV1HeaderCells()
{
    const TemplateSchema s = standardV1(false);
    const HeaderFieldDef* v = s.headerField("voltage");
    QVERIFY(v);
    QCOMPARE(v->row, 3); QCOMPARE(v->col, 6);
    const HeaderFieldDef* p = s.headerField("power");
    QVERIFY(p);
    QCOMPARE(p->calculator, QStringLiteral("power_v1"));
    QCOMPARE(s.headerFields.size(), 12);
}

void TestV3Model::standardV1HeaderAliasesMatchRealTemplate()
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/false);
    QVERIFY(s.column("tpm")->headerAliases.contains(QStringLiteral("TPM (mg/puff)")));
    QVERIFY(s.column("oil_consumed")->headerAliases.contains(QStringLiteral("Oil Consumed (Cumulative, g)")));
}

void TestV3Model::standardV1AggregatesLocked()
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/false);
    QCOMPARE(s.aggregates.size(), 8);
    const auto it = std::find_if(s.aggregates.begin(), s.aggregates.end(), [](const AggregateDef& a) {
        return a.key == QStringLiteral("efficiency_percent");
    });
    QVERIFY(it != s.aggregates.end());
    QCOMPARE(it->calculator, QStringLiteral("efficiency_v1"));
    QCOMPARE(it->inputs, QStringList({QStringLiteral("total_oil_consumed"), QStringLiteral("header:initial_oil_mass")}));
}

void TestV3Model::missingKeysReturnSentinels()
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/false);
    QVERIFY(s.column("nope") == nullptr);
    QCOMPARE(s.columnPos("nope"), -1);
    QVERIFY(s.headerField("nope") == nullptr);
    QCOMPARE(Sample{}.rowCount(), 0);
}

QTEST_MAIN(TestV3Model)
#include "tst_v3model.moc"
