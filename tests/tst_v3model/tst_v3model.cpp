#include <QtTest>
#include "model/MetricSample.h"
#include "model/StandardSchema.h"
#include "model/SchemaDrivenReader.h"
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
    void readerParsesTwoBlocks();
    void readerMatchesReorderedColumnsByName();
    void readerFallsBackPositionally();
    void readerStopsAtAllEmptyRow();
    void readerMatchesRealTemplateAliases();
    void normalizeHeaderStripsToAlnum();
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

// Helper: build a standard-shaped grid for `blocks` samples, `rows` data rows.
static QVector<QVector<QVariant>> makeStandardGrid(int blocks, int rows,
                                                   bool swapSmellNotes = false)
{
    const TemplateSchema s = standardV1(false);
    QVector<QVector<QVariant>> g(4 + rows);
    for (int b = 0; b < blocks; ++b) {
        const int off = b * s.blockCols;
        auto set = [&](int r, int c, const QVariant& v) {
            if (g[r].size() < off + s.blockCols) g[r].resize(off + s.blockCols);
            g[r][off + c] = v;
        };
        // header band (values only where the schema looks; skip derived fields)
        for (const HeaderFieldDef& h : s.headerFields)
            if (h.calculator.isEmpty())
                set(h.row - 1, h.col - 1, QStringLiteral("%1_%2").arg(h.key).arg(b));
        // column header row (row index 3)
        for (int c = 0; c < s.columns.size(); ++c) {
            int cc = c;
            if (swapSmellNotes && s.columns[c].key == "smell") cc = s.columnPos("notes");
            else if (swapSmellNotes && s.columns[c].key == "notes") cc = s.columnPos("smell");
            set(3, cc, s.columns[c].displayName);
        }
        // data rows
        for (int r = 0; r < rows; ++r) {
            set(4 + r, 0, (r + 1) * 10);            // puffs
            set(4 + r, 1, 25.0 - r * 0.03);         // before
            set(4 + r, 2, 25.0 - r * 0.03 - 0.035); // after
            int smellCol = swapSmellNotes ? s.columnPos("notes") : s.columnPos("smell");
            set(4 + r, smellCol, QStringLiteral("ok"));
        }
    }
    return g;
}

void TestV3Model::readerParsesTwoBlocks()
{
    const TemplateSchema s = standardV1(false);
    const Sheet sheet = SchemaDrivenReader::parseSheet(makeStandardGrid(2, 5), "Lifetime Test", s, false);
    QCOMPARE(sheet.samples.size(), 2);
    QCOMPARE(sheet.samples[0].rowCount(), 5);
    QCOMPARE(sheet.samples[0].headers.value("tester").toString(), QStringLiteral("tester_0"));
    QCOMPARE(sheet.samples[1].headers.value("tester").toString(), QStringLiteral("tester_1"));
    QCOMPARE(sheet.samples[0].series("puffs")->values[2].toInt(), 30);
}

void TestV3Model::readerMatchesReorderedColumnsByName()
{
    const TemplateSchema s = standardV1(false);
    const Sheet sheet = SchemaDrivenReader::parseSheet(makeStandardGrid(1, 3, /*swap*/true), "Lifetime Test", s, false);
    // smell data was PHYSICALLY written into the notes position, but the header
    // text moved with it - name matching must still key it as "smell".
    QCOMPARE(sheet.samples[0].series("smell")->values[0].toString(), QStringLiteral("ok"));
}

void TestV3Model::readerFallsBackPositionally()
{
    const TemplateSchema s = standardV1(false);
    auto g = makeStandardGrid(1, 3);
    g[3].fill(QVariant(), g[3].size());   // wipe every column header
    const Sheet sheet = SchemaDrivenReader::parseSheet(g, "Lifetime Test", s, false);
    QCOMPARE(sheet.samples.size(), 1);
    QCOMPARE(sheet.samples[0].series("puffs")->values.size(), 3);
}

void TestV3Model::readerStopsAtAllEmptyRow()
{
    auto g = makeStandardGrid(1, 3);
    g.append(QVector<QVariant>());                 // blank row
    g.append(QVector<QVariant>{QVariant(999)});    // stray data AFTER blank - ignored
    const Sheet sheet = SchemaDrivenReader::parseSheet(g, "Lifetime Test", standardV1(false), false);
    QCOMPARE(sheet.samples[0].rowCount(), 3);
}

void TestV3Model::readerMatchesRealTemplateAliases()
{
    // Real Dec-2025 header text (an alias, not the displayName) must still match.
    const TemplateSchema s = standardV1(false);
    auto g = makeStandardGrid(1, 2);
    g[3][s.columnPos("tpm")] = QStringLiteral("TPM (mg/puff)");
    const Sheet sheet = SchemaDrivenReader::parseSheet(g, "Lifetime Test", s, false);
    QVERIFY(sheet.samples[0].series("tpm") != nullptr);
}

void TestV3Model::normalizeHeaderStripsToAlnum()
{
    QCOMPARE(SchemaDrivenReader::normalizeHeader(" Before Weight (g) "), QStringLiteral("beforeweightg"));
}

QTEST_MAIN(TestV3Model)
#include "tst_v3model.moc"
