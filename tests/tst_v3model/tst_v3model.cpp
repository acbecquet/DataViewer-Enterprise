#include <QtTest>
#include "model/MetricSample.h"
#include "model/StandardSchema.h"
#include "model/SchemaDrivenReader.h"
#include "model/LegacyAdapter.h"
#include "model/RegimeParser.h"
#include "model/MetricRegistry.h"
#include "model/Manifest.h"
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
    void adapterLowersMetadataAndGrid();
    void regimeParserStandard();
    void regimeParserFourPart();
    void regimeParserRealVariants();
    void regimeParserRejectsGarbage();
    void registryResolvesEveryObservedSpelling();
    void registryPerPuffAliases();
    void registryNamespacesAreCollisionFree();
    void registryNewTypesAndTags();
    void standardSchemaDrawsFromRegistry();
    void manifestRoundTripsSchema();
    void manifestParsesCustomColumnsAndTags();
    void manifestPokaYokesWarnNeverFail();
    void manifestSheetScoping();
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
    g[4][s.columnPos("tpm")] = 42.5;
    const Sheet sheet = SchemaDrivenReader::parseSheet(g, "Lifetime Test", s, false);
    QCOMPARE(sheet.samples[0].series("tpm")->values[0].toDouble(), 42.5);
    QVERIFY(sheet.samples[0].series("variation_tpm")->values[0].toString().isEmpty());
}

void TestV3Model::normalizeHeaderStripsToAlnum()
{
    QCOMPARE(SchemaDrivenReader::normalizeHeader(" Before Weight (g) "), QStringLiteral("beforeweightg"));
}

void TestV3Model::adapterLowersMetadataAndGrid()
{
    const TemplateSchema s = standardV1(false);
    Sample m;
    m.headers.insert("test_name", "My Test");
    m.headers.insert("voltage", "3.7");
    m.headers.insert("resistance", "1.2 Ohm");        // tolerant parse must apply
    m.headers.insert("heating_technology", "CCELL3.0");
    for (const MetricDef& c : s.columns) m.data.append(MetricSeries{c.key, {}});
    // 2 rows of puffs/before/after
    m.data[0].values = {10, 20};
    m.data[1].values = {25.10, 25.065};
    m.data[2].values = {25.065, 25.032};

    const ExcelReader::SampleData raw = LegacyAdapter::lowerSample(m, s, /*blockIndex=*/1);
    QCOMPARE(raw.metadata.testName, QStringLiteral("My Test"));
    QCOMPARE(raw.metadata.voltage, 3.7);
    QCOMPARE(raw.metadata.resistance, 1.2);
    // power = V^2 / (R + CCELL3.0 offset 0.78) - mirrors ExcelReader::extractMetadata
    QVERIFY(qAbs(raw.metadata.power - (3.7 * 3.7) / (1.2 + 0.78)) < 1e-9);
    QCOMPARE(raw.startColumn, 12);
    QCOMPARE(raw.dataRows.size(), 2);
    QCOMPARE(raw.dataRows[0].size(), 12);             // full block width restored
    QCOMPARE(raw.dataRows[0][0].toInt(), 10);
    QCOMPARE(raw.dataRows[1][2].toDouble(), 25.032);
}

void TestV3Model::regimeParserStandard()
{
    const auto r = RegimeParser::parse(QStringLiteral("60mL/3s/30s"));
    QVERIFY(r.valid);
    QCOMPARE(r.puffVolumeMl, 60.0);
    QCOMPARE(r.puffTimeS, 3.0);
    QCOMPARE(r.puffRestS, 30.0);
    QCOMPARE(r.sessionRestS, 0.0);   // 3-part regime defaults session rest to 0 (registry 8.2)
}

void TestV3Model::regimeParserFourPart()
{
    const auto r = RegimeParser::parse(QStringLiteral("60mL/3s/30s/5minute"));
    QVERIFY(r.valid);
    QCOMPARE(r.sessionRestS, 300.0);
}

void TestV3Model::regimeParserRealVariants()
{
    QCOMPARE(RegimeParser::parse(QStringLiteral("100mL/2.5s/15s")).puffTimeS, 2.5);
    QCOMPARE(RegimeParser::parse(QStringLiteral("200mL/10s/60s")).puffRestS, 60.0);
    QVERIFY(RegimeParser::parse(QStringLiteral("60 ml / 3 s / 30 s")).valid);  // spacing + case tolerant
    QCOMPARE(RegimeParser::parse(QStringLiteral("60mL/3s/30s/90s")).sessionRestS, 90.0);
    QCOMPARE(RegimeParser::parse(QStringLiteral("60mL/3s/30s/2min")).sessionRestS, 120.0);
}

void TestV3Model::regimeParserRejectsGarbage()
{
    QVERIFY(!RegimeParser::parse(QString()).valid);
    QVERIFY(!RegimeParser::parse(QStringLiteral("as needed")).valid);
    QVERIFY(!RegimeParser::parse(QStringLiteral("60mL/3s")).valid);          // too few parts
    QVERIFY(!RegimeParser::parse(QStringLiteral("60mL/3s/30s/x/y")).valid);  // too many parts
}

void TestV3Model::registryResolvesEveryObservedSpelling()
{
    // Data-driven: verbatim spellings from the RATIFIED registry doc -> canonical key.
    const struct { const char* spelling; const char* key; } metricRows[] = {
        {"puffs", "puffs"},
        {"Before weight/g", "before_weight"},
        {"Before Weight/g", "before_weight"},
        {"Before weight (g)", "before_weight"},
        {"After weight/g", "after_weight"},
        {"Draw Pressure (kpa)", "draw_pressure"},
        {"Resistance", "resistance"},
        {"Puffing Regime", "puffing_regime"},
        {"Smell (1-4)", "smell"},
        {"Smell (0-4)", "smell"},
        {"Clog (Y/N)", "clog"},
        {"Clog?", "clog"},
        {"Notes", "notes"},
        {"TPM (mg/puff)", "tpm"},
        {"TPM Power Density (mg/(W*s))", "tpm_power_density"},   // current era
        {"TPM Power Density (mg/puff/W)", "tpm_puff_density"},   // old era = TPM/P (registry 9.1)
        {"Variation in TPM (%)", "variation_tpm"},
        {"TPM Consistency", "tpm_consistency"},
        {"Rolling Average TPM", "rolling_avg_tpm"},
        {"Oil Consumed (Cumulative, g)", "oil_consumed"},
        {"Chronology", "chronology"},
        {"Failure? (if yes, add detailed notes)", "failure"},
    };
    for (const auto& row : metricRows) {
        const MetricDef* m = MetricRegistry::metricByAlias(
            SchemaDrivenReader::normalizeHeader(QString::fromUtf8(row.spelling)));
        QVERIFY2(m, row.spelling);
        QCOMPARE(m->key, QString::fromUtf8(row.key));
    }

    const struct { const char* label; const char* key; } headerRows[] = {
        {"Date:", "date"},
        {"Tester:", "tester"},
        {"Sample ID:", "sample_id"},
        {"Cart #", "sample_id"},
        {"Project:", "project_name"},
        {"Sample:", "sample_suffix"},
        {"Distributor:", "distributor"},
        {"Resistance (\xCE\xA9):", "resistance"},
        {"Resistance (Ohms):", "resistance"},
        {"Ri (Ohms)", "resistance_initial"},
        {"Rf (Ohms)", "resistance_final"},
        {"Voltage:", "voltage"},
        {"Power:", "power"},
        {"Heating Technology:", "heating_technology"},
        {"Heater Technology:", "heating_technology"},
        {"Media:", "media"},
        {"Viscosity:", "viscosity"},
        {"Initial Oil Mass:", "initial_oil_mass"},
        {"Fill Volume:", "fill_volume"},
        {"Number of Samples", "number_of_samples"},
        {"Puffing Regime:", "puffing_regime"},
        {"Puff Regime", "puffing_regime"},
        {"Did this burn?", "did_burn"},
        {"Did this clog?", "did_clog"},
        {"Did this leak?", "did_leak"},
        {"Coil Material", "coil_material"},
        {"Thermal Conductivity", "thermal_conductivity"},
        {"Column inner diameter", "column_inner_diameter"},
        {"Column length", "column_length"},
        {"Coil shape", "coil_shape"},
        {"Cotton length (if applicable)", "cotton_length"},
    };
    for (const auto& row : headerRows) {
        const HeaderFieldDef* h = MetricRegistry::headerByLabel(
            SchemaDrivenReader::normalizeHeader(QString::fromUtf8(row.label)));
        QVERIFY2(h, row.label);
        QCOMPARE(h->key, QString::fromUtf8(row.key));
    }
}

void TestV3Model::registryPerPuffAliases()
{
    for (int i = 1; i <= 5; ++i) {
        const auto hit = MetricRegistry::perPuffAlias(QStringLiteral("pv%1").arg(i));
        QCOMPARE(hit.targetKey, QStringLiteral("draw_pressure_per_puff"));
        QCOMPARE(hit.index, i);
    }
    QCOMPARE(MetricRegistry::perPuffAlias(QStringLiteral("puffs")).index, 0);  // miss
}

void TestV3Model::registryNamespacesAreCollisionFree()
{
    // Naming-policy rule 1/7: within a namespace no two ENTRIES may claim the
    // same normalized alias (displayName, key, and headerAliases all count).
    // Duplicates WITHIN one entry (key == normalized displayName, redundant
    // punctuation-only alias spellings) are harmless and deduped first.
    QSet<QString> seen;
    for (const MetricDef& m : MetricRegistry::allMetrics()) {
        QStringList names{m.displayName, m.key};
        names += m.headerAliases;
        QSet<QString> mine;
        for (const QString& n : names) {
            const QString norm = SchemaDrivenReader::normalizeHeader(n);
            if (norm.isEmpty()) continue;
            mine.insert(norm);
        }
        for (const QString& norm : mine) {
            QVERIFY2(!seen.contains(norm), qPrintable(m.key + ": " + norm));
            seen.insert(norm);
        }
    }
    QSet<QString> seenH;
    for (const HeaderFieldDef& h : MetricRegistry::allHeaderFields()) {
        QStringList names{h.displayName, h.key};
        names += MetricRegistry::headerAliasesFor(h.key);
        QSet<QString> mine;
        for (const QString& n : names) {
            const QString norm = SchemaDrivenReader::normalizeHeader(n);
            if (norm.isEmpty()) continue;
            mine.insert(norm);
        }
        for (const QString& norm : mine) {
            QVERIFY2(!seenH.contains(norm), qPrintable(h.key + ": " + norm));
            seenH.insert(norm);
        }
    }
}

void TestV3Model::registryNewTypesAndTags()
{
    const MetricDef* dp = MetricRegistry::metric(QStringLiteral("draw_pressure_per_puff"));
    QVERIFY(dp);
    QCOMPARE(dp->type, ValueType::NumberList);
    QCOMPARE(dp->unit, QStringLiteral("kPa"));
    QVERIFY(dp->tags.contains(QStringLiteral("source")));

    const MetricDef* v = MetricRegistry::metric(QStringLiteral("voltage"));
    QVERIFY(v);
    QCOMPARE(v->type, ValueType::Mixed);         // number OR curve-name string (owner D1 addendum)

    const MetricDef* img = MetricRegistry::metric(QStringLiteral("image"));
    QVERIFY(img);
    QCOMPARE(img->type, ValueType::Image);

    const HeaderFieldDef* burn = MetricRegistry::headerField(QStringLiteral("did_burn"));
    QVERIFY(burn);
    QCOMPARE(burn->type, ValueType::Bool);

    // Regime-split quartet registered with the documented default.
    const MetricDef* srt = MetricRegistry::metric(QStringLiteral("session_rest_time"));
    QVERIFY(srt);
    QCOMPARE(srt->unit, QStringLiteral("s"));
}

void TestV3Model::standardSchemaDrawsFromRegistry()
{
    const TemplateSchema s = standardV1(false);
    // Column KEY ORDER is the byte-identity contract - it must never change here.
    const QStringList expectedKeys{
        "puffs", "before_weight", "after_weight", "draw_pressure", "resistance",
        "smell", "clog", "notes", "tpm", "tpm_power_density", "variation_tpm", "oil_consumed"};
    QCOMPARE(s.columns.size(), expectedKeys.size());
    for (int i = 0; i < expectedKeys.size(); ++i)
        QCOMPARE(s.columns[i].key, expectedKeys[i]);
    // Registry aliases flow through (superset of the old hand-maintained sets).
    QVERIFY(s.columns[1].headerAliases.contains(QStringLiteral("Before weight/g")));
    // Layout-side flags still applied by the builder.
    QVERIFY(s.columns[8].plottable);                       // tpm
    QVERIFY(s.column("smell") && s.column("smell")->editable);
}

void TestV3Model::manifestRoundTripsSchema()
{
    // standardV1 -> grid -> parse must reproduce keys, order, geometry,
    // header positions, and per-row-regime column.
    const TemplateSchema src = standardV1(true);
    const QVector<QVector<QVariant>> grid = Manifest::gridFor(src, {QStringLiteral("*")});
    const Manifest::ParseResult pr = Manifest::parse(grid);
    QVERIFY(pr.warnings.isEmpty());
    QCOMPARE(pr.blocks.size(), 1);
    const TemplateSchema& back = pr.blocks[0].schema;
    QCOMPARE(back.blockCols, src.blockCols);
    QCOMPARE(back.dataStartRow, src.dataStartRow);
    QCOMPARE(back.columnHeaderRow, src.columnHeaderRow);
    QCOMPARE(back.columns.size(), src.columns.size());
    for (int i = 0; i < src.columns.size(); ++i) {
        QCOMPARE(back.columns[i].key,  src.columns[i].key);
        QCOMPARE(back.columns[i].role, src.columns[i].role);
    }
    QCOMPARE(back.headerFields.size(), src.headerFields.size());
    const HeaderFieldDef* media = back.headerField(QStringLiteral("media"));
    QVERIFY(media);
    QCOMPARE(media->row, 2);
    QCOMPARE(media->col, 2);
    // Registry inheritance: known keys regain their alias sets.
    QVERIFY(back.column(QStringLiteral("before_weight"))->headerAliases.contains(
        QStringLiteral("Before weight/g")));
}

void TestV3Model::manifestParsesCustomColumnsAndTags()
{
    TemplateSchema s;
    s.schemaId = QStringLiteral("demo");
    s.version = 1;
    s.headerRows = 3; s.columnHeaderRow = 4; s.dataStartRow = 5; s.blockCols = 3;
    MetricDef custom;
    custom.key = QStringLiteral("coil_temp");
    custom.displayName = QStringLiteral("Coil Temp (C)");
    custom.type = ValueType::Number;
    custom.unit = QStringLiteral("C");
    custom.role = Role::Measured;
    custom.tags.insert(QStringLiteral("source"), QStringLiteral("thermocouple"));
    s.columns = { *MetricRegistry::metric("puffs"), *MetricRegistry::metric("tpm"), custom };
    const auto pr = Manifest::parse(Manifest::gridFor(s, {QStringLiteral("Demo Sheet")}));
    QVERIFY(pr.warnings.isEmpty());
    QCOMPARE(pr.blocks[0].sheets, QStringList{QStringLiteral("Demo Sheet")});
    const MetricDef* back = pr.blocks[0].schema.column(QStringLiteral("coil_temp"));
    QVERIFY(back);
    QCOMPARE(back->unit, QStringLiteral("C"));
    QCOMPARE(back->tags.value(QStringLiteral("source")), QStringLiteral("thermocouple"));
}

void TestV3Model::manifestPokaYokesWarnNeverFail()
{
    TemplateSchema s = standardV1(false);
    QVector<QVector<QVariant>> grid = Manifest::gridFor(s, {QStringLiteral("*")});
    // Corrupt: an unknown role and a duplicate key row.
    // Locate the [column] row for "smell" and set its role cell to "banana";
    // duplicate the row after it ("clog"). (Walk the grid: rows after [column].)
    bool corrupted = false;
    for (int r = 0; r < grid.size(); ++r) {
        if (grid[r].value(0).toString() == QLatin1String("smell")) {
            grid[r][4] = QStringLiteral("banana");        // role cell
            grid.insert(r + 1, grid[r + 1]);              // duplicate next row once
            corrupted = true;
            break;
        }
    }
    QVERIFY(corrupted);
    const auto pr = Manifest::parse(grid);
    QVERIFY(!pr.warnings.isEmpty());                      // warned...
    QCOMPARE(pr.blocks.size(), 1);                        // ...but parsed
    QVERIFY(pr.blocks[0].schema.column(QStringLiteral("smell")));  // role defaulted
}

void TestV3Model::manifestSheetScoping()
{
    const TemplateSchema a = standardV1(false);
    const TemplateSchema b = standardV1(true);
    QVector<QVector<QVariant>> grid = Manifest::gridFor(a, {QStringLiteral("Special")});
    grid += Manifest::gridFor(b, {QStringLiteral("*")});
    const auto pr = Manifest::parse(grid);
    QCOMPARE(pr.blocks.size(), 2);
    QCOMPARE(Manifest::blockForSheet(pr, QStringLiteral("Special"))->schema.columns[4].key,
             QStringLiteral("resistance"));
    QCOMPARE(Manifest::blockForSheet(pr, QStringLiteral("Anything"))->schema.columns[4].key,
             QStringLiteral("puffing_regime"));
    QVERIFY(!Manifest::blockForSheet(Manifest::ParseResult{}, QStringLiteral("x")));
}

QTEST_MAIN(TestV3Model)
#include "tst_v3model.moc"
