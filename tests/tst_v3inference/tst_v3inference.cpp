#include <QtTest>
#include <QFileInfo>
#include "model/LegacyAdapter.h"
#include "model/Manifest.h"
#include "model/MetricRegistry.h"
#include "model/SchemaDrivenReader.h"
#include "model/SchemaInference.h"
#include "model/SchemaResolver.h"
#include "model/StandardSchema.h"
#include "model/TemplateSchema.h"
#include "pipeline/CellAddress.h"
#include "pipeline/DataProcessor.h"
#include "pipeline/ReportData.h"
#include "CorpusUtils.h"
#include "TestHelpers.h"

using namespace DVE::model;

class TestV3Inference : public QObject {
    Q_OBJECT
private slots:
    void inferBlockColsS26();
    void inferBlockColsUserSim();
    void standardFitsTrueForStandardGrid();
    void standardFitsTrueForEmptyHeaderRow();
    void standardFitsFalseForS26();
    void standardFitsFalseForUserSim();
    void standardFitsTrueForPaddedSingleBlockGrid();
    void standardFitsFalseForPaddedS26();
    void inferSchemaS26Columns();
    void inferSchemaS26HeaderFields();
    void inferSchemaUserSimColumns();
    void inferSchemaUserSimHeaderFields();
    void registryLabelsRecognized();
    void registryRfLabelCanonicalized();
    void resistanceInitialLowersToSampleResistance();
    void pvColumnsAssemblePerPuffList();
    void writeProvenanceRecordedOnInferredParse();
    // W3b (smoke-fix Task 7): the lowering's puff extrapolation must be
    // disabled when the reader flagged the sheet's caches as destroyed, and
    // the count must ride out on the SheetResult (warning + DB-save block).
    void strippedCountGatesLoweringPuffExtrapolation();

    // Phase 2c Task 3: NameFirst resolved-slot exposure + the generalized
    // key-based lowering (manifest-era; inference outputs stay byte-unchanged).
    void parseSheetExposesResolvedSlots();
    void manifestSheetLowersByKeyWithProvenanceInSlotOrder();
    // Phase 2c Task 6 hardening: over-padded grids must not grow phantom
    // trailing samples on the schema (NameFirst) paths.
    void phantomTrailingSamplesDroppedOnSchemaPaths();
    // Phase 2c Task 4: grid-level rehearsal of the processSheet manifest
    // wiring (manifest grid -> resolver rung 1 -> NameFirst parse ->
    // generalized lowering). processSheet itself needs an ExcelReader with a
    // loaded xlsx, so the FILE-level E2E lands with Task 5's demo fixture.
    void manifestResolutionRoutesShuffledGrid();

    // End-to-end value tests: run the REAL DataProcessor::processFile pipeline
    // (openpyxl subprocess) over the inference fixtures and assert the KNOWN
    // fields land correctly - the exact regression class the owner hit.
    void inferenceE2EPv13();
    void inferenceE2EUserSim8();
    void standardFixtureDoesNotInfer();
    void realS26SampleIdentity();
    // Phase 2c Task 5: the demo workbook - shuffled standard columns + a
    // custom coil_temp column, described by a veryHidden _dve_schema sheet.
    void manifestDemoEndToEnd();
};

namespace {

void setCell(QVector<QVector<QVariant>>& g, int r, int c, const QVariant& v)
{
    if (g.size() <= r) g.resize(r + 1);
    if (g[r].size() <= c) g[r].resize(c + 1);
    g[r][c] = v;
}

// Real ExcelReader grids pad every row out to the sheet-wide max column
// (unrelated content elsewhere on the sheet can push that width well past
// any one block's real content) - simulate that by widening every row with
// trailing empty cells, never shrinking one.
void padRowsTo(QVector<QVector<QVariant>>& g, int width)
{
    for (QVector<QVariant>& row : g)
        if (row.size() < width) row.resize(width);
}

// A minimal standard-shaped (12-col) grid - just enough header/column-row
// text for standardFits' width + first-3-column checks; mirrors
// tst_v3model's makeStandardGrid at reduced scope (this suite doesn't need
// the full header band).
QVector<QVector<QVariant>> makeStandardGridLocal(int blocks, int rows)
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/false);
    QVector<QVector<QVariant>> g(4 + rows);
    for (int b = 0; b < blocks; ++b) {
        const int off = b * s.blockCols;
        for (int c = 0; c < s.columns.size(); ++c)
            setCell(g, 3, off + c, s.columns[c].displayName);
        for (int r = 0; r < rows; ++r) {
            setCell(g, 4 + r, off + 0, (r + 1) * 10);
            setCell(g, 4 + r, off + 1, 25.0 - r * 0.03);
            setCell(g, 4 + r, off + 2, 25.0 - r * 0.03 - 0.035);
        }
    }
    return g;
}

// S26 4D/4E/4F "Cart" layout (13-col block, PV1..PV5 inserted before
// Resistance): puffs, Before weight/g, After weight/g, PV1..PV5, Resistance,
// Smell (0-4), Clog?, Notes, TPM (mg/puff). Header band: row1 exotic
// "Coil Material:" label, row2 "Cart #" + "Ri (Ohms)", row3 "Media".
QVector<QVector<QVariant>> makeS26Grid(int blocks, int rows)
{
    QVector<QVector<QVariant>> g;
    const int blockCols = 13;
    const QStringList headers{
        QStringLiteral("puffs"), QStringLiteral("Before weight/g"), QStringLiteral("After weight/g"),
        QStringLiteral("PV1"), QStringLiteral("PV2"), QStringLiteral("PV3"), QStringLiteral("PV4"), QStringLiteral("PV5"),
        QStringLiteral("Resistance"), QStringLiteral("Smell (0-4)"), QStringLiteral("Clog?"),
        QStringLiteral("Notes"), QStringLiteral("TPM (mg/puff)")
    };
    for (int b = 0; b < blocks; ++b) {
        const int off = b * blockCols;
        setCell(g, 0, off + 0, QStringLiteral("Coil Material:"));
        setCell(g, 0, off + 1, QStringLiteral("Kanthal-%1").arg(b));
        setCell(g, 1, off + 0, QStringLiteral("Cart #"));
        setCell(g, 1, off + 1, QStringLiteral("S26-Cart-%1").arg(b));
        setCell(g, 1, off + 2, QStringLiteral("Ri (Ohms)"));
        setCell(g, 1, off + 3, 1.10 + b * 0.05);
        setCell(g, 2, off + 0, QStringLiteral("Media"));
        setCell(g, 2, off + 1, QStringLiteral("PG70/VG30-%1").arg(b));
        for (int c = 0; c < headers.size(); ++c)
            setCell(g, 3, off + c, headers[c]);
        for (int r = 0; r < rows; ++r) {
            const int row = 4 + r;
            setCell(g, row, off + 0, (r + 1) * 10);                    // puffs
            setCell(g, row, off + 1, 25.0 - r * 0.03);                 // Before weight/g
            setCell(g, row, off + 2, 25.0 - r * 0.03 - 0.035);         // After weight/g
            setCell(g, row, off + 3, 1.00 + r * 0.01);                 // PV1
            setCell(g, row, off + 4, 1.05 + r * 0.01);                 // PV2
            setCell(g, row, off + 5, 1.10 + r * 0.01);                 // PV3
            setCell(g, row, off + 6, 1.15 + r * 0.01);                 // PV4
            setCell(g, row, off + 7, 1.20 + r * 0.01);                 // PV5
            setCell(g, row, off + 8, 1.15 + b * 0.05);                 // Resistance (per-row)
            setCell(g, row, off + 9, QStringLiteral("1"));             // Smell (0-4)
            setCell(g, row, off + 10, QStringLiteral("N"));            // Clog?
            setCell(g, row, off + 11, QStringLiteral(""));             // Notes
            setCell(g, row, off + 12, 3.50 - r * 0.01);                // TPM (mg/puff)
        }
    }
    return g;
}

// CPS2920 "User Test Simulation" layout (8-col block, leading Chronology
// column): Chronology, puffs, Before Weight/g, After Weight/g,
// Draw Pressure (kpa), Failure? (if yes, when), Notes, TPM (mg/puff).
// Header band: row1 "Resistance:"/"Sample ID:"/"Initial Oil Mass:", row2
// "Media:"/"Tester:"/"Voltage:"/"Power:".
QVector<QVector<QVariant>> makeUserSimGrid(int blocks, int rows)
{
    QVector<QVector<QVariant>> g;
    const int blockCols = 8;
    const QStringList headers{
        QStringLiteral("Chronology"), QStringLiteral("puffs"),
        QStringLiteral("Before Weight/g"), QStringLiteral("After Weight/g"),
        QStringLiteral("Draw Pressure (kpa)"), QStringLiteral("Failure? (if yes, when)"),
        QStringLiteral("Notes"), QStringLiteral("TPM (mg/puff)")
    };
    for (int b = 0; b < blocks; ++b) {
        const int off = b * blockCols;
        setCell(g, 0, off + 0, QStringLiteral("Resistance:"));
        setCell(g, 0, off + 1, 1.50 + b * 0.1);
        setCell(g, 0, off + 2, QStringLiteral("Sample ID:"));
        setCell(g, 0, off + 3, QStringLiteral("US-%1").arg(b));
        setCell(g, 0, off + 4, QStringLiteral("Initial Oil Mass:"));
        setCell(g, 0, off + 5, 5.20 + b * 0.1);
        setCell(g, 1, off + 0, QStringLiteral("Media:"));
        setCell(g, 1, off + 1, QStringLiteral("PG"));
        setCell(g, 1, off + 2, QStringLiteral("Tester:"));
        setCell(g, 1, off + 3, QStringLiteral("Alice"));
        setCell(g, 1, off + 4, QStringLiteral("Voltage:"));
        setCell(g, 1, off + 5, 3.70);
        setCell(g, 1, off + 6, QStringLiteral("Power:"));
        setCell(g, 1, off + 7, 11.40 + b * 0.2);
        for (int c = 0; c < headers.size(); ++c)
            setCell(g, 3, off + c, headers[c]);
        for (int r = 0; r < rows; ++r) {
            const int row = 4 + r;
            setCell(g, row, off + 0, QStringLiteral("Day %1").arg(r + 1));   // Chronology
            setCell(g, row, off + 1, (r + 1) * 10);                          // puffs
            setCell(g, row, off + 2, 25.0 - r * 0.03);                       // Before Weight/g
            setCell(g, row, off + 3, 25.0 - r * 0.03 - 0.035);               // After Weight/g
            setCell(g, row, off + 4, 12.5 + r * 0.1);                        // Draw Pressure (kpa)
            setCell(g, row, off + 5, QStringLiteral("No"));                  // Failure? (if yes, when)
            setCell(g, row, off + 6, QStringLiteral(""));                    // Notes
            setCell(g, row, off + 7, 3.20 - r * 0.01);                       // TPM (mg/puff)
        }
    }
    return g;
}

// Shuffled 3-col single-block grid for the NameFirst slot tests: PHYSICAL
// column order is [TPM (mg/puff), puffs, Notes] while the matching schema
// (makeShuffled3ColSchema) orders its columns [puffs, notes, tpm] - name
// matching must resolve block-relative slots {1, 2, 0}.
QVector<QVector<QVariant>> makeShuffled3ColGrid()
{
    QVector<QVector<QVariant>> g;
    setCell(g, 3, 0, QStringLiteral("TPM (mg/puff)"));
    setCell(g, 3, 1, QStringLiteral("puffs"));
    setCell(g, 3, 2, QStringLiteral("Notes"));
    for (int r = 0; r < 3; ++r) {
        setCell(g, 4 + r, 0, 999.0);                                 // bogus TPM - derived, recomputed
        setCell(g, 4 + r, 1, (r + 1) * 10);                          // puffs
        setCell(g, 4 + r, 2, QStringLiteral("note-%1").arg(r + 1));  // Notes
    }
    return g;
}

// The matching 3-col schema, columns in SCHEMA order [puffs, notes, tpm]
// (deliberately NOT the grid's physical order), built from registry copies -
// exactly how a manifest block inherits its known keys.
TemplateSchema makeShuffled3ColSchema()
{
    TemplateSchema s;
    s.schemaId   = QStringLiteral("shuffle3");
    s.version    = 1;
    s.headerRows = 3; s.columnHeaderRow = 4; s.dataStartRow = 5; s.blockCols = 3;
    s.columns = { *MetricRegistry::metric(QStringLiteral("puffs")),
                  *MetricRegistry::metric(QStringLiteral("notes")),
                  *MetricRegistry::metric(QStringLiteral("tpm")) };
    return s;
}

} // namespace

void TestV3Inference::inferBlockColsS26()
{
    const auto g = makeS26Grid(2, 3);
    QCOMPARE(SchemaInference::inferBlockCols(g[3]), 13);
}

void TestV3Inference::inferBlockColsUserSim()
{
    const auto g = makeUserSimGrid(2, 3);
    QCOMPARE(SchemaInference::inferBlockCols(g[3]), 8);
}

void TestV3Inference::standardFitsTrueForStandardGrid()
{
    const TemplateSchema std0 = standardV1(/*perRowRegime=*/false);
    QVERIFY(SchemaInference::standardFits(makeStandardGridLocal(2, 5), std0));
    QVERIFY(SchemaInference::standardFits(makeStandardGridLocal(1, 3), std0));
}

void TestV3Inference::standardFitsTrueForEmptyHeaderRow()
{
    // Legacy positional-only file: row 4 (the column-header row) is blank.
    // An all-empty header row can't be inferred from at all, so it must be
    // treated as fitting the standard (existing byte-identical path).
    const TemplateSchema std0 = standardV1(/*perRowRegime=*/false);
    QVector<QVector<QVariant>> g(5);
    g[3].resize(12);
    QVERIFY(SchemaInference::standardFits(g, std0));
}

void TestV3Inference::standardFitsFalseForS26()
{
    const TemplateSchema std0 = standardV1(/*perRowRegime=*/false);
    QVERIFY(!SchemaInference::standardFits(makeS26Grid(2, 3), std0));
}

void TestV3Inference::standardFitsFalseForUserSim()
{
    const TemplateSchema std0 = standardV1(/*perRowRegime=*/false);
    QVERIFY(!SchemaInference::standardFits(makeUserSimGrid(2, 3), std0));
}

void TestV3Inference::standardFitsTrueForPaddedSingleBlockGrid()
{
    // A genuinely standard single-block sheet whose row 4 has 12 real
    // headers followed by trailing blanks (padded by unrelated content
    // elsewhere on the sheet, exactly how ExcelReader grids come back) must
    // still be routed to the existing byte-identical path - raw
    // QVector::size() over-counts the padding and must NOT be used to
    // measure header width.
    const TemplateSchema std0 = standardV1(/*perRowRegime=*/false);
    auto g = makeStandardGridLocal(1, 3);
    padRowsTo(g, 40);   // 12 real cols + 28 trailing blanks
    QVERIFY(SchemaInference::standardFits(g, std0));
}

void TestV3Inference::standardFitsFalseForPaddedS26()
{
    // Negative twin: a genuinely non-standard (13-col) sheet padded the
    // same way must still be correctly rejected.
    const TemplateSchema std0 = standardV1(/*perRowRegime=*/false);
    auto g = makeS26Grid(1, 3);
    padRowsTo(g, 41);   // 13 real cols + 28 trailing blanks
    QVERIFY(!SchemaInference::standardFits(g, std0));
}

void TestV3Inference::inferSchemaS26Columns()
{
    const auto g = makeS26Grid(2, 3);
    const TemplateSchema s = SchemaInference::inferSchema(g, QStringLiteral("S26 4D, 4E, 4F Designs"));
    QCOMPARE(s.columns.size(), 13);
    QCOMPARE(s.blockCols, 13);

    const MetricDef* puffs = s.column(QStringLiteral("puffs"));
    QVERIFY(puffs);
    QVERIFY(puffs->role == Role::Measured);
    QVERIFY(puffs->type == ValueType::Number);

    const MetricDef* before = s.column(QStringLiteral("before_weight"));
    QVERIFY(before);
    QVERIFY(before->role == Role::Measured);

    const MetricDef* after = s.column(QStringLiteral("after_weight"));
    QVERIFY(after);
    QVERIFY(after->role == Role::Measured);

    const QStringList pvKeys{QStringLiteral("pv1"), QStringLiteral("pv2"), QStringLiteral("pv3"),
                             QStringLiteral("pv4"), QStringLiteral("pv5")};
    for (const QString& key : pvKeys) {
        const MetricDef* pv = s.column(key);
        QVERIFY2(pv, qPrintable(key));
        QVERIFY2(pv->type == ValueType::Number, qPrintable(key));
        QVERIFY2(pv->role == Role::Measured, qPrintable(key));
    }

    const MetricDef* tpm = s.column(QStringLiteral("tpm"));
    QVERIFY(tpm);
    QVERIFY(tpm->role == Role::Derived);
    QCOMPARE(tpm->calculator, QStringLiteral("tpm_v1"));
}

void TestV3Inference::inferSchemaS26HeaderFields()
{
    const auto g = makeS26Grid(2, 3);
    const TemplateSchema s = SchemaInference::inferSchema(g, QStringLiteral("S26 4D, 4E, 4F Designs"));

    const HeaderFieldDef* sampleId = s.headerField(QStringLiteral("sample_id"));
    QVERIFY(sampleId);   // from "Cart #"

    const HeaderFieldDef* ri = s.headerField(QStringLiteral("resistance_initial"));
    QVERIFY(ri);   // from "Ri (Ohms)" - ratified key resistance_initial (D11)
    QVERIFY(ri->type == ValueType::Number);

    const HeaderFieldDef* media = s.headerField(QStringLiteral("media"));
    QVERIFY(media);

    const HeaderFieldDef* coil = s.headerField(QStringLiteral("coil_material"));
    QVERIFY(coil);   // registry-known label (D6); displayName stays the sheet's raw text
    QCOMPARE(coil->displayName, QStringLiteral("Coil Material:"));
}

void TestV3Inference::inferSchemaUserSimColumns()
{
    const auto g = makeUserSimGrid(2, 3);
    const TemplateSchema s = SchemaInference::inferSchema(g, QStringLiteral("User Test Simulation"));
    QCOMPARE(s.columns.size(), 8);
    QCOMPARE(s.blockCols, 8);

    const MetricDef* chron = s.column(QStringLiteral("chronology"));
    QVERIFY(chron);
    QVERIFY(chron->type == ValueType::Text);

    const MetricDef* puffs = s.column(QStringLiteral("puffs"));
    QVERIFY(puffs);
    QVERIFY(puffs->role == Role::Measured);

    // "Failure? (if yes, when)" -> registry metric "failure" (ratified alias).
    const MetricDef* failure = s.column(QStringLiteral("failure"));
    QVERIFY(failure);
    QVERIFY(failure->type == ValueType::Text);

    const MetricDef* tpm = s.column(QStringLiteral("tpm"));
    QVERIFY(tpm);
    QVERIFY(tpm->role == Role::Derived);
}

void TestV3Inference::inferSchemaUserSimHeaderFields()
{
    const auto g = makeUserSimGrid(2, 3);
    const TemplateSchema s = SchemaInference::inferSchema(g, QStringLiteral("User Test Simulation"));

    const HeaderFieldDef* sampleId = s.headerField(QStringLiteral("sample_id"));
    QVERIFY(sampleId);
    QCOMPARE(sampleId->row, 1);

    const HeaderFieldDef* oilMass = s.headerField(QStringLiteral("initial_oil_mass"));
    QVERIFY(oilMass);
    QCOMPARE(oilMass->row, 1);
    QVERIFY(oilMass->type == ValueType::Number);

    const HeaderFieldDef* power = s.headerField(QStringLiteral("power"));
    QVERIFY(power);   // from "Power:" on row 2
    QCOMPARE(power->row, 2);
    QVERIFY(power->type == ValueType::Number);
}

void TestV3Inference::registryLabelsRecognized()
{
    // Cart-era design-spec labels have NO trailing ':' and were previously
    // skipped as plain text; the registry now knows them (D6). Reuse the
    // S26-style grid, but strip the colon off its "Coil Material:" label and
    // add a bare-position "Did this burn?" Q&A pair to the header band.
    auto g = makeS26Grid(1, 3);
    setCell(g, 0, 0, QStringLiteral("Coil Material"));   // bare label; value stays at c1
    setCell(g, 0, 2, QStringLiteral("Did this burn?"));
    setCell(g, 0, 3, QStringLiteral("no"));
    const TemplateSchema s = SchemaInference::inferSchema(g, QStringLiteral("t"));
    const HeaderFieldDef* cm = s.headerField(QStringLiteral("coil_material"));
    QVERIFY(cm);
    QCOMPARE(cm->displayName, QStringLiteral("Coil Material"));
    // "Did this burn?" previously landed under snake key did_this_burn; the
    // registry canonicalizes it (D8).
    QVERIFY(s.headerField(QStringLiteral("did_burn")));
    QVERIFY(!s.headerField(QStringLiteral("did_this_burn")));
}

void TestV3Inference::registryRfLabelCanonicalized()
{
    // "Rf (Ohms)" used to map to rf_ohms; ratified key is resistance_final (D11).
    auto g = makeS26Grid(1, 3);
    setCell(g, 2, 2, QStringLiteral("Rf (Ohms)"));
    setCell(g, 2, 3, 1.31);
    const TemplateSchema s = SchemaInference::inferSchema(g, QStringLiteral("t"));
    QVERIFY(s.headerField(QStringLiteral("resistance_final")));
    QVERIFY(!s.headerField(QStringLiteral("rf_ohms")));
}

void TestV3Inference::resistanceInitialLowersToSampleResistance()
{
    // The registry re-keys "Ri (Ohms)" to resistance_initial (D11), but the
    // legacy SampleResult::resistance member must still be populated from it:
    // Cart-era sheets carry no plain "Resistance:" label, and the initial
    // reading IS the sample resistance in the legacy model.
    const auto g = makeS26Grid(2, 3);
    const TemplateSchema schema = SchemaInference::inferSchema(g, QStringLiteral("t"));
    const Sheet sheet = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), schema, /*perRowRegime=*/false,
        ColumnResolution::NameFirst);
    const DVE::SheetResult sr = LegacyAdapter::lowerInferredSheet(
        sheet, QStringLiteral("t"), QStringLiteral("new"));
    QCOMPARE(sr.samples.size(), 2);
    QCOMPARE(sr.samples[0].resistance, 1.10);   // header "Ri (Ohms)", NOT the per-row column (1.15)
    QCOMPARE(sr.samples[1].resistance, 1.15);   // block-2 header value
}

void TestV3Inference::pvColumnsAssemblePerPuffList()
{
    // 13-col Cart-era grid: PV1..PV5 at cols 3-7. The lowering must assemble
    // the five per-row values into ONE draw_pressure_per_puff QVariantList
    // (in PV order) instead of five individual pvN extra keys (registry D5/8.2).
    auto g = makeS26Grid(1, 3);
    for (int i = 0; i < 5; ++i)
        setCell(g, 4, 3 + i, 10.0 + i);   // first data row: PV1..PV5 = 10..14
    const TemplateSchema schema = SchemaInference::inferSchema(g, QStringLiteral("t"));
    const Sheet sheet = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), schema, /*perRowRegime=*/false,
        ColumnResolution::NameFirst);
    const DVE::SheetResult sr = LegacyAdapter::lowerInferredSheet(
        sheet, QStringLiteral("t"), QStringLiteral("new"));
    QCOMPARE(sr.samples.size(), 1);
    QVERIFY(!sr.samples[0].rows.isEmpty());
    const DVE::DataRow& row = sr.samples[0].rows[0];
    QVERIFY(!row.extra.contains(QStringLiteral("pv1")));   // individual pvN keys replaced
    const QVariant v = row.extra.value(QStringLiteral("draw_pressure_per_puff"));
    QCOMPARE(v.typeId(), static_cast<int>(QMetaType::QVariantList));
    const QVariantList l = v.toList();
    QCOMPARE(l.size(), 5);
    QCOMPARE(l[0].toDouble(), 10.0);
    QCOMPARE(l[4].toDouble(), 14.0);
}

void TestV3Inference::writeProvenanceRecordedOnInferredParse()
{
    // Write provenance (Phase 2b): the inferred lowering must record the
    // sheet's own geometry (13-wide S26 shape) + per-sample block origins so
    // write-back can derive addresses on these layouts too.
    const auto g = makeS26Grid(2, 3);
    const TemplateSchema schema = SchemaInference::inferSchema(g, QStringLiteral("t"));
    const Sheet sheet = SchemaDrivenReader::parseSheet(g, QStringLiteral("t"), schema,
                                                       /*perRowRegime=*/false,
                                                       ColumnResolution::NameFirst);
    const DVE::SheetResult sr = LegacyAdapter::lowerInferredSheet(
        sheet, QStringLiteral("t"), QStringLiteral("new"));
    QCOMPARE(sr.blockCols, 13);
    QCOMPARE(sr.dataStartRow, 5);
    QCOMPARE(sr.columnKeys.size(), 13);
    QVERIFY(sr.columnKeys.contains(QStringLiteral("smell")));
    QVERIFY(sr.headerCells.contains(QStringLiteral("media")));
    QCOMPARE(sr.samples.size(), 2);
    for (int i = 0; i < sr.samples.size(); ++i)
        QCOMPARE(sr.samples[i].startColumn, i * 13);
}

void TestV3Inference::strippedCountGatesLoweringPuffExtrapolation()
{
    // S26 grid with the puffs of data rows 2-3 blanked - the shape a
    // cache-stripped workbook presents (formula cells read as None).
    auto g = makeS26Grid(1, 3);
    setCell(g, 5, 0, QVariant());
    setCell(g, 6, 0, QVariant());
    const TemplateSchema schema = SchemaInference::inferSchema(g, QStringLiteral("t"));
    const Sheet sheet = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), schema, /*perRowRegime=*/false,
        ColumnResolution::NameFirst);

    // Control: the legacy repair fills the gaps on a healthy sheet (count 0).
    const DVE::SheetResult healthy = LegacyAdapter::lowerInferredSheet(
        sheet, QStringLiteral("t"), QStringLiteral("new"));
    QCOMPARE(healthy.strippedFormulaCells, 0);
    QCOMPARE(healthy.samples[0].rows[1].puffs, 20.0);
    QCOMPARE(healthy.samples[0].rows[2].puffs, 30.0);

    // Flagged: the zeros are destroyed data - no fabrication, count rides out.
    const DVE::SheetResult flagged = LegacyAdapter::lowerInferredSheet(
        sheet, QStringLiteral("t"), QStringLiteral("new"), /*strippedFormulaCells=*/7);
    QCOMPARE(flagged.strippedFormulaCells, 7);
    QCOMPARE(flagged.samples[0].rows[0].puffs, 10.0);   // literal seed survives
    QCOMPARE(flagged.samples[0].rows[1].puffs, 0.0);
    QCOMPARE(flagged.samples[0].rows[2].puffs, 0.0);
}

void TestV3Inference::parseSheetExposesResolvedSlots()
{
    const auto g = makeShuffled3ColGrid();
    const TemplateSchema s = makeShuffled3ColSchema();

    // NameFirst: schema column i lives at block-relative physical slot
    // columnSlots[i] - puffs at 1, notes at 2, tpm at 0.
    const Sheet byName = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), s, /*perRowRegime=*/false,
        ColumnResolution::NameFirst);
    QCOMPARE(byName.columnSlots, (QVector<int>{1, 2, 0}));

    // Positional: the identity map (the byte-identity guarantee of the slot
    // exposure on the legacy standard path).
    const Sheet positional = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), s, /*perRowRegime=*/false,
        ColumnResolution::Positional);
    QCOMPARE(positional.columnSlots, (QVector<int>{0, 1, 2}));
}

void TestV3Inference::manifestSheetLowersByKeyWithProvenanceInSlotOrder()
{
    const auto g = makeShuffled3ColGrid();
    const TemplateSchema s = makeShuffled3ColSchema();
    const Sheet sheet = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), s, /*perRowRegime=*/false,
        ColumnResolution::NameFirst);
    const DVE::SheetResult sr = LegacyAdapter::lowerSchemaSheet(
        sheet, QStringLiteral("t"), QStringLiteral("new"),
        /*fromInference=*/false, /*perRowRegime=*/false);

    // Manifest sheets are declared, not inferred.
    QVERIFY(!sr.fromInferredSchema);
    QVERIFY(!sr.hasPerRowRegime);

    // Values land by KEY despite the physical shuffle.
    QCOMPARE(sr.samples.size(), 1);
    const DVE::SampleResult& s0 = sr.samples[0];
    QCOMPARE(s0.rows.size(), 3);
    QCOMPARE(s0.rows[0].puffs, 10.0);
    QCOMPARE(s0.rows[2].puffs, 30.0);
    QCOMPARE(s0.rows[0].notes, QStringLiteral("note-1"));
    QCOMPARE(s0.rows[2].notes, QStringLiteral("note-3"));
    // tpm is Derived: the sheet's bogus 999 is dropped and recomputed (zero
    // weights on this minimal grid -> clean 0.0, never 999).
    QCOMPARE(s0.rows[0].tpm, 0.0);

    // Provenance columnKeys in PHYSICAL slot order, not schema order.
    QCOMPARE(sr.columnKeys,
             (QStringList{QStringLiteral("tpm"), QStringLiteral("puffs"),
                          QStringLiteral("notes")}));

    // CellAddress derives the SHUFFLED physical column: puffs is slot 1.
    const DVE::CellAddress addr =
        DVE::CellAddress::dataCell(sr, s0, QStringLiteral("puffs"), 0);
    QVERIFY(addr.valid);
    QCOMPARE(addr.row, 5);                        // dataStartRow + data row 0
    QCOMPARE(addr.col, s0.startColumn + 1 + 1);   // slot 1 -> 1-based col 2
}

void TestV3Inference::phantomTrailingSamplesDroppedOnSchemaPaths()
{
    // One REAL 13-wide S26 block + 13 empty padded columns (rows padded to
    // the sheet-wide max width, exactly how ExcelReader grids come back):
    // hdrWidth / blockCols over-counts the blocks, producing a trailing
    // sample whose every header cell and every data value is empty. The
    // schema paths (NameFirst: inference + manifest) must drop it.
    auto g = makeS26Grid(1, 3);
    padRowsTo(g, 26);
    const TemplateSchema schema = SchemaInference::inferSchema(g, QStringLiteral("t"));
    QCOMPARE(schema.blockCols, 13);

    const Sheet byName = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), schema, /*perRowRegime=*/false,
        ColumnResolution::NameFirst);
    QCOMPARE(byName.samples.size(), 1);       // phantom dropped

    // Positional keeps the legacy phantom behavior untouched (byte-identity;
    // the shadow suite referees the standard path).
    const Sheet positional = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("t"), schema, /*perRowRegime=*/false,
        ColumnResolution::Positional);
    QCOMPARE(positional.samples.size(), 2);   // legacy phantom preserved
}

void TestV3Inference::manifestResolutionRoutesShuffledGrid()
{
    // The schema serialized to a _dve_schema grid and parsed back - exactly
    // what DataProcessor::processFile hands the resolver for a manifest
    // workbook (registry inheritance restores the tpm alias that name-matches
    // the shuffled "TPM (mg/puff)" header).
    const TemplateSchema s = makeShuffled3ColSchema();
    const Manifest::ParseResult pr =
        Manifest::parse(Manifest::gridFor(s, {QStringLiteral("*")}));
    QVERIFY(pr.warnings.isEmpty());

    const auto g = makeShuffled3ColGrid();
    const SchemaResolver::Resolution res = SchemaResolver::resolve(
        &pr, QStringLiteral("Custom Coil Test"), g, QStringLiteral("new"));
    QVERIFY(res.source == SchemaResolver::Source::Manifest);
    QVERIFY(res.columnResolution == ColumnResolution::NameFirst);
    QVERIFY(!res.perRowRegime);   // no puffing_regime column declared

    // The Task 4 processSheet seam: NameFirst parse with the manifest schema,
    // then the generalized lowering with fromInference=false.
    const Sheet sheet = SchemaDrivenReader::parseSheet(
        g, QStringLiteral("Custom Coil Test"), res.schema, res.perRowRegime,
        res.columnResolution);
    const DVE::SheetResult sr = LegacyAdapter::lowerSchemaSheet(
        sheet, QStringLiteral("Custom Coil Test"), QStringLiteral("new"),
        /*fromInference=*/false, res.perRowRegime);

    QVERIFY(!sr.fromInferredSchema);
    QCOMPARE(sr.samples.size(), 1);
    QCOMPARE(sr.samples[0].rows.size(), 3);
    QCOMPARE(sr.samples[0].rows[1].puffs, 20.0);   // physical col 1, by name
    QCOMPARE(sr.columnKeys,
             (QStringList{QStringLiteral("tpm"), QStringLiteral("puffs"),
                          QStringLiteral("notes")}));
}

// ── E2E helpers ────────────────────────────────────────────────────────────
namespace {

// The first sheet that carries sample data (fixtures put data on sheet 0, but
// be defensive so an extra summary sheet wouldn't break the test).
const DVE::SheetResult* firstSheetWithSamples(const DVE::FileResult& f)
{
    for (const DVE::SheetResult& sh : f.sheets)
        if (sh.hasSamples()) return &sh;
    return nullptr;
}

} // namespace

void TestV3Inference::inferenceE2EPv13()
{
    DVE::DataProcessor dp;
    const DVE::FileResult f = dp.processFile(testDataFile(QStringLiteral("pv13.xlsx")));
    if (f.filePath.isEmpty())
        QSKIP("pipeline unavailable (bundled/system python + openpyxl not found)");

    QVERIFY2(dp.lastFileUsedInference(), "13-col pv13 must take the inference path");

    const DVE::SheetResult* sh = firstSheetWithSamples(f);
    QVERIFY(sh);
    QCOMPARE(sh->samples.size(), 2);
    QVERIFY(sh->fromInferredSchema);

    // Sample 1 (block 1): tpm[0] recomputed from the fixture weights (the bogus
    // 999 in the sheet's TPM column must be overwritten); PV1..PV5 -> the
    // assembled draw_pressure_per_puff row-extra list (registry D5/8.2);
    // the exotic "Coil Material:" label -> sample extra.
    const DVE::SampleResult& s0 = sh->samples[0];
    QVERIFY(!s0.rows.isEmpty());
    QVERIFY2(fuzzyEqual(s0.rows[0].tpm, 3.5, 1e-4),
             qPrintable(QStringLiteral("tpm[0]=%1, expected 3.5").arg(s0.rows[0].tpm)));
    QVERIFY2(!s0.rows[0].extra.contains(QStringLiteral("pv1")),
             "individual pvN keys must be replaced by the assembled list");
    const QVariant ppList = s0.rows[0].extra.value(QStringLiteral("draw_pressure_per_puff"));
    QVERIFY2(ppList.typeId() == QMetaType::QVariantList,
             "row 0 extra should carry the assembled draw_pressure_per_puff list");
    QCOMPARE(ppList.toList().size(), 5);
    QVERIFY2(fuzzyEqual(ppList.toList()[0].toDouble(), 1.00, 1e-6),
             qPrintable(QStringLiteral("pp[0]=%1, expected 1.00 (fixture PV1)")
                        .arg(ppList.toList()[0].toDouble())));
    QCOMPARE(s0.extra.value(QStringLiteral("coil_material")).toString(),
             QStringLiteral("Kanthal-A1"));

    // Sample 2 (block 2): sampleID / resistance / media MUST be the block-2
    // values (the owner's exact regression - previously read a column short).
    const DVE::SampleResult& s1 = sh->samples[1];
    QCOMPARE(s1.sampleID, QStringLiteral("S26 4E-1"));
    QVERIFY2(fuzzyEqual(s1.resistance, 1.84, 1e-6),
             qPrintable(QStringLiteral("resistance=%1, expected 1.84").arg(s1.resistance)));
    QCOMPARE(s1.media, QStringLiteral("10kcp D8"));
}

void TestV3Inference::inferenceE2EUserSim8()
{
    DVE::DataProcessor dp;
    const DVE::FileResult f = dp.processFile(testDataFile(QStringLiteral("usersim8.xlsx")));
    if (f.filePath.isEmpty())
        QSKIP("pipeline unavailable (bundled/system python + openpyxl not found)");

    QVERIFY2(dp.lastFileUsedInference(), "8-col usersim8 must take the inference path");

    const DVE::SheetResult* sh = firstSheetWithSamples(f);
    QVERIFY(sh);
    QCOMPARE(sh->samples.size(), 2);

    const DVE::SampleResult& s0 = sh->samples[0];
    QVERIFY(!s0.rows.isEmpty());
    QVERIFY2(fuzzyEqual(s0.rows[0].tpm, 3.5, 1e-4),
             qPrintable(QStringLiteral("tpm[0]=%1, expected 3.5").arg(s0.rows[0].tpm)));
    // Leading Chronology column -> open per-row metric.
    QVERIFY2(s0.rows[0].extra.contains(QStringLiteral("chronology")),
             "row 0 extra should carry chronology");
    QCOMPARE(s0.rows[0].extra.value(QStringLiteral("chronology")).toString(),
             QStringLiteral("Day 1"));
    // Exotic "Firmware:" header label -> sample extra.
    QCOMPARE(s0.extra.value(QStringLiteral("firmware")).toString(), QStringLiteral("v2.1"));
    // A present numeric Power header value wins over deriving.
    QVERIFY2(fuzzyEqual(s0.power, 11.40, 1e-6),
             qPrintable(QStringLiteral("power=%1, expected 11.40 (present value)").arg(s0.power)));

    const DVE::SampleResult& s1 = sh->samples[1];
    QCOMPARE(s1.sampleID, QStringLiteral("US-B"));
    QVERIFY2(fuzzyEqual(s1.resistance, 1.60, 1e-6),
             qPrintable(QStringLiteral("resistance=%1, expected 1.60").arg(s1.resistance)));
    QCOMPARE(s1.media, QStringLiteral("PG"));
}

void TestV3Inference::standardFixtureDoesNotInfer()
{
    // A genuinely standard 12-wide fixture must keep the byte-identical
    // positional path - lastFileUsedInference() stays false.
    DVE::DataProcessor dp;
    const DVE::FileResult f = dp.processFile(testDataFile(QStringLiteral("format_e.xlsx")));
    if (f.filePath.isEmpty())
        QSKIP("pipeline unavailable (bundled/system python + openpyxl not found)");
    QVERIFY(!f.sheets.isEmpty());
    QVERIFY2(!dp.lastFileUsedInference(),
             "standard 12-wide format_e must NOT take the inference path");
}

void TestV3Inference::realS26SampleIdentity()
{
    // Acceptance evidence on the owner's actual complaint file (corpus-only;
    // gitignored). Skips cleanly when DVE_TEST_CORPUS_DIR is unset / the file
    // is absent, per CorpusUtils conventions.
    QString s26;
    const QStringList corpus = DVE::testutil::corpusFiles();
    for (const QString& f : corpus) {
        if (QFileInfo(f).fileName().startsWith(QStringLiteral("S26 4D"))) { s26 = f; break; }
    }
    if (s26.isEmpty())
        QSKIP("real S26 corpus file absent (set DVE_TEST_CORPUS_DIR)");

    DVE::DataProcessor dp;
    const DVE::FileResult f = dp.processFile(s26);
    if (f.filePath.isEmpty())
        QSKIP("pipeline unavailable (bundled/system python + openpyxl not found)");

    QVERIFY2(dp.lastFileUsedInference(), "real S26 is 13-wide - must take the inference path");
    const DVE::SheetResult* sh = firstSheetWithSamples(f);
    QVERIFY(sh);
    QVERIFY2(sh->samples.size() >= 3,
             qPrintable(QStringLiteral("expected >=3 samples, got %1").arg(sh->samples.size())));

    // The exact values from the plan's Trigger probe (block-2 / block-3 reads
    // that the 12-wide parsers mangled).
    QCOMPARE(sh->samples[1].sampleID, QStringLiteral("S26 4E-1"));
    QCOMPARE(sh->samples[1].media,    QStringLiteral("10kcp D8"));
    QCOMPARE(sh->samples[2].sampleID, QStringLiteral("S26 4F-1"));
}

void TestV3Inference::manifestDemoEndToEnd()
{
    // The Phase 2c acceptance demo: a workbook whose data sheet reorders the
    // standard columns AND adds a custom "Coil Temp (C)" column, made
    // self-describing by a veryHidden _dve_schema manifest sheet. Production
    // must parse it NameFirst via the declared schema - reordering + adding
    // columns Just Works.
    const QString path = testDataFile(QStringLiteral("manifest_demo.xlsx"));
    QVERIFY2(QFileInfo::exists(path),
             "manifest_demo.xlsx fixture missing - run tests/generate_fixtures.py");

    DVE::DataProcessor dp;
    const DVE::FileResult f = dp.processFile(path);
    if (f.filePath.isEmpty())
        QSKIP("pipeline unavailable (bundled/system python + openpyxl not found)");

    // The manifest sheet is an internal artifact: excluded from display.
    QVERIFY2(!f.sheetNames.contains(QStringLiteral("_dve_schema")),
             "_dve_schema must never appear in sheetNames");
    QCOMPARE(f.sheets.size(), 1);
    const DVE::SheetResult& sh = f.sheets[0];
    QCOMPARE(sh.sheetName, QStringLiteral("Custom Coil Test"));

    // Declared, not inferred; the manifest lists a puffing_regime column.
    QVERIFY(!sh.fromInferredSchema);
    QVERIFY(sh.hasPerRowRegime);
    QCOMPARE(sh.samples.size(), 2);

    // Sample 1: every series lands by NAME from its shuffled physical slot.
    const DVE::SampleResult& s0 = sh.samples[0];
    QCOMPARE(s0.sampleID, QStringLiteral("Coil-1"));
    QCOMPARE(s0.rows.size(), 6);
    for (int r = 0; r < s0.rows.size(); ++r) {
        // puffs from PHYSICAL column 2 (slot 1), not the standard slot 0.
        QVERIFY2(fuzzyEqual(s0.rows[r].puffs, (r + 1) * 10.0, 1e-9),
                 qPrintable(QStringLiteral("row %1 puffs=%2, expected %3")
                            .arg(r).arg(s0.rows[r].puffs).arg((r + 1) * 10)));
        // Custom column -> DataRow::extra under its manifest key.
        QVERIFY2(fuzzyEqual(s0.rows[r].extra.value(QStringLiteral("coil_temp")).toDouble(),
                            200.0 + r, 1e-9),
                 qPrintable(QStringLiteral("row %1 coil_temp=%2, expected %3")
                            .arg(r).arg(s0.rows[r].extra.value(QStringLiteral("coil_temp")).toDouble())
                            .arg(200 + r)));
        QCOMPARE(s0.rows[r].puffingRegime, QStringLiteral("60mL/3s/30s"));
    }
    // tpm recomputed from the weight pairs (Derived recompute unchanged - the
    // sheet's bogus 999 TPM column is overwritten).
    QVERIFY2(fuzzyEqual(s0.rows[0].tpm, 3.5, 1e-4),
             qPrintable(QStringLiteral("tpm[0]=%1, expected 3.5").arg(s0.rows[0].tpm)));
    // Qualitative columns from their shuffled slots.
    QVERIFY(s0.rows[0].smell.isEmpty());
    QVERIFY(s0.rows[0].clog.isEmpty());
    QCOMPARE(s0.rows[3].notes, QStringLiteral("no burn"));

    // Write provenance in PHYSICAL slot order (not schema/standard order).
    QCOMPARE(sh.columnKeys.size(), 10);
    QCOMPARE(sh.columnKeys[0], QStringLiteral("tpm"));
    QCOMPARE(sh.columnKeys[4], QStringLiteral("coil_temp"));
    const DVE::CellAddress addr =
        DVE::CellAddress::dataCell(sh, s0, QStringLiteral("puffs"), 0);
    QVERIFY(addr.valid);
    QCOMPARE(addr.row, 5);                      // dataStartRow + data row 0
    QCOMPARE(addr.col, s0.startColumn + 2);     // physical slot 1 -> 1-based col 2

    // Sample 2: block 2 reads block-2 cells.
    const DVE::SampleResult& s1 = sh.samples[1];
    QCOMPARE(s1.startColumn, 10);
    QCOMPARE(s1.sampleID, QStringLiteral("Coil-2"));
    QVERIFY2(fuzzyEqual(s1.rows[0].extra.value(QStringLiteral("coil_temp")).toDouble(),
                        210.0, 1e-9),
             qPrintable(QStringLiteral("block-2 coil_temp[0]=%1, expected 210")
                        .arg(s1.rows[0].extra.value(QStringLiteral("coil_temp")).toDouble())));
}

QTEST_MAIN(TestV3Inference)
#include "tst_v3inference.moc"
