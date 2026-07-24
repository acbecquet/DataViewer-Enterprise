#include <QtTest>
#include "model/SchemaInference.h"
#include "model/StandardSchema.h"
#include "model/TemplateSchema.h"

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
    void inferSchemaS26Columns();
    void inferSchemaS26HeaderFields();
    void inferSchemaUserSimColumns();
    void inferSchemaUserSimHeaderFields();
};

namespace {

void setCell(QVector<QVector<QVariant>>& g, int r, int c, const QVariant& v)
{
    if (g.size() <= r) g.resize(r + 1);
    if (g[r].size() <= c) g[r].resize(c + 1);
    g[r][c] = v;
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

    const HeaderFieldDef* resistance = s.headerField(QStringLiteral("resistance"));
    QVERIFY(resistance);   // from "Ri (Ohms)"
    QVERIFY(resistance->type == ValueType::Number);

    const HeaderFieldDef* media = s.headerField(QStringLiteral("media"));
    QVERIFY(media);

    const HeaderFieldDef* coil = s.headerField(QStringLiteral("coil_material"));
    QVERIFY(coil);   // exotic label -> new open key, snake_case of "Coil Material:"
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

    // "Failure? (if yes, when)" -> new open key, snake_case.
    const MetricDef* failure = s.column(QStringLiteral("failure_if_yes_when"));
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

QTEST_MAIN(TestV3Inference)
#include "tst_v3inference.moc"
