#include <QtTest/QtTest>

#include <QJsonObject>
#include <QRectF>
#include <QVariant>

#include "pipeline/ReportData.h"
#include "pipeline/ReportDataJson.h"

using namespace DVE;

// Round-trip test for the FileResult <-> JSON serializer (Plan C recovery blob).
// Builds a fully-populated FileResult, serializes it, parses it back, and
// asserts every PERSISTED field at every level survives unchanged. Fields the
// serializer intentionally omits (tpmTrend / puffCounts / images /
// dbDataIncomplete) are asserted to come back empty/default, which documents
// the contract: those are re-derived on restore, never carried in the blob.
class TstReportDataJson : public QObject
{
    Q_OBJECT

private:
    // ---- Builders for a richly-populated tree -----------------------------

    static DataRow makeRow(double base)
    {
        DataRow r;
        r.puffs           = base + 1.0;
        r.beforeWeight    = base + 2.5;
        r.afterWeight     = base + 1.5;
        r.drawPressure    = base + 0.25;
        r.resistance      = base + 0.75;
        r.puffingRegime   = QStringLiteral("CRM %1").arg(base);
        r.smell           = QStringLiteral("smell-%1").arg(base);
        r.clog            = QStringLiteral("N");
        r.notes           = QStringLiteral("note for row %1").arg(base);
        r.tpm             = base + 3.14;
        r.tpmPowerDensity = base + 4.2;
        r.variationTPM    = base + 0.05;
        r.oilConsumed     = base + 6.0;
        r.id              = static_cast<qint64>(900000000000LL + static_cast<qint64>(base));
        r.version         = static_cast<int>(base) + 2;
        return r;
    }

    static SampleResult makeSample()
    {
        SampleResult s;
        s.sampleName          = QStringLiteral("Sample Alpha");
        s.sampleID            = QStringLiteral("SA-001");
        s.date                = QStringLiteral("2026-06-04");
        s.tester              = QStringLiteral("CB");
        s.media               = QStringLiteral("Distillate X");
        s.viscosity           = 1234.5;
        s.resistance          = 1.2;
        s.voltage             = 3.7;
        s.power               = 11.4;
        s.heatingTechnology   = QStringLiteral("Ceramic");
        s.puffingRegime       = QStringLiteral("CRM");
        s.initialOilMass      = 500.5;
        s.averageTPM          = 5.55;
        s.stdDevTPM           = 0.44;
        s.averagePowerDensity = 2.22;
        s.efficiencyPercent   = 88.8;
        s.totalOilConsumed    = 123.4;
        s.totalPuffs          = 250;
        s.normalizedTPM       = 1.11;
        s.burnStatus          = QStringLiteral("N");
        s.clogStatus          = QStringLiteral("Y");
        s.leakStatus          = QStringLiteral("N");

        s.rows.append(makeRow(10.0));
        s.rows.append(makeRow(20.0));

        s.extra.insert(QStringLiteral("customMetric"), QVariant(42.0));
        s.extra.insert(QStringLiteral("label"), QVariant(QStringLiteral("hello")));
        s.extra.insert(QStringLiteral("flag"), QVariant(true));

        s.imagePaths   << QStringLiteral("C:/imgs/a.png") << QStringLiteral("C:/imgs/b.png");
        s.imageLayouts << QRectF(0.5, 1.5, 2.5, 3.5) << QRectF(0.0, 0.0, 0.0, 0.0);
        s.imageCrops   << QRectF(0.1, 0.2, 0.3, 0.4) << QRectF(0.0, 0.0, 1.0, 1.0);
        s.imageIds     << static_cast<qint64>(700000000001LL) << static_cast<qint64>(-1);
        s.imageVersions << 5 << 0;

        s.id      = static_cast<qint64>(800000000099LL);
        s.version = 7;
        return s;
    }

    static SheetResult makeNormalSheet()
    {
        SheetResult sh;
        sh.sheetName        = QStringLiteral("Test 1");
        sh.templateVersion  = QStringLiteral("new");
        sh.hasPerRowRegime  = true;
        sh.columnHeaders    << QStringLiteral("Puffs") << QStringLiteral("Before")
                            << QStringLiteral("After");
        sh.overallAvgTPM    = 6.66;
        sh.overallStdDevTPM = 0.33;
        sh.isRawTable       = false;
        sh.samples.append(makeSample());

        // OMITTED fields populated on purpose, to prove they DON'T survive.
        sh.tpmTrend   << 1.0 << 2.0 << 3.0;
        sh.puffCounts << 50.0 << 100.0;
        sh.images.append(QByteArray("\x89PNG-fake-bytes", 15));
        sh.dbDataIncomplete = true;

        sh.id      = static_cast<qint64>(600000000050LL);
        sh.version = 4;
        return sh;
    }

    static SheetResult makeRawSheet()
    {
        SheetResult sh;
        sh.sheetName       = QStringLiteral("Test SOP's");
        sh.templateVersion = QStringLiteral("new");
        sh.isRawTable      = true;
        sh.rawHeaders      << QStringLiteral("Step") << QStringLiteral("Instruction");
        QStringList row1;  row1 << QStringLiteral("1") << QStringLiteral("Do the thing");
        QStringList row2;  row2 << QStringLiteral("2") << QStringLiteral("Do the next thing");
        sh.rawRows << row1 << row2;
        sh.id      = static_cast<qint64>(600000000051LL);
        sh.version = 1;
        return sh;
    }

    static FileResult makeFile()
    {
        FileResult f;
        f.filePath        = QStringLiteral("C:/data/test.xlsx");
        f.fileName        = QStringLiteral("test.xlsx");
        f.templateVersion = QStringLiteral("new");
        f.sheetNames      << QStringLiteral("Test 1") << QStringLiteral("Test SOP's");
        f.sheets.append(makeNormalSheet());
        f.sheets.append(makeRawSheet());
        f.id      = static_cast<qint64>(500000000123LL);
        f.version = 9;
        return f;
    }

    // ---- Level-by-level comparison helpers --------------------------------

    static void compareRow(const DataRow& a, const DataRow& b)
    {
        QCOMPARE(a.puffs,           b.puffs);
        QCOMPARE(a.beforeWeight,    b.beforeWeight);
        QCOMPARE(a.afterWeight,     b.afterWeight);
        QCOMPARE(a.drawPressure,    b.drawPressure);
        QCOMPARE(a.resistance,      b.resistance);
        QCOMPARE(a.puffingRegime,   b.puffingRegime);
        QCOMPARE(a.smell,           b.smell);
        QCOMPARE(a.clog,            b.clog);
        QCOMPARE(a.notes,           b.notes);
        QCOMPARE(a.tpm,             b.tpm);
        QCOMPARE(a.tpmPowerDensity, b.tpmPowerDensity);
        QCOMPARE(a.variationTPM,    b.variationTPM);
        QCOMPARE(a.oilConsumed,     b.oilConsumed);
        QCOMPARE(a.id,              b.id);
        QCOMPARE(a.version,         b.version);
    }

    static void compareSample(const SampleResult& a, const SampleResult& b)
    {
        QCOMPARE(a.sampleName,          b.sampleName);
        QCOMPARE(a.sampleID,            b.sampleID);
        QCOMPARE(a.date,                b.date);
        QCOMPARE(a.tester,              b.tester);
        QCOMPARE(a.media,               b.media);
        QCOMPARE(a.viscosity,           b.viscosity);
        QCOMPARE(a.resistance,          b.resistance);
        QCOMPARE(a.voltage,             b.voltage);
        QCOMPARE(a.power,               b.power);
        QCOMPARE(a.heatingTechnology,   b.heatingTechnology);
        QCOMPARE(a.puffingRegime,       b.puffingRegime);
        QCOMPARE(a.initialOilMass,      b.initialOilMass);
        QCOMPARE(a.averageTPM,          b.averageTPM);
        QCOMPARE(a.stdDevTPM,           b.stdDevTPM);
        QCOMPARE(a.averagePowerDensity, b.averagePowerDensity);
        QCOMPARE(a.efficiencyPercent,   b.efficiencyPercent);
        QCOMPARE(a.totalOilConsumed,    b.totalOilConsumed);
        QCOMPARE(a.totalPuffs,          b.totalPuffs);
        QCOMPARE(a.normalizedTPM,       b.normalizedTPM);
        QCOMPARE(a.burnStatus,          b.burnStatus);
        QCOMPARE(a.clogStatus,          b.clogStatus);
        QCOMPARE(a.leakStatus,          b.leakStatus);

        QCOMPARE(a.rows.size(), b.rows.size());
        for (int i = 0; i < a.rows.size(); ++i)
            compareRow(a.rows[i], b.rows[i]);

        QCOMPARE(a.extra, b.extra);

        QCOMPARE(a.imagePaths,    b.imagePaths);
        QCOMPARE(a.imageLayouts,  b.imageLayouts);
        QCOMPARE(a.imageCrops,    b.imageCrops);
        QCOMPARE(a.imageIds,      b.imageIds);
        QCOMPARE(a.imageVersions, b.imageVersions);

        QCOMPARE(a.id,      b.id);
        QCOMPARE(a.version, b.version);
    }

    static void compareSheet(const SheetResult& a, const SheetResult& b)
    {
        QCOMPARE(a.sheetName,        b.sheetName);
        QCOMPARE(a.templateVersion,  b.templateVersion);
        QCOMPARE(a.hasPerRowRegime,  b.hasPerRowRegime);
        QCOMPARE(a.columnHeaders,    b.columnHeaders);
        QCOMPARE(a.overallAvgTPM,    b.overallAvgTPM);
        QCOMPARE(a.overallStdDevTPM, b.overallStdDevTPM);
        QCOMPARE(a.isRawTable,       b.isRawTable);
        QCOMPARE(a.rawHeaders,       b.rawHeaders);
        QCOMPARE(a.rawRows,          b.rawRows);
        QCOMPARE(a.id,               b.id);
        QCOMPARE(a.version,          b.version);

        QCOMPARE(a.samples.size(), b.samples.size());
        for (int i = 0; i < a.samples.size(); ++i)
            compareSample(a.samples[i], b.samples[i]);
    }

private slots:
    void roundTripPreservesEveryPersistedField()
    {
        const FileResult original = makeFile();

        const QJsonObject json = fileResultToJson(original);
        const FileResult restored = fileResultFromJson(json);

        // ---- File level ----
        QCOMPARE(restored.filePath,        original.filePath);
        QCOMPARE(restored.fileName,        original.fileName);
        QCOMPARE(restored.templateVersion, original.templateVersion);
        QCOMPARE(restored.sheetNames,      original.sheetNames);
        QCOMPARE(restored.id,              original.id);
        QCOMPARE(restored.version,         original.version);

        // ---- Sheets (recurses into samples + rows) ----
        QCOMPARE(restored.sheets.size(), original.sheets.size());
        for (int i = 0; i < original.sheets.size(); ++i)
            compareSheet(original.sheets[i], restored.sheets[i]);
    }

    void omittedFieldsAreNotCarried()
    {
        // The normal sheet was built with non-empty tpmTrend / puffCounts /
        // images and dbDataIncomplete=true. These are re-derived on restore,
        // so the blob must NOT carry them: after round-trip they are
        // empty/default. This locks in the coverage contract.
        const FileResult original = makeFile();
        const FileResult restored = fileResultFromJson(fileResultToJson(original));

        QVERIFY(!original.sheets[0].tpmTrend.isEmpty());      // sanity: builder populated it
        QVERIFY(!original.sheets[0].images.isEmpty());        // sanity
        QVERIFY(original.sheets[0].dbDataIncomplete);         // sanity

        const SheetResult& sh = restored.sheets[0];
        QVERIFY(sh.tpmTrend.isEmpty());
        QVERIFY(sh.puffCounts.isEmpty());
        QVERIFY(sh.images.isEmpty());
        QCOMPARE(sh.dbDataIncomplete, false);
    }
};

QTEST_MAIN(TstReportDataJson)
#include "tst_reportdatajson.moc"
