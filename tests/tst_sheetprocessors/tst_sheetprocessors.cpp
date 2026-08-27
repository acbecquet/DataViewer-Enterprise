// ── DataViewer Enterprise — SheetProcessors Unit Tests ──────────────────────
#include <QtTest>
#include <memory>
#include "TestHelpers.h"
#include "ReportData.h"
#include "SheetProcessors.h"
#include "ExcelReader.h"

class tst_SheetProcessors : public QObject
{
    Q_OBJECT

private:
    // Build a minimal ExcelReader::SampleData with the given number of rows.
    // Each row gets sequential puffs (10, 20, 30, ...) and weights that yield
    // a known TPM value.
    ExcelReader::SampleData makeSampleData(int rowCount,
                                           const QString& sampleID = "Test-1",
                                           double beforeWeight = 25.1,
                                           double oilConsumedPerRow = 0.035)
    {
        ExcelReader::SampleData sd;
        sd.metadata.sampleID    = sampleID;
        sd.metadata.testName    = "Test";
        sd.metadata.date        = "2025-01-01";
        sd.metadata.tester      = "QA";
        sd.metadata.media       = "E-liquid";
        sd.metadata.resistance  = 1.0;
        sd.metadata.voltage     = 3.0;
        sd.metadata.power       = 9.0;
        sd.metadata.viscosity   = 50.0;
        sd.metadata.initialOilMass = 1.0;
        sd.metadata.puffingRegime  = "CORESTA";
        sd.metadata.heatingTechnology = "Mesh";
        sd.startColumn = 0;

        double weight = beforeWeight;
        for (int i = 0; i < rowCount; ++i) {
            QVector<QVariant> row;
            double puffs       = (i + 1) * 10.0;
            double after       = weight - oilConsumedPerRow;
            double drawP       = 100.0;
            double resistance  = 1.0;
            QString smell      = "";
            QString clog       = "";
            QString notes      = "";

            // Columns: puffs, beforeWeight, afterWeight, drawPressure,
            //          resistance, smell, clog, notes
            row << QVariant(puffs)
                << QVariant(weight)
                << QVariant(after)
                << QVariant(drawP)
                << QVariant(resistance)
                << QVariant(smell)
                << QVariant(clog)
                << QVariant(notes);

            sd.dataRows.append(row);
            weight = after;  // next row's beforeWeight = this row's afterWeight
        }

        return sd;
    }

private slots:
    // ── Factory ─────────────────────────────────────────────────────────
    void testCreateProcessor()
    {
        std::unique_ptr<DVE::SheetProcessor> p1(DVE::createProcessor("Long Puff lifetime Test"));
        QVERIFY(p1 != nullptr);

        std::unique_ptr<DVE::SheetProcessor> p2(DVE::createProcessor("Unknown Sheet"));
        QVERIFY(p2 != nullptr);  // should return GenericSheetProcessor
    }

    // ── Basic sample building ───────────────────────────────────────────
    void testBuildSampleResult()
    {
        auto sd = makeSampleData(3, "Sample-1");
        QVector<ExcelReader::SampleData> rawSamples;
        rawSamples.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(rawSamples, "Generic", "new");

        QCOMPARE(result.samples.size(), 1);
        QCOMPARE(result.samples[0].sampleID, QStringLiteral("Sample-1"));
        QCOMPARE(result.samples[0].rows.size(), 3);

        // Verify puffs values were copied correctly
        QFUZZY_COMPARE(result.samples[0].rows[0].puffs, 10.0, 0.01);
        QFUZZY_COMPARE(result.samples[0].rows[1].puffs, 20.0, 0.01);
        QFUZZY_COMPARE(result.samples[0].rows[2].puffs, 30.0, 0.01);
    }

    // ── Metric calculation ──────────────────────────────────────────────
    void testCalculateMetrics()
    {
        // before=25.1, after=25.065 → consumed=0.035g = 35mg, interval=10 puffs
        // TPM = consumed_mg / puff_interval = 35 / 10 = 3.5 mg/puff
        // Use 3 rows so average() has >= 2 finite values (its minimum requirement)
        auto sd = makeSampleData(3, "MetricTest", 25.1, 0.035);
        QVector<ExcelReader::SampleData> raw;
        raw.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QCOMPARE(result.samples.size(), 1);
        QFUZZY_COMPARE(result.samples[0].averageTPM, 3.5, 0.5);
    }

    // ── Before-weight repair ────────────────────────────────────────────
    void testBeforeWeightRepair()
    {
        auto sd = makeSampleData(2, "RepairBW");
        // Corrupt row 1: set beforeWeight = 0
        sd.dataRows[1][1] = QVariant(0.0);

        QVector<ExcelReader::SampleData> raw;
        raw.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QCOMPARE(result.samples.size(), 1);
        QVERIFY(result.samples[0].rows.size() >= 2);

        // Row 1 beforeWeight should be repaired to row 0 afterWeight
        double row0After = result.samples[0].rows[0].afterWeight;
        QFUZZY_COMPARE(result.samples[0].rows[1].beforeWeight, row0After, 0.001);
    }

    // ── Puffs repair ────────────────────────────────────────────────────
    void testPuffsRepair()
    {
        auto sd = makeSampleData(2, "RepairPuffs");
        // Corrupt row 1: set puffs = 0
        sd.dataRows[1][0] = QVariant(0.0);

        QVector<ExcelReader::SampleData> raw;
        raw.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QCOMPARE(result.samples.size(), 1);
        QVERIFY(result.samples[0].rows.size() >= 2);

        // Row 1 puffs should be repaired to row 0 puffs + interval (20)
        QFUZZY_COMPARE(result.samples[0].rows[1].puffs, 20.0, 0.01);
    }

    // ── W1 poka-yoke (2026-08-27): puffs repair is DISABLED on stripped
    // sheets. When the reader flags the sheet as cache-stripped (a workbook
    // whose formula caches were destroyed by a save), the zero puffs are
    // destroyed data - the old unconditional fill fabricated plausible +N
    // chains straight into the database during the v2.10.6 smoke. Healthy
    // sheets (count 0, testPuffsRepair above) keep the historical fill. ──
    void strippedSheet_zeroPuffsAreNotFabricated()
    {
        auto sd = makeSampleData(3, "Stripped-1");
        sd.dataRows[1][0] = QVariant(0.0);   // puffs destroyed
        sd.dataRows[2][0] = QVariant(0.0);

        QVector<ExcelReader::SampleData> raw;
        raw.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        proc->setStrippedFormulaCells(7);
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QCOMPARE(result.samples.size(), 1);
        QCOMPARE(result.samples[0].rows[1].puffs, 0.0);   // visible damage,
        QCOMPARE(result.samples[0].rows[2].puffs, 0.0);   // not fiction
    }

    // ── Burn/clog/leak from notes ───────────────────────────────────────
    void testBurnClogLeakFromNotes()
    {
        auto sd = makeSampleData(1, "FlagTest");
        // Set notes to contain burn and clog keywords
        sd.dataRows[0][7] = QVariant(QStringLiteral("Device burned, clog observed"));

        QVector<ExcelReader::SampleData> raw;
        raw.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QCOMPARE(result.samples.size(), 1);
        QCOMPARE(result.samples[0].burnStatus, QStringLiteral("Y"));
        QCOMPARE(result.samples[0].clogStatus, QStringLiteral("Y"));
    }

    // ── Sheet aggregates ────────────────────────────────────────────────
    void testSheetAggregates()
    {
        // Two samples with 3 rows each so average() has enough data
        auto sd1 = makeSampleData(3, "Agg-1", 25.1, 0.030);  // 30mg / 10 = 3.0
        auto sd2 = makeSampleData(3, "Agg-2", 25.1, 0.040);  // 40mg / 10 = 4.0

        QVector<ExcelReader::SampleData> raw;
        raw.append(sd1);
        raw.append(sd2);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QCOMPARE(result.samples.size(), 2);
        // Overall average should be ~3.5
        QFUZZY_COMPARE(result.overallAvgTPM, 3.5, 1.0);
    }

    // ── Empty input ─────────────────────────────────────────────────────
    void testEmptyInput()
    {
        QVector<ExcelReader::SampleData> raw;  // empty

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QVERIFY(result.samples.isEmpty());
    }

    // ── Efficiency calculation ───────────────────────────────────────────
    void testEfficiency()
    {
        // initialOilMass = 1.0g, 5 rows consuming 0.035g each = 0.175g total
        // efficiency = 0.175 / 1.0 * 100 = 17.5%
        auto sd = makeSampleData(5, "Eff-1", 25.1, 0.035);
        sd.metadata.initialOilMass = 1.0;

        QVector<ExcelReader::SampleData> raw;
        raw.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");

        QCOMPARE(result.samples.size(), 1);
        QFUZZY_COMPARE(result.samples[0].efficiencyPercent, 17.5, 2.0);
    }

    // ── Oil is GRAMS, not milligrams (owner ruling 2026-08-03, index D10) ──
    // testEfficiency above passes under EITHER unit, because the old code
    // cancelled its own milligram scaling by dividing initialOilMass by 1000.
    // These assertions pin the actual magnitude, so a regression to milligrams
    // is a failure here rather than a silent 1000x error in metric_defs.unit,
    // the reports and the DB.
    void oilQuantitiesAreGrams()
    {
        // 5 rows x 0.035 g consumed each, weights in grams.
        auto sd = makeSampleData(5, "Oil-1", 25.1, 0.035);
        sd.metadata.initialOilMass = 1.0;

        QVector<ExcelReader::SampleData> raw;
        raw.append(sd);

        std::unique_ptr<DVE::SheetProcessor> proc(DVE::createProcessor("Generic"));
        DVE::SheetResult result = proc->process(raw, "Generic", "new");
        QCOMPARE(result.samples.size(), 1);
        const DVE::SampleResult& s = result.samples[0];

        // Per-row oilConsumed is the CUMULATIVE grams consumed so far.
        QCOMPARE(s.rows.size(), 5);
        for (int i = 0; i < 5; ++i) {
            QFUZZY_COMPARE(s.rows[i].oilConsumed, 0.035 * (i + 1), 1e-6);
            // Guard the specific regression: milligrams would be 1000x this.
            QVERIFY2(s.rows[i].oilConsumed < 1.0,
                     "oilConsumed looks like milligrams - it must be grams");
        }

        // Sample total is grams, and is the same scale as initialOilMass so
        // the two can be divided directly.
        QFUZZY_COMPARE(s.totalOilConsumed, 0.175, 1e-6);
        QVERIFY2(s.totalOilConsumed < s.initialOilMass,
                 "totalOilConsumed must be comparable to initialOilMass (both grams)");
    }

    // ── Per-row regime: new template ────────────────────────────────────
    void perRowRegime_newTemplate_readsStringIntoPuffingRegime()
    {
        ExcelReader::SampleData raw;
        raw.metadata.sampleID = "S1";
        QVector<QVariant> r0 = {10, 25.10, 25.06, 0.45, QString("60mL/3s/30s"),
                                "", "", "", 0, 0, 0, 0};
        QVector<QVariant> r1 = {20, 25.06, 25.02, 0.42, QString("200mL/9s/300s"),
                                "", "", "", 0, 0, 0, 0};
        raw.dataRows << r0 << r1;

        DVE::GenericSheetProcessor proc;
        proc.setPerRowRegime(true);
        DVE::SheetResult sheet = proc.process({raw}, "Lifetime Test", "new");

        QCOMPARE(sheet.hasPerRowRegime, true);
        QCOMPARE(sheet.samples[0].rows[0].puffingRegime, QString("60mL/3s/30s"));
        QCOMPARE(sheet.samples[0].rows[1].puffingRegime, QString("200mL/9s/300s"));
        QCOMPARE(sheet.samples[0].rows[0].resistance, 0.0);
    }

    // ── Per-row regime: old template ────────────────────────────────────
    void perRowRegime_oldTemplate_readsResistanceNumber()
    {
        ExcelReader::SampleData raw;
        raw.metadata.sampleID = "S1";
        QVector<QVariant> r0 = {10, 25.10, 25.06, 0.45, 1.25 /*resistance*/,
                                "", "", "", 0, 0, 0, 0};
        raw.dataRows << r0;

        DVE::GenericSheetProcessor proc;          // setPerRowRegime defaults false
        DVE::SheetResult sheet = proc.process({raw}, "Lifetime Test", "old");

        QCOMPARE(sheet.hasPerRowRegime, false);
        QCOMPARE(sheet.samples[0].rows[0].resistance, 1.25);
        QVERIFY(sheet.samples[0].rows[0].puffingRegime.isEmpty());
    }
};

QTEST_APPLESS_MAIN(tst_SheetProcessors)
#include "tst_sheetprocessors.moc"
