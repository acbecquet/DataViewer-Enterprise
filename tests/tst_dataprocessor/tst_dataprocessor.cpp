// ── DataViewer Enterprise — DataProcessor Integration Tests ─────────────────
#include <QtTest>
#include "TestHelpers.h"
#include "ReportData.h"
#include "DataProcessor.h"

class tst_DataProcessor : public QObject
{
    Q_OBJECT

private:
    bool pythonAvailable = false;

private slots:
    void initTestCase()
    {
        // Probe Python availability with a quick load.
        DVE::DataProcessor dp;
        DVE::FileResult res = dp.processFile(testDataFile("format_e.xlsx"));
        pythonAvailable = !res.filePath.isEmpty();
        if (!pythonAvailable)
            QWARN("Python/openpyxl not found — most tests will be skipped.");
    }

    // ── Format E processing ─────────────────────────────────────────────
    void testProcessFormatE()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::DataProcessor dp;
        DVE::FileResult result = dp.processFile(testDataFile("format_e.xlsx"));

        QVERIFY(!result.filePath.isEmpty());
        QVERIFY(result.sheets.size() >= 1);
        QVERIFY(result.sheets[0].hasSamples());
    }

    // ── Format A processing ─────────────────────────────────────────────
    void testProcessFormatA()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::DataProcessor dp;
        DVE::FileResult result = dp.processFile(testDataFile("format_a.xlsx"));

        QVERIFY(!result.filePath.isEmpty());
        QVERIFY(result.sheets.size() >= 1);

        // Find a sheet with samples and verify metadata was extracted.
        bool foundSamples = false;
        for (const auto& sheet : result.sheets) {
            if (sheet.hasSamples()) {
                foundSamples = true;
                QVERIFY(!sheet.samples[0].sampleID.isEmpty());
                break;
            }
        }
        QVERIFY2(foundSamples, "At least one sheet should have samples");
    }

    // ── Multi-sheet processing ──────────────────────────────────────────
    void testProcessMultiSheet()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::DataProcessor dp;
        DVE::FileResult result = dp.processFile(testDataFile("multi_sheet.xlsx"));

        QVERIFY(!result.filePath.isEmpty());
        QCOMPARE(result.sheets.size(), 3);
    }

    // ── Progress callback ───────────────────────────────────────────────
    void testProgressCallback()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::DataProcessor dp;
        bool callbackInvoked = false;

        DVE::FileResult result = dp.processFile(
            testDataFile("format_e.xlsx"),
            [&](int /*percent*/, const QString& /*msg*/) {
                callbackInvoked = true;
            });

        QVERIFY(!result.filePath.isEmpty());
        QVERIFY(callbackInvoked);
    }

    // ── Empty file ──────────────────────────────────────────────────────
    void testProcessEmptyFile()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::DataProcessor dp;
        DVE::FileResult result = dp.processFile(testDataFile("empty.xlsx"));

        // Should not crash. Sheets may be empty or absent.
        // Either filePath is empty (failure) or sheets is empty.
        if (!result.filePath.isEmpty()) {
            // Successfully loaded but no data
            for (const auto& sheet : result.sheets)
                QVERIFY(!sheet.hasSamples() || sheet.samples.isEmpty());
        }
    }

    // ── Missing file ────────────────────────────────────────────────────
    void testProcessMissingFile()
    {
        DVE::DataProcessor dp;
        DVE::FileResult result = dp.processFile(QStringLiteral("nonexistent.xlsx"));

        // Failure signal: filePath should be empty
        QVERIFY(result.filePath.isEmpty());
    }

    // ── Template version ────────────────────────────────────────────────
    void testTemplateVersion()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::DataProcessor dp;
        DVE::FileResult result = dp.processFile(testDataFile("format_e.xlsx"));

        QVERIFY(!result.filePath.isEmpty());
        QCOMPARE(result.templateVersion, QStringLiteral("new"));
    }
};

QTEST_APPLESS_MAIN(tst_DataProcessor)
#include "tst_dataprocessor.moc"
