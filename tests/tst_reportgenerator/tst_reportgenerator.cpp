// ── DataViewer Enterprise — ReportGenerator Integration Tests ───────────────
//
// NOTE: Plot rendering (QPainter on QPixmap) can deadlock in headless/MSYS
// environments on Windows.  Tests that exercise report generation disable
// plots (config.includePlots = false) so they run reliably in CI/CLI.
// Plot rendering is separately validated by tst_plotengine.
// ────────────────────────────────────────────────────────────────────────────
#include <QtTest>
#include <QTemporaryDir>
#include <QFile>
#include "TestHelpers.h"
#include "ReportData.h"
#include "DataProcessor.h"
#include "ReportGenerator.h"

class tst_ReportGenerator : public QObject
{
    Q_OBJECT

private:
    QTemporaryDir m_tempDir;
    bool pythonAvailable = false;
    DVE::FileResult m_formatE;  // cached processed result

    bool isValidZip(const QString& path) {
        QFile f(path);
        if (!f.open(QIODevice::ReadOnly)) return false;
        QByteArray header = f.read(2);
        f.close();
        return header.size() == 2 && header[0] == 'P' && header[1] == 'K';
    }

private slots:
    void initTestCase()
    {
        QVERIFY(m_tempDir.isValid());

        DVE::DataProcessor dp;
        m_formatE = dp.processFile(testDataFile("format_e.xlsx"));
        pythonAvailable = !m_formatE.filePath.isEmpty();
        if (!pythonAvailable)
            qWarning("Python/openpyxl not found — most tests will be skipped.");
    }

    // ── Full report (table-only, no plots — plots tested in tst_plotengine) ──
    void testFullReport()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::ReportGenerator gen;
        DVE::ReportConfig config;
        config.outputPath    = m_tempDir.path() + "/full_report.pptx";
        config.reportTitle   = "Full Report Test";
        config.includePlots  = false;
        config.includeImages = false;

        bool ok = gen.generateFullReport(m_formatE, config);
        QVERIFY2(ok, qPrintable(gen.lastError()));
        QVERIFY(QFile::exists(config.outputPath));
        QVERIFY(isValidZip(config.outputPath));
    }

    // ── Single-sheet report ─────────────────────────────────────────────
    void testSingleSheetReport()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::ReportGenerator gen;
        DVE::ReportConfig config;
        config.outputPath      = m_tempDir.path() + "/single_sheet.pptx";
        config.reportTitle     = "Single Sheet Test";
        config.singleSheetOnly = true;
        config.includePlots    = false;
        config.includeImages   = false;

        QVERIFY(!m_formatE.sheetNames.isEmpty());
        config.singleSheetName = m_formatE.sheetNames.first();

        bool ok = gen.generateTestReport(m_formatE,
                                         config.singleSheetName,
                                         config);
        QVERIFY2(ok, qPrintable(gen.lastError()));
        QVERIFY(QFile::exists(config.outputPath));
        QVERIFY(isValidZip(config.outputPath));
    }

    // ── Progress callback ───────────────────────────────────────────────
    void testProgressCallback()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::ReportGenerator gen;
        DVE::ReportConfig config;
        config.outputPath    = m_tempDir.path() + "/progress_test.pptx";
        config.reportTitle   = "Progress Test";
        config.includePlots  = false;
        config.includeImages = false;

        bool callbackInvoked = false;
        gen.generateFullReport(m_formatE, config,
            [&](int /*pct*/, const QString& /*msg*/) {
                callbackInvoked = true;
            });

        QVERIFY(callbackInvoked);
    }

    // ── Empty data ──────────────────────────────────────────────────────
    void testEmptyData()
    {
        DVE::ReportGenerator gen;
        DVE::ReportConfig config;
        config.outputPath  = m_tempDir.path() + "/empty_report.pptx";
        config.reportTitle = "Empty Data Test";

        DVE::FileResult emptyResult;
        bool ok = gen.generateFullReport(emptyResult, config);

        // Empty data should return false
        QVERIFY(!ok);
    }

    // ── Output path verification ────────────────────────────────────────
    void testOutputPath()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        QString expectedPath = m_tempDir.path() + "/specific_path_report.pptx";

        DVE::ReportGenerator gen;
        DVE::ReportConfig config;
        config.outputPath    = expectedPath;
        config.reportTitle   = "Path Test";
        config.includePlots  = false;
        config.includeImages = false;

        bool ok = gen.generateFullReport(m_formatE, config);
        QVERIFY2(ok, qPrintable(gen.lastError()));
        QVERIFY2(QFile::exists(expectedPath),
                 qPrintable("File not created at expected path: " + expectedPath));
    }
};

QTEST_MAIN(tst_ReportGenerator)
#include "tst_reportgenerator.moc"
