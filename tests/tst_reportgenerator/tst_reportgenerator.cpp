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
#include <QColor>
#include "TestHelpers.h"
#include "ReportData.h"
#include "DataProcessor.h"
#include "ReportGenerator.h"
#include "SopLoader.h"
#include "AppTheme.h"

class tst_ReportGenerator : public QObject
{
    Q_OBJECT

private:
    QTemporaryDir m_tempDir;
    bool pythonAvailable = false;
    DVE::FileResult m_formatE;  // cached processed result

    QString findTemplateDir() const
    {
        // Walk up from executable directory looking for resources/templates
        QDir d(QCoreApplication::applicationDirPath());
        for (int i = 0; i < 6; ++i) {
            QString candidate = d.absoluteFilePath("resources/templates");
            if (QDir(candidate).exists()) return candidate;
            candidate = d.absoluteFilePath("../resources/templates");
            if (QDir(candidate).exists()) return candidate;
            d.cdUp();
        }
        return QString();
    }

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

    // ── isLongPuff tests ────────────────────────────────────────────────
    void isLongPuff_sheetNameMatch();
    void isLongPuff_regimeRegexMatch();
    void isLongPuff_neither();

    // ── computeTpmYMax tests ────────────────────────────────────────────
    void yMax_nonLongPuff_underCeiling();
    void yMax_nonLongPuff_overCeiling();
    void yMax_longPuff_inRange();
    void yMax_longPuff_belowMin();
    void yMax_longPuff_aboveMax();

    // ── adaptiveDotRadius tests ─────────────────────────────────────────
    void adaptiveDotRadius_cases();

    // ── lifetimeBarColor tests ──────────────────────────────────────────
    void lifetimeBarColor_distinctPerFile();
    void lifetimeBarColor_progressiveShading();

    // ── loadSopRows tests ────────────────────────────────────────────
    void loadSopRows_filtersToRequestedTests();

    // ── regimeSlideCount tests ────────────────────────────────────────
    void regimeSlideCount_fansOutByRegime();
};

void tst_ReportGenerator::isLongPuff_sheetNameMatch()
{
    DVE::ReportGenerator gen;
    DVE::SheetResult sheet;
    sheet.sheetName = "Long Puff Lifetime Test";
    QVERIFY(gen.isLongPuffForTesting(sheet));
}

void tst_ReportGenerator::isLongPuff_regimeRegexMatch()
{
    DVE::ReportGenerator gen;
    DVE::SheetResult sheet;
    sheet.sheetName = "Lifetime Test";
    DVE::SampleResult s;
    s.puffingRegime = "200mL/10s/60s";
    sheet.samples.append(s);
    QVERIFY(gen.isLongPuffForTesting(sheet));
}

void tst_ReportGenerator::isLongPuff_neither()
{
    DVE::ReportGenerator gen;
    DVE::SheetResult sheet;
    sheet.sheetName = "Lifetime Test";
    DVE::SampleResult s;
    s.puffingRegime = "55mL/3s/30s";
    sheet.samples.append(s);
    QVERIFY(!gen.isLongPuffForTesting(sheet));
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper for constructing test SheetResult
static DVE::SheetResult sheetWith(const QString& sheetName,
                                  const QString& regime,
                                  const QVector<double>& tpmValues)
{
    DVE::SheetResult sh;
    sh.sheetName = sheetName;
    DVE::SampleResult s;
    s.puffingRegime = regime;
    double sum = 0;
    for (double t : tpmValues) {
        DVE::DataRow r;
        r.tpm = t;
        s.rows.append(r);
        sum += t;
    }
    s.averageTPM = tpmValues.isEmpty() ? 0.0 : sum / tpmValues.size();
    sh.samples.append(s);
    return sh;
}

void tst_ReportGenerator::yMax_nonLongPuff_underCeiling()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Lifetime Test", "55mL/3s/30s", {3, 4, 5, 6, 6.5});
    QVERIFY(qAbs(gen.computeTpmYMaxForTesting(sh) - 7.0) < 1e-6);
}

void tst_ReportGenerator::yMax_nonLongPuff_overCeiling()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Lifetime Test", "55mL/3s/30s", {7.5, 8.0, 8.5});
    QVERIFY(qAbs(gen.computeTpmYMaxForTesting(sh) - 9.5) < 1e-6);
}

void tst_ReportGenerator::yMax_longPuff_inRange()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Long Puff Lifetime Test", "200mL/10s/60s", {15, 18, 22, 24});
    QVERIFY(qAbs(gen.computeTpmYMaxForTesting(sh) - 25.0) < 1e-6);
}

void tst_ReportGenerator::yMax_longPuff_belowMin()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Long Puff Lifetime Test", "200mL/10s/60s", {8, 10, 12, 14});
    QVERIFY(qAbs(gen.computeTpmYMaxForTesting(sh) - 15.0) < 1e-6);
}

void tst_ReportGenerator::yMax_longPuff_aboveMax()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Long Puff Lifetime Test", "200mL/10s/60s", {18, 22, 27});
    QVERIFY(qAbs(gen.computeTpmYMaxForTesting(sh) - 28.0) < 1e-6);
}

void tst_ReportGenerator::adaptiveDotRadius_cases()
{
    DVE::ReportGenerator gen;
    QCOMPARE(gen.adaptiveDotRadiusForTesting(0),    5);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(30),   5);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(150),  2);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(1000), 2);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(90),   4);
}

void tst_ReportGenerator::lifetimeBarColor_distinctPerFile()
{
    // Each file gets a distinct base hue from AppTheme::seriesColor; single-sample
    // files (totalSamples=1) return the base color directly via AppTheme::shade.
    QColor f0 = DVE::ReportGenerator::lifetimeBarColor(AppTheme::seriesColor(0), 0, 1);
    QColor f1 = DVE::ReportGenerator::lifetimeBarColor(AppTheme::seriesColor(1), 0, 1);
    QColor f2 = DVE::ReportGenerator::lifetimeBarColor(AppTheme::seriesColor(2), 0, 1);
    QVERIFY(f0 != f1);
    QVERIFY(f1 != f2);
    QVERIFY(f0 != f2);
}

void tst_ReportGenerator::lifetimeBarColor_progressiveShading()
{
    // Three samples within the same file → shade() must produce 3 distinct colors
    // all derived from the same base hue. seriesColor(0) is a chromatic blue, so
    // its hue angle is well-defined and preserved by shade().
    const QColor base = AppTheme::seriesColor(0);
    QColor a = DVE::ReportGenerator::lifetimeBarColor(base, 0, 3);
    QColor b = DVE::ReportGenerator::lifetimeBarColor(base, 1, 3);
    QColor c = DVE::ReportGenerator::lifetimeBarColor(base, 2, 3);
    QVERIFY(a != b && b != c);
    int ha, sa, va, hb, sb, vb, hc, sc, vc;
    a.getHsv(&ha, &sa, &va);
    b.getHsv(&hb, &sb, &vb);
    c.getHsv(&hc, &sc, &vc);
    QCOMPARE(ha, hb);
    QCOMPARE(hb, hc);
    QVERIFY(va != vb || sa != sb);
}

void tst_ReportGenerator::loadSopRows_filtersToRequestedTests()
{
    const QString templateDir = findTemplateDir();
    QVERIFY2(!templateDir.isEmpty(), "Could not find resources/templates directory");

    // v2.0.9 — ReportGenerator now reads SOPs from the TPM template (not the
    // removed standalone sops.xlsx). Skip on MIP-developer machines where
    // QXlsx sees ciphertext. Runs cleanly on CI / deployment self-test machines.
    const QString templateXlsx = templateDir +
        "/Standardized Test Template - December 2025.xlsx";
    QFile probe(templateXlsx);
    if (probe.open(QIODevice::ReadOnly)) {
        const QByteArray head = probe.read(16);
        probe.close();
        if (head.startsWith("%TSD-Header")) {
            QSKIP("Template is MIP-encrypted on this developer machine; runs on clean CI/prod.");
        }
    }

    DVE::ReportGenerator gen;
    gen.setResourcePath(templateDir + "/..");
    const QStringList request = {"Lifetime Test"};
    QVector<DVE::SopEntry> filtered = gen.loadSopRowsForTesting(request);
    QCOMPARE(filtered.size(), 1);
    QCOMPARE(filtered[0].test, QString("Lifetime Test"));
}

void tst_ReportGenerator::regimeSlideCount_fansOutByRegime()
{
    DVE::SheetResult two; two.sheetName = "Lifetime Test"; two.hasPerRowRegime = true;
    DVE::SampleResult s;
    DVE::DataRow a; a.beforeWeight=25.1; a.afterWeight=25.06; a.puffingRegime="60mL/3s/30s";
    DVE::DataRow b; b.beforeWeight=25.06; b.afterWeight=25.02; b.puffingRegime="200mL/9s/300s";
    s.rows << a << b; two.samples << s;
    QCOMPARE(DVE::ReportGenerator::regimeSlideCount(two), 2);

    DVE::SheetResult old; old.sheetName = "Lifetime Test";
    DVE::SampleResult so; DVE::DataRow c; c.beforeWeight=25.1; c.afterWeight=25.06; so.rows << c;
    old.samples << so;
    QCOMPARE(DVE::ReportGenerator::regimeSlideCount(old), 1);   // no regime -> single slide
}

QTEST_MAIN(tst_ReportGenerator)
#include "tst_reportgenerator.moc"
