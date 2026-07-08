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
#include "DataCleanup.h"
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

    // ── adaptiveDotRadius tests ─────────────────────────────────────────
    void adaptiveDotRadius_cases();

    // ── lifetimeBarColor tests ──────────────────────────────────────────
    void lifetimeBarColor_distinctPerFile();
    void lifetimeBarColor_progressiveShading();

    // ── loadSopRows tests ────────────────────────────────────────────
    void loadSopRows_filtersToRequestedTests();

    // ── regimeSlideCount tests ────────────────────────────────────────
    void regimeSlideCount_fansOutByRegime();

    // ── DATAVIEWER-16 data-cleanup robustness ──────────────────────────
    // GAP-A: each file's report must drop ONLY that file's excluded rows.
    void buildTable_drawPressureIsMeanOfNonZeroRows();

    void cleanup_buildCleanedFile_usesPerFileIndex();
    // GAP-B: closing a file re-keys surviving exclusions to the shifted indices.
    void cleanup_rekeyAfterClose_shiftsHigherFiles();
};

// isLongPuff / computeTpmYMax tests removed with those helpers: the standing
// axis rule (y anchored at 0, max data value + 1, on every x-y plot) replaced
// the per-test-type y-axis floors. See PlotEngine::applyAnchoredYRange tests
// in tst_plotengine.

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

// ─────────────────────────────────────────────────────────────────────────────
// DATAVIEWER-16 — data-cleanup robustness (GAP-A / GAP-B)
//
// Build N independent single-sheet files whose only sample has a handful of
// valid rows; record an exclusion against ONE file and verify the cleaned
// output drops the row only for that file, and that closing an earlier file
// re-keys the surviving exclusion onto the shifted index.
static DVE::FileResult makeCleanupFile(const QString& name, int rowCount)
{
    DVE::SampleResult s;
    s.sampleName = "Sample 1";
    double w = 25.10;
    for (int i = 0; i < rowCount; ++i) {
        DVE::DataRow r;
        r.puffs        = (i + 1) * 10;
        r.beforeWeight = w;
        w -= 0.04;
        r.afterWeight  = w;
        s.rows.append(r);
    }

    DVE::SheetResult sh;
    sh.sheetName = "Lifetime Test";
    sh.samples.append(s);

    DVE::FileResult f;
    f.fileName = name;
    f.filePath = "C:/data/" + name + ".xlsx";
    f.sheets.append(sh);
    return f;
}

void tst_ReportGenerator::buildTable_drawPressureIsMeanOfNonZeroRows()
{
    // Values {1.0, 2.0, 6.0} plus a zero (empty) row: the table cell must show
    // the mean of the non-zero values, 3.00. The old median would give 2.00.
    DVE::SheetResult sh;
    sh.sheetName = "Big Headspace";
    DVE::SampleResult s;
    s.sampleName = "S1";
    for (double dp : {1.0, 2.0, 6.0, 0.0}) {
        DVE::DataRow r;
        r.drawPressure = dp;
        s.rows.append(r);
    }
    sh.samples.append(s);

    DVE::ReportGenerator gen;
    DVE::ReportConfig config;   // empty selectedColumns → default column set
    DVE::SlideTable tbl = gen.buildTableForTesting(sh, config);

    int dpCol = -1;
    for (int i = 0; i < tbl.headers.size(); ++i)
        if (tbl.headers[i].contains("Draw", Qt::CaseInsensitive)) { dpCol = i; break; }
    QVERIFY2(dpCol >= 0, "Draw Pressure column missing from default headers");
    QCOMPARE(tbl.rows.size(), 1);
    QCOMPARE(tbl.rows[0][dpCol], QString("3.00"));
}

void tst_ReportGenerator::cleanup_buildCleanedFile_usesPerFileIndex()
{
    // Three files, each with one sample of 5 valid rows.
    DVE::FileResult f0 = makeCleanupFile("file0", 5);
    DVE::FileResult f1 = makeCleanupFile("file1", 5);
    DVE::FileResult f2 = makeCleanupFile("file2", 5);

    const int rowsPerSample = f1.sheets[0].samples[0].rows.size();
    QCOMPARE(rowsPerSample, 5);

    // Exclude row index 2 in FILE 1 only (file:sheet:sample = 1:0:0).
    QMap<QString, QSet<int>> excl;
    excl.insert(DVE::DataCleanup::key(1, 0, 0), QSet<int>{2});

    DVE::FileResult c0 = DVE::DataCleanup::buildCleanedFile(f0, excl, 0);
    DVE::FileResult c1 = DVE::DataCleanup::buildCleanedFile(f1, excl, 1);
    DVE::FileResult c2 = DVE::DataCleanup::buildCleanedFile(f2, excl, 2);

    // GAP-A: only file 1's report drops the excluded row. The other two are
    // untouched. (The pre-fix code hardcoded m_currentFileIndex, so whichever
    // file happened to be current got its exclusion applied to ALL files.)
    QCOMPARE(c0.sheets[0].samples[0].rows.size(), rowsPerSample);     // untouched
    QCOMPARE(c1.sheets[0].samples[0].rows.size(), rowsPerSample - 1); // row dropped
    QCOMPARE(c2.sheets[0].samples[0].rows.size(), rowsPerSample);     // untouched

    // The dropped file records a cleanup note; the others do not.
    QVERIFY(c1.sheets[0].samples[0].extra.contains("cleanupNote"));
    QVERIFY(!c0.sheets[0].samples[0].extra.contains("cleanupNote"));
    QVERIFY(!c2.sheets[0].samples[0].extra.contains("cleanupNote"));
}

void tst_ReportGenerator::cleanup_rekeyAfterClose_shiftsHigherFiles()
{
    // Exclusions across three files: file 0, file 1 and file 2 each have one.
    QMap<QString, QSet<int>> excl;
    excl.insert(DVE::DataCleanup::key(0, 0, 0), QSet<int>{1});
    excl.insert(DVE::DataCleanup::key(1, 0, 0), QSet<int>{2});
    excl.insert(DVE::DataCleanup::key(2, 1, 3), QSet<int>{4, 5});

    // Close file 0: it is removed; files 1 and 2 shift down to 0 and 1.
    QMap<QString, QSet<int>> after = DVE::DataCleanup::rekeyAfterClose(excl, 0);

    // GAP-B: the closed file's key is gone…
    QVERIFY(!after.contains(DVE::DataCleanup::key(0, 0, 0)) ||
            after.value(DVE::DataCleanup::key(0, 0, 0)) != QSet<int>{1});
    // …former file 1 now lives at index 0 with its exclusion intact…
    QCOMPARE(after.value(DVE::DataCleanup::key(0, 0, 0)), (QSet<int>{2}));
    // …and former file 2 now lives at index 1, sheet/sample/rows preserved.
    QCOMPARE(after.value(DVE::DataCleanup::key(1, 1, 3)), (QSet<int>{4, 5}));
    // Nothing lingers at the old, now-wrong index 2.
    QVERIFY(!after.contains(DVE::DataCleanup::key(2, 1, 3)));
    QCOMPARE(after.size(), 2);

    // Verify the re-keyed exclusion still selects the right row in the surviving
    // file (former file 1 → index 0): build it and confirm the row is dropped.
    DVE::FileResult survivor = makeCleanupFile("file1", 5);
    DVE::FileResult cleaned  = DVE::DataCleanup::buildCleanedFile(survivor, after, 0);
    QCOMPARE(cleaned.sheets[0].samples[0].rows.size(), 4);  // one of 5 dropped
}

QTEST_MAIN(tst_ReportGenerator)
#include "tst_reportgenerator.moc"
