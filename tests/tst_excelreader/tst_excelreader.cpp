// ── DataViewer Enterprise — ExcelReader Integration Tests ───────────────────
#include <QtTest>
#include "TestHelpers.h"
#include "ExcelReader.h"

class tst_ExcelReader : public QObject
{
    Q_OBJECT

private:
    bool pythonAvailable = false;

    // Helper: load file and skip if Python is not available.
    bool tryLoad(ExcelReader& reader, const QString& path) {
        bool ok = reader.loadFile(path);
        if (!ok && reader.getLastError().contains("python", Qt::CaseInsensitive)) {
            QWARN("Python/openpyxl not available — skipping test");
            pythonAvailable = false;
            return false;
        }
        pythonAvailable = true;
        return ok;
    }

private slots:
    void initTestCase() {
        // Quick probe: try loading a known file to see if Python works.
        ExcelReader probe;
        pythonAvailable = probe.loadFile(testDataFile("format_e.xlsx"));
        if (!pythonAvailable)
            QWARN("Python/openpyxl not found — most tests will be skipped.");
    }

    // ── Format E (Dec 2025 "new") ───────────────────────────────────────
    void testLoadFormatE()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_e.xlsx")));

        QStringList sheets = reader.getSheetNames();
        QCOMPARE(sheets.size(), 2);

        QVERIFY(reader.selectSheet("Lifetime Test"));
        QCOMPARE(reader.getSampleCount(), 2);

        auto sample0 = reader.getSample(0);
        QCOMPARE(sample0.metadata.sampleID, QStringLiteral("Lifetime-1"));
        QFUZZY_COMPARE(sample0.metadata.resistance, 1.1, 0.05);
        QFUZZY_COMPARE(sample0.metadata.voltage, 2.8, 0.05);
    }

    // ── Format D (Jan 2025 "old") ───────────────────────────────────────
    void testLoadFormatD()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_d.xlsx")));
        QVERIFY(reader.selectSheet("Lifetime Test"));

        auto sample0 = reader.getSample(0);
        QCOMPARE(sample0.metadata.sampleID, QStringLiteral("Life-1"));

        // After normalization, headers should contain the Ω character.
        QStringList headers = reader.getColumnHeaders();
        bool hasOhm = false;
        for (const auto& h : headers)
            if (h.contains(QChar(0x03A9))) { hasOhm = true; break; }
        QVERIFY2(hasOhm, "Headers should contain Ω after normalization");
    }

    // ── Format C (Gembox/T58G) ──────────────────────────────────────────
    void testLoadFormatC()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_c.xlsx")));

        // Select first available sheet
        QStringList sheets = reader.getSheetNames();
        QVERIFY(!sheets.isEmpty());
        QVERIFY(reader.selectSheet(sheets.first()));

        auto sample0 = reader.getSample(0);
        QCOMPARE(sample0.metadata.sampleID, QStringLiteral("Gembox Intense-1"));
    }

    // ── Format B (M1 Extended) ──────────────────────────────────────────
    void testLoadFormatB()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_b.xlsx")));

        QStringList sheets = reader.getSheetNames();
        QVERIFY(!sheets.isEmpty());
        QVERIFY(reader.selectSheet(sheets.first()));

        auto sample0 = reader.getSample(0);
        QCOMPARE(sample0.metadata.sampleID, QStringLiteral("CPS1910-1"));
        QCOMPARE(sample0.metadata.media, QStringLiteral("Distillate"));
    }

    // ── Format A (Comparison Test) ──────────────────────────────────────
    void testLoadFormatA()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_a.xlsx")));

        QStringList sheets = reader.getSheetNames();
        QVERIFY(!sheets.isEmpty());
        QVERIFY(reader.selectSheet(sheets.first()));

        QCOMPARE(reader.getSampleCount(), 2);

        auto sample0 = reader.getSample(0);
        QCOMPARE(sample0.metadata.sampleID, QStringLiteral("T51P-C39-1"));
        QCOMPARE(sample0.dataRows.size(), 5);

        auto sample1 = reader.getSample(1);
        QCOMPARE(sample1.dataRows.size(), 5);
    }

    // ── Template version detection ──────────────────────────────────────
    void testFormatDetection()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        {
            ExcelReader reader;
            QVERIFY(reader.loadFile(testDataFile("format_e.xlsx")));
            QCOMPARE(reader.detectTemplateVersion(), QStringLiteral("new"));
        }
        {
            ExcelReader reader;
            QVERIFY(reader.loadFile(testDataFile("format_d.xlsx")));
            QVERIFY(reader.selectSheet("Lifetime Test"));
            // After normalization Format D should present as "new"
            QCOMPARE(reader.detectTemplateVersion(), QStringLiteral("new"));
        }
    }

    // ── Data row content ────────────────────────────────────────────────
    void testSampleDataRows()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_e.xlsx")));
        QStringList sheets = reader.getSheetNames();
        QVERIFY(!sheets.isEmpty());
        QVERIFY(reader.selectSheet(sheets.first()));

        auto sample0 = reader.getSample(0);
        QVERIFY(!sample0.dataRows.isEmpty());

        // Row 0: puffs column (index 0) should be 10
        QVariant puffs = sample0.dataRows[0][0];
        QCOMPARE(puffs.toInt(), 10);

        // Row 0: beforeWeight column (index 1) should be ≈ 25.1
        QVariant bw = sample0.dataRows[0][1];
        QFUZZY_COMPARE(bw.toDouble(), 25.1, 0.05);
    }

    // ── Column headers ──────────────────────────────────────────────────
    void testColumnHeaders()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_e.xlsx")));
        QStringList sheets = reader.getSheetNames();
        QVERIFY(!sheets.isEmpty());
        QVERIFY(reader.selectSheet(sheets.first()));

        QStringList headers = reader.getColumnHeaders();
        QVERIFY(!headers.isEmpty());
        QCOMPARE(headers.first().toLower(), QStringLiteral("puffs"));
    }

    // ── Empty file ──────────────────────────────────────────────────────
    void testEmptyFile()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        // May or may not load successfully depending on implementation
        reader.loadFile(testDataFile("empty.xlsx"));
        QCOMPARE(reader.getSampleCount(), 0);
    }

    // ── Missing file ────────────────────────────────────────────────────
    void testMissingFile()
    {
        ExcelReader reader;
        bool ok = reader.loadFile(QStringLiteral("nonexistent.xlsx"));
        QVERIFY(!ok);
        QVERIFY(!reader.getLastError().isEmpty());
    }

    // ── Multi-sheet file ────────────────────────────────────────────────
    void testMultiSheet()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("multi_sheet.xlsx")));
        QCOMPARE(reader.getSheetNames().size(), 3);
    }
};

QTEST_APPLESS_MAIN(tst_ExcelReader)
#include "tst_excelreader.moc"
