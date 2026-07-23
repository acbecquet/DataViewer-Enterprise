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

    // ── v3: raw typed grid accessor ──────────────────────────────────────
    void currentSheetCellsExposesGrid()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader r;
        QVERIFY(r.loadFile(testDataFile("format_e.xlsx")));
        QVERIFY(r.selectSheet(r.getSheetNames().first()));
        const auto cells = r.currentSheetCells();
        QVERIFY(!cells.isEmpty());
        QVERIFY(cells.size() >= 5);          // header band + col headers + data
        QVERIFY(!cells[3].isEmpty());        // row 4 (0-based 3) carries col headers

        // No sheet selected (fresh, unloaded reader) → empty grid.
        ExcelReader r2;
        QVERIFY(r2.currentSheetCells().isEmpty());
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

    // ── v2.2.1 new-template: per-row Puffing Regime column ──────────────
    // format_e_regime.xlsx has col-E header "Puffing Regime" (not "Resistance
    // (Ω)") and per-row regime strings in the data cells.  ExcelReader returns
    // this verbatim; DataProcessor uses it to set hasPerRowRegime.
    void testPerRowPuffingRegime()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        ExcelReader reader;
        QVERIFY(reader.loadFile(testDataFile("format_e_regime.xlsx")));
        QVERIFY(reader.selectSheet("Lifetime Test"));

        // Col-E header (index 4) must read "Puffing Regime" — not "Resistance (Ω)"
        QStringList headers = reader.getColumnHeaders();
        QVERIFY2(headers.size() > 4, "Expected at least 5 column headers");
        QCOMPARE(headers.at(4), QStringLiteral("Puffing Regime"));

        // Data rows: col-E (index 4) must carry the per-row regime string
        ExcelReader::SampleData sample = reader.getSample(0);
        QVERIFY2(!sample.dataRows.isEmpty(), "Expected at least one data row");

        QString firstRegime = sample.dataRows.first().at(4).toString();
        QCOMPARE(firstRegime, QStringLiteral("60mL/3s/30s"));

        // Mid-session regime change: row 4 (0-based index 3) switched to new regime
        QVERIFY2(sample.dataRows.size() >= 5, "Expected 5 data rows");
        QString lastRegime = sample.dataRows.last().at(4).toString();
        QCOMPARE(lastRegime, QStringLiteral("200mL/9s/300s"));
    }

    // ── Tolerant numeric header cells ───────────────────────────────────
    // Testers type units into numeric cells; the real-world case was a B3
    // viscosity of "800kcp" that read as 0 in every report.
    void testTolerantCellDouble()
    {
        using ER = ExcelReader;
        QCOMPARE(ER::tolerantCellDouble(QVariant(2.09)), 2.09);
        QCOMPARE(ER::tolerantCellDouble(QVariant("0.9")), 0.9);
        QCOMPARE(ER::tolerantCellDouble(QVariant("800kcp")), 800000.0);
        QCOMPARE(ER::tolerantCellDouble(QVariant("800 kcP")), 800000.0);
        QCOMPARE(ER::tolerantCellDouble(QVariant("100 cP")), 100.0);
        QCOMPARE(ER::tolerantCellDouble(QVariant("2.09 Ohm")), 2.09);
        QCOMPARE(ER::tolerantCellDouble(QVariant("3.4V")), 3.4);
        QCOMPARE(ER::tolerantCellDouble(QVariant(".5")), 0.5);
        QCOMPARE(ER::tolerantCellDouble(QVariant("N/A")), 0.0);
        QCOMPARE(ER::tolerantCellDouble(QVariant()), 0.0);
    }
};

QTEST_APPLESS_MAIN(tst_ExcelReader)
#include "tst_excelreader.moc"
