// ── DataViewer Enterprise — DataProcessor Integration Tests ─────────────────
#include <QtTest>
#include <QFileInfo>
#include <QProcess>
#include <QScopeGuard>
#include <QTemporaryDir>
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

    // ── Write provenance (Phase 2b) ─────────────────────────────────────
    // The standard schema-driven path must record where each sheet's cells
    // came from (block geometry, column keys, header value cells, per-sample
    // block origin) so write-back can derive addresses instead of assuming
    // the 12-wide standardized layout.
    void writeProvenanceRecordedOnStandardParse()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        DVE::DataProcessor dp;
        DVE::FileResult f = dp.processFile(testDataFile("format_e.xlsx"));
        QVERIFY(!f.filePath.isEmpty());
        QVERIFY(!f.sheets.isEmpty());

        const DVE::SheetResult& sh = f.sheets[0];
        QVERIFY(sh.hasSamples());
        QCOMPARE(sh.blockCols, 12);
        QCOMPARE(sh.dataStartRow, 5);
        QCOMPARE(sh.columnKeys.size(), 12);
        QCOMPARE(sh.columnKeys[0], QStringLiteral("puffs"));
        QCOMPARE(sh.columnKeys[5], QStringLiteral("smell"));
        QVERIFY(sh.headerCells.contains(QStringLiteral("media")));
        QCOMPARE(sh.headerCells.value(QStringLiteral("media")), QPoint(2, 2));
        QVERIFY(sh.samples.size() >= 2);
        for (int i = 0; i < sh.samples.size(); ++i)
            QCOMPARE(sh.samples[i].startColumn, i * 12);
    }

    void writeProvenanceRecordsProjectLayoutHeaders()
    {
        if (!pythonAvailable) QSKIP("Python not available");

        // format_c is the Project-variant fixture ("Project:" landmark at
        // row 1 col 6). Expected value cells come from StandardSchema.cpp's
        // Project layout hf(...) table: hf(key, ..., row, col) -> QPoint(col, row).
        DVE::DataProcessor dp;
        DVE::FileResult f = dp.processFile(testDataFile("format_c.xlsx"));
        QVERIFY(!f.filePath.isEmpty());
        QVERIFY(!f.sheets.isEmpty());

        const DVE::SheetResult& sh = f.sheets[0];
        QVERIFY(sh.hasSamples());
        // Project layout: media (2,2); tester value cell on ROW 1 col 5; NO
        // sample_id entry (assembled from project+sample - not single-cell
        // addressable); project_name present instead.
        QCOMPARE(sh.headerCells.value(QStringLiteral("media")), QPoint(2, 2));
        QCOMPARE(sh.headerCells.value(QStringLiteral("tester")), QPoint(5, 1));
        QVERIFY(!sh.headerCells.contains(QStringLiteral("sample_id")));
        QVERIFY(sh.headerCells.contains(QStringLiteral("project_name")));
    }

    // ── W3b stripped-lineage loads (smoke-fix Task 7, 2026-08-31) ───────
    //
    // makeStrippedVariant (make_stripped_variant.py) copies a fixture with
    // the puffs-column literals of data rows 2+ replaced by UNCACHED
    // formulas: the on-disk state of every app-template-lineage workbook.
    // The bundled New File template is openpyxl-born with =prev+K puff
    // chains pre-filled and zero caches by construction, which makes it
    // BYTE-INDISTINGUISHABLE from a workbook the OLD openpyxl write-back
    // destroyed (same standard layout - the Dec 2025 template is itself
    // per-row-regime - same formula shapes, same zero-cache state). A
    // load-time gate therefore cannot exist without misfiring on legitimate
    // files (v2.10.7 shipped one and broke the New File lineage), so BOTH
    // standard variants must load exactly like their fully-literal
    // originals: legacy repairs reconstruct the template chains, count 0,
    // no warning, no DB-save block. Wreck protection lives at the root
    // (surgical write-back preserves caches - tst_excelsurgery gates).
    // DVE_READER_NO_COM forces the openpyxl fallback so the detector
    // actually runs (COM computes live and always reports 0).

    void strippedLineage_standardTemplate_repairsSilently()
    {
        verifyStrippedLineageLoadsLikeOriginal(QStringLiteral("format_e.xlsx"));
    }

    void strippedLineage_regimeTemplate_repairsSilently()
    {
        verifyStrippedLineageLoadsLikeOriginal(QStringLiteral("format_e_regime.xlsx"));
    }

private:
    QString makeStrippedVariant(const QString& fixture, const QString& outDir)
    {
        const QString script = QFINDTESTDATA("make_stripped_variant.py");
        if (script.isEmpty()) return QString();
        const QString out =
            outDir + QStringLiteral("/") + QFileInfo(fixture).fileName();
        QProcess p;
        p.setProcessChannelMode(QProcess::MergedChannels);
        p.start(QStringLiteral("python"), {script, fixture, out});
        if (!p.waitForStarted(10000)) return QString();
        p.waitForFinished(60000);
        return (p.exitCode() == 0) ? out : QString();
    }

    void verifyStrippedLineageLoadsLikeOriginal(const QString& fixtureName)
    {
        if (!pythonAvailable) QSKIP("Python not available");
        qputenv("DVE_READER_NO_COM", "1");
        const auto restore = qScopeGuard([] { qunsetenv("DVE_READER_NO_COM"); });

        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());
        const QString variant =
            makeStrippedVariant(testDataFile(fixtureName), tmp.path());
        QVERIFY2(!variant.isEmpty(), "stripped-variant generation failed");

        DVE::DataProcessor dpRef;
        const DVE::FileResult ref = dpRef.processFile(testDataFile(fixtureName));
        DVE::DataProcessor dp;
        const DVE::FileResult got = dp.processFile(variant);
        QCOMPARE(got.sheets.size(), ref.sheets.size());

        for (int sh = 0; sh < got.sheets.size(); ++sh) {
            const DVE::SheetResult& g = got.sheets[sh];
            const DVE::SheetResult& r = ref.sheets[sh];
            // Template lineage is NOT destroyed data: no warning, no DB block.
            QCOMPARE(g.strippedFormulaCells, 0);
            // The legacy repairs reconstruct the template's puff chain, so
            // the parse matches the fully-literal original exactly.
            QCOMPARE(g.samples.size(), r.samples.size());
            for (int s = 0; s < g.samples.size(); ++s) {
                QCOMPARE(g.samples[s].rows.size(), r.samples[s].rows.size());
                for (int i = 0; i < g.samples[s].rows.size(); ++i)
                    QCOMPARE(g.samples[s].rows[i].puffs,
                             r.samples[s].rows[i].puffs);
            }
        }
    }
};

QTEST_APPLESS_MAIN(tst_DataProcessor)
#include "tst_dataprocessor.moc"
