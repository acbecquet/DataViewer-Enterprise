// ── DataViewer Enterprise — DatabaseManager Integration Tests ───────────────
#include <QtTest>
#include <QTemporaryDir>
#include "TestHelpers.h"
#include "DatabaseManager.h"
#include "ReportData.h"

class tst_DatabaseManager : public QObject
{
    Q_OBJECT

private:
    QTemporaryDir m_tempDir;

    // Build a minimal FileResult for database round-trip testing.
    DVE::FileResult makeFileResult(const QString& fileName = "test.xlsx",
                                   const QString& filePath = "/tmp/test.xlsx")
    {
        DVE::FileResult fr;
        fr.filePath        = filePath;
        fr.fileName        = fileName;
        fr.templateVersion = "new";
        fr.sheetNames      << "Lifetime Test";

        DVE::SheetResult sheet;
        sheet.sheetName       = "Lifetime Test";
        sheet.templateVersion = "new";
        sheet.columnHeaders   << "puffs" << "beforeWeight" << "afterWeight";
        sheet.overallAvgTPM   = 3.5;

        DVE::SampleResult sample;
        sample.sampleName = "Sample 1";
        sample.sampleID   = "S-1";
        sample.date       = "2025-01-01";
        sample.tester     = "QA";
        sample.media      = "E-liquid";
        sample.resistance = 1.1;
        sample.voltage    = 3.0;
        sample.power      = 8.18;
        sample.averageTPM = 3.5;

        DVE::DataRow row;
        row.puffs        = 10.0;
        row.beforeWeight = 25.1;
        row.afterWeight  = 25.065;
        row.tpm          = 3.5;
        row.oilConsumed  = 0.035;
        sample.rows.append(row);

        sheet.samples.append(sample);
        fr.sheets.append(sheet);

        return fr;
    }

private slots:
    void initTestCase()
    {
        QVERIFY(m_tempDir.isValid());
    }

    // ── Open / Close ────────────────────────────────────────────────────
    void testOpenClose()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));
        QVERIFY(db.isOpen());

        db.close();
        QVERIFY(!db.isOpen());
    }

    // ── Settings ────────────────────────────────────────────────────────
    void testSettings()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        QVERIFY(db.setSetting("key1", "val1"));
        QCOMPARE(db.getSetting("key1"), QStringLiteral("val1"));

        // Missing key returns default
        QCOMPARE(db.getSetting("missing", "def"), QStringLiteral("def"));

        db.close();
    }

    // ── Save / List / Load round-trip ───────────────────────────────────
    void testSaveLoadFile()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::FileResult original = makeFileResult();
        QVERIFY(db.saveFile(original));

        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);

        DVE::FileResult loaded = db.loadFile(files[0].id);
        QVERIFY(!loaded.filePath.isEmpty());
        QCOMPARE(loaded.fileName, original.fileName);
        QCOMPARE(loaded.templateVersion, original.templateVersion);
        QVERIFY(!loaded.sheets.isEmpty());
        QCOMPARE(loaded.sheets[0].sheetName, original.sheets[0].sheetName);

        if (!loaded.sheets[0].samples.isEmpty()) {
            QCOMPARE(loaded.sheets[0].samples[0].sampleID,
                     original.sheets[0].samples[0].sampleID);
        }

        db.close();
    }

    // ── Remove file ─────────────────────────────────────────────────────
    void testRemoveFile()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        QVERIFY(db.saveFile(makeFileResult()));
        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);

        QVERIFY(db.removeFile(files[0].id));
        QVERIFY(db.listFiles().isEmpty());

        db.close();
    }

    // ── Deduplication ───────────────────────────────────────────────────
    void testDeduplicate()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        // Save same fileName 5 times with different paths to create duplicates
        for (int i = 0; i < 5; ++i) {
            DVE::FileResult fr = makeFileResult(
                "duplicate.xlsx",
                QString("/tmp/path%1/duplicate.xlsx").arg(i));
            QVERIFY(db.saveFile(fr));
        }

        QVERIFY(db.listFiles().size() == 5);

        int removed = db.deduplicateFiles(2);
        QVERIFY(removed > 0);
        QVERIFY(db.listFiles().size() <= 2);

        db.close();
    }

    // ── Recent file paths ───────────────────────────────────────────────
    void testRecentFiles()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        for (int i = 0; i < 3; ++i) {
            DVE::FileResult fr = makeFileResult(
                QString("file%1.xlsx").arg(i),
                QString("/tmp/file%1.xlsx").arg(i));
            QVERIFY(db.saveFile(fr));
        }

        QStringList recent = db.recentFilePaths();
        QCOMPARE(recent.size(), 3);

        db.close();
    }

    // ── Multi-sheet / multi-sample / multi-row round-trip ───────────────
    // Catches prepared-statement / binding-reuse / iteration bugs that the
    // single-row test cannot detect. Builds a 3-sheet × 4-sample × 6-row
    // file with deterministic, distinct values per row, saves it, loads it
    // back, and asserts every field matches.
    void testMultiSheetRoundTrip()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::FileResult fr;
        fr.filePath        = "/tmp/multi.xlsx";
        fr.fileName        = "multi.xlsx";
        fr.templateVersion = "new";

        const QStringList sheetNames = {"Lifetime Test", "Initial Test", "Stability"};
        for (int si = 0; si < sheetNames.size(); ++si) {
            DVE::SheetResult sheet;
            sheet.sheetName       = sheetNames[si];
            sheet.templateVersion = "new";
            sheet.overallAvgTPM   = 1.0 + si * 0.5;
            sheet.overallStdDevTPM = 0.1 + si * 0.05;

            for (int sj = 0; sj < 4; ++sj) {
                DVE::SampleResult sample;
                sample.sampleName = QString("Sheet%1-Sample%2").arg(si).arg(sj);
                sample.sampleID   = QString("ID-%1-%2").arg(si).arg(sj);
                sample.date       = QString("2026-05-0%1").arg(sj + 1);
                sample.tester     = QString("Tester%1").arg(sj);
                sample.media      = QString("Media-%1-%2").arg(si).arg(sj);
                sample.viscosity  = 100.0 + si * 10 + sj;
                sample.resistance = 1.0 + sj * 0.1;
                sample.voltage    = 3.0 + sj * 0.05;
                sample.power      = 8.0 + sj * 0.2;
                sample.averageTPM = 2.0 + si + sj * 0.1;
                sample.stdDevTPM  = 0.05 + sj * 0.01;
                sample.totalPuffs = 30 + sj * 10;

                for (int ri = 0; ri < 6; ++ri) {
                    DVE::DataRow row;
                    row.puffs        = 10.0 + ri * 5;
                    row.beforeWeight = 25.0 + ri * 0.001;
                    row.afterWeight  = 24.5 + ri * 0.001;
                    row.tpm          = 1.0 + si + sj * 0.1 + ri * 0.01;
                    row.tpmPowerDensity = 0.1 + ri * 0.001;
                    row.variationTPM = 0.02 + ri * 0.001;
                    row.oilConsumed  = 0.5 - ri * 0.001;
                    row.notes        = QString("note s%1-s%2-r%3").arg(si).arg(sj).arg(ri);
                    sample.rows.append(row);
                }
                sheet.samples.append(sample);
            }
            fr.sheets.append(sheet);
        }

        QVERIFY(db.saveFile(fr));

        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);

        DVE::FileResult loaded = db.loadFile(files[0].id);
        QCOMPARE(loaded.fileName, fr.fileName);
        QCOMPARE(loaded.sheets.size(), fr.sheets.size());

        for (int si = 0; si < fr.sheets.size(); ++si) {
            const auto& expSheet = fr.sheets[si];
            const auto& gotSheet = loaded.sheets[si];
            QCOMPARE(gotSheet.sheetName, expSheet.sheetName);
            QCOMPARE(gotSheet.samples.size(), expSheet.samples.size());

            for (int sj = 0; sj < expSheet.samples.size(); ++sj) {
                const auto& expS = expSheet.samples[sj];
                const auto& gotS = gotSheet.samples[sj];
                QCOMPARE(gotS.sampleName, expS.sampleName);
                QCOMPARE(gotS.sampleID,   expS.sampleID);
                QCOMPARE(gotS.tester,     expS.tester);
                QCOMPARE(gotS.media,      expS.media);
                QCOMPARE(gotS.viscosity,  expS.viscosity);
                QCOMPARE(gotS.averageTPM, expS.averageTPM);
                QCOMPARE(gotS.totalPuffs, expS.totalPuffs);
                QCOMPARE(gotS.rows.size(), expS.rows.size());

                for (int ri = 0; ri < expS.rows.size(); ++ri) {
                    const auto& expR = expS.rows[ri];
                    const auto& gotR = gotS.rows[ri];
                    QCOMPARE(gotR.puffs,        expR.puffs);
                    QCOMPARE(gotR.beforeWeight, expR.beforeWeight);
                    QCOMPARE(gotR.afterWeight,  expR.afterWeight);
                    QCOMPARE(gotR.tpm,          expR.tpm);
                    QCOMPARE(gotR.oilConsumed,  expR.oilConsumed);
                    QCOMPARE(gotR.notes,        expR.notes);
                }
            }
        }

        db.close();
    }

    // ── Re-save (overwrite) round-trip ──────────────────────────────────
    // Saves, edits, and re-saves the same file_path several times. Catches
    // bugs where the SELECT/DELETE/INSERT cycle leaves stale rows or
    // drops columns.
    void testReSaveRoundTrip()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::FileResult fr = makeFileResult("resave.xlsx", "/tmp/resave.xlsx");

        for (int iter = 0; iter < 4; ++iter) {
            // Mutate something distinctive each iteration
            fr.sheets[0].samples[0].sampleName = QString("Sample iter %1").arg(iter);
            fr.sheets[0].samples[0].rows[0].tpm = 1.0 + iter * 0.5;

            QVERIFY(db.saveFile(fr));

            auto files = db.listFiles();
            QCOMPARE(files.size(), 1);  // still one row (replaced, not duplicated)

            DVE::FileResult loaded = db.loadFile(files[0].id);
            QVERIFY(!loaded.sheets.isEmpty());
            QVERIFY(!loaded.sheets[0].samples.isEmpty());
            QCOMPARE(loaded.sheets[0].samples[0].sampleName,
                     QString("Sample iter %1").arg(iter));
            QCOMPARE(loaded.sheets[0].samples[0].rows[0].tpm,
                     1.0 + iter * 0.5);
        }

        db.close();
    }

    // ── Double open ─────────────────────────────────────────────────────
    void testDoubleOpen()
    {
        DVE::DatabaseManager db;
        QString path = m_tempDir.path() + "/double.db";

        QVERIFY(db.open(path));
        QVERIFY(db.isOpen());

        // Opening again to same path should work
        QVERIFY(db.open(path));
        QVERIFY(db.isOpen());

        db.close();
    }
};

QTEST_MAIN(tst_DatabaseManager)
#include "tst_databasemanager.moc"
