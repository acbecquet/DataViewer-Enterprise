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
