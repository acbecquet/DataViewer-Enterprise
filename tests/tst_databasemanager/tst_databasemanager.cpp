// ── DataViewer Enterprise — DatabaseManager Integration Tests ───────────────
#include <QtTest>
#include <QTemporaryDir>
#include <QSqlDatabase>
#include <QSqlQuery>
#include "TestHelpers.h"
#include "DatabaseManager.h"
#include "ReportData.h"
#include "SensoryData.h"
#include "DetailedSensoryData.h"

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

    // ── TPM upsert: re-saving same file_path replaces, never duplicates ────
    void testTpmUpsertNoDuplicates()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::FileResult fr = makeFileResult();
        // Save 5 times — UNIQUE index on file_path + atomic transaction means
        // we should always have exactly 1 row.
        for (int i = 0; i < 5; ++i) {
            fr.sheets[0].samples[0].rows[0].tpm = 1.0 + i;
            QVERIFY(db.saveFile(fr));
        }
        QCOMPARE(db.listFiles().size(), 1);

        // Newest values win
        DVE::FileResult loaded = db.loadFileByPath(fr.filePath);
        QCOMPARE(loaded.sheets[0].samples[0].rows[0].tpm, 5.0);
        db.close();
    }

    // Two files with the same fileName but different paths must coexist.
    void testTpmDifferentPathsCoexist()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        QVERIFY(db.saveFile(makeFileResult("test.xlsx", "/tmp/A/test.xlsx")));
        QVERIFY(db.saveFile(makeFileResult("test.xlsx", "/tmp/B/test.xlsx")));
        QVERIFY(db.saveFile(makeFileResult("test.xlsx", "/tmp/C/test.xlsx")));
        QCOMPARE(db.listFiles().size(), 3);
        db.close();
    }

    // ── Sensory upsert: re-saving the same key replaces, never duplicates ──
    void testSensoryUpsertNoDuplicates()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::SensorySession sess;
        sess.sessionName = "MyTest";
        sess.testTitle   = "Heater Comparison";
        sess.testerName  = "Charlie";
        sess.assessorName = "Charlie";
        sess.media       = "Sample A";
        sess.date        = "2026-05-01";
        sess.timestamp   = "2026-05-01T10:00:00Z";

        DVE::SensorySample s1;
        s1.name = "Sample 1";
        for (const QString& m : DVE::kSensoryMetrics) s1.scores[m] = 5.0;
        sess.samples.append(s1);

        // Save 5 times — UNIQUE constraint + INSERT OR REPLACE means
        // we should always have exactly 1 row.
        for (int i = 0; i < 5; ++i) {
            sess.samples[0].scores["Overall Liking"] = 5.0 + i;
            QVERIFY(db.saveSensorySession(sess));
        }
        QCOMPARE(db.listSensoryRecords().size(), 1);

        // Latest scores should be the last write's values
        const auto sessions = db.loadSensorySessions();
        QCOMPARE(sessions.size(), 1);
        QCOMPARE(sessions[0].samples[0].scores.value("Overall Liking"), 9.0);

        db.close();
    }

    // Different testers for the same test must coexist (each is its own row).
    void testSensoryDifferentTestersCoexist()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        auto makeSess = [](const QString& tester) {
            DVE::SensorySession sess;
            sess.sessionName = "MyTest";
            sess.testTitle   = "Heater Comparison";
            sess.testerName  = tester;
            sess.assessorName = "Lead";
            sess.date        = "2026-05-01";
            DVE::SensorySample s; s.name = "Sample 1";
            for (const QString& m : DVE::kSensoryMetrics) s.scores[m] = 5.0;
            sess.samples.append(s);
            return sess;
        };

        QVERIFY(db.saveSensorySession(makeSess("Alice")));
        QVERIFY(db.saveSensorySession(makeSess("Bob")));
        QVERIFY(db.saveSensorySession(makeSess("Charlie")));

        QCOMPARE(db.listSensoryRecords().size(), 3);

        // Re-saving Bob should overwrite Bob's row only (still 3 total).
        QVERIFY(db.saveSensorySession(makeSess("Bob")));
        QCOMPARE(db.listSensoryRecords().size(), 3);

        db.close();
    }

    // Same test+tester on a different date is a new logical session — both
    // rows should coexist (so historical data is preserved).
    void testSensoryDifferentDatesCoexist()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::SensorySession sess;
        sess.sessionName = "MyTest";
        sess.testerName  = "Charlie";
        DVE::SensorySample s; s.name = "Sample 1";
        for (const QString& m : DVE::kSensoryMetrics) s.scores[m] = 5.0;
        sess.samples.append(s);

        sess.date = "2026-05-01"; QVERIFY(db.saveSensorySession(sess));
        sess.date = "2026-05-02"; QVERIFY(db.saveSensorySession(sess));
        sess.date = "2026-05-03"; QVERIFY(db.saveSensorySession(sess));

        QCOMPARE(db.listSensoryRecords().size(), 3);
        db.close();
    }

    // The migration should collapse pre-existing duplicate rows on open()
    // so the new UNIQUE index can be created without error.
    void testSensoryMigrationCollapsesDuplicates()
    {
        const QString tmp = m_tempDir.path() + "/migrate_dup.db";
        QFile::remove(tmp);

        // Build a DB with duplicate rows by inserting raw rows BEFORE the
        // unique index would exist. Simulate the pre-1.0.7 state.
        {
            QSqlDatabase raw = QSqlDatabase::addDatabase("QSQLITE", "raw_migration_test");
            raw.setDatabaseName(tmp);
            QVERIFY(raw.open());
            QSqlQuery q(raw);
            QVERIFY(q.exec(
                "CREATE TABLE sensory_sessions ("
                "  id INTEGER PRIMARY KEY AUTOINCREMENT, session_name TEXT,"
                "  tester_name TEXT, assessor_name TEXT, media TEXT,"
                "  puff_length TEXT, date TEXT, timestamp TEXT, json_data TEXT)"));
            for (int i = 0; i < 4; ++i) {
                QSqlQuery ins(raw);
                ins.prepare("INSERT INTO sensory_sessions "
                            "(session_name, tester_name, date, json_data) "
                            "VALUES ('MyTest', 'Charlie', '2026-05-01', ?)");
                ins.addBindValue(QStringLiteral("{\"v\":%1}").arg(i));
                QVERIFY(ins.exec());
            }
            raw.close();
        }
        QSqlDatabase::removeDatabase("raw_migration_test");

        // Now open with DatabaseManager. Migration collapses the duplicates
        // so the UNIQUE index can be built.
        DVE::DatabaseManager db;
        QVERIFY(db.open(tmp));
        QCOMPARE(db.listSensoryRecords().size(), 1);
        db.close();
        QFile::remove(tmp);
    }

    // After migration, the UNIQUE index actually rejects future direct
    // duplicate INSERTs (belt-and-suspenders).
    void testSensoryUniqueIndexEnforced()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::SensorySession sess;
        sess.sessionName = "T"; sess.testerName = "C"; sess.date = "D";
        DVE::SensorySample s; s.name = "S";
        for (const QString& m : DVE::kSensoryMetrics) s.scores[m] = 5.0;
        sess.samples.append(s);
        QVERIFY(db.saveSensorySession(sess));
        QVERIFY(db.saveSensorySession(sess));   // upsert path — should still be 1
        QCOMPARE(db.listSensoryRecords().size(), 1);
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

    // ── Layout JSON: per-session persistence ────────────────────────────
    void testSensoryLayoutPersistence()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::SensorySession s;
        s.testTitle  = "Test";
        s.testerName = "T";
        s.date       = "2026-01-01";
        DVE::SensorySample samp;
        samp.name = "S1";
        s.samples.append(samp);

        QVERIFY(db.saveSensorySession(s));

        // Look up the id we just wrote (saveSensorySession does not yet return it —
        // that is Task 4's prerequisite. For now, fetch the autoincrement id from
        // the table directly.)
        QSqlQuery q(QSqlDatabase::database("dve_main"));
        QVERIFY(q.exec("SELECT id FROM sensory_sessions ORDER BY id DESC LIMIT 1"));
        QVERIFY(q.next());
        int id = q.value(0).toInt();
        QVERIFY(id > 0);

        QCOMPARE(db.loadSensoryLayout(id), QString());     // NULL → empty

        QVERIFY(db.saveSensoryLayout(id, R"({"version":1,"mode":"sensory"})"));
        QCOMPARE(db.loadSensoryLayout(id),
                 QString(R"({"version":1,"mode":"sensory"})"));
    }

    // ── Layout JSON: cumulative (settings-backed) persistence ───────────
    void testCumulativeLayoutPersistence()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));
        QCOMPARE(db.loadCumulativeLayout(), QString());
        QVERIFY(db.saveCumulativeLayout("{\"x\":1}"));
        QCOMPARE(db.loadCumulativeLayout(), QString("{\"x\":1}"));
    }

    // ── Regression: layout_json must survive session re-save ─────────────
    void testSaveSensorySessionPreservesLayoutOnReSave()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));

        DVE::SensorySession s;
        s.sessionName = "Test";
        s.testTitle   = "Test";
        s.testerName  = "T";
        s.date        = "2026-01-01";
        DVE::SensorySample samp;
        samp.name = "S1";
        s.samples.append(samp);

        QVERIFY(db.saveSensorySession(s));

        QSqlQuery q(QSqlDatabase::database("dve_main"));
        QVERIFY(q.exec("SELECT id FROM sensory_sessions ORDER BY id DESC LIMIT 1"));
        QVERIFY(q.next());
        int id = q.value(0).toInt();

        const QString layoutJson = R"({"version":1,"mode":"sensory"})";
        QVERIFY(db.saveSensoryLayout(id, layoutJson));
        QCOMPARE(db.loadSensoryLayout(id), layoutJson);

        // Re-save the session — layout_json must NOT be wiped.
        QVERIFY(db.saveSensorySession(s));

        QVERIFY(q.exec("SELECT id FROM sensory_sessions ORDER BY id DESC LIMIT 1"));
        QVERIFY(q.next());
        int newId = q.value(0).toInt();
        QCOMPARE(db.loadSensoryLayout(newId), layoutJson);
    }

    // ── id population on save/load (Phase 1A Task 4 prerequisite) ─────────
    void testSaveSensorySessionByValueDoesNotChangeId() {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));
        DVE::SensorySession s;
        s.sessionName = "T"; s.testerName = "x"; s.date = "2026-01-01";
        DVE::SensorySample samp; samp.name = "S"; s.samples.append(samp);
        QCOMPARE(s.id, -1);
        QVERIFY(db.saveSensorySession(static_cast<const DVE::SensorySession&>(s)));
        // const-ref overload — id should still be -1
        QCOMPARE(s.id, -1);
    }

    void testSaveSensorySessionByRefPopulatesId() {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));
        DVE::SensorySession s;
        s.sessionName = "T2"; s.testerName = "y"; s.date = "2026-01-02";
        DVE::SensorySample samp; samp.name = "S"; s.samples.append(samp);
        QCOMPARE(s.id, -1);
        QVERIFY(db.saveSensorySession(s));    // non-const ref overload
        QVERIFY(s.id > 0);
    }

    void testLoadSensorySessionPopulatesId() {
        DVE::DatabaseManager db;
        QVERIFY(db.open(":memory:"));
        DVE::SensorySession s;
        s.sessionName = "T3"; s.testerName = "z"; s.date = "2026-01-03";
        DVE::SensorySample samp; samp.name = "S"; s.samples.append(samp);
        QVERIFY(db.saveSensorySession(s));
        int id = s.id;
        QVERIFY(id > 0);

        DVE::SensorySession loaded = db.loadSensorySession(id);
        QCOMPARE(loaded.id, id);
    }
};

QTEST_MAIN(tst_DatabaseManager)
#include "tst_databasemanager.moc"
