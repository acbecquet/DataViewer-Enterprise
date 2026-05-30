// -- DataViewer Enterprise -- DatabaseManager Integration Tests (Postgres) --
//
// This suite exercises the Postgres-backed DatabaseManager. It REQUIRES an
// ephemeral Postgres reachable via DVE_TEST_PG_CONN; otherwise every slot is
// skipped at initTestCase. Each slot starts from a wiped schema via init()
// so cross-test contamination (UNIQUE constraint violations on file_path or
// natural-key collisions on sensory_sessions) cannot occur.
//
// Coverage:
//   * open / close / settings / per-file CRUD
//   * multi-sheet / multi-sample / multi-row round-trip
//   * re-save (UPDATE) round-trip preserves id+version semantics
//   * sensory upsert + natural-key + layout_json persistence
//   * id+version population on save (by-ref) and load
//   * WriteResult::VersionMismatch (UPDATE with stale version)
//   * WriteResult::RowDeleted (UPDATE on missing row)
//   * WriteResult::UniqueViolation (INSERT collides on UNIQUE)
//
// Dropped vs. the previous SQLite suite:
//   * testDoubleOpen -- lock-file specific; Postgres has no equivalent.
//   * testSensoryMigrationCollapsesDuplicates -- one-shot Plan A migration,
//     not a runtime DatabaseManager concern.

#include <QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QJsonDocument>
#include <QJsonObject>
#include <QTemporaryDir>
#include "TestHelpers.h"
#include "DatabaseManager.h"
#include "OfflineSnapshot.h"
#include "PostgresConnection.h"
#include "ConfigLoader.h"
#include "IdentityManager.h"
#include "ReportData.h"
#include "SensoryData.h"
#include "DetailedSensoryData.h"

namespace {

// Build a DbConfig from DVE_TEST_PG_CONN (space-separated key=value pairs).
DVE::DbConfig pgConfig() {
    DVE::DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    if (env.isEmpty()) return c;
    for (const QString& part : env.split(' ', Qt::SkipEmptyParts)) {
        const QStringList kv = part.split('=');
        if (kv.size() != 2) continue;
        const QString k = kv[0], v = kv[1];
        if      (k == "host")     c.host = v;
        else if (k == "port")     c.port = v.toInt();
        else if (k == "dbname")   c.database = v;
        else if (k == "user")     c.user = v;
        else if (k == "password") c.password = v;
    }
    return c;
}

} // anonymous

class tst_DatabaseManager : public QObject
{
    Q_OBJECT

private:
    DVE::IdentityManager m_identity;

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

    // Build a minimal SensorySession with one sample, all metrics scored 5.0.
    DVE::SensorySession makeSensorySession(const QString& sessionName = "MyTest",
                                            const QString& tester      = "Charlie",
                                            const QString& date        = "2026-05-01")
    {
        DVE::SensorySession s;
        s.sessionName = sessionName;
        s.testTitle   = "Heater Comparison";
        s.testerName  = tester;
        s.assessorName= "Charlie";
        s.media       = "Sample A";
        s.date        = date;
        s.timestamp   = date + "T10:00:00Z";

        DVE::SensorySample sample;
        sample.name = "Sample 1";
        for (const QString& m : DVE::kSensoryMetrics) sample.scores[m] = 5.0;
        s.samples.append(sample);
        return s;
    }

    // Build a minimal DetailedSensorySession for offline-routing tests.
    // The natural key is (session_name, tester_name, date) per the UNIQUE
    // index on detailed_sensory_sessions in init.sql.
    DVE::DetailedSensorySession makeDetailedSensorySession(
        const QString& sessionName = "DetMyTest",
        const QString& tester      = "Charlie",
        const QString& date        = "2026-05-01")
    {
        DVE::DetailedSensorySession s;
        s.sessionName  = sessionName;
        s.testTitle    = "Detailed Heater Test";
        s.testerName   = tester;
        s.assessorName = "Charlie";
        s.media        = "Sample A";
        s.date         = date;
        s.timestamp    = date + "T10:00:00Z";

        DVE::DetailedSensorySample sample;
        sample.name = "Sample 1";
        // Don't bother scoring every metric -- serialization defaults
        // missing keys to 0.0, which is fine for write-side guard tests.
        s.samples.append(sample);
        return s;
    }

    // -- Helpers: open a fresh DatabaseManager bound to the test Postgres --
    bool openDb(DVE::DatabaseManager& db) {
        return db.open(pgConfig(), &m_identity);
    }

    // -- Helper: directly bump the row version of an id to simulate a
    //    concurrent writer. Uses a dedicated QSqlDatabase connection so the
    //    bump_version trigger fires exactly once per call. Returns the new
    //    version on success, -1 on failure.
    int bumpRowVersionOutOfBand(const QString& table, qint64 id) {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_oob_" + QString::number(id);
        int newVersion = -1;
        {
            QSqlDatabase oob = QSqlDatabase::addDatabase("QPSQL", cname);
            oob.setHostName(cfg.host);
            oob.setPort(cfg.port);
            oob.setDatabaseName(cfg.database);
            oob.setUserName(cfg.user);
            oob.setPassword(cfg.password);
            if (oob.open()) {
                QSqlQuery q(oob);
                // updated_by is NOT NULL; bump_version trigger sets version+1.
                q.prepare(QString("UPDATE %1 SET updated_by = 'oob-tester' "
                                  "WHERE id = ? RETURNING version").arg(table));
                q.addBindValue(static_cast<qlonglong>(id));
                if (q.exec() && q.next()) {
                    newVersion = q.value(0).toInt();
                }
                oob.close();
            }
        }
        QSqlDatabase::removeDatabase(cname);
        return newVersion;
    }

    // -- Helper: delete a row out-of-band so the test can verify RowDeleted.
    bool deleteRowOutOfBand(const QString& table, qint64 id) {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_del_" + QString::number(id);
        bool ok = false;
        {
            QSqlDatabase oob = QSqlDatabase::addDatabase("QPSQL", cname);
            oob.setHostName(cfg.host);
            oob.setPort(cfg.port);
            oob.setDatabaseName(cfg.database);
            oob.setUserName(cfg.user);
            oob.setPassword(cfg.password);
            if (oob.open()) {
                QSqlQuery q(oob);
                q.prepare(QString("DELETE FROM %1 WHERE id = ?").arg(table));
                q.addBindValue(static_cast<qlonglong>(id));
                ok = q.exec();
                oob.close();
            }
        }
        QSqlDatabase::removeDatabase(cname);
        return ok;
    }

private slots:
    void initTestCase()
    {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
            QSKIP("DVE_TEST_PG_CONN not set; Postgres integration tests skipped");
    }

    // Per-slot cleanup: wipe every table the suite touches. Uses a separate
    // connection so it never collides with whatever the test code is doing.
    void init()
    {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) return;
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_wipe";
        {
            QSqlDatabase wipe = QSqlDatabase::addDatabase("QPSQL", cname);
            wipe.setHostName(cfg.host);
            wipe.setPort(cfg.port);
            wipe.setDatabaseName(cfg.database);
            wipe.setUserName(cfg.user);
            wipe.setPassword(cfg.password);
            if (wipe.open()) {
                QSqlQuery q(wipe);
                // Order matters: children first (FK CASCADE handles most of
                // it but explicit is safer here). settings has no FK; wipe it
                // too so getSetting starts fresh per test.
                for (const QString& t : QStringList{
                         "data_rows", "images", "samples", "tests", "files",
                         "sensory_images", "sensory_sessions",
                         "detailed_sensory_images", "detailed_sensory_sessions",
                         "sensory_header_presets",
                         "settings"
                     }) {
                    q.exec("DELETE FROM " + t);
                }
                wipe.close();
            }
        }
        QSqlDatabase::removeDatabase(cname);
    }

    // -- Open / Close ---------------------------------------------------------
    void testOpenClose()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        QVERIFY(db.isOpen());
        db.close();
        QVERIFY(!db.isOpen());
    }

    // -- Settings round-trip --------------------------------------------------
    void testSettings()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        QVERIFY(db.setSetting("key1", "val1"));
        QCOMPARE(db.getSetting("key1"), QStringLiteral("val1"));

        QCOMPARE(db.getSetting("missing", "def"), QStringLiteral("def"));

        db.close();
    }

    // -- saveFile() bool shim + listFiles + loadFile (round-trip) ------------
    void testSaveLoadFile()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

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
        QVERIFY(!loaded.sheets[0].samples.isEmpty());
        QCOMPARE(loaded.sheets[0].samples[0].sampleID,
                 original.sheets[0].samples[0].sampleID);

        db.close();
    }

    // -- tryWriteFile() returns Success on first INSERT ----------------------
    void testTryWriteFile_returnsSuccessOnInsert()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("twf.xlsx", "/tmp/twf.xlsx");
        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);
        QCOMPARE(db.listFiles().size(), 1);

        db.close();
    }

    // -- Remove file ----------------------------------------------------------
    void testRemoveFile()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        QVERIFY(db.saveFile(makeFileResult()));
        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);

        QVERIFY(db.removeFile(files[0].id));
        QVERIFY(db.listFiles().isEmpty());

        db.close();
    }

    // -- Deduplication: keep most-recent-N + drop "unknown" template ---------
    void testDeduplicate()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        for (int i = 0; i < 5; ++i) {
            DVE::FileResult fr = makeFileResult(
                "duplicate.xlsx",
                QString("/tmp/path%1/duplicate.xlsx").arg(i));
            QVERIFY(db.saveFile(fr));
        }
        QCOMPARE(db.listFiles().size(), 5);

        const int removed = db.deduplicateFiles(2);
        QVERIFY(removed > 0);
        QVERIFY(db.listFiles().size() <= 2);

        db.close();
    }

    // -- Recent file paths ----------------------------------------------------
    void testRecentFiles()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

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

    // -- Multi-sheet / multi-sample / multi-row round-trip -------------------
    // Catches prepared-statement / binding-reuse / iteration bugs that the
    // single-row test cannot detect. Builds a 3-sheet x 4-sample x 6-row
    // file with deterministic, distinct values per row, saves it, loads it
    // back, and asserts every field matches.
    void testMultiSheetRoundTrip()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

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

    // -- Re-save (overwrite) round-trip: load id+version, mutate, save -------
    // Drives the UPDATE branch of tryWriteFile end-to-end. After each save
    // the id/version on the FileResult must point at the same row, and the
    // version must monotonically increase.
    void testReSaveRoundTrip()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("resave.xlsx", "/tmp/resave.xlsx");
        QVERIFY(db.saveFile(fr));

        // Re-load to populate fr.id and fr.version (the original save by-value
        // doesn't write them back).
        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);
        fr = db.loadFile(files[0].id);
        QVERIFY(fr.id > 0);
        QVERIFY(fr.version >= 1);

        int prevVer = fr.version;
        for (int iter = 0; iter < 4; ++iter) {
            fr.sheets[0].samples[0].sampleName = QString("Sample iter %1").arg(iter);
            fr.sheets[0].samples[0].rows[0].tpm = 1.0 + iter * 0.5;

            // Save via tryWriteFile so we can assert the WriteResult.
            QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);

            QCOMPARE(db.listFiles().size(), 1);  // still one row

            DVE::FileResult loaded = db.loadFile(fr.id);
            QVERIFY(!loaded.sheets.isEmpty());
            QCOMPARE(loaded.sheets[0].samples[0].sampleName,
                     QString("Sample iter %1").arg(iter));
            QCOMPARE(loaded.sheets[0].samples[0].rows[0].tpm,
                     1.0 + iter * 0.5);
            QVERIFY(loaded.version > prevVer);
            prevVer = loaded.version;
            fr.version = loaded.version;
        }

        db.close();
    }

    // -- TPM upsert: re-saving same fresh FileResult (id=-1) is rejected by
    //    UNIQUE on file_path -- bool saveFile thus returns false on the 2nd
    //    call. (The previous SQLite test relied on INSERT-OR-REPLACE; under
    //    optimistic concurrency that path is dead -- callers must load first.)
    void testTpmUpsertOnFreshStructIsUniqueViolation()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult();
        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);
        QCOMPARE(db.listFiles().size(), 1);

        // Second INSERT with id=-1 hits the UNIQUE index on file_path.
        DVE::FileResult fr2 = makeFileResult();
        QCOMPARE(db.tryWriteFile(fr2), DVE::WriteResult::UniqueViolation);
        // Original row is untouched.
        QCOMPARE(db.listFiles().size(), 1);
        db.close();
    }

    // -- Two files with the same fileName but different paths coexist --------
    void testTpmDifferentPathsCoexist()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        QVERIFY(db.saveFile(makeFileResult("test.xlsx", "/tmp/A/test.xlsx")));
        QVERIFY(db.saveFile(makeFileResult("test.xlsx", "/tmp/B/test.xlsx")));
        QVERIFY(db.saveFile(makeFileResult("test.xlsx", "/tmp/C/test.xlsx")));
        QCOMPARE(db.listFiles().size(), 3);
        db.close();
    }

    // -- Sensory: re-saving the same loaded session UPDATEs the row ----------
    void testSensoryUpsertNoDuplicates()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession sess = makeSensorySession();
        QVERIFY(db.saveSensorySession(sess));   // populates sess.id, sess.version
        QVERIFY(sess.id > 0);
        QCOMPARE(db.listSensoryRecords().size(), 1);

        for (int i = 0; i < 4; ++i) {
            sess.samples[0].scores["Overall Liking"] = 5.0 + i;
            QVERIFY(db.saveSensorySession(sess));
        }
        QCOMPARE(db.listSensoryRecords().size(), 1);

        const auto sessions = db.loadSensorySessions();
        QCOMPARE(sessions.size(), 1);
        QCOMPARE(sessions[0].samples[0].scores.value("Overall Liking"), 8.0);

        db.close();
    }

    // -- Different testers for the same test coexist (each is its own row) --
    void testSensoryDifferentTestersCoexist()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        auto makeSess = [&](const QString& tester) {
            return makeSensorySession("MyTest", tester, "2026-05-01");
        };

        QVERIFY(db.saveSensorySession(makeSess("Alice")));
        QVERIFY(db.saveSensorySession(makeSess("Bob")));
        QVERIFY(db.saveSensorySession(makeSess("Charlie")));

        QCOMPARE(db.listSensoryRecords().size(), 3);

        // Save Bob again -- must collide on (session_name, tester_name, date).
        // Under optimistic concurrency the fresh-struct path INSERTs and
        // therefore hits UniqueViolation; row count must stay at 3.
        QCOMPARE(db.tryWriteSensorySession(makeSess("Bob")),
                 DVE::WriteResult::UniqueViolation);
        QCOMPARE(db.listSensoryRecords().size(), 3);

        db.close();
    }

    // -- Same test+tester on different dates coexist (preserves history) ----
    void testSensoryDifferentDatesCoexist()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        QVERIFY(db.saveSensorySession(makeSensorySession("MyTest", "Charlie", "2026-05-01")));
        QVERIFY(db.saveSensorySession(makeSensorySession("MyTest", "Charlie", "2026-05-02")));
        QVERIFY(db.saveSensorySession(makeSensorySession("MyTest", "Charlie", "2026-05-03")));

        QCOMPARE(db.listSensoryRecords().size(), 3);
        db.close();
    }

    // -- UNIQUE index enforced by Postgres on the natural key ----------------
    void testSensoryUniqueIndexEnforced()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // Two fresh structs (id=-1) with the same natural key. First wins,
        // second is rejected -- and the row count stays at 1.
        DVE::SensorySession a = makeSensorySession("T", "C", "D");
        DVE::SensorySession b = makeSensorySession("T", "C", "D");
        QCOMPARE(db.tryWriteSensorySession(a), DVE::WriteResult::Success);
        QCOMPARE(db.tryWriteSensorySession(b), DVE::WriteResult::UniqueViolation);
        QCOMPARE(db.listSensoryRecords().size(), 1);
        db.close();
    }

    // -- Layout JSON: per-session persistence --------------------------------
    void testSensoryLayoutPersistence()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s = makeSensorySession("layout-test");
        QVERIFY(db.saveSensorySession(s));   // populates s.id
        QVERIFY(s.id > 0);
        const int id = s.id;

        QCOMPARE(db.loadSensoryLayout(id), QString());     // NULL -> empty

        const QString layoutJson = R"({"version":1,"mode":"sensory"})";
        QVERIFY(db.saveSensoryLayout(id, layoutJson));
        // Postgres stores JSONB as a normalized representation -- key order
        // and whitespace are not preserved on round-trip. Compare parsed
        // JSON objects instead of raw byte strings.
        const QJsonObject got = QJsonDocument::fromJson(
            db.loadSensoryLayout(id).toUtf8()).object();
        const QJsonObject want = QJsonDocument::fromJson(layoutJson.toUtf8()).object();
        QCOMPARE(got, want);

        db.close();
    }

    // -- Layout JSON: cumulative (settings-backed) persistence ---------------
    void testCumulativeLayoutPersistence()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        QCOMPARE(db.loadCumulativeLayout(), QString());
        QVERIFY(db.saveCumulativeLayout("{\"x\":1}"));
        QCOMPARE(db.loadCumulativeLayout(), QString("{\"x\":1}"));
        db.close();
    }

    // -- Regression: layout_json must survive session re-save ----------------
    void testSaveSensorySessionPreservesLayoutOnReSave()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s = makeSensorySession("layout-resave");
        QVERIFY(db.saveSensorySession(s));
        QVERIFY(s.id > 0);

        const QString layoutJson = R"({"version":1,"mode":"sensory"})";
        QVERIFY(db.saveSensoryLayout(s.id, layoutJson));
        // See testSensoryLayoutPersistence for the JSONB normalization story.
        const QJsonObject want = QJsonDocument::fromJson(layoutJson.toUtf8()).object();
        QCOMPARE(QJsonDocument::fromJson(db.loadSensoryLayout(s.id).toUtf8()).object(),
                 want);

        // Re-save -- layout_json must be preserved (saveSensorySession only
        // touches json_data, not layout_json).
        s = db.loadSensorySession(s.id);   // reload to refresh version
        QVERIFY(db.saveSensorySession(s));
        QCOMPARE(QJsonDocument::fromJson(db.loadSensoryLayout(s.id).toUtf8()).object(),
                 want);

        db.close();
    }

    // -- Phase 1A Task 4: id population on save/load -------------------------
    void testSaveSensorySessionByValueDoesNotChangeId() {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        DVE::SensorySession s = makeSensorySession("byval", "x", "2026-01-01");
        QCOMPARE(s.id, -1);
        QVERIFY(db.saveSensorySession(static_cast<const DVE::SensorySession&>(s)));
        // const-ref overload -- id should still be -1
        QCOMPARE(s.id, -1);
        db.close();
    }

    void testSaveSensorySessionByRefPopulatesId() {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        DVE::SensorySession s = makeSensorySession("byref", "y", "2026-01-02");
        QCOMPARE(s.id, -1);
        QCOMPARE(s.version, 0);
        QVERIFY(db.saveSensorySession(s));
        QVERIFY(s.id > 0);
        QVERIFY(s.version > 0);   // server-assigned version (>= 1 after insert)
        db.close();
    }

    void testLoadSensorySessionPopulatesId() {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        DVE::SensorySession s = makeSensorySession("load", "z", "2026-01-03");
        QVERIFY(db.saveSensorySession(s));
        const int id = s.id;
        QVERIFY(id > 0);

        DVE::SensorySession loaded = db.loadSensorySession(id);
        QCOMPARE(loaded.id, id);
        QVERIFY(loaded.version > 0);
        db.close();
    }

    // -- WriteResult coverage: VersionMismatch on file UPDATE ---------------
    // 1. Save + load file -> got id, version.
    // 2. Bump version out-of-band (simulating a concurrent writer).
    // 3. Try to save with the stale version -> VersionMismatch.
    void testTryWriteFile_VersionMismatch()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("vm.xlsx", "/tmp/vm.xlsx");
        QVERIFY(db.saveFile(fr));
        fr = db.loadFile(db.listFiles()[0].id);
        QVERIFY(fr.id > 0 && fr.version > 0);

        const int newVer = bumpRowVersionOutOfBand("files", fr.id);
        QVERIFY(newVer > fr.version);

        // fr.version is now stale -> VersionMismatch.
        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::VersionMismatch);
        db.close();
    }

    // -- WriteResult coverage: RowDeleted on file UPDATE --------------------
    void testTryWriteFile_RowDeleted()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("rd.xlsx", "/tmp/rd.xlsx");
        QVERIFY(db.saveFile(fr));
        fr = db.loadFile(db.listFiles()[0].id);
        QVERIFY(fr.id > 0 && fr.version > 0);

        QVERIFY(deleteRowOutOfBand("files", fr.id));

        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::RowDeleted);
        db.close();
    }

    // -- WriteResult coverage: UniqueViolation on file INSERT ----------------
    void testTryWriteFile_UniqueViolation()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult a = makeFileResult("a.xlsx", "/tmp/dup.xlsx");
        QCOMPARE(db.tryWriteFile(a), DVE::WriteResult::Success);

        // Fresh (id=-1) struct with the SAME file_path -> INSERT collides.
        DVE::FileResult b = makeFileResult("b.xlsx", "/tmp/dup.xlsx");
        QCOMPARE(db.tryWriteFile(b), DVE::WriteResult::UniqueViolation);
        QCOMPARE(db.listFiles().size(), 1);
        db.close();
    }

    // -- WriteResult coverage: VersionMismatch on sensory UPDATE -------------
    void testTryWriteSensorySession_VersionMismatch()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s = makeSensorySession("vm-sens", "C", "2026-05-01");
        QVERIFY(db.saveSensorySession(s));   // populates s.id, s.version
        QVERIFY(s.id > 0 && s.version > 0);

        const int newVer = bumpRowVersionOutOfBand("sensory_sessions", s.id);
        QVERIFY(newVer > s.version);

        // s.version is stale -> VersionMismatch.
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::VersionMismatch);
        db.close();
    }

    // -- WriteResult coverage: RowDeleted on sensory UPDATE ------------------
    void testTryWriteSensorySession_RowDeleted()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s = makeSensorySession("rd-sens", "C", "2026-05-01");
        QVERIFY(db.saveSensorySession(s));
        QVERIFY(s.id > 0);

        QVERIFY(deleteRowOutOfBand("sensory_sessions", s.id));

        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::RowDeleted);
        db.close();
    }

    // -- WriteResult coverage: UniqueViolation on sensory INSERT -------------
    void testTryWriteSensorySession_UniqueViolation()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // INSERT first session.
        DVE::SensorySession a = makeSensorySession("uv-sens", "C", "2026-05-01");
        QCOMPARE(db.tryWriteSensorySession(a), DVE::WriteResult::Success);

        // INSERT another fresh struct with the same natural key.
        DVE::SensorySession b = makeSensorySession("uv-sens", "C", "2026-05-01");
        QCOMPARE(db.tryWriteSensorySession(b), DVE::WriteResult::UniqueViolation);
        QCOMPARE(db.listSensoryRecords().size(), 1);
        db.close();
    }

    // -- loadFile round-trips id and version cleanly -------------------------
    void testLoadFile_RoundTripsIdAndVersion()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("rt.xlsx", "/tmp/rt.xlsx");
        QVERIFY(db.saveFile(fr));

        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);
        const int id = files[0].id;

        DVE::FileResult loaded = db.loadFile(id);
        QCOMPARE(loaded.id, id);
        QVERIFY(loaded.version >= 1);   // server-assigned, INSERT starts at 1

        // Capture the pre-save version because the mutable-ref overload of
        // tryWriteFile now writes the post-save version back into `loaded`.
        const int preVer = loaded.version;
        // Saving the loaded struct must succeed because version is current.
        QCOMPARE(db.tryWriteFile(loaded), DVE::WriteResult::Success);
        // Writeback: tryWriteFile's mutable overload should have bumped
        // loaded.version to match what the server actually has now.
        QVERIFY(loaded.version > preVer);
        DVE::FileResult after = db.loadFile(id);
        QVERIFY(after.version > preVer);          // bump_version trigger fired
        QCOMPARE(after.version, loaded.version);  // writeback matches DB
        db.close();
    }

    // ────────────────────────────────────────────────────────────────────────
    // Plan C T3: Offline routing
    // ────────────────────────────────────────────────────────────────────────
    //
    // DatabaseManager owns a soft "online" flag separate from the underlying
    // PostgresConnection. The tests below verify:
    //   1) open() flips online to true.
    //   2) setOnline() round-trips.
    //   3) Every write path returns OfflineReadOnly / false when offline.
    //   4) Reads route to OfflineSnapshot when offline + snapshot is set+open.
    //   5) Reads degrade gracefully (empty result + error) when offline +
    //      no snapshot.
    //   6) tryWriteDetailedSensorySession's mutable overload writes id+
    //      version back on Success (parity with FileResult / SensorySession).

    void testIsOnlineDefaultsTrueAfterOpen()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        QVERIFY(db.isOnline());
        db.close();
    }

    void testSetOnlineToggle()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        QVERIFY(db.isOnline());
        db.setOnline(false);
        QVERIFY(!db.isOnline());
        db.setOnline(true);
        QVERIFY(db.isOnline());
        db.close();
    }

    void testTryWriteFileReturnsOfflineReadOnlyWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        db.setOnline(false);

        DVE::FileResult fr = makeFileResult("off.xlsx", "/tmp/off.xlsx");
        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::OfflineReadOnly);
        QVERIFY(db.lastError().contains("offline", Qt::CaseInsensitive));

        db.close();
    }

    void testTryWriteSensorySessionReturnsOfflineReadOnlyWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        db.setOnline(false);

        DVE::SensorySession s = makeSensorySession("off-sens", "C", "2026-05-01");
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::OfflineReadOnly);
        QVERIFY(db.lastError().contains("offline", Qt::CaseInsensitive));

        db.close();
    }

    void testTryWriteDetailedSensorySessionConstRefReturnsOfflineReadOnlyWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        db.setOnline(false);

        const DVE::DetailedSensorySession s = makeDetailedSensorySession(
            "off-det-const", "C", "2026-05-01");
        QCOMPARE(db.tryWriteDetailedSensorySession(s),
                 DVE::WriteResult::OfflineReadOnly);
        QVERIFY(db.lastError().contains("offline", Qt::CaseInsensitive));

        db.close();
    }

    void testTryWriteDetailedSensorySessionMutableRefReturnsOfflineReadOnlyWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        db.setOnline(false);

        DVE::DetailedSensorySession s = makeDetailedSensorySession(
            "off-det-mut", "C", "2026-05-01");
        QCOMPARE(db.tryWriteDetailedSensorySession(s),
                 DVE::WriteResult::OfflineReadOnly);
        QVERIFY(db.lastError().contains("offline", Qt::CaseInsensitive));
        // id/version must NOT be written back when the call is refused.
        QCOMPARE(s.id, -1);
        QCOMPARE(s.version, 0);

        db.close();
    }

    void testSaveFileBoolShimReturnsFalseWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        db.setOnline(false);

        DVE::FileResult fr = makeFileResult("off-shim.xlsx", "/tmp/off-shim.xlsx");
        QVERIFY(!db.saveFile(fr));
        QVERIFY(db.lastError().contains("offline", Qt::CaseInsensitive));

        db.close();
    }

    void testRemoveFileReturnsFalseWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        db.setOnline(false);

        QVERIFY(!db.removeFile(123));
        QVERIFY(db.lastError().contains("offline", Qt::CaseInsensitive));

        db.close();
    }

    void testSetSettingReturnsFalseWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        db.setOnline(false);

        QVERIFY(!db.setSetting("k", "v"));
        QVERIFY(db.lastError().contains("offline", Qt::CaseInsensitive));

        db.close();
    }

    void testListFilesRoutesToSnapshotWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // Seed Postgres with two known files via the online path.
        DVE::FileResult a = makeFileResult("snap-a.xlsx", "/tmp/snap-a.xlsx");
        DVE::FileResult b = makeFileResult("snap-b.xlsx", "/tmp/snap-b.xlsx");
        QCOMPARE(db.tryWriteFile(a), DVE::WriteResult::Success);
        QCOMPARE(db.tryWriteFile(b), DVE::WriteResult::Success);
        QCOMPARE(db.listFiles().size(), 2);

        // Regenerate snapshot against the live Postgres, then open read-only.
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(tmp.path());

        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        pg.close();
        QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));

        db.setOfflineSnapshot(&snap);
        db.setOnline(false);

        const auto offlineList = db.listFiles();
        QCOMPARE(offlineList.size(), 2);
        // Spot-check one filename round-trips through the snapshot.
        QStringList names;
        for (const auto& r : offlineList) names << r.fileName;
        QVERIFY(names.contains("snap-a.xlsx"));
        QVERIFY(names.contains("snap-b.xlsx"));

        // Reset state for subsequent tests.
        db.setOfflineSnapshot(nullptr);
        db.setOnline(true);
        db.close();
    }

    void testLoadFileByPathRoutesToSnapshotWhenOffline()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // Seed Postgres with one known file.
        DVE::FileResult fr = makeFileResult("snap-lookup.xlsx", "/tmp/snap-lookup.xlsx");
        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);

        // Snapshot the live state.
        QTemporaryDir tmp;
        QVERIFY(tmp.isValid());
        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(tmp.path());

        DVE::PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
        pg.close();
        QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));

        db.setOfflineSnapshot(&snap);
        db.setOnline(false);

        const DVE::FileResult loaded = db.loadFileByPath("/tmp/snap-lookup.xlsx");
        QVERIFY(!loaded.filePath.isEmpty());
        QCOMPARE(loaded.filePath, QStringLiteral("/tmp/snap-lookup.xlsx"));
        QCOMPARE(loaded.fileName, QStringLiteral("snap-lookup.xlsx"));
        QCOMPARE(loaded.templateVersion, QStringLiteral("new"));

        // Reset state for subsequent tests.
        db.setOfflineSnapshot(nullptr);
        db.setOnline(true);
        db.close();
    }

    void testListFilesReturnsEmptyWhenOfflineAndNoSnapshot()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));
        // No setOfflineSnapshot -- the pointer stays null.
        db.setOnline(false);

        const auto records = db.listFiles();
        QVERIFY(records.isEmpty());
        QVERIFY(db.lastError().contains("snapshot", Qt::CaseInsensitive));

        db.setOnline(true);
        db.close();
    }

    void testTryWriteDetailedSensorySessionMutableRefWritesBackIdVersion()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::DetailedSensorySession s = makeDetailedSensorySession(
            "writeback-det", "C", "2026-05-01");
        QCOMPARE(s.id, -1);
        QCOMPARE(s.version, 0);

        QCOMPARE(db.tryWriteDetailedSensorySession(s),
                 DVE::WriteResult::Success);

        // Mutable-ref overload must populate id+version on Success
        // (parity with tryWriteFile / tryWriteSensorySession).
        QVERIFY(s.id > 0);
        QVERIFY(s.version >= 1);

        db.close();
    }

    // v2.2.1 Task 3: per-row puffing_regime column on data_rows ------------------
    // Verifies:
    //   * new-template: regime strings stored and loaded back correctly;
    //     hasPerRowRegime derived true from non-NULL DB rows.
    //   * old-template: hasPerRowRegime = false, regime stays empty,
    //     other fields (resistance) round-trip unchanged.
    void perRowRegime_roundTrip_newAndOld()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // --- new-template file with two rows each carrying a regime string ---
        DVE::FileResult fNew;
        fNew.filePath = "regime_new.xlsx";
        fNew.fileName = "regime_new.xlsx";
        DVE::SheetResult shNew;
        shNew.sheetName = "Lifetime Test";
        shNew.hasPerRowRegime = true;
        DVE::SampleResult smNew;
        smNew.sampleName = "S1";
        DVE::DataRow a; a.puffs = 10; a.beforeWeight = 25.1; a.afterWeight = 25.06; a.puffingRegime = "60mL/3s/30s";
        DVE::DataRow b; b.puffs = 20; b.beforeWeight = 25.06; b.afterWeight = 25.02; b.puffingRegime = "200mL/9s/300s";
        smNew.rows << a << b;
        shNew.samples << smNew;
        fNew.sheets << shNew;
        QVERIFY(db.saveFile(fNew));

        DVE::FileResult loaded = db.loadFileByPath("regime_new.xlsx");
        QCOMPARE(loaded.sheets.size(), 1);
        QCOMPARE(loaded.sheets[0].hasPerRowRegime, true);
        QCOMPARE(loaded.sheets[0].samples[0].rows[0].puffingRegime, QString("60mL/3s/30s"));
        QCOMPARE(loaded.sheets[0].samples[0].rows[1].puffingRegime, QString("200mL/9s/300s"));

        // --- old-template file: no regime, hasPerRowRegime must stay false ---
        DVE::FileResult fOld;
        fOld.filePath = "regime_old.xlsx";
        fOld.fileName = "regime_old.xlsx";
        DVE::SheetResult shOld;
        shOld.sheetName = "Lifetime Test";
        shOld.hasPerRowRegime = false;
        DVE::SampleResult smOld;
        smOld.sampleName = "S1";
        DVE::DataRow c; c.puffs = 10; c.beforeWeight = 25.1; c.afterWeight = 25.06; c.resistance = 1.25;
        smOld.rows << c;
        shOld.samples << smOld;
        fOld.sheets << shOld;
        QVERIFY(db.saveFile(fOld));

        DVE::FileResult loadedOld = db.loadFileByPath("regime_old.xlsx");
        QCOMPARE(loadedOld.sheets[0].hasPerRowRegime, false);
        QVERIFY(loadedOld.sheets[0].samples[0].rows[0].puffingRegime.isEmpty());
        QCOMPARE(loadedOld.sheets[0].samples[0].rows[0].resistance, 1.25);

        db.close();
    }

    // v2.0.4: round-trip the sensory header presets and confirm
    //  - saveSensoryHeaderPresets inserts each non-empty value
    //  - empty/whitespace values are silently skipped (not stored)
    //  - re-saving the same trio is idempotent (no UNIQUE violation)
    //  - loadSensoryHeaderPresets returns the right kind, alphabetised,
    //    and is empty for a kind that has no entries.
    void sensoryHeaderPresets_roundTrip()
    {
        DVE::DatabaseManager db;
        QVERIFY(db.open(pgConfig(), &m_identity));

        QVERIFY(db.saveSensoryHeaderPresets(
            "Beta Panel 12",
            "VG/PG 70/30",
            { "Sample A", "Sample B", "  ", "Sample C", "" }));

        // Second save: same names + a new media. Idempotent on dups,
        // additive on new entries.
        QVERIFY(db.saveSensoryHeaderPresets(
            "Beta Panel 12",
            "VG/PG 50/50",
            { "Sample A", "Sample D" }));

        const QStringList tests   = db.loadSensoryHeaderPresets("test_name");
        const QStringList medias  = db.loadSensoryHeaderPresets("media");
        const QStringList samples = db.loadSensoryHeaderPresets("sample_name");

        QCOMPARE(tests,  QStringList{ "Beta Panel 12" });
        QCOMPARE(medias, (QStringList{ "VG/PG 50/50", "VG/PG 70/30" }));
        QCOMPARE(samples,
                 (QStringList{ "Sample A", "Sample B", "Sample C", "Sample D" }));

        // Unknown kind → empty list, no error to the caller.
        QCOMPARE(db.loadSensoryHeaderPresets("not_a_kind"), QStringList{});

        db.close();
    }
};

QTEST_MAIN(tst_DatabaseManager)
#include "tst_databasemanager.moc"
