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
//   * fresh-version OCC: a STALE in-memory version still saves (RC1, v2.4.0
//     data-loss regression) -- the write cores re-read the current committed
//     version inside the save transaction (SELECT ... FOR UPDATE) and bind
//     THAT in the OCC WHERE clause, so a whole-row save is row-level
//     last-writer-wins by design. With the row locked and the fresh version
//     bound, a genuine WriteResult::VersionMismatch from the write path is
//     intentionally no longer reachable -- the `AND version = ?` clause is a
//     defensive invariant only -- so this suite has no VersionMismatch
//     assertion. Coverage for RowDeleted and UniqueViolation remains below.
//   * WriteResult::RowDeleted (UPDATE on missing row)
//   * WriteResult::UniqueViolation (INSERT collides on UNIQUE) + the shared
//     connection stays usable afterwards (rollback hygiene, RC3)
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
#include <QImage>
#include "TestHelpers.h"
#include "DatabaseManager.h"
#include "DbRepair.h"
#include "DataProcessor.h"
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

    // -- Helper: simulate a LiveSync per-cell score edit out-of-band. Uses
    //    jsonb_set to overwrite ONE metric on samples[sampleIdx] of the given
    //    sensory_sessions row, exactly as the live-sync path does. The UPDATE
    //    fires the bump_version trigger, so the row version advances; the new
    //    version is returned (or -1 on failure). Mirrors the connection
    //    lifecycle of bumpRowVersionOutOfBand.
    int setSensoryScoreOutOfBand(qint64 id, int sampleIdx,
                                 const QString& metric, double value) {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_score_" + QString::number(id);
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
                // jsonb_set(target, '{samples,<idx>,<metric>}', '<value>'::jsonb).
                // The path is bound as a text[] literal ('{samples,0,Smoothness}');
                // metric names contain spaces ("Burnt Taste"), which the array
                // literal tolerates unquoted here since they have no commas or
                // braces. value is bound, not interpolated.
                const QString path = QString("{samples,%1,%2}")
                                         .arg(sampleIdx).arg(metric);
                q.prepare("UPDATE sensory_sessions "
                          "SET json_data = jsonb_set(json_data, ?::text[], "
                          "                           to_jsonb(?::double precision), true), "
                          "    updated_by = 'livesync-sim' "
                          "WHERE id = ? RETURNING version");
                q.addBindValue(path);
                q.addBindValue(value);
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

    // -- Helper: detailed-sensory analogue of setSensoryScoreOutOfBand. Uses
    //    jsonb_set to overwrite ONE metric on samples[sampleIdx] of the given
    //    detailed_sensory_sessions row, exactly as the live-sync path does. The
    //    UPDATE fires the bump_version trigger, so the row version advances; the
    //    new version is returned (or -1 on failure). Mirrors the connection
    //    lifecycle of setSensoryScoreOutOfBand.
    int setDetailedScoreOutOfBand(qint64 id, int sampleIdx,
                                  const QString& metric, double value) {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_detscore_" + QString::number(id);
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
                // jsonb_set(target, '{samples,<idx>,<metric>}', '<value>'::jsonb).
                // The path is bound as a text[] literal ('{samples,0,Flavor Intensity}');
                // metric names contain spaces, which the array literal tolerates
                // unquoted here since they have no commas or braces. value is
                // bound, not interpolated.
                const QString path = QString("{samples,%1,%2}")
                                         .arg(sampleIdx).arg(metric);
                q.prepare("UPDATE detailed_sensory_sessions "
                          "SET json_data = jsonb_set(json_data, ?::text[], "
                          "                           to_jsonb(?::double precision), true), "
                          "    updated_by = 'livesync-sim' "
                          "WHERE id = ? RETURNING version");
                q.addBindValue(path);
                q.addBindValue(value);
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

    // -- Helper: does <table>.<column> exist? (out-of-band catalog read). Used
    //    by the schema-drift self-heal test to assert columns before/after. --
    bool columnExistsOob(const QString& table, const QString& column) {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_colchk";
        bool exists = false;
        {
            QSqlDatabase oob = QSqlDatabase::addDatabase("QPSQL", cname);
            oob.setHostName(cfg.host);
            oob.setPort(cfg.port);
            oob.setDatabaseName(cfg.database);
            oob.setUserName(cfg.user);
            oob.setPassword(cfg.password);
            if (oob.open()) {
                QSqlQuery q(oob);
                q.prepare("SELECT 1 FROM information_schema.columns "
                          "WHERE table_name = ? AND column_name = ?");
                q.addBindValue(table);
                q.addBindValue(column);
                exists = q.exec() && q.next();
                oob.close();
            }
        }
        QSqlDatabase::removeDatabase(cname);
        return exists;
    }

    // -- Helper: drop <table>.<column> out-of-band to simulate a pre-migration
    //    (v2.0-baseline) live DB that never received the additive column. --
    bool dropColumnOob(const QString& table, const QString& column) {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_dropcol";
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
                ok = q.exec(QString("ALTER TABLE %1 DROP COLUMN IF EXISTS %2")
                                .arg(table, column));
                oob.close();
            }
        }
        QSqlDatabase::removeDatabase(cname);
        return ok;
    }

    // -- Helper: run an arbitrary out-of-band query block on a fresh second
    //    QPSQL connection (same host/db/credentials as the suite). Mirrors the
    //    connection lifecycle of the helpers above so schema-drift tests can
    //    set up / inspect arbitrary catalog state without entangling the
    //    DatabaseManager under test. The callback receives a live QSqlQuery
    //    bound to the OOB connection. --
    template <typename Fn>
    void runOob(Fn&& fn) {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_runoob";
        {
            QSqlDatabase oob = QSqlDatabase::addDatabase("QPSQL", cname);
            oob.setHostName(cfg.host);
            oob.setPort(cfg.port);
            oob.setDatabaseName(cfg.database);
            oob.setUserName(cfg.user);
            oob.setPassword(cfg.password);
            if (oob.open()) {
                QSqlQuery q(oob);
                fn(q);
                oob.close();
            }
        }
        QSqlDatabase::removeDatabase(cname);
    }

    // -- Helper: count rows out-of-band, optionally filtered to a single
    //    parent id (e.g. sensory_images WHERE session_id = ?). Returns -1 on
    //    query failure so callers can distinguish "0 rows" from "couldn't ask".
    int countRowsOob(const QString& table,
                     const QString& parentCol = QString(),
                     qint64 parentId = -1) {
        int n = -1;
        runOob([&](QSqlQuery& q) {
            QString sql = QString("SELECT count(*) FROM %1").arg(table);
            if (!parentCol.isEmpty())
                sql += QString(" WHERE %1 = %2").arg(parentCol).arg(parentId);
            if (q.exec(sql) && q.next())
                n = q.value(0).toInt();
        });
        return n;
    }

    // -- Helper: write a tiny valid PNG to a temp file and return its path.
    //    upsertImagesFor reads the file from disk, so an image-bearing session
    //    needs a real readable file. The QTemporaryDir lives in m_imageTmpDir
    //    so the file outlives the call but is cleaned up at suite teardown.
    QString writeTempImage(const QString& baseName = "img.png") {
        if (!m_imageTmpDir) m_imageTmpDir.reset(new QTemporaryDir);
        const QString path = m_imageTmpDir->path() + "/" + baseName;
        QImage img(4, 4, QImage::Format_RGB32);
        img.fill(Qt::red);
        img.save(path, "PNG");
        return path;
    }

    QScopedPointer<QTemporaryDir> m_imageTmpDir;

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
    // DATAVIEWER-4: a whole-session save is now score-immutable -- LiveSync owns
    // per-cell scores. So the score change is driven through a per-cell jsonb_set
    // (the live-sync path), and each subsequent whole-session save must (a) keep
    // UPDATEing the SAME row (no duplicates) and (b) PRESERVE the live-synced
    // score via the merge instead of clobbering it back to the 5.0 default.
    void testSensoryUpsertNoDuplicates()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession sess = makeSensorySession();
        QVERIFY(db.saveSensorySession(sess));   // populates sess.id, sess.version
        QVERIFY(sess.id > 0);
        QCOMPARE(db.listSensoryRecords().size(), 1);

        for (int i = 0; i < 4; ++i) {
            // LiveSync-style per-cell edit bumps the row version; refresh the
            // in-memory version so the next whole-session UPDATE is not stale.
            const int v = setSensoryScoreOutOfBand(sess.id, 0,
                                                   "Overall Liking", 5.0 + i);
            QVERIFY(v > 0);
            sess.version = v;
            // A whole-session re-save must UPDATE in place and NOT reset the
            // live-synced score (the merge keeps the DB-authoritative value).
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

    // -- RC1 (v2.4.0 data-loss regression): a STALE in-memory version must
    //    not fail a whole-file save. files.version routinely advances past
    //    the in-memory copy (LiveSync commits; this client's own prior
    //    saves), so tryWriteFile re-reads the current version inside the save
    //    transaction (SELECT ... FOR UPDATE) and binds THAT in the OCC WHERE
    //    clause: a whole-file save is row-level last-writer-wins by design.
    //    The FOR UPDATE lock serializes concurrent whole-file savers, so the
    //    SELECT->UPDATE window is closed and a genuine VersionMismatch is
    //    unreachable (the `AND version = ?` clause is a defensive invariant).
    //    RowDeleted recovery is handled by the wrapper (see
    //    tpmWholeSave_fileRowDeleted_reinsertsWholeFile).
    //    [Replaces testTryWriteFile_VersionMismatch, whose stale-version
    //    setup now -- by design -- expects Success.]
    void tpmUpdate_staleInMemoryVersion_succeeds()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("stale-tpm.xlsx", "/tmp/stale-tpm.xlsx");
        QVERIFY(db.saveFile(fr));
        fr = db.loadFile(db.listFiles()[0].id);
        QVERIFY(fr.id > 0 && fr.version > 0);

        // Advance files.version twice out-of-band; fr keeps the stale value.
        QVERIFY(bumpRowVersionOutOfBand("files", fr.id) > fr.version);
        QVERIFY(bumpRowVersionOutOfBand("files", fr.id) > fr.version);

        // Mutate a cell value. Child-row versions are fresh from loadFile;
        // only the files-row version is stale here.
        fr.sheets[0].samples[0].rows[0].tpm = 9.25;

        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);

        const DVE::FileResult loaded = db.loadFile(fr.id);
        QVERIFY(!loaded.sheets.isEmpty());
        QCOMPARE(loaded.sheets[0].samples[0].rows[0].tpm, 9.25);
        db.close();
    }

    // -- RC1 (v2.4.0 data-loss regression), CHILD-row twin: a whole-file save
    //    must land even when the in-memory CHILD-row versions (tests / samples
    //    / data_rows) are stale. The Task-1 fix only freshened the files-row
    //    version; the per-child UPDATEs in tryWriteFile still bound the
    //    in-memory sheet.version / sr.version / dr.version, so a LiveSync
    //    per-cell commit (or this client's own prior save) that bumped a child
    //    row's version made the whole-file save fail with VersionMismatch on
    //    that child -- "tryWriteFile(UPDATE test id=142): version mismatch" in
    //    the production log -- which callers then dropped. Each child UPDATE
    //    now re-reads its current version inside the save transaction
    //    (SELECT version ... FOR UPDATE) and binds THAT: child rows are
    //    row-level last-writer-wins by design, same as the file row.
    void tpmWholeSave_staleChildVersions_succeeds()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("stale-child.xlsx", "/tmp/stale-child.xlsx");
        QVERIFY(db.saveFile(fr));
        fr = db.loadFile(db.listFiles()[0].id);
        QVERIFY(fr.id > 0 && fr.version > 0);
        QVERIFY(!fr.sheets.isEmpty());
        QVERIFY(!fr.sheets[0].samples.isEmpty());
        QVERIFY(!fr.sheets[0].samples[0].rows.isEmpty());

        const qint64 testId   = fr.sheets[0].id;
        const qint64 sampleId = fr.sheets[0].samples[0].id;
        const qint64 rowId    = fr.sheets[0].samples[0].rows[0].id;
        QVERIFY(testId > 0 && sampleId > 0 && rowId > 0);

        // Advance each child row's version twice out-of-band (LiveSync-style
        // per-cell commits). The in-memory fr keeps the versions from loadFile.
        QVERIFY(bumpRowVersionOutOfBand("tests",     testId)   > fr.sheets[0].version);
        QVERIFY(bumpRowVersionOutOfBand("tests",     testId)   > fr.sheets[0].version);
        QVERIFY(bumpRowVersionOutOfBand("samples",   sampleId) > fr.sheets[0].samples[0].version);
        QVERIFY(bumpRowVersionOutOfBand("samples",   sampleId) > fr.sheets[0].samples[0].version);
        QVERIFY(bumpRowVersionOutOfBand("data_rows", rowId)    > fr.sheets[0].samples[0].rows[0].version);
        QVERIFY(bumpRowVersionOutOfBand("data_rows", rowId)    > fr.sheets[0].samples[0].rows[0].version);

        // Mutate a cell value on the stale struct.
        fr.sheets[0].samples[0].rows[0].tpm = 7.77;

        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);

        const DVE::FileResult loaded = db.loadFile(fr.id);
        QVERIFY(!loaded.sheets.isEmpty());
        QVERIFY(!loaded.sheets[0].samples.isEmpty());
        QVERIFY(!loaded.sheets[0].samples[0].rows.isEmpty());
        QCOMPARE(loaded.sheets[0].samples[0].rows[0].tpm, 7.77);
        db.close();
    }

    // -- RC1 wrapper resilience: a whole-file save whose file row was deleted
    //    out-of-band must NOT vanish. The byRef tryWriteFile wrapper detects
    //    the RowDeleted return from the core, sees the file row is gone, and
    //    re-INSERTs the whole tree as a fresh file -> Success with a new id.
    //    [Was: this test asserted the raw core RowDeleted return; the wrapper
    //    now recovers, so it expects Success + a fresh row.]
    void tpmWholeSave_fileRowDeleted_reinsertsWholeFile()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr = makeFileResult("rd.xlsx", "/tmp/rd.xlsx");
        QVERIFY(db.saveFile(fr));
        fr = db.loadFile(db.listFiles()[0].id);
        QVERIFY(fr.id > 0 && fr.version > 0);
        const qint64 oldId = fr.id;

        // An in-memory edit the user must keep, then a delete out from under us.
        fr.sheets[0].samples[0].rows[0].tpm = 6.66;
        QVERIFY(deleteRowOutOfBand("files", fr.id));

        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);
        QVERIFY(fr.id > 0);            // byRef back-fill gave us the new file id
        QVERIFY(fr.id != oldId);       // it is a NEW row, not the deleted one

        const DVE::FileResult loaded = db.loadFile(fr.id);
        QVERIFY(!loaded.sheets.isEmpty());
        QCOMPARE(loaded.sheets[0].samples[0].rows[0].tpm, 6.66);   // edit survived
        db.close();
    }

    // -- v2.5.0 review (Task 2): when a CHILD row (here a data_rows row) is
    //    deleted out-of-band but the FILE row survives, tryWriteFile's
    //    RowDeleted recovery must re-INSERT the missing children WITHOUT
    //    duplicating the surviving ones. The recovery resets every child
    //    anchor (resetFileIdsForReinsert with resetFileRow=false) and re-runs
    //    the core down the INSERT path; the core DELETEs the file's existing
    //    children first, so the final per-table row counts must match the
    //    in-memory FileResult exactly -- no duplicates of the rows that were
    //    never deleted.
    void tpmChildDeleted_childrenReinsertedNoDuplicates()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // Build a file with 1 sheet, 2 samples, 2 data_rows each (4 rows total).
        DVE::FileResult fr;
        fr.filePath        = "/tmp/childdel.xlsx";
        fr.fileName        = "childdel.xlsx";
        fr.templateVersion = "new";
        fr.sheetNames      << "Lifetime Test";
        {
            DVE::SheetResult sheet;
            sheet.sheetName       = "Lifetime Test";
            sheet.templateVersion = "new";
            sheet.columnHeaders   << "puffs" << "beforeWeight" << "afterWeight";
            for (int si = 0; si < 2; ++si) {
                DVE::SampleResult sample;
                sample.sampleName = QString("Sample %1").arg(si + 1);
                sample.sampleID   = QString("S-%1").arg(si + 1);
                sample.averageTPM = 3.5;
                for (int ri = 0; ri < 2; ++ri) {
                    DVE::DataRow row;
                    row.puffs = 10.0 * (ri + 1);
                    row.tpm   = 3.0 + si + ri;
                    sample.rows.append(row);
                }
                sheet.samples.append(sample);
            }
            fr.sheets.append(sheet);
        }
        QVERIFY(db.saveFile(fr));
        fr = db.loadFile(db.listFiles()[0].id);
        QVERIFY(fr.id > 0 && fr.version > 0);

        // Sanity: the in-memory tree and the DB agree pre-delete.
        QCOMPARE(fr.sheets.size(), 1);
        QCOMPARE(fr.sheets[0].samples.size(), 2);
        QCOMPARE(countRowsOob("tests"),     1);
        QCOMPARE(countRowsOob("samples"),   2);
        QCOMPARE(countRowsOob("data_rows"), 4);

        // Delete ONE data_rows row out-of-band (file + siblings survive).
        runOob([](QSqlQuery& q) {
            q.exec("DELETE FROM data_rows "
                   "WHERE id = (SELECT id FROM data_rows ORDER BY id LIMIT 1)");
        });
        QCOMPARE(countRowsOob("data_rows"), 3);

        // Mutate a cell so this is a genuine whole-file save.
        fr.sheets[0].samples[0].rows[0].tpm = 9.99;

        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);

        // Per-table counts must match the in-memory FileResult EXACTLY -- the
        // surviving children were not duplicated, the deleted one was re-INSERTED.
        int memSamples = 0, memRows = 0;
        for (const auto& sh : fr.sheets)
            for (const auto& sa : sh.samples) {
                ++memSamples;
                memRows += sa.rows.size();
            }
        QCOMPARE(countRowsOob("tests"),     fr.sheets.size());
        QCOMPARE(countRowsOob("samples"),   memSamples);   // 2
        QCOMPARE(countRowsOob("data_rows"), memRows);       // 4 again, no dupes

        // The edit survived the round-trip.
        const DVE::FileResult loaded = db.loadFile(fr.id);
        QCOMPARE(loaded.sheets[0].samples[0].rows[0].tpm, 9.99);
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

    // -- RC1 (v2.4.0 data-loss regression): sensory twin of
    //    tpmUpdate_staleInMemoryVersion_succeeds. The DB version advances via
    //    LiveSync per-cell commits while the in-memory struct keeps the
    //    version from its INSERT; the whole-session save must still land
    //    instead of returning VersionMismatch (which callers used to treat as
    //    "already up to date" and silently drop the user's edits). Scores
    //    stay DB-authoritative under the DATAVIEWER-4 merge until the
    //    dirty-aware merge ships (v2.5.0 plan Task 3); metadata edits must
    //    persist. [Replaces testTryWriteSensorySession_VersionMismatch.]
    void sensoryUpdate_staleInMemoryVersion_succeeds()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s = makeSensorySession("stale-sens", "C", "2026-06-10");
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);
        QVERIFY(s.id > 0 && s.version > 0);   // byRef back-fills id/version

        // LiveSync-style per-cell commit + a second writer: the DB version
        // moves +2 while the struct stays at its INSERT-time version.
        QVERIFY(setSensoryScoreOutOfBand(s.id, 0, "Smoothness", 8.0) > s.version);
        QVERIFY(bumpRowVersionOutOfBand("sensory_sessions", s.id) > s.version);

        // Mutate the stale struct: metadata + a score the merge arbitrates.
        s.media = "StaleMedia";
        s.samples[0].scores["Overall Liking"] = 7.5;

        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);

        const DVE::SensorySession loaded = db.loadSensorySession(s.id);
        QCOMPARE(loaded.media, QString("StaleMedia"));   // metadata landed
        QVERIFY(!loaded.samples.isEmpty());
        // LiveSync-owned value preserved by the DATAVIEWER-4 merge.
        QCOMPARE(loaded.samples[0].scores.value("Smoothness"), 8.0);
        // DB-authoritative merge still wins for scores (Task 3 will make the
        // locally-dirty 7.5 authoritative; until then it must not land).
        QCOMPARE(loaded.samples[0].scores.value("Overall Liking"), 5.0);
        db.close();
    }

    // -- WriteResult coverage: the CORE still classifies RowDeleted -----------
    // The byRef wrapper now recovers from RowDeleted by re-INSERTing (see
    // sensoryUpdate_rowDeleted_reinsertsFreshRow), so to keep coverage of the
    // raw classification we exercise the const-ref / fire-and-forget overload,
    // which surfaces the core's WriteResult unchanged. (saveSensorySession's
    // const-ref bool shim returns false; tryWriteSensorySession(const&) returns
    // the WriteResult.)
    void testTryWriteSensorySession_RowDeleted_coreClassification()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s = makeSensorySession("rd-sens", "C", "2026-05-01");
        QVERIFY(db.saveSensorySession(s));
        QVERIFY(s.id > 0);

        QVERIFY(deleteRowOutOfBand("sensory_sessions", s.id));

        // const-ref overload: no wrapper recovery, raw core classification.
        const DVE::SensorySession& cref = s;
        QCOMPARE(db.tryWriteSensorySession(cref), DVE::WriteResult::RowDeleted);
        db.close();
    }

    // -- RC1 (v2.4.0 data-loss regression), wrapper resilience: if another
    //    client (or a stale handle) deleted the session row out-of-band, the
    //    user's in-memory edits must NOT vanish. The wrapper detects RowDeleted
    //    from the core, resets id/version on a local copy, and re-INSERTs the
    //    in-memory data as a fresh row -- never dropping the work. byRef
    //    back-fill still populates the new id/version so the panel keeps
    //    tracking the row. [Was: testTryWriteSensorySession_RowDeleted observed
    //    the raw core return; this locks the recover-by-reinsert behavior.]
    void sensoryUpdate_rowDeleted_reinsertsFreshRow()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s =
            makeSensorySession("reinsert-sens", "Charlie", "2026-06-10");
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);
        const qint64 oldId = s.id;
        QVERIFY(oldId > 0 && s.version > 0);

        // Make an in-memory edit the user expects to keep.
        s.media = "ReinsertMedia";

        // Another client deletes the row out from under us.
        QVERIFY(deleteRowOutOfBand("sensory_sessions", oldId));

        // The wrapper must recover by re-INSERTing the in-memory data.
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);
        QVERIFY(s.id > 0);                 // byRef back-fill gave us a fresh id
        QVERIFY(s.id != oldId);            // it is a NEW row, not the deleted one
        QVERIFY(s.version > 0);

        const DVE::SensorySession loaded = db.loadSensorySession(s.id);
        QCOMPARE(loaded.id, s.id);
        QCOMPARE(loaded.sessionName, QString("reinsert-sens"));
        QCOMPARE(loaded.media, QString("ReinsertMedia"));   // the edit survived
        db.close();
    }

    // -- v2.5.0 review (Task 2): the RowDeleted re-INSERT recovery must reset
    //    the IMAGE anchors too, not just the session id/version. When a session
    //    row is deleted out-of-band its sensory_images rows are CASCADE-removed,
    //    yet the in-memory struct still holds the old image ids/versions. If the
    //    recovery left those set, upsertImagesFor would take its UPDATE branch
    //    (WHERE id=<old>), match no row, and the whole save would fail forever
    //    (OtherError, dirty flag retries indefinitely). After the fix every
    //    image takes the INSERT branch and re-lands under the new session id.
    //    [RED before the fix: this returned OtherError; GREEN after.]
    void sensoryUpdate_rowDeletedWithImages_reinsertsImagesToo()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession s =
            makeSensorySession("reinsert-img", "Charlie", "2026-06-10");
        s.imagePaths << writeTempImage("sens.png");
        s.imageLayouts << QRectF();
        s.imageCrops   << QRectF(0, 0, 1, 1);

        // INSERT the session + its one image; byRef back-fills the image
        // anchors with the server-assigned id/version.
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);
        const qint64 oldId = s.id;
        QVERIFY(oldId > 0);
        QVERIFY(!s.imageIds.isEmpty() && s.imageIds[0] > 0);
        QVERIFY(!s.imageVersions.isEmpty() && s.imageVersions[0] > 0);
        QCOMPARE(countRowsOob("sensory_images", "session_id", oldId), 1);

        // Another client deletes the session row; CASCADE removes its images.
        QVERIFY(deleteRowOutOfBand("sensory_sessions", oldId));
        QCOMPARE(countRowsOob("sensory_images", "session_id", oldId), 0);

        // The wrapper must recover: fresh session row AND fresh image row.
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);
        QVERIFY(s.id > 0 && s.id != oldId);           // fresh session row
        QVERIFY(!s.imageIds.isEmpty() && s.imageIds[0] > 0);
        QVERIFY(s.imageIds[0] != -1);
        // The image row lives under the NEW session id, nowhere under the old.
        QCOMPARE(countRowsOob("sensory_images", "session_id", s.id), 1);
        QCOMPARE(countRowsOob("sensory_images", "session_id", oldId), 0);

        const DVE::SensorySession loaded = db.loadSensorySession(s.id);
        QCOMPARE(loaded.imagePaths.size(), 1);        // image survived the round-trip
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

    // -- RC3 (June 10 production log): a failed INSERT (23505) inside the
    //    write transaction must never leave the SHARED connection in
    //    "current transaction is aborted, commands ignored until end of
    //    transaction block". Every error return path in the tryWrite*
    //    wrappers must roll the transaction back, so an immediate read AND a
    //    subsequent write on the same DatabaseManager both work.
    void sensoryInsert_uniqueViolation_leavesConnectionUsable()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::SensorySession a = makeSensorySession("uv-conn", "C", "2026-06-10");
        QCOMPARE(db.tryWriteSensorySession(a), DVE::WriteResult::Success);
        QVERIFY(a.id > 0);

        // Second fresh struct (id=-1) with the SAME natural key -> 23505.
        DVE::SensorySession b = makeSensorySession("uv-conn", "C", "2026-06-10");
        QCOMPARE(db.tryWriteSensorySession(b), DVE::WriteResult::UniqueViolation);

        // Read path: must execute, not be "ignored until end of transaction".
        QCOMPARE(db.listSensoryRecords().size(), 1);
        const DVE::SensorySession reloaded = db.loadSensorySession(a.id);
        QCOMPARE(reloaded.id, a.id);
        QCOMPARE(reloaded.sessionName, QString("uv-conn"));

        // Write path: a NEW transaction must begin cleanly on the same
        // connection after the rollback.
        DVE::SensorySession c = makeSensorySession("uv-conn-2", "C", "2026-06-10");
        QCOMPARE(db.tryWriteSensorySession(c), DVE::WriteResult::Success);
        QCOMPARE(db.listSensoryRecords().size(), 2);
        db.close();
    }

    // -- DATAVIEWER-4: a whole-session save must NOT clobber LiveSync-owned
    //    per-cell scores ------------------------------------------------------
    // Repro of the score-reset bug: the live-sync path writes a single metric
    // straight into json_data; the wholesale save then serialized a stale
    // in-memory copy (unset metrics default to 5.0) and overwrote it. The fix
    // reads the current blob in the same txn and re-applies DB-authoritative
    // scores before the UPDATE. Metadata edits still land.
    void sensoryWholeSessionSave_preservesLiveSyncScores()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // 1. INSERT a session (by-ref fills s.id / s.version).
        DVE::SensorySession s =
            makeSensorySession("ClobberTest", "Charlie", "2026-06-06");
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);
        const qint64 sid = s.id;
        QVERIFY(sid > 0);
        QVERIFY(s.version > 0);

        // 2. Simulate a LiveSync per-cell edit out-of-band: Smoothness -> 8 on
        //    samples[0]. The UPDATE bumps the row version via the trigger.
        const int bumpedVersion =
            setSensoryScoreOutOfBand(sid, 0, "Smoothness", 8.0);
        QVERIFY(bumpedVersion > s.version);

        // 3. Build a STALE in-memory copy: current id + bumped version, all
        //    scores reset to the serializer default (5.0), plus a metadata
        //    change (media) to prove non-score writes still apply.
        DVE::SensorySession stale =
            makeSensorySession("ClobberTest", "Charlie", "2026-06-06");
        stale.id      = sid;
        stale.version = bumpedVersion;
        stale.media   = "MergedMedia";
        QVERIFY(!stale.samples.isEmpty());
        for (const QString& m : DVE::kSensoryMetrics)
            stale.samples[0].scores[m] = 5.0;

        // 4. Whole-session save: UPDATE branch must land (version is current).
        QCOMPARE(db.tryWriteSensorySession(stale), DVE::WriteResult::Success);

        // 5. Reload and assert: the live-synced score survived (NOT reset to
        //    5), an untouched metric stays at its DB value, and the metadata
        //    change applied.
        DVE::SensorySession loaded = db.loadSensorySession(sid);
        QVERIFY(!loaded.samples.isEmpty());
        QCOMPARE(loaded.samples[0].scores.value("Smoothness"), 8.0);
        QCOMPARE(loaded.samples[0].scores.value("Burnt Taste"), 5.0);
        QCOMPARE(loaded.media, QString("MergedMedia"));
        db.close();
    }

    // -- DATAVIEWER-4: the detailed-sensory twin of the clobber test ---------
    // Same bug, detailed mode: the live-sync path writes a single metric
    // straight into json_data; the wholesale save then serialized a stale
    // in-memory copy (unset metrics default to 0.0 here -- the detailed
    // serializer's default) and overwrote it. The fix reads the current blob
    // in the same txn and re-applies DB-authoritative scores before the
    // UPDATE. Metadata edits still land.
    void detailedWholeSessionSave_preservesLiveSyncScores()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // 1. INSERT a detailed session (by-ref fills s.id / s.version).
        DVE::DetailedSensorySession s =
            makeDetailedSensorySession("DetClobberTest", "Charlie", "2026-06-06");
        QCOMPARE(db.tryWriteDetailedSensorySession(s), DVE::WriteResult::Success);
        const qint64 sid = s.id;
        QVERIFY(sid > 0);
        QVERIFY(s.version > 0);

        // 2. Simulate a LiveSync per-cell edit out-of-band: Flavor Intensity
        //    -> 6 on samples[0]. 6 is strictly inside fromJson's [1, max=9]
        //    clamp for this metric (so the assert below is unambiguous). The
        //    UPDATE bumps the row version via the trigger.
        const int bumpedVersion =
            setDetailedScoreOutOfBand(sid, 0, "Flavor Intensity", 6.0);
        QVERIFY(bumpedVersion > s.version);

        // 3. Build a STALE in-memory copy: current id + bumped version, every
        //    metric reset to the serializer default (0.0), plus a metadata
        //    change (media) to prove non-score writes still apply.
        DVE::DetailedSensorySession stale =
            makeDetailedSensorySession("DetClobberTest", "Charlie", "2026-06-06");
        stale.id      = static_cast<int>(sid);
        stale.version = bumpedVersion;
        stale.media   = "MergedMedia";
        QVERIFY(!stale.samples.isEmpty());
        for (const QString& m : DVE::kDetailedAllMetrics)
            stale.samples[0].scores[m] = 0.0;

        // 4. Whole-session save: UPDATE branch must land (version is current).
        QCOMPARE(db.tryWriteDetailedSensorySession(stale), DVE::WriteResult::Success);

        // 5. Reload and assert: the live-synced score survived (== 6, NOT reset
        //    to the serializer default), an untouched metric reloads at the
        //    post-clamp default (0.0 -> qBound(1.0, .., 9) == 1.0), and the
        //    metadata change applied.
        DVE::DetailedSensorySession loaded = db.loadDetailedSensorySession(static_cast<int>(sid));
        QVERIFY(!loaded.samples.isEmpty());
        QCOMPARE(loaded.samples[0].scores.value("Flavor Intensity"), 6.0);
        QCOMPARE(loaded.samples[0].scores.value("Burn Taste"), 1.0);
        QCOMPARE(loaded.media, QString("MergedMedia"));
        db.close();
    }

    // -- RC1 (v2.4.0 data-loss regression): detailed twin of
    //    sensoryUpdate_staleInMemoryVersion_succeeds. Stale in-memory version
    //    + LiveSync-advanced DB version -> the whole-session save must land.
    void detailedUpdate_staleInMemoryVersion_succeeds()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::DetailedSensorySession s =
            makeDetailedSensorySession("stale-det", "C", "2026-06-10");
        QCOMPARE(db.tryWriteDetailedSensorySession(s), DVE::WriteResult::Success);
        QVERIFY(s.id > 0 && s.version > 0);   // byRef back-fills id/version

        // LiveSync-style per-cell commit + a second writer: the DB version
        // moves +2 while the struct stays at its INSERT-time version.
        QVERIFY(setDetailedScoreOutOfBand(s.id, 0, "Flavor Intensity", 6.0)
                > s.version);
        QVERIFY(bumpRowVersionOutOfBand("detailed_sensory_sessions", s.id)
                > s.version);

        // Mutate the stale struct: metadata + a score the merge arbitrates.
        s.media = "StaleMedia";
        s.samples[0].scores["Flavor Intensity"] = 9.0;

        QCOMPARE(db.tryWriteDetailedSensorySession(s), DVE::WriteResult::Success);

        const DVE::DetailedSensorySession loaded =
            db.loadDetailedSensorySession(s.id);
        QCOMPARE(loaded.media, QString("StaleMedia"));   // metadata landed
        QVERIFY(!loaded.samples.isEmpty());
        // LiveSync-owned value preserved by the DATAVIEWER-4 merge; the
        // in-memory 9.0 stays DB-arbitrated until the Task-3 dirty-aware merge.
        QCOMPARE(loaded.samples[0].scores.value("Flavor Intensity"), 6.0);
        db.close();
    }

    // -- v2.5.0 wrapper resilience, detailed twin of
    //    sensoryUpdate_rowDeleted_reinsertsFreshRow: if another client deleted
    //    the detailed-session row out-of-band, the user's in-memory edits must
    //    not vanish. The byRef wrapper detects RowDeleted from the core, resets
    //    id/version (and image anchors) on a local copy, and re-INSERTs the
    //    in-memory data as a fresh row -- never dropping the work. byRef
    //    back-fill repopulates the new id/version so the panel keeps tracking it.
    void detailedUpdate_rowDeleted_reinsertsFreshRow()
    {
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::DetailedSensorySession s =
            makeDetailedSensorySession("reinsert-det", "Charlie", "2026-06-10");
        QCOMPARE(db.tryWriteDetailedSensorySession(s), DVE::WriteResult::Success);
        const qint64 oldId = s.id;
        QVERIFY(oldId > 0 && s.version > 0);

        // An in-memory edit the user expects to keep.
        s.media = "ReinsertDetMedia";

        // Another client deletes the row out from under us.
        QVERIFY(deleteRowOutOfBand("detailed_sensory_sessions", oldId));

        // The wrapper must recover by re-INSERTing the in-memory data.
        QCOMPARE(db.tryWriteDetailedSensorySession(s), DVE::WriteResult::Success);
        QVERIFY(s.id > 0);                 // byRef back-fill gave us a fresh id
        QVERIFY(s.id != oldId);            // it is a NEW row, not the deleted one
        QVERIFY(s.version > 0);

        const DVE::DetailedSensorySession loaded =
            db.loadDetailedSensorySession(s.id);
        QCOMPARE(loaded.id, s.id);
        QCOMPARE(loaded.sessionName, QString("reinsert-det"));
        QCOMPARE(loaded.media, QString("ReinsertDetMedia"));   // the edit survived
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

        // --- new-template with a BLANK regime cell: persists as '' (non-NULL),
        //     so the sheet still loads as new-template (hasPerRowRegime true). ---
        DVE::FileResult fBlank;
        fBlank.filePath = "regime_blank.xlsx";
        fBlank.fileName = "regime_blank.xlsx";
        DVE::SheetResult shBlank;
        shBlank.sheetName = "Lifetime Test";
        shBlank.hasPerRowRegime = true;
        DVE::SampleResult smBlank;
        smBlank.sampleName = "S1";
        DVE::DataRow d; d.puffs = 10; d.beforeWeight = 25.1; d.afterWeight = 25.06; d.puffingRegime = "";
        smBlank.rows << d;
        shBlank.samples << smBlank;
        fBlank.sheets << shBlank;
        QVERIFY(db.saveFile(fBlank));

        DVE::FileResult loadedBlank = db.loadFileByPath("regime_blank.xlsx");
        QCOMPARE(loadedBlank.sheets[0].hasPerRowRegime, true);   // '' is non-NULL -> new-template
        QVERIFY(loadedBlank.sheets[0].samples[0].rows[0].puffingRegime.isEmpty());

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

    // ── Schema-drift self-heal (v2.3.3 hotfix) ──────────────────────────────
    // A live DB created from an OLDER init.sql is missing the additive columns
    // added by later migrations (data_rows.puffing_regime [v2.2.1],
    // tests.raw_grid [v2.x]). tryWriteFile references both unconditionally, so
    // every file upload fails at prepare time with SQLSTATE 42703
    // ("column ... does not exist"). ensureSchema() — run on every connect —
    // must reconcile the missing columns idempotently and non-destructively so
    // uploads recover with no manual NAS migration step. Verifies the method
    // directly, idempotency, the connect-hook (reopen), and end-to-end write.
    void ensureSchema_reAddsDroppedAdditiveColumns()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // (1) Direct ensureSchema() reconciles both dropped columns. NOTE: every
        //     QVERIFY is deferred to the end so a failed assertion can never
        //     leave the shared test schema missing a column (and even if it
        //     did, the next slot's open() re-runs ensureSchema()).
        dropColumnOob("data_rows", "puffing_regime");
        dropColumnOob("tests", "raw_grid");
        const bool droppedOk = !columnExistsOob("data_rows", "puffing_regime")
                            && !columnExistsOob("tests", "raw_grid");

        db.ensureSchema();
        const bool reAdded = columnExistsOob("data_rows", "puffing_regime")
                          && columnExistsOob("tests", "raw_grid");

        db.ensureSchema();   // idempotent — second pass is a harmless no-op
        const bool stillThere = columnExistsOob("data_rows", "puffing_regime")
                             && columnExistsOob("tests", "raw_grid");

        // (2) The connect hook also reconciles: drop again, reopen(), verify.
        dropColumnOob("data_rows", "puffing_regime");
        const bool reopenOk   = db.reopen();
        const bool hookHealed = columnExistsOob("data_rows", "puffing_regime");

        // (3) End-to-end: the upload path that 42703'd now succeeds.
        DVE::FileResult fr = makeFileResult("schemaheal.xlsx", "/tmp/schemaheal.xlsx");
        const DVE::WriteResult wr = db.tryWriteFile(fr);

        db.close();

        QVERIFY2(droppedOk,  "precondition: both additive columns should be dropped");
        QVERIFY2(reAdded,    "ensureSchema() must re-add data_rows.puffing_regime + tests.raw_grid");
        QVERIFY2(stillThere, "ensureSchema() must be idempotent");
        QVERIFY2(reopenOk,   "reopen() should reconnect to the test DB");
        QVERIFY2(hookHealed, "the connect hook must run ensureSchema() on reopen()");
        QCOMPARE(wr, DVE::WriteResult::Success);
    }

    // ── Bug-2 diagnostic: NORMAL-SHEET round-trip ───────────────────────────
    // Prove (or disprove) that DataRow fields beforeWeight/afterWeight survive
    // a Postgres save→reload with the exact non-zero values written.
    // Also counts how many rows would pass the UI gate
    //   displayCurrentSample: if (dr.beforeWeight == 0 || dr.afterWeight == 0) skip
    // so we can confirm the UI would actually render the rows.
    void testNormalSheetRoundTrip_nonZeroWeights()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // Build a FileResult that looks like a real "Lifetime Test" sheet:
        // one sheet, two samples, three rows each.  Every row has non-zero
        // before/afterWeight so ALL should survive the UI gate.
        DVE::FileResult fr;
        fr.filePath        = "/tmp/bug2_normal.xlsx";
        fr.fileName        = "bug2_normal.xlsx";
        fr.templateVersion = "new";

        DVE::SheetResult sheet;
        sheet.sheetName       = "Lifetime Test";
        sheet.templateVersion = "new";
        sheet.overallAvgTPM   = 4.2;

        for (int si = 0; si < 2; ++si) {
            DVE::SampleResult sample;
            sample.sampleName = QString("Sample %1").arg(si + 1);
            sample.sampleID   = QString("S-%1").arg(si + 1);
            sample.date       = "2026-01-01";
            sample.tester     = "QA";
            sample.media      = "E-liquid";
            sample.resistance = 1.2 + si * 0.05;
            sample.voltage    = 3.7;
            sample.power      = 11.4;
            sample.averageTPM = 4.0 + si * 0.2;

            for (int ri = 0; ri < 3; ++ri) {
                DVE::DataRow row;
                row.puffs        = 10.0 + ri * 10;   // 10, 20, 30
                row.beforeWeight = 25.100 + ri * 0.05; // 25.100 25.150 25.200
                row.afterWeight  = 25.065 + ri * 0.05; // 25.065 25.115 25.165
                row.tpm          = 3.5 + ri * 0.1;
                row.oilConsumed  = 0.035 + ri * 0.001;
                row.notes        = QString("row %1/%2").arg(si).arg(ri);
                sample.rows.append(row);
            }
            sheet.samples.append(sample);
        }
        fr.sheets.append(sheet);

        // Save
        QVERIFY(db.saveFile(fr));

        // Reload
        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);
        DVE::FileResult loaded = db.loadFile(files[0].id);

        QVERIFY(!loaded.sheets.isEmpty());
        const DVE::SheetResult& gotSheet = loaded.sheets[0];
        QCOMPARE(gotSheet.sheetName, sheet.sheetName);
        QCOMPARE(gotSheet.samples.size(), 2);

        int rowsPassingUiGate = 0;

        for (int si = 0; si < 2; ++si) {
            const DVE::SampleResult& expS = fr.sheets[0].samples[si];
            const DVE::SampleResult& gotS = gotSheet.samples[si];

            QCOMPARE(gotS.sampleName, expS.sampleName);

            // ── Row-count assertion: must not drop any rows ──────────────
            QCOMPARE(gotS.rows.size(), expS.rows.size());

            for (int ri = 0; ri < expS.rows.size(); ++ri) {
                const DVE::DataRow& expR = expS.rows[ri];
                const DVE::DataRow& gotR = gotS.rows[ri];

                // Weight fields intact?
                QCOMPARE(gotR.beforeWeight, expR.beforeWeight);
                QCOMPARE(gotR.afterWeight,  expR.afterWeight);
                QCOMPARE(gotR.puffs,        expR.puffs);

                // Simulate the UI gate: displayCurrentSample skips rows
                // where beforeWeight == 0 || afterWeight == 0.
                if (gotR.beforeWeight != 0.0 && gotR.afterWeight != 0.0)
                    ++rowsPassingUiGate;
            }
        }

        // All 6 rows must pass (none zeroed by the DB round-trip).
        QCOMPARE(rowsPassingUiGate, 6);

        // Also verify tpmTrend was rebuilt from the rows (it's reconstructed
        // in loadFile, not stored separately).
        QCOMPARE(gotSheet.tpmTrend.size(), 6);
        QCOMPARE(gotSheet.puffCounts.size(), 6);

        db.close();
    }

    // ── Bug-2: RAW/SOP-sheet round-trip ────────────────────────────────────
    // Verifies that rawHeaders and rawRows survive a Postgres save→reload via
    // the raw_grid JSONB column on the tests table (Plan B Task 4).
    // isRawTable must also remain true after reload.
    void testRawSheetRoundTrip_gridSurvives()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        DVE::FileResult fr;
        fr.filePath        = "/tmp/bug2_raw.xlsx";
        fr.fileName        = "bug2_raw.xlsx";
        fr.templateVersion = "new";

        DVE::SheetResult sheet;
        sheet.sheetName  = "Test SOP's";
        sheet.isRawTable = true;
        sheet.rawHeaders << "Step" << "Description" << "Notes";
        sheet.rawRows.append(QStringList{"1", "Prepare device", "Check resistance"});
        sheet.rawRows.append(QStringList{"2", "Run test",       "Record weights"});

        fr.sheets.append(sheet);

        QVERIFY(db.saveFile(fr));

        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);
        DVE::FileResult loaded = db.loadFile(files[0].id);

        QVERIFY(!loaded.sheets.isEmpty());
        const DVE::SheetResult& gotSheet = loaded.sheets[0];

        // isRawTable flag must survive (stored in tests.is_raw_table column).
        QVERIFY(gotSheet.isRawTable);

        // rawHeaders and rawRows must survive via the raw_grid JSONB column.
        QCOMPARE(gotSheet.rawHeaders, sheet.rawHeaders);
        QCOMPARE(gotSheet.rawRows,    sheet.rawRows);

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

    // ── DATAVIEWER-2: ensureSchema heals sensory_header_presets shape ────────
    // The sample-name dropdown becomes test-scoped, which requires the preset
    // table to carry a test_name column and swap its uniqueness from
    // (kind, value) to a UNIQUE index on (kind, value, COALESCE(test_name,'')).
    // A live DB (or the test container, whose init.sql predates the table)
    // must self-heal to that shape on connect via ensureSchema() — create the
    // table fresh in the new shape if absent, add test_name if the legacy
    // table exists, and replace the old unique constraint with the expression
    // index. Idempotent, catalog-guarded, best-effort (never throws).
    void sensoryHeaderPresets_ensureSchemaHealsStructure()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // Start from a clean slate: drop the table so ensureSchema must create
        // it fresh in the new (test-scoped) shape, exercising the create path.
        runOob([](QSqlQuery& q){
            QVERIFY(q.exec("DROP TABLE IF EXISTS sensory_header_presets CASCADE"));
        });

        db.ensureSchema();

        // Collect observations first; defer QVERIFY to the end so a failure can
        // never leave the shared test schema half-healed for the next slot
        // (and even if it did, the next open() re-runs ensureSchema()).
        bool tableExists = false, hasTestName = false,
             hasCoalesceIdx = false, hasOldConstraint = true;

        runOob([&](QSqlQuery& q){
            if (q.exec("SELECT to_regclass('public.sensory_header_presets') "
                       "IS NOT NULL") && q.next())
                tableExists = q.value(0).toBool();
        });
        runOob([&](QSqlQuery& q){
            q.exec("SELECT 1 FROM information_schema.columns "
                   "WHERE table_name='sensory_header_presets' "
                   "AND column_name='test_name'");
            hasTestName = q.next();
        });
        runOob([&](QSqlQuery& q){
            if (q.exec("SELECT indexdef FROM pg_indexes "
                       "WHERE tablename='sensory_header_presets'"))
                while (q.next())
                    if (q.value(0).toString().contains("COALESCE(test_name"))
                        hasCoalesceIdx = true;
        });
        runOob([&](QSqlQuery& q){
            q.exec("SELECT 1 FROM pg_constraint "
                   "WHERE conname='sensory_header_presets_kind_value_key'");
            hasOldConstraint = q.next();
        });

        // Idempotent: a second pass must be a harmless no-op.
        db.ensureSchema();
        bool stillCoalesceIdx = false;
        runOob([&](QSqlQuery& q){
            if (q.exec("SELECT indexdef FROM pg_indexes "
                       "WHERE tablename='sensory_header_presets'"))
                while (q.next())
                    if (q.value(0).toString().contains("COALESCE(test_name"))
                        stillCoalesceIdx = true;
        });

        // The healed table is immediately usable by the save path.
        const bool saveOk =
            db.saveSensoryHeaderPresets("TestA", "Media1",
                                        QStringList{ "S1", "S2" });
        const QString saveErr = db.lastError();

        db.close();

        QVERIFY2(tableExists,
                 "ensureSchema() must create sensory_header_presets when absent");
        QVERIFY2(hasTestName,
                 "ensureSchema() must add the test_name column");
        QVERIFY2(hasCoalesceIdx,
                 "ensureSchema() must create the COALESCE(test_name,'') unique index");
        QVERIFY2(!hasOldConstraint,
                 "ensureSchema() must drop the legacy UNIQUE (kind, value) constraint");
        QVERIFY2(stillCoalesceIdx, "ensureSchema() must be idempotent");
        QVERIFY2(saveOk, qPrintable("healed table must be usable: " + saveErr));
    }

    // ── DATAVIEWER-2: sample-name presets are scoped per test ────────────────
    // saveSensoryHeaderPresets stores each sample_name row stamped with the
    // owning test_name (test_name/media rows stay global with a NULL test_name).
    // loadSampleNamesForTest returns only the names saved under one test, and
    // re-saving the same trio is a harmless no-op (the (kind,value,
    // COALESCE(test_name,'')) unique index matches the ON CONFLICT target).
    void sampleNames_scopedByTest()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db; QVERIFY(openDb(db));
        db.ensureSchema();
        runOob([](QSqlQuery& q){ q.exec("DELETE FROM sensory_header_presets"); });

        QVERIFY(db.saveSensoryHeaderPresets("TestA", "M", QStringList{"Alpha","Bravo"}));
        QVERIFY(db.saveSensoryHeaderPresets("TestB", "M", QStringList{"Charlie"}));
        QVERIFY(db.saveSensoryHeaderPresets("TestA", "M", QStringList{"Alpha","Bravo"}));  // repeat: must NOT error

        QStringList a = db.loadSampleNamesForTest("TestA");
        QStringList b = db.loadSampleNamesForTest("TestB");
        QCOMPARE(a, (QStringList{"Alpha","Bravo"}));    // scoped to A (ORDER BY lower(value))
        QCOMPARE(b, (QStringList{"Charlie"}));          // scoped to B; A's names absent
        QVERIFY(db.loadSensoryHeaderPresets("test_name").contains("TestA"));   // global kinds intact
        QVERIFY(db.loadSensoryHeaderPresets("test_name").contains("TestB"));
        QVERIFY(db.loadSensoryHeaderPresets("media").contains("M"));

        db.close();
    }

    // ── DATAVIEWER-2: one-time backfill of sample-name presets from history ──
    // So the scoped dropdown is useful from day one, ensureSchema() runs a
    // ONE-TIME pass that reconstructs (test_title -> sample-name) preset rows
    // from existing session history (sensory_sessions + detailed_sensory_sessions
    // json_data). The scoping key is the *test title* — exactly what the live
    // save/load path keys on (SensoryPanel/DetailedSensoryPanel pass
    // m_testTitleEdit; loadSampleNamesForTest matches the test_name column on
    // that title), NOT the session name. The pass is gated by a schema_meta
    // marker (runs once, not on every connect), idempotent (ON CONFLICT DO
    // NOTHING), and best-effort (logs, never throws). Covers both session tables.
    void sampleNames_backfilledFromSessions()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db; QVERIFY(openDb(db));        // ensureSchema runs (may set marker)
        db.ensureSchema();

        // Seed a SENSORY session under test title "HistSensory" with samples
        // Hist1/Hist2. makeSensorySession gives sessionName==title? No — its
        // testTitle is hardcoded; set it explicitly to the scoping key we assert.
        DVE::SensorySession s = makeSensorySession("HistSensorySess", "Charlie", "2026-06-07");
        s.testTitle = "HistSensory";
        QVERIFY(s.samples.size() >= 1);
        s.samples[0].name = "Hist1";
        DVE::SensorySample s2; s2.name = "Hist2";
        for (const QString& m : DVE::kSensoryMetrics) s2.scores[m] = 5.0;
        s.samples.append(s2);
        QCOMPARE(db.tryWriteSensorySession(s), DVE::WriteResult::Success);

        // Seed a DETAILED session under test title "HistDetailed" with sample
        // HistD1, to exercise the detailed_sensory_sessions backfill branch too.
        DVE::DetailedSensorySession d =
            makeDetailedSensorySession("HistDetailedSess", "Charlie", "2026-06-07");
        d.testTitle = "HistDetailed";
        QVERIFY(d.samples.size() >= 1);
        d.samples[0].name = "HistD1";
        QCOMPARE(db.tryWriteDetailedSensorySession(d), DVE::WriteResult::Success);

        // Force the one-time backfill to re-run: drop the marker + any presets
        // we expect it to (re)create, so this slot is self-contained regardless
        // of whether openDb's ensureSchema already ran the pass once.
        runOob([](QSqlQuery& q){ q.exec("DELETE FROM schema_meta WHERE key='dv2_sample_name_backfill'"); });
        runOob([](QSqlQuery& q){
            q.exec("DELETE FROM sensory_header_presets "
                   "WHERE kind='sample_name' "
                   "AND test_name IN ('HistSensory','HistDetailed')");
        });

        db.ensureSchema();                                   // backfill runs, scans sessions

        // Historical names appear, scoped to their owning test title.
        const QStringList sensoryNames  = db.loadSampleNamesForTest("HistSensory");
        const QStringList detailedNames = db.loadSampleNamesForTest("HistDetailed");
        QVERIFY2(sensoryNames.contains("Hist1"),
                 qPrintable("sensory backfill missing Hist1; got: " + sensoryNames.join(',')));
        QVERIFY2(sensoryNames.contains("Hist2"),
                 qPrintable("sensory backfill missing Hist2; got: " + sensoryNames.join(',')));
        QVERIFY2(detailedNames.contains("HistD1"),
                 qPrintable("detailed backfill missing HistD1; got: " + detailedNames.join(',')));

        // Idempotent: re-run without clearing the marker → skipped, no error,
        // no duplicate rows.
        db.ensureSchema();
        QCOMPARE(db.loadSampleNamesForTest("HistSensory").count("Hist1"), 1);
        QCOMPARE(db.loadSampleNamesForTest("HistDetailed").count("HistD1"), 1);

        db.close();
    }

    // ── Plan B Task 6: dbDataIncomplete detection ───────────────────────────
    // Verifies that DatabaseManager::loadFile sets dbDataIncomplete correctly:
    //   * COMPLETE normal sheet (samples with rows) → false
    //   * INCOMPLETE normal sheet (sample has averageTPM > 0 but no data_rows)
    //     → true  (mimics legacy/migration-gap state)
    void testDbDataIncomplete_normalSheet()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // --- COMPLETE: sample has averageTPM > 0 AND rows present ---
        {
            DVE::FileResult fr;
            fr.filePath        = "/tmp/incomplete_complete.xlsx";
            fr.fileName        = "incomplete_complete.xlsx";
            fr.templateVersion = "new";

            DVE::SheetResult sheet;
            sheet.sheetName       = "Lifetime Test";
            sheet.templateVersion = "new";

            DVE::SampleResult sample;
            sample.sampleName = "S1";
            sample.sampleID   = "S-1";
            sample.date       = "2026-01-01";
            sample.tester     = "QA";
            sample.averageTPM = 3.5;   // > 0

            DVE::DataRow row;
            row.puffs        = 10.0;
            row.beforeWeight = 25.1;
            row.afterWeight  = 25.065;
            row.tpm          = 3.5;
            sample.rows.append(row);   // rows present → complete

            sheet.samples.append(sample);
            fr.sheets.append(sheet);

            // Use the mutable-ref write so the server-assigned id is stamped
            // back onto fr, then load THAT exact row. listFiles() is ORDER BY
            // loaded_at DESC with second granularity, so .last() ties (and
            // returns an arbitrary file) when this slot saves two files in the
            // same second — the source of intermittent failures here.
            QVERIFY(db.tryWriteFile(fr) == DVE::WriteResult::Success);
            QVERIFY(fr.id != -1);
            DVE::FileResult loaded = db.loadFile(fr.id);
            QVERIFY(!loaded.sheets.isEmpty());
            QVERIFY(!loaded.sheets[0].dbDataIncomplete);   // complete → false
        }

        // --- INCOMPLETE: insert sample row directly, no data_rows ---
        // We save a FileResult with an EMPTY rows vector but averageTPM > 0,
        // which is exactly what a legacy record looks like after migration.
        {
            DVE::FileResult fr;
            fr.filePath        = "/tmp/incomplete_legacy.xlsx";
            fr.fileName        = "incomplete_legacy.xlsx";
            fr.templateVersion = "new";

            DVE::SheetResult sheet;
            sheet.sheetName       = "Lifetime Test";
            sheet.templateVersion = "new";

            DVE::SampleResult sample;
            sample.sampleName = "S1";
            sample.sampleID   = "S-1";
            sample.date       = "2026-01-01";
            sample.tester     = "QA";
            sample.averageTPM = 3.5;   // > 0
            // sample.rows intentionally left empty → mimics legacy state

            sheet.samples.append(sample);
            fr.sheets.append(sheet);

            // Load by the stamped-back id (not listFiles().last()) so this
            // assertion targets the legacy-shaped file it just saved, never the
            // complete one from the block above when both land in the same
            // loaded_at second.
            QVERIFY(db.tryWriteFile(fr) == DVE::WriteResult::Success);
            QVERIFY(fr.id != -1);
            DVE::FileResult loaded = db.loadFile(fr.id);
            QVERIFY(!loaded.sheets.isEmpty());
            QVERIFY(loaded.sheets[0].dbDataIncomplete);    // incomplete → true
        }

        db.close();
    }

    // ── Plan B Task 7: DbRepair backfill ────────────────────────────────────
    // Primary test (no Python required): exercises the testable backfillFile
    // core with an in-memory complete FileResult and a DB stub that has
    // averageTPM > 0 but no data_rows (the legacy migration-gap state).
    //
    // Verifies:
    //   * before repair: loadFile returns dbDataIncomplete = true
    //   * backfillFile returns true and detail contains "healed"
    //   * after repair: loadFile returns non-empty rows and dbDataIncomplete = false
    //   * dryRun=true does NOT modify the DB
    void testDbRepair_backfillCore()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // --- Seed an incomplete DB record (averageTPM > 0, no data_rows) ---
        DVE::FileResult incomplete;
        incomplete.filePath        = "/tmp/repair_test.xlsx";
        incomplete.fileName        = "repair_test.xlsx";
        incomplete.templateVersion = "new";

        DVE::SheetResult sheet;
        sheet.sheetName       = "Lifetime Test";
        sheet.templateVersion = "new";
        sheet.overallAvgTPM   = 4.2;

        DVE::SampleResult sample;
        sample.sampleName = "Sample 1";
        sample.sampleID   = "S-1";
        sample.date       = "2026-01-01";
        sample.tester     = "QA";
        sample.averageTPM = 4.2;   // > 0 but...
        // ...sample.rows is intentionally left empty → mimics legacy state

        sheet.samples.append(sample);
        incomplete.sheets.append(sheet);

        QVERIFY(db.saveFile(incomplete));

        // Confirm the DB record is seen as incomplete by loadFile.
        auto files = db.listFiles();
        QCOMPARE(files.size(), 1);
        DVE::FileResult dbStub = db.loadFile(files[0].id);
        QVERIFY(!dbStub.filePath.isEmpty());
        QVERIFY(dbStub.sheets[0].dbDataIncomplete);

        // --- Build a "live" complete FileResult as if processFile returned it ---
        DVE::FileResult live;
        live.filePath        = incomplete.filePath;
        live.fileName        = incomplete.fileName;
        live.templateVersion = incomplete.templateVersion;

        DVE::SheetResult liveSheet;
        liveSheet.sheetName       = "Lifetime Test";
        liveSheet.templateVersion = "new";
        liveSheet.overallAvgTPM   = 4.2;

        DVE::SampleResult liveSample;
        liveSample.sampleName = "Sample 1";
        liveSample.sampleID   = "S-1";
        liveSample.date       = "2026-01-01";
        liveSample.tester     = "QA";
        liveSample.averageTPM = 4.2;

        DVE::DataRow row;
        row.puffs        = 10.0;
        row.beforeWeight = 25.1;
        row.afterWeight  = 25.065;
        row.tpm          = 4.2;
        row.oilConsumed  = 0.035;
        liveSample.rows.append(row);

        liveSheet.samples.append(liveSample);
        live.sheets.append(liveSheet);

        // --- dryRun=true must NOT modify the DB ---
        {
            QString detail;
            const bool ok = DVE::backfillFile(db, live, dbStub, /*dryRun=*/true, detail);
            QVERIFY(ok);
            QVERIFY(detail.contains("dryRun", Qt::CaseInsensitive));
            // DB must still show incomplete after dry run
            DVE::FileResult afterDry = db.loadFile(dbStub.id);
            QVERIFY(afterDry.sheets[0].dbDataIncomplete);
        }

        // --- dryRun=false: actual backfill ---
        {
            QString detail;
            const bool ok = DVE::backfillFile(db, live, dbStub, /*dryRun=*/false, detail);
            QVERIFY2(ok, qPrintable(detail));
            QVERIFY(detail.toLower().contains("heal"));
        }

        // After backfill: reload must show rows and dbDataIncomplete = false.
        DVE::FileResult healed = db.loadFile(dbStub.id);
        QVERIFY(!healed.sheets.isEmpty());
        QVERIFY(!healed.sheets[0].samples.isEmpty());
        QVERIFY(!healed.sheets[0].samples[0].rows.isEmpty());
        QVERIFY(!healed.sheets[0].dbDataIncomplete);

        // Row values must match what we supplied.
        const DVE::DataRow& got = healed.sheets[0].samples[0].rows[0];
        QCOMPARE(got.puffs,        10.0);
        QCOMPARE(got.beforeWeight, 25.1);
        QCOMPARE(got.afterWeight,  25.065);
        QCOMPARE(got.tpm,          4.2);

        db.close();
    }

    // ── Plan B Task 7: runDbRepair classification ────────────────────────────
    // Verifies the RepairSummary counters for all four outcomes:
    //   alreadyComplete, skippedNoXlsx, healed (via backfillFile), failed (not tested here).
    // Does NOT require Python. Uses backfillFile directly for the healed case.
    void testDbRepair_summaryCounters()
    {
        if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
        DVE::DatabaseManager db;
        QVERIFY(openDb(db));

        // --- File A: already complete (has rows) ---
        {
            DVE::FileResult fr;
            fr.filePath        = "/tmp/repair_complete.xlsx";
            fr.fileName        = "repair_complete.xlsx";
            fr.templateVersion = "new";
            DVE::SheetResult sheet;
            sheet.sheetName = "Lifetime Test";
            DVE::SampleResult samp;
            samp.sampleName = "S1";
            samp.averageTPM = 3.0;
            DVE::DataRow r; r.puffs = 10; r.beforeWeight = 25.0; r.afterWeight = 24.9;
            samp.rows << r;
            sheet.samples << samp;
            fr.sheets << sheet;
            QVERIFY(db.saveFile(fr));
        }

        // --- File B: incomplete, .xlsx does NOT exist on disk → skippedNoXlsx ---
        {
            DVE::FileResult fr;
            fr.filePath        = "/tmp/nonexistent_repair.xlsx";
            fr.fileName        = "nonexistent_repair.xlsx";
            fr.templateVersion = "new";
            DVE::SheetResult sheet;
            sheet.sheetName = "Lifetime Test";
            DVE::SampleResult samp;
            samp.sampleName = "S1";
            samp.averageTPM = 2.0;
            // rows intentionally empty
            sheet.samples << samp;
            fr.sheets << sheet;
            QVERIFY(db.saveFile(fr));
        }

        // runDbRepair with dryRun=false but no sourceDir — file B can't be found.
        // We use a DataProcessor stub; since no files are healable via processFile
        // in this test, the loop never calls it.
        DVE::DataProcessor proc;
        DVE::RepairOptions opts;
        opts.dryRun    = false;
        opts.sourceDir = "";  // no search dir

        DVE::RepairSummary summary = DVE::runDbRepair(db, proc, opts);

        QCOMPARE(summary.alreadyComplete, 1);
        QCOMPARE(summary.skippedNoXlsx,   1);
        QCOMPARE(summary.healed,          0);
        QCOMPARE(summary.failed,          0);
        QCOMPARE(summary.details.size(),  2);

        db.close();
    }
};

QTEST_MAIN(tst_DatabaseManager)
#include "tst_databasemanager.moc"
