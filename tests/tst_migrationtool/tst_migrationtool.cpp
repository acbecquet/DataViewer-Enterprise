#include <QtTest/QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QDir>
#include <QUuid>
#include "../../src/database/MigrationTool.h"

using DVE::MigrationTool;
using DVE::DbConfig;

namespace {
DbConfig pgConfig() {
    DbConfig c;
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

// Returns a unique path under QDir::tempPath() of the form
// "<temp>/tst_mig_<random>.sqlite". The file does NOT yet exist; the caller
// creates it via QSqlDatabase. Using a plain string (not QTemporaryFile)
// avoids the Windows file-handle/lock interaction QTemporaryFile has even
// after close() — that interaction blocks QFile::rename in finalizeSource.
QString tempSqlitePath() {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    return QDir::tempPath() + "/tst_mig_" + tag + ".sqlite";
}

QString makeEmptySqlite() {
    const QString path = tempSqlitePath();
    {
        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "tst_mig_seed");
        db.setDatabaseName(path);
        db.open();
        QSqlQuery q(db);
        q.exec("CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL, "
               "file_name TEXT NOT NULL, loaded_at TEXT NOT NULL, "
               "template_version TEXT, sheet_count INT, sample_count INT)");
        db.close();
    }
    QSqlDatabase::removeDatabase("tst_mig_seed");
    return path;
}
}

class TstMigrationTool : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) {
            QSKIP("DVE_TEST_PG_CONN not set; integration tests skipped");
        }
    }

    // Per-slot setup: wipe all Postgres tables so each test starts from a
    // known empty state. Otherwise leftover rows (e.g., file_path conflicts
    // from run_fullRoundTrip) cause UNIQUE-constraint failures in later runs.
    void init() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) return;
        DbConfig cfg = pgConfig();
        QSqlDatabase wipe = QSqlDatabase::addDatabase("QPSQL", "tst_init_wipe");
        wipe.setHostName(cfg.host); wipe.setPort(cfg.port);
        wipe.setDatabaseName(cfg.database); wipe.setUserName(cfg.user);
        wipe.setPassword(cfg.password);
        if (wipe.open()) {
            QSqlQuery q(wipe);
            // Delete child tables first to respect FK CASCADE order.
            for (const QString& t : QStringList{
                     "data_rows", "images", "samples", "tests", "files",
                     "sensory_images", "sensory_sessions",
                     "detailed_sensory_images", "detailed_sensory_sessions",
                     "settings", "schema_meta"
                 }) {
                q.exec("DELETE FROM " + t);
            }
            wipe.close();
        }
        QSqlDatabase::removeDatabase("tst_init_wipe");
    }

    void open_returnsTrue_withValidPaths() {
        MigrationTool m;
        const QString sqlitePath = makeEmptySqlite();
        QVERIFY(m.open(sqlitePath, pgConfig()));
        // Cleanup
        QFile::remove(sqlitePath);
    }

    void open_returnsFalse_whenSqliteMissing() {
        MigrationTool m;
        QVERIFY(!m.open("C:/no/such/file.sqlite", pgConfig()));
        QVERIFY(m.lastError().contains("not found", Qt::CaseInsensitive));
    }

    void migrate_files_roundTripsAllRowsWithIds() {
        const QString sqlitePath = tempSqlitePath();
        {
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "seed_files");
            db.setDatabaseName(sqlitePath);
            db.open();
            QSqlQuery q(db);
            q.exec("CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL, "
                   "file_name TEXT NOT NULL, loaded_at TEXT NOT NULL, "
                   "template_version TEXT, sheet_count INT, sample_count INT)");
            q.exec("INSERT INTO files VALUES (10, 'C:/a.xlsx', 'a.xlsx', '2026-01-01', 'v1', 2, 5)");
            q.exec("INSERT INTO files VALUES (20, 'C:/b.xlsx', 'b.xlsx', '2026-01-02', 'v1', 3, 6)");
            q.exec("INSERT INTO files VALUES (30, 'C:/c.xlsx', 'c.xlsx', '2026-01-03', 'v2', 1, 2)");
            db.close();
        }
        QSqlDatabase::removeDatabase("seed_files");

        MigrationTool m;
        QVERIFY(m.open(sqlitePath, pgConfig()));
        QVERIFY2(m.migrateTable("files"), qPrintable(m.lastError()));

        DbConfig cfg = pgConfig();
        QSqlDatabase chk = QSqlDatabase::addDatabase("QPSQL", "tst_chk_files");
        chk.setHostName(cfg.host); chk.setPort(cfg.port);
        chk.setDatabaseName(cfg.database); chk.setUserName(cfg.user);
        chk.setPassword(cfg.password);
        QVERIFY(chk.open());
        QSqlQuery q(chk);
        QVERIFY(q.exec("SELECT id, file_path FROM files WHERE id IN (10,20,30) ORDER BY id"));
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 10);
        QCOMPARE(q.value(1).toString(), QString("C:/a.xlsx"));
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 20);
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 30);
        QVERIFY(!q.next());
        chk.close();
        QSqlDatabase::removeDatabase("tst_chk_files");

        // Cleanup Postgres rows from this test
        QSqlDatabase cln = QSqlDatabase::addDatabase("QPSQL", "tst_cln_files");
        cln.setHostName(cfg.host); cln.setPort(cfg.port);
        cln.setDatabaseName(cfg.database); cln.setUserName(cfg.user);
        cln.setPassword(cfg.password);
        cln.open();
        QSqlQuery cq(cln);
        cq.exec("DELETE FROM files WHERE id IN (10, 20, 30)");
        cln.close();
        QSqlDatabase::removeDatabase("tst_cln_files");

        // Cleanup the SQLite source file
        QFile::remove(sqlitePath);
    }

    void run_fullRoundTrip_succeedsAndVerifies() {
        const QString sqlitePath = tempSqlitePath();
        {
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "seed_full");
            db.setDatabaseName(sqlitePath);
            db.open();
            QSqlQuery q(db);
            q.exec("CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL, "
                   "file_name TEXT NOT NULL, loaded_at TEXT NOT NULL, "
                   "template_version TEXT, sheet_count INT, sample_count INT)");
            q.exec("CREATE TABLE tests (id INTEGER PRIMARY KEY, file_id INT, "
                   "sheet_name TEXT, template_version TEXT, overall_avg_tpm REAL, "
                   "overall_stddev_tpm REAL, is_raw_table INT, sort_order INT)");
            q.exec("CREATE TABLE samples (id INTEGER PRIMARY KEY, test_id INT, "
                   "sort_order INT, sample_name TEXT, sample_id TEXT, date TEXT, "
                   "tester TEXT, media TEXT, viscosity REAL, resistance REAL, "
                   "voltage REAL, power REAL, heating_technology TEXT, "
                   "puffing_regime TEXT, initial_oil_mass REAL, average_tpm REAL, "
                   "stddev_tpm REAL, avg_power_density REAL, efficiency_percent REAL, "
                   "total_oil_consumed REAL, total_puffs INT, normalized_tpm REAL, "
                   "burn_status TEXT, clog_status TEXT, leak_status TEXT)");
            q.exec("CREATE TABLE data_rows (id INTEGER PRIMARY KEY, sample_id INT, "
                   "sort_order INT, puffs REAL, before_weight REAL, after_weight REAL, "
                   "draw_pressure REAL, resistance REAL, smell TEXT, clog TEXT, "
                   "notes TEXT, tpm REAL, tpm_power_density REAL, variation_tpm REAL, "
                   "oil_consumed REAL)");
            q.exec("CREATE TABLE images (id INTEGER PRIMARY KEY, sample_id INT, "
                   "sort_order INT, file_name TEXT, image_data BLOB, "
                   "layout_x REAL, layout_y REAL, layout_w REAL, layout_h REAL, "
                   "crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL)");
            q.exec("CREATE TABLE sensory_sessions (id INTEGER PRIMARY KEY, "
                   "session_name TEXT, tester_name TEXT, assessor_name TEXT, "
                   "media TEXT, puff_length TEXT, date TEXT, timestamp TEXT, "
                   "json_data TEXT, layout_json TEXT)");
            q.exec("CREATE TABLE sensory_images (id INTEGER PRIMARY KEY, session_id INT, "
                   "sort_order INT, file_name TEXT, image_data BLOB, "
                   "layout_x REAL, layout_y REAL, layout_w REAL, layout_h REAL, "
                   "crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL)");
            q.exec("CREATE TABLE detailed_sensory_sessions (id INTEGER PRIMARY KEY, "
                   "session_name TEXT, tester_name TEXT, assessor_name TEXT, "
                   "media TEXT, date TEXT, timestamp TEXT, json_data TEXT)");
            q.exec("CREATE TABLE detailed_sensory_images (id INTEGER PRIMARY KEY, "
                   "session_id INT, sort_order INT, file_name TEXT, image_data BLOB, "
                   "layout_x REAL, layout_y REAL, layout_w REAL, layout_h REAL, "
                   "crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL)");
            q.exec("CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT)");

            q.exec("INSERT INTO files VALUES (1, 'C:/a.xlsx', 'a.xlsx', "
                   "'2026-01-01', 'v1', 1, 1)");
            q.exec("INSERT INTO tests VALUES (1, 1, 'Sheet1', 'v1', 1.0, 0.1, 0, 0)");
            q.exec("INSERT INTO samples VALUES (1, 1, 0, 'S1', 'sid1', "
                   "'2026-01-01', 'Bob', 'oil', 1, 2, 3, 4, 'ceramic', 'CRM', "
                   "5, 6, 7, 8, 90, 9, 10, 11, 'no', 'no', 'no')");
            q.exec("INSERT INTO data_rows VALUES (1, 1, 0, 50, 1.0, 1.05, "
                   "100, 2, '', '', '', 1.0, 0.5, 0.05, 0.05)");
            q.exec("INSERT INTO sensory_sessions VALUES (1, 'Trial1', 'Bob', "
                   "'Sarah', 'oil', '3s', '2026-01-01', '2026-01-01T10:00:00', "
                   "'{\"k\":1}', '{}')");
            q.exec("INSERT INTO detailed_sensory_sessions VALUES (1, 'DTrial1', "
                   "'Bob', 'Sarah', 'oil', '2026-01-01', "
                   "'2026-01-01T10:00:00', '{\"q\":1}')");
            q.exec("INSERT INTO settings VALUES ('theme', 'dark')");
            db.close();
        }
        QSqlDatabase::removeDatabase("seed_full");

        MigrationTool m;
        QVERIFY(m.open(sqlitePath, pgConfig()));
        QVERIFY2(m.run(/*force=*/true), qPrintable(m.lastError()));

        const auto& r = m.report();
        QVERIFY2(r.summary().contains("status=success"), qPrintable(r.summary()));
        QVERIFY2(r.summary().contains("match=yes"), qPrintable(r.summary()));

        // Source SQLite must have been renamed by finalizeSource().
        const QString renamed = sqlitePath + ".pre-migration.sqlite";
        QVERIFY2(!QFile::exists(sqlitePath),
                 qPrintable("Source still exists at " + sqlitePath));
        QVERIFY2(QFile::exists(renamed),
                 qPrintable("Renamed file missing at " + renamed));

        // Cleanup
        QFile::remove(renamed);
    }
};

QTEST_MAIN(TstMigrationTool)
#include "tst_migrationtool.moc"
