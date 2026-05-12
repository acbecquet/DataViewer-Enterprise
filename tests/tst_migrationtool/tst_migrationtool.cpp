#include <QtTest/QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QTemporaryFile>
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

QString makeEmptySqlite() {
    auto* tmp = new QTemporaryFile;
    tmp->setAutoRemove(false);
    tmp->open();
    tmp->close();
    {
        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "tst_mig_seed");
        db.setDatabaseName(tmp->fileName());
        db.open();
        QSqlQuery q(db);
        q.exec("CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL, "
               "file_name TEXT NOT NULL, loaded_at TEXT NOT NULL, "
               "template_version TEXT, sheet_count INT, sample_count INT)");
        db.close();
    }
    QSqlDatabase::removeDatabase("tst_mig_seed");
    return tmp->fileName();
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

    void open_returnsTrue_withValidPaths() {
        MigrationTool m;
        const QString sqlitePath = makeEmptySqlite();
        QVERIFY(m.open(sqlitePath, pgConfig()));
    }

    void open_returnsFalse_whenSqliteMissing() {
        MigrationTool m;
        QVERIFY(!m.open("C:/no/such/file.sqlite", pgConfig()));
        QVERIFY(m.lastError().contains("not found", Qt::CaseInsensitive));
    }

    void migrate_files_roundTripsAllRowsWithIds() {
        // Seed SQLite with 3 rows
        auto* tmp = new QTemporaryFile;
        tmp->setAutoRemove(false);
        tmp->open(); tmp->close();
        {
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "seed_files");
            db.setDatabaseName(tmp->fileName());
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
        QVERIFY(m.open(tmp->fileName(), pgConfig()));
        QVERIFY2(m.migrateTable("files"), qPrintable(m.lastError()));

        DbConfig cfg = pgConfig();
        QSqlDatabase chk = QSqlDatabase::addDatabase("QPSQL", "tst_chk_files");
        chk.setHostName(cfg.host); chk.setPort(cfg.port);
        chk.setDatabaseName(cfg.database); chk.setUserName(cfg.user);
        chk.setPassword(cfg.password);
        QVERIFY(chk.open());
        QSqlQuery q(chk);
        QVERIFY(q.exec("SELECT id, file_path FROM files ORDER BY id"));
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 10);
        QCOMPARE(q.value(1).toString(), QString("C:/a.xlsx"));
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 20);
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 30);
        QVERIFY(!q.next());
        chk.close();
        QSqlDatabase::removeDatabase("tst_chk_files");

        // Cleanup for next test runs: DELETE the rows we inserted
        QSqlDatabase cln = QSqlDatabase::addDatabase("QPSQL", "tst_cln_files");
        cln.setHostName(cfg.host); cln.setPort(cfg.port);
        cln.setDatabaseName(cfg.database); cln.setUserName(cfg.user);
        cln.setPassword(cfg.password);
        cln.open();
        QSqlQuery cq(cln);
        cq.exec("DELETE FROM files WHERE id IN (10, 20, 30)");
        cln.close();
        QSqlDatabase::removeDatabase("tst_cln_files");
    }
};

QTEST_MAIN(TstMigrationTool)
#include "tst_migrationtool.moc"
