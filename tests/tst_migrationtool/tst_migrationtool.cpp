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
};

QTEST_MAIN(TstMigrationTool)
#include "tst_migrationtool.moc"
