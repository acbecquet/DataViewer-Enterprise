#include <QtTest/QtTest>
#include <QSqlError>
#include "../../src/database/PostgresConnection.h"

using DVE::PostgresConnection;
using DVE::DbConfig;

namespace {
DbConfig configFromEnv() {
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
}

class TstPostgresConnection : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) {
            QSKIP("DVE_TEST_PG_CONN not set; integration tests skipped");
        }
    }

    void open_returnsTrue_whenServerAvailable() {
        PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));
        QVERIFY(pg.isOpen());
        pg.close();
        QVERIFY(!pg.isOpen());
    }

    void open_returnsFalse_whenServerUnreachable() {
        DbConfig bad = configFromEnv();
        bad.port = 1;
        PostgresConnection pg;
        QVERIFY(!pg.open(bad));
        QVERIFY(!pg.lastError().isEmpty());
    }

    void open_returnsFalse_withBadCredentials() {
        DbConfig bad = configFromEnv();
        bad.password = "wrong-password";
        PostgresConnection pg;
        QVERIFY(!pg.open(bad));
    }

    void ping_returnsTrue_whenConnected() {
        PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));
        QVERIFY(pg.ping());
        pg.close();
    }
};

QTEST_MAIN(TstPostgresConnection)
#include "tst_postgresconnection.moc"
