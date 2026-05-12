#include <QtTest/QtTest>
#include <QSqlQuery>
#include <QSettings>
#include <QCoreApplication>
#include "../../src/database/PresenceManager.h"
#include "../../src/database/PostgresConnection.h"
#include "../../src/database/IdentityManager.h"
#include "../../src/database/ConfigLoader.h"

using namespace DVE;

namespace {
DbConfig pgConfig() {
    DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
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

class TstPresenceManager : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
            QSKIP("DVE_TEST_PG_CONN not set");
        QCoreApplication::setOrganizationName("DataViewerTest");
        QCoreApplication::setApplicationName("tst_presencemanager");
    }

    // Per-slot setup: clear QSettings (so each test gets a fresh UUID) and
    // wipe presence rows so visibility windows don't leak across tests.
    void init() {
        QSettings s;
        s.clear();
        PostgresConnection pg;
        if (pg.open(pgConfig())) {
            QSqlQuery q(pg.queryDb());
            q.exec("DELETE FROM presence");
        }
    }

    void activate_insertsRow() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        IdentityManager id;
        id.setDisplayName("Alice");
        id.setColor("#ef4444");
        PresenceManager pm(&pg, &id);
        QVERIFY(pm.activate("file", 42, "viewing"));

        QSqlQuery q(pg.queryDb());
        QVERIFY(q.exec("SELECT user_name, intent FROM presence "
                       "WHERE resource_type='file' AND resource_id=42"));
        QVERIFY(q.next());
        QCOMPARE(q.value(0).toString(), QString("Alice"));
        QCOMPARE(q.value(1).toString(), QString("viewing"));
    }

    void activate_deletesPreviousActiveOfSameType() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        IdentityManager id;
        id.setDisplayName("Alice");
        id.setColor("#ef4444");
        PresenceManager pm(&pg, &id);

        QVERIFY(pm.activate("file", 1));
        QVERIFY(pm.activate("file", 2));

        QSqlQuery q(pg.queryDb());
        QVERIFY(q.exec("SELECT COUNT(*) FROM presence "
                       "WHERE resource_type='file'"));
        QVERIFY(q.next());
        QCOMPARE(q.value(0).toInt(), 1);   // only the latest survives
    }

    void setIntent_updatesRow() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        IdentityManager id;
        id.setDisplayName("Alice");
        id.setColor("#ef4444");
        PresenceManager pm(&pg, &id);
        QVERIFY(pm.activate("file", 1, "viewing"));
        QVERIFY(pm.setIntent("editing"));

        QSqlQuery q(pg.queryDb());
        q.exec("SELECT intent FROM presence WHERE resource_type='file' AND resource_id=1");
        QVERIFY(q.next());
        QCOMPARE(q.value(0).toString(), QString("editing"));
    }

    void deactivate_removesRow() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        IdentityManager id;
        id.setDisplayName("Alice");
        id.setColor("#ef4444");
        PresenceManager pm(&pg, &id);
        QVERIFY(pm.activate("file", 1));
        QVERIFY(pm.deactivate());

        QSqlQuery q(pg.queryDb());
        q.exec("SELECT COUNT(*) FROM presence WHERE resource_type='file' AND resource_id=1");
        QVERIFY(q.next());
        QCOMPARE(q.value(0).toInt(), 0);
    }

    void activeFor_returnsRowsForGivenResource() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        IdentityManager id;
        id.setDisplayName("Alice");
        id.setColor("#ef4444");
        PresenceManager pm(&pg, &id);
        QVERIFY(pm.activate("file", 7, "editing"));

        auto rows = pm.activeFor("file", 7);
        QCOMPARE(rows.size(), 1);
        QCOMPARE(rows[0].userName, QString("Alice"));
        QCOMPARE(rows[0].intent, QString("editing"));
        QCOMPARE(rows[0].userColor, QString("#ef4444"));
    }

    void activeFor_filtersStaleHeartbeats() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        IdentityManager id;
        id.setDisplayName("Alice");
        id.setColor("#ef4444");
        PresenceManager pm(&pg, &id);
        QVERIFY(pm.activate("file", 8));

        // Manually backdate the heartbeat
        QSqlQuery q(pg.queryDb());
        q.exec("UPDATE presence SET last_heartbeat = now() - INTERVAL '2 minutes' "
               "WHERE resource_type='file' AND resource_id=8");

        auto rows = pm.activeFor("file", 8);
        QCOMPARE(rows.size(), 0);   // stale row filtered out
    }
};

QTEST_MAIN(TstPresenceManager)
#include "tst_presencemanager.moc"
