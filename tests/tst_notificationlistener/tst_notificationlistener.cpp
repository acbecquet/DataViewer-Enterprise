#include <QtTest/QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSignalSpy>
#include "../../src/database/NotificationListener.h"
#include "../../src/database/PostgresConnection.h"
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

class TstNotificationListener : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
            QSKIP("DVE_TEST_PG_CONN not set");
    }

    // Per-slot setup: wipe presence + the test file_path so tests
    // start from a known state. `files` is used (not `settings`) because
    // settings is intentionally excluded from the NOTIFY trigger array
    // — its PK is `key TEXT`, not `id`, so the trigger's NEW.id ref fails.
    void init() {
        PostgresConnection pg;
        if (pg.open(pgConfig())) {
            QSqlQuery q(pg.queryDb());
            q.exec("DELETE FROM presence");
            q.exec("DELETE FROM files WHERE file_path='C:/nl-test.xlsx'");
        }
    }

    void subscribe_returnsTrue_whenConnected() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        NotificationListener nl(&pg);
        QVERIFY(nl.subscribe());
        QVERIFY(nl.isSubscribed());
    }

    void rowChange_emits_on_insert() {
        PostgresConnection listener;
        QVERIFY(listener.open(pgConfig()));
        NotificationListener nl(&listener);
        QVERIFY(nl.subscribe());

        QSignalSpy spy(&nl, &NotificationListener::rowChanged);

        // Write from a separate connection so the listener sees a foreign change.
        // Use `files` (not `settings`) — settings is excluded from the NOTIFY
        // trigger because it lacks an `id` column.
        PostgresConnection writer;
        QVERIFY(writer.open(pgConfig()));
        QSqlQuery q(writer.queryDb());
        QVERIFY(q.exec("INSERT INTO files(file_path, file_name, loaded_at, "
                       "template_version, updated_by) "
                       "VALUES ('C:/nl-test.xlsx', 'nl-test.xlsx', '2026-01-01', 'v1', 'bob')"));

        QVERIFY(spy.wait(2000));
        QVERIFY(spy.count() >= 1);

        const RowChange c = spy.at(0).at(0).value<RowChange>();
        QCOMPARE(c.table, QString("files"));
        QCOMPARE(c.op, QString("INSERT"));
        QCOMPARE(c.updatedBy, QString("bob"));
    }

    void presenceChange_emits_on_insert() {
        PostgresConnection listener;
        QVERIFY(listener.open(pgConfig()));
        NotificationListener nl(&listener);
        QVERIFY(nl.subscribe());

        QSignalSpy spy(&nl, &NotificationListener::presenceChanged);

        PostgresConnection writer;
        QVERIFY(writer.open(pgConfig()));
        QSqlQuery q(writer.queryDb());
        const QString uuid = "00000000-0000-0000-0000-000000000001";
        QVERIFY(q.exec(QString(
            "INSERT INTO presence(user_uuid, user_name, user_color, "
            "resource_type, resource_id, intent) VALUES "
            "('%1'::uuid, 'B', '#3b82f6', 'file', 1, 'viewing')").arg(uuid)));

        QVERIFY(spy.wait(2000));
        QVERIFY(spy.count() >= 1);

        const PresenceChange p = spy.at(0).at(0).value<PresenceChange>();
        QCOMPARE(p.op, QString("INSERT"));
        QCOMPARE(p.resourceType, QString("file"));
        QCOMPARE(p.resourceId, qint64(1));
        QCOMPARE(p.intent, QString("viewing"));
    }
};

QTEST_MAIN(TstNotificationListener)
#include "tst_notificationlistener.moc"
