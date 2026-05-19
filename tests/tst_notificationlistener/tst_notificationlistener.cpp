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
            q.exec("DELETE FROM cell_focus");
            q.exec("DELETE FROM files WHERE file_path='C:/nl-test.xlsx'");
            q.exec("DELETE FROM files WHERE file_path='C:/nl-row-test.xlsx'");
        }
    }

    void subscribe_returnsTrue_whenConnected() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        NotificationListener nl(&pg);
        QVERIFY(nl.subscribe());
        QVERIFY(nl.isSubscribed());
    }

    // H5: per-channel subscription tracking. isSubscribedTo returns the
    // truth for each named channel after subscribe(), and false for
    // anything that wasn't attempted. The aggregate isSubscribed() is
    // still true when at least one channel attached.
    void subscribe_tracksEachChannelIndividually() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        NotificationListener nl(&pg);
        QVERIFY(nl.subscribe());
        QVERIFY(nl.isSubscribed());
        QVERIFY(nl.isSubscribedTo("dataviewer_changes"));
        QVERIFY(nl.isSubscribedTo("dataviewer_presence"));
        QVERIFY(nl.isSubscribedTo("dataviewer_cell_focus"));
        QVERIFY(!nl.isSubscribedTo("dataviewer_does_not_exist"));
    }

    // H5: unsubscribe clears the recorded set so isSubscribed flips back
    // to false and re-subscribing starts from a clean slate.
    void unsubscribe_clearsAllChannels() {
        PostgresConnection pg;
        QVERIFY(pg.open(pgConfig()));
        NotificationListener nl(&pg);
        QVERIFY(nl.subscribe());
        QVERIFY(nl.isSubscribed());
        nl.unsubscribe();
        QVERIFY(!nl.isSubscribed());
        QVERIFY(!nl.isSubscribedTo("dataviewer_changes"));
        QVERIFY(!nl.isSubscribedTo("dataviewer_presence"));
        QVERIFY(!nl.isSubscribedTo("dataviewer_cell_focus"));
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

    void rowChangedPayloadIncludesColumn() {
        PostgresConnection listener;
        QVERIFY(listener.open(pgConfig()));
        NotificationListener nl(&listener);
        QVERIFY(nl.subscribe());

        QSignalSpy spy(&nl, &NotificationListener::rowChanged);

        // Build a data_rows fixture row to UPDATE: needs file -> test -> sample.
        PostgresConnection writer;
        QVERIFY(writer.open(pgConfig()));
        QSqlQuery q(writer.queryDb());
        QVERIFY(q.exec("INSERT INTO files(file_path, file_name, loaded_at, "
                       "template_version, updated_by) VALUES "
                       "('C:/nl-row-test.xlsx', 'nl-row-test.xlsx', '2026-01-01', "
                       "'v1', 'fixture') RETURNING id"));
        QVERIFY(q.next());
        const qint64 fileId = q.value(0).toLongLong();

        QVERIFY(q.exec(QString("INSERT INTO tests(file_id, sheet_name, updated_by) "
                               "VALUES (%1, 'Sheet1', 'fixture') RETURNING id").arg(fileId)));
        QVERIFY(q.next());
        const qint64 testId = q.value(0).toLongLong();

        QVERIFY(q.exec(QString("INSERT INTO samples(test_id, sample_name, updated_by) "
                               "VALUES (%1, 'S1', 'fixture') RETURNING id").arg(testId)));
        QVERIFY(q.next());
        const qint64 sampleId = q.value(0).toLongLong();

        QVERIFY(q.exec(QString("INSERT INTO data_rows(sample_id, updated_by) "
                               "VALUES (%1, 'fixture') RETURNING id").arg(sampleId)));
        QVERIFY(q.next());
        const qint64 rowId = q.value(0).toLongLong();

        // Drain the spy of the fixture INSERT notifications.
        spy.wait(500);
        spy.clear();

        // Single-column UPDATE: trigger reads dve.live_column / dve.live_value
        // session vars to attach them to the payload.
        QVERIFY(q.exec("SELECT set_config('dve.live_column', 'draw_pressure', false)"));
        QVERIFY(q.exec("SELECT set_config('dve.live_value',  '1.42', false)"));
        QVERIFY(q.exec(QString("UPDATE data_rows SET draw_pressure = 1.42 "
                               "WHERE id = %1").arg(rowId)));

        QVERIFY(spy.wait(2000));
        QVERIFY(spy.count() >= 1);
        const RowChange c = spy.takeLast().at(0).value<RowChange>();
        QCOMPARE(c.table,  QString("data_rows"));
        QCOMPARE(c.column, QString("draw_pressure"));
        QCOMPARE(c.newValue.toDouble(), 1.42);
    }

    void cellFocusChannelEmitsSignal() {
        PostgresConnection listener;
        QVERIFY(listener.open(pgConfig()));
        NotificationListener nl(&listener);
        QVERIFY(nl.subscribe());

        QSignalSpy spy(&nl, &NotificationListener::cellFocusChanged);

        PostgresConnection writer;
        QVERIFY(writer.open(pgConfig()));
        QSqlQuery q(writer.queryDb());
        QVERIFY(q.exec(
            "INSERT INTO cell_focus(user_uuid, table_name, row_id, "
            "column_name, user_name, user_color) "
            "VALUES('11111111-1111-1111-1111-111111111111'::uuid, "
            "'data_rows', 1, 'draw_pressure', 'Tina', '#16a34a')"));

        QVERIFY(spy.wait(2000));
        QVERIFY(spy.count() >= 1);
        const CellFocusChange f = spy.takeLast().at(0).value<CellFocusChange>();
        QCOMPARE(f.op,         QString("INSERT"));
        QCOMPARE(f.userName,   QString("Tina"));
        QCOMPARE(f.userColor,  QString("#16a34a"));
        QCOMPARE(f.tableName,  QString("data_rows"));
        QCOMPARE(f.rowId,      qint64(1));
        QCOMPARE(f.columnName, QString("draw_pressure"));
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
