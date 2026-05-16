#include <QtTest/QtTest>
#include <QSignalSpy>
#include <QSqlQuery>
#include <QCoreApplication>
#include <QSettings>

#include "../../src/database/LiveSync.h"
#include "../../src/database/NotificationListener.h"
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

class TstLiveSync : public QObject
{
    Q_OBJECT

private slots:
    void initTestCase();
    void cleanupTestCase();

    void commitCell_writesScalarColumnAndBumpsVersion();
    void commitCell_writesJsonPathWithoutClobberingSiblings();
    void commitCell_notifyPayloadCarriesColumnAndValue();
    void focusCell_writesRowAndBlurDeletes();

private:
    PostgresConnection* m_conn = nullptr;
    IdentityManager*    m_identity = nullptr;
    LiveSync*           m_sync = nullptr;
    qint64              m_sampleId = -1;
    qint64              m_dataRowId = -1;
    qint64              m_sensorySessionId = -1;
};

void TstLiveSync::initTestCase()
{
    if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
        QSKIP("DVE_TEST_PG_CONN not set; skipping");
    QCoreApplication::setOrganizationName("DataViewerTest");
    QCoreApplication::setApplicationName("tst_livesync");

    QSettings s;
    s.clear();

    m_conn = new PostgresConnection(this);
    QVERIFY(m_conn->open(pgConfig()));

    m_identity = new IdentityManager(this);
    m_identity->setDisplayName("TestUser");
    m_identity->setColor("#ff0000");
    m_sync = new LiveSync(m_conn, m_identity, this);

    // Wipe any stale fixture rows from prior aborted runs.
    QSqlQuery cleanup(m_conn->queryDb());
    cleanup.exec("DELETE FROM files WHERE file_path = '/tmp/t.xlsx'");
    cleanup.exec("DELETE FROM sensory_sessions WHERE session_name = 'test'");
    cleanup.exec("DELETE FROM cell_focus");

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec("INSERT INTO files(file_path, file_name, loaded_at) "
                   "VALUES('/tmp/t.xlsx', 't.xlsx', '2026-05-16') RETURNING id"));
    QVERIFY(q.next());
    const qint64 fileId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString("INSERT INTO tests(file_id, sheet_name) "
                           "VALUES(%1, 'S1') RETURNING id").arg(fileId)));
    QVERIFY(q.next());
    const qint64 testId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString("INSERT INTO samples(test_id, sample_name) "
                           "VALUES(%1, 'A') RETURNING id").arg(testId)));
    QVERIFY(q.next());
    m_sampleId = q.value(0).toLongLong();

    QVERIFY(q.exec(QString(
        "INSERT INTO data_rows(sample_id, sort_order, draw_pressure) "
        "VALUES(%1, 0, 0.0) RETURNING id").arg(m_sampleId)));
    QVERIFY(q.next());
    m_dataRowId = q.value(0).toLongLong();

    QVERIFY(q.exec(
        "INSERT INTO sensory_sessions(session_name, json_data) "
        "VALUES('test', '{\"samples\":[{\"name\":\"D\",\"voltage\":3.5,"
        "\"power_type\":\"Constant Voltage\"}]}'::jsonb) RETURNING id"));
    QVERIFY(q.next());
    m_sensorySessionId = q.value(0).toLongLong();
}

void TstLiveSync::cleanupTestCase()
{
    if (m_conn && m_conn->isOpen()) {
        QSqlQuery q(m_conn->queryDb());
        q.exec("DELETE FROM files WHERE file_path = '/tmp/t.xlsx'");
        q.exec("DELETE FROM sensory_sessions WHERE session_name = 'test'");
        q.exec("DELETE FROM cell_focus");
    }
}

void TstLiveSync::commitCell_writesScalarColumnAndBumpsVersion()
{
    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString("SELECT version FROM data_rows WHERE id=%1")
                   .arg(m_dataRowId)));
    QVERIFY(q.next());
    const int beforeVersion = q.value(0).toInt();

    bool ok = m_sync->commitCell("data_rows", m_dataRowId,
                                 "draw_pressure", 1.42);
    QVERIFY(ok);

    QVERIFY(q.exec(QString(
        "SELECT draw_pressure, version FROM data_rows WHERE id=%1")
        .arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), 1.42);
    QCOMPARE(q.value(1).toInt(), beforeVersion + 1);
}

void TstLiveSync::commitCell_writesJsonPathWithoutClobberingSiblings()
{
    bool ok = m_sync->commitCell(
        "sensory_sessions", m_sensorySessionId,
        "json_path:samples[0].voltage", 4.2);
    QVERIFY(ok);

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString(
        "SELECT json_data->'samples'->0->>'voltage', "
        "       json_data->'samples'->0->>'name', "
        "       json_data->'samples'->0->>'power_type' "
        "FROM sensory_sessions WHERE id=%1").arg(m_sensorySessionId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), 4.2);
    QCOMPARE(q.value(1).toString(), QStringLiteral("D"));
    QCOMPARE(q.value(2).toString(), QStringLiteral("Constant Voltage"));
}

void TstLiveSync::commitCell_notifyPayloadCarriesColumnAndValue()
{
    // Subscribe to the channel via a second connection so we can verify
    // the trigger emits a column-aware payload. The whole point of the
    // set_config() calls in runScalarUpdate is to drive this payload --
    // without this test, a future refactor that drops those calls would
    // still pass the version/value assertions in the other slots.
    PostgresConnection foreign;
    QVERIFY(foreign.open(pgConfig()));
    NotificationListener listener(&foreign);
    QVERIFY(listener.subscribe());
    QSignalSpy spy(&listener, &NotificationListener::rowChanged);

    bool ok = m_sync->commitCell("data_rows", m_dataRowId,
                                 "draw_pressure", 2.5);
    QVERIFY(ok);

    QVERIFY(spy.wait(2000));
    bool found = false;
    while (spy.count() > 0) {
        const auto args = spy.takeFirst();
        const RowChange c = args.first().value<RowChange>();
        if (c.table == QStringLiteral("data_rows")
            && c.id == m_dataRowId
            && c.column == QStringLiteral("draw_pressure")) {
            QCOMPARE(c.newValue.toDouble(), 2.5);
            found = true;
            break;
        }
    }
    QVERIFY(found);
}

void TstLiveSync::focusCell_writesRowAndBlurDeletes()
{
    QVERIFY(m_sync->focusCell("data_rows", m_dataRowId, "draw_pressure"));

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString(
        "SELECT count(*) FROM cell_focus "
        "WHERE table_name='data_rows' AND row_id=%1 "
        "AND column_name='draw_pressure'").arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toInt(), 1);

    QVERIFY(m_sync->blurCell());

    QVERIFY(q.exec(QString(
        "SELECT count(*) FROM cell_focus "
        "WHERE table_name='data_rows' AND row_id=%1").arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toInt(), 0);
}

QTEST_MAIN(TstLiveSync)
#include "tst_livesync.moc"
