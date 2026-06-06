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
    // v2.0.2 P1 OCC chain — C1 (stored fn enforces version) + C2
    // (caller reads BOOLEAN return). Hits the sync fallback path.
    void commitCell_occGuardRejectsStaleVersion();
    void commitCell_occGuardAcceptsCurrentVersion();
    // v2.0.2 P3 H1/H9 — destructor drains throttle queue and guards
    // against a dead worker thread.
    void destructor_flushesPendingCommits();
    void destructor_safeWhenWorkerThreadStopped();
    // DATAVIEWER-4 Task 3 — flushNowAndWait drains the throttle queue NOW
    // and blocks (bounded) until the worker has written everything queued.
    void flushNowAndWait_drainsPendingToDbSynchronously();

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
    // LiveSync queues the write on a 200ms throttle timer; wait for it
    // to fire (sync fallback runs from onThrottleTick).
    QTest::qWait(300);

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
    QTest::qWait(300);

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

void TstLiveSync::commitCell_occGuardRejectsStaleVersion()
{
    // Drive the row to a known state through the throttle path. qWait > 200ms
    // so the throttle timer's single-shot tick fires and the sync fallback
    // runs the actual UPDATE.
    QVERIFY(m_sync->commitCell("data_rows", m_dataRowId,
                               "draw_pressure", 5.0));
    QTest::qWait(300);

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString(
        "SELECT draw_pressure, version FROM data_rows WHERE id=%1")
        .arg(m_dataRowId)));
    QVERIFY(q.next());
    const double valueBefore   = q.value(0).toDouble();
    const int    versionBefore = q.value(1).toInt();
    QCOMPARE(valueBefore, 5.0);

    // Register a lookup that returns a STALE version (one less than current).
    // The stored function's WHERE id = ? AND version = ? matches zero rows
    // and the sync fallback emits commitConflict.
    m_sync->setVersionLookup(
        [versionBefore](const QString&, qint64) -> qint64 {
            return versionBefore - 1;  // stale
        });

    QSignalSpy conflictSpy(m_sync, &LiveSync::commitConflict);

    m_sync->commitCell("data_rows", m_dataRowId, "draw_pressure", 9.99);
    QTest::qWait(300);

    QCOMPARE(conflictSpy.count(), 1);
    const auto args = conflictSpy.takeFirst();
    QCOMPARE(args.at(0).toString(), QStringLiteral("data_rows"));
    QCOMPARE(args.at(1).toLongLong(), m_dataRowId);
    QCOMPARE(args.at(2).toString(), QStringLiteral("draw_pressure"));
    QCOMPARE(args.at(3).toDouble(), 9.99);
    QCOMPARE(args.at(4).toLongLong(), static_cast<qint64>(versionBefore - 1));

    // Row must be unchanged — stale write was rejected.
    QVERIFY(q.exec(QString(
        "SELECT draw_pressure, version FROM data_rows WHERE id=%1")
        .arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), valueBefore);
    QCOMPARE(q.value(1).toInt(), versionBefore);

    m_sync->setVersionLookup({});
}

void TstLiveSync::commitCell_occGuardAcceptsCurrentVersion()
{
    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString(
        "SELECT version FROM data_rows WHERE id=%1").arg(m_dataRowId)));
    QVERIFY(q.next());
    const int versionBefore = q.value(0).toInt();

    // Lookup returns the CURRENT version — OCC accepts the write.
    m_sync->setVersionLookup(
        [versionBefore](const QString&, qint64) -> qint64 {
            return versionBefore;
        });

    QSignalSpy conflictSpy(m_sync, &LiveSync::commitConflict);

    QVERIFY(m_sync->commitCell("data_rows", m_dataRowId,
                               "draw_pressure", 7.77));
    QTest::qWait(300);

    QCOMPARE(conflictSpy.count(), 0);

    QVERIFY(q.exec(QString(
        "SELECT draw_pressure, version FROM data_rows WHERE id=%1")
        .arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), 7.77);
    QCOMPARE(q.value(1).toInt(), versionBefore + 1);

    m_sync->setVersionLookup({});
}

void TstLiveSync::destructor_flushesPendingCommits()
{
    // H1: the 200 ms throttle queue must drain before the destructor
    // returns. Without the fix, deleting LiveSync while m_pendingCommits
    // has entries silently drops the last burst of edits.
    //
    // Use a dedicated LiveSync (not m_sync) so we can delete it cleanly.
    auto* tempIdentity = new IdentityManager(this);
    tempIdentity->setDisplayName("DtorUser");
    tempIdentity->setColor("#00ff00");
    LiveSync* sync = new LiveSync(m_conn, tempIdentity);

    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec(QString("SELECT version FROM data_rows WHERE id=%1")
                   .arg(m_dataRowId)));
    QVERIFY(q.next());
    const int beforeVersion = q.value(0).toInt();

    // Queue a commit but DO NOT call qWait — the 200 ms timer has not
    // fired yet, so the value sits in m_pendingCommits.
    QVERIFY(sync->commitCell("data_rows", m_dataRowId,
                             "draw_pressure", 8.88));

    delete sync;   // destructor must drain before returning

    QVERIFY(q.exec(QString(
        "SELECT draw_pressure, version FROM data_rows WHERE id=%1")
        .arg(m_dataRowId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), 8.88);
    QCOMPARE(q.value(1).toInt(), beforeVersion + 1);

    delete tempIdentity;
}

void TstLiveSync::destructor_safeWhenWorkerThreadStopped()
{
    // H9: the destructor invokes the worker stop slot via
    // BlockingQueuedConnection. If the worker thread already exited (e.g.
    // DB drop during shutdown), the blocking call would deadlock waiting
    // for a slot that never runs. The fix gates the invoke on
    // m_workerThread->isRunning().
    //
    // No worker is set on this LiveSync (no setWorkerConfig call), so
    // m_workerThread is null and the destructor short-circuits past the
    // worker stop. This proves the null-guard path is safe; the harder
    // case (worker present but thread already quit) would need a worker
    // PG config that we don't have a clean way to construct in-test.
    auto* tempIdentity = new IdentityManager(this);
    tempIdentity->setDisplayName("DtorUser2");
    tempIdentity->setColor("#0000ff");
    LiveSync* sync = new LiveSync(m_conn, tempIdentity);

    // Should complete without hanging or crashing.
    delete sync;
    delete tempIdentity;

    QVERIFY(true);
}

void TstLiveSync::flushNowAndWait_drainsPendingToDbSynchronously()
{
    if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
        QSKIP("DVE_TEST_PG_CONN not set; skipping");

    // This is the ONLY slot in this suite that exercises the async worker
    // path: flushNowAndWait's whole contract is "drain the throttle queue
    // and block until the BACKGROUND worker has actually written". A
    // dedicated worker-wired LiveSync (not m_sync, which uses the sync
    // fallback) keeps the assertion honest.
    auto* tempIdentity = new IdentityManager(this);
    tempIdentity->setDisplayName("FlushUser");
    tempIdentity->setColor("#abcdef");
    LiveSync* sync = new LiveSync(m_conn, tempIdentity);
    sync->setWorkerConfig(pgConfig());

    QSqlQuery q(m_conn->queryDb());

    // Avoid the worker-startup race: the worker opens its own QSqlDatabase
    // on its thread asynchronously. Prime it with one commit and QTRY for
    // the value to land, so by the time we do the *timed* assertion below
    // the worker connection is provably open. Without this, flushNowAndWait
    // could win the race against an unopened connection and the synced()
    // barrier would fire before the very first write.
    QVERIFY(sync->commitCell("sensory_sessions", m_sensorySessionId,
                             "json_path:samples[0].voltage", 4.4));
    QTRY_VERIFY2(
        ([&]() {
            return q.exec(QString("SELECT json_data->'samples'->0->>'voltage' "
                                  "FROM sensory_sessions WHERE id=%1")
                              .arg(m_sensorySessionId))
                && q.next() && qFuzzyCompare(q.value(0).toDouble(), 4.4);
        })(),
        "worker did not open its connection / first async write never landed");

    // Now the real test: queue an edit and confirm it is NOT yet written
    // (still in the in-process throttle queue), then flushNowAndWait must
    // make pendingCount()==0 AND have the value in the DB with NO polling.
    QVERIFY(sync->commitCell("sensory_sessions", m_sensorySessionId,
                             "json_path:samples[0].voltage", 8.0));
    QVERIFY(sync->pendingCount() > 0);

    sync->flushNowAndWait(4000);

    // Deterministic half: throttle queue drained synchronously.
    QCOMPARE(sync->pendingCount(), 0);

    // Proof the wait actually blocked until the background write landed:
    // a single read-back with NO qWait / QTRY.
    QVERIFY(q.exec(QString("SELECT json_data->'samples'->0->>'voltage' "
                           "FROM sensory_sessions WHERE id=%1")
                       .arg(m_sensorySessionId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toDouble(), 8.0);

    delete sync;          // joins/stops the worker thread
    delete tempIdentity;
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
