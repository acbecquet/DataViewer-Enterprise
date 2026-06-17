#include <QtTest/QtTest>
#include <QSignalSpy>
#include <QSqlQuery>
#include <QSqlError>
#include <QCoreApplication>
#include <QSettings>
#include <QElapsedTimer>
#include <QScopeGuard>
#include <QTemporaryDir>
#include <QFile>
#include <QDir>
#include <QFileInfo>

#include "../../src/database/LiveSync.h"
#include "../../src/database/OfflineSnapshot.h"
#include "../../src/utils/MipFallback.h"
#include "../../src/database/LiveSyncWorker.h"
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
    // v2.5.0 Task 4 (RC3) — worker recovers from a killed backend connection:
    // after pg_terminate_backend on the worker's PID, the next commit must
    // reconnect-and-retry and land in the DB. Without the fix the worker's
    // connection stays permanently broken (the 26000 "prepared statement does
    // not exist" loop seen in production) and the second commit never lands.
    void worker_reconnectsAfterBackendKill();
    // v2.5.0 Task 4 review (2d) — pure classifier coverage: feed synthetic
    // QSqlError values to LiveSyncWorker::isConnectionError and pin which ones
    // count as connection-shaped (trigger reconnect) vs statement-level (must
    // NOT reconnect). No live DB needed — guards the reconnect trigger logic
    // independently of the backend-kill integration test above.
    void isConnectionError_classifiesNeedles();
    // v2.4.2 Task 1 (R2): the app advertises a space-free application_name in
    // the connect string and it SURVIVES a reconnect (close+reopen rebuilds the
    // connection from the same string). No prior test asserts this.
    void connection_advertisesApplicationName_survivesReopen();
    // v2.4.2 Task 1 (R1): every connection sets statement_timeout via a SET in
    // the open path, reconnect-durable, and a long query fails fast instead of
    // hanging the thread forever (the GFW half-open-socket UI freeze).
    void connection_statementTimeoutSetAndBounded();
    // v2.4.2 SP2 Task 1 (R4 prereq): a writing session that has set the
    // dve.maintenance='1' GUC (the legacy-score normalizer) must NOT fan out a
    // NOTIFY per row — notify_row_change short-circuits — so a bulk normalize
    // can't storm every live client. Positive control: with the GUC off the
    // NOTIFY still arrives.
    void notify_suppressedWhenMaintenanceGucSet();
    // v2.4.2 SP2 Task 6 (R4b): a half-open NOTIFY socket (query socket still
    // answering SELECT 1) silently stopped live updates on flaky/GFW networks
    // because ConnectionMonitor only pinged the query socket. pingListen()
    // probes the LISTEN socket so a dead NOTIFY channel flips offline ->
    // reconnect. On a healthy connection BOTH ping() and pingListen() return
    // true.
    void pingListen_detectsListenSocket();
    // v2.4.2 SP2 Task 6 (R4b): resubscribeMissing() heals a partial channel
    // drop. Drop one channel (test-only helper), then resubscribeMissing()
    // must refill exactly that channel — replacing the old all-or-nothing
    // subscribe guard that left the surviving channels untouched but never
    // re-attached the dropped one.
    void resubscribeMissing_fillsDroppedChannel();
    // v2.4.4 R5: counter-pinning regression. With no live connection and an
    // undecodable (MIP-encrypted) offline queue, three offline commitCell()s
    // must surface an unsynced count of exactly 3 — one bump per failed enqueue.
    // The old quadratic bug double-bumped (handleEnqueueResult bumped AND the
    // caller bumped), which would surface 6. No live DB needed.
    void offlineEnqueueFailure_bumpsUnsyncedExactlyOncePerEdit();

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

void TstLiveSync::worker_reconnectsAfterBackendKill()
{
    if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
        QSKIP("DVE_TEST_PG_CONN not set; skipping");

    // Worker-wired LiveSync (the async path that owns its own connection).
    auto* tempIdentity = new IdentityManager(this);
    tempIdentity->setDisplayName("KillUser");
    tempIdentity->setColor("#123456");
    LiveSync* sync = new LiveSync(m_conn, tempIdentity);
    sync->setWorkerConfig(pgConfig());

    QSqlQuery q(m_conn->queryDb());

    // Snapshot the backend PIDs that exist BEFORE the worker opens its
    // connection: this set is { m_conn.query, m_conn.listen, ... }. The worker's
    // backend will be the one PID that is NOT in this set, letting us terminate
    // exactly the worker's connection and leave m_conn (used for readback) alive.
    // Filter to real client backends only: pg_cron (deployed via the
    // postgresql-16-cron image) runs a launcher + per-job background workers
    // whose PIDs flicker in and out of pg_stat_activity. Counting those in the
    // "new PID" set-difference could mis-target the terminate at a transient
    // cron worker instead of the LiveSync worker. backend_type='client backend'
    // + application_name NOT LIKE 'pg_cron%' restricts to genuine libpq
    // sessions, so the only NEW client backend after the prime is the worker's.
    static const char* kClientBackendFilter =
        "SELECT pid FROM pg_stat_activity "
        "WHERE datname = current_database() AND pid <> pg_backend_pid() "
        "AND backend_type = 'client backend' "
        "AND coalesce(application_name, '') NOT LIKE 'pg_cron%'";

    QSet<int> preWorkerPids;
    QVERIFY(q.exec(kClientBackendFilter));
    while (q.next()) preWorkerPids.insert(q.value(0).toInt());

    // Prime the worker: commit a known value and QTRY for it to land. This both
    // proves the worker connection is open and ensures a backend now exists for
    // us to kill. Use distinct sentinel values so the assertions are unambiguous.
    QVERIFY(sync->commitCell("sensory_sessions", m_sensorySessionId,
                             "json_path:samples[0].voltage", 1.1));
    QTRY_VERIFY2(
        ([&]() {
            return q.exec(QString("SELECT json_data->'samples'->0->>'voltage' "
                                  "FROM sensory_sessions WHERE id=%1")
                              .arg(m_sensorySessionId))
                && q.next() && qFuzzyCompare(q.value(0).toDouble(), 1.1);
        })(),
        "worker did not open its connection / first write never landed");

    // Identify the worker's backend PID (the new one not in the pre-worker set)
    // and terminate it. This mirrors production: libpq's socket dies and the
    // next statement on that connection fails with a connection-shaped error.
    int workerPid = -1;
    QVERIFY(q.exec(kClientBackendFilter));   // same client-backend filter as the pre-set
    while (q.next()) {
        const int pid = q.value(0).toInt();
        if (!preWorkerPids.contains(pid)) { workerPid = pid; break; }
    }
    QVERIFY2(workerPid > 0, "could not identify the worker's backend PID");

    QVERIFY(q.exec(QString("SELECT pg_terminate_backend(%1)").arg(workerPid)));
    QVERIFY(q.next());
    QVERIFY2(q.value(0).toBool(), "pg_terminate_backend reported failure");

    // Give Postgres a beat to actually tear the backend down so the worker's
    // next exec() genuinely hits a dead socket (not a still-closing one).
    QTRY_VERIFY2(
        ([&]() {
            return q.exec(QString("SELECT count(*) FROM pg_stat_activity "
                                  "WHERE pid = %1").arg(workerPid))
                && q.next() && q.value(0).toInt() == 0;
        })(),
        "worker backend did not terminate");

    // THE TEST: commit a NEW value. The worker's connection is now broken; the
    // fix must detect the connection-shaped error, stop()/start() a fresh
    // connection, replay the statement once, and land the value. Without the
    // fix this commit dies silently and the value stays 1.1.
    QVERIFY(sync->commitCell("sensory_sessions", m_sensorySessionId,
                             "json_path:samples[0].voltage", 2.2));

    QTRY_VERIFY2(
        ([&]() {
            return q.exec(QString("SELECT json_data->'samples'->0->>'voltage' "
                                  "FROM sensory_sessions WHERE id=%1")
                              .arg(m_sensorySessionId))
                && q.next() && qFuzzyCompare(q.value(0).toDouble(), 2.2);
        })(),
        "second commit never landed — worker did not reconnect after backend kill");

    // Sibling fields must be intact (the reconnect replayed the SAME jsonb_set
    // statement, not a clobbering whole-row write).
    QVERIFY(q.exec(QString(
        "SELECT json_data->'samples'->0->>'name', "
        "       json_data->'samples'->0->>'power_type' "
        "FROM sensory_sessions WHERE id=%1").arg(m_sensorySessionId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toString(), QStringLiteral("D"));
    QCOMPARE(q.value(1).toString(), QStringLiteral("Constant Voltage"));

    delete sync;          // joins/stops the worker thread
    delete tempIdentity;
}

void TstLiveSync::isConnectionError_classifiesNeedles()
{
    // No DB connection required — isConnectionError is a pure static classifier.
    // QSqlError(driverText, databaseText, type, nativeCode): text() == the two
    // text fields joined; nativeErrorCode() == the 4th arg.

    // !dbOpen short-circuits to true regardless of the error contents.
    QVERIFY2(LiveSyncWorker::isConnectionError(QSqlError(), /*dbOpen=*/false),
             "a closed db handle must classify as a connection error");

    // SQLSTATE 26000 — the production "unnamed prepared statement does not
    // exist" loop. Connection-shaped.
    {
        QSqlError e("", "prepared statement does not exist",
                    QSqlError::StatementError, "26000");
        QVERIFY2(LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "26000 must classify as a connection error");
    }

    // SQLSTATE class 08 — connection_exception (08006 = connection failure).
    {
        QSqlError e("", "connection failure", QSqlError::ConnectionError, "08006");
        QVERIFY2(LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "08006 must classify as a connection error");
    }

    // v2.4.2: SQLSTATE 25P02 (in_failed_sql_transaction) — a connection wedged
    // in an aborted transaction (e.g. after a statement_timeout cancel). Must
    // reconnect (a fresh backend clears it), not fail forever.
    {
        QSqlError e("", "current transaction is aborted",
                    QSqlError::StatementError, "25P02");
        QVERIFY2(LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "25P02 must classify as a connection error");
    }
    // 57014 (query_canceled / statement_timeout) is NOT connection-shaped: the
    // server is reachable and cancelled a slow query; reconnecting won't help.
    {
        QSqlError e("", "canceling statement due to statement timeout",
                    QSqlError::StatementError, "57014");
        QVERIFY2(!LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "57014 statement_timeout must NOT classify as a connection error");
    }

    // libpq text needle with no SQLSTATE the driver surfaces — "server
    // closed the connection unexpectedly".
    {
        QSqlError e("server closed the connection unexpectedly", "",
                    QSqlError::ConnectionError, "");
        QVERIFY2(LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "'server closed' text must classify as a connection error");
    }

    // NON-connection error: 23505 unique_violation is a constraint/statement
    // error. Must NOT trigger a reconnect (it would just fail again).
    {
        QSqlError e("", "duplicate key value violates unique constraint",
                    QSqlError::StatementError, "23505");
        QVERIFY2(!LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "23505 unique_violation must NOT classify as a connection error");
    }

    // Another non-connection error: 42601 syntax_error. Statement-level.
    {
        QSqlError e("", "syntax error at or near", QSqlError::StatementError, "42601");
        QVERIFY2(!LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "42601 syntax_error must NOT classify as a connection error");
    }
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

void TstLiveSync::connection_advertisesApplicationName_survivesReopen()
{
    auto appName = [&]() -> QString {
        QSqlQuery q(m_conn->queryDb());
        if (!q.exec("SELECT application_name FROM pg_stat_activity "
                    "WHERE pid = pg_backend_pid()") || !q.next())
            return QStringLiteral("<query-failed>");
        return q.value(0).toString();
    };
    // The connection opened in initTestCase must already advertise the name.
    QVERIFY2(appName().startsWith("DataViewer/"),
             qPrintable("application_name not advertised: " + appName()));

    // Reconnect-durable: close + reopen rebuilds from the same connect string.
    m_conn->close();
    QVERIFY(m_conn->open(pgConfig()));
    QVERIFY2(appName().startsWith("DataViewer/"),
             qPrintable("application_name lost after reopen: " + appName()));
}

void TstLiveSync::connection_statementTimeoutSetAndBounded()
{
    // statement_timeout applied by applyPgSessionSettings() in the open path.
    auto timeoutMs = [&]() -> int {
        QSqlQuery q(m_conn->queryDb());
        if (!q.exec("SHOW statement_timeout") || !q.next()) return -1;
        // SHOW returns e.g. "10s". Normalize to ms.
        const QString v = q.value(0).toString();
        if (v.endsWith("ms")) return v.left(v.size() - 2).toInt();
        if (v.endsWith("s"))  return v.left(v.size() - 1).toInt() * 1000;
        return v.toInt();
    };
    QCOMPARE(timeoutMs(), 10000);

    // A query longer than the timeout fails fast (does NOT hang). pg_sleep(20)
    // must be cancelled by statement_timeout (~10s) well under 20s.
    QElapsedTimer t; t.start();
    QSqlQuery q(m_conn->queryDb());
    const bool ok = q.exec("SELECT pg_sleep(20)");
    QVERIFY2(!ok, "pg_sleep(20) should have been cancelled by statement_timeout");
    QVERIFY2(t.elapsed() < 15000,
             qPrintable(QString("statement_timeout did not bound the query: %1ms")
                            .arg(t.elapsed())));
    // SQLSTATE 57014 = query_canceled (statement_timeout).
    QCOMPARE(q.lastError().nativeErrorCode(), QString("57014"));
}

void TstLiveSync::notify_suppressedWhenMaintenanceGucSet()
{
    PostgresConnection foreign;
    QVERIFY(foreign.open(pgConfig()));
    NotificationListener listener(&foreign);
    QVERIFY(listener.subscribe());
    QSignalSpy spy(&listener, &NotificationListener::rowChanged);

    QSqlQuery q(m_conn->queryDb());
    // RESET dve.maintenance on every return path. dve.maintenance is a
    // SESSION-level GUC on the shared m_conn, so if any QVERIFY/QCOMPARE below
    // fails (which does a plain `return;` out of the slot) the inline reset
    // would be skipped, leaving the shared connection suppressing NOTIFYs for
    // every slot declared after this one. The scope guard runs on all paths.
    auto gucGuard = qScopeGuard([&]{
        QSqlQuery r(m_conn->queryDb());
        r.exec(QStringLiteral("RESET dve.maintenance"));
    });

    // With dve.maintenance='1' set on the WRITING session, the trigger must
    // NOT emit a NOTIFY for this UPDATE.
    QVERIFY(q.exec("SET dve.maintenance = '1'"));
    QVERIFY(q.exec(QString("UPDATE data_rows SET draw_pressure = 3.21 "
                           "WHERE id=%1").arg(m_dataRowId)));
    QVERIFY2(!spy.wait(1500),
             "no rowChanged NOTIFY should arrive while dve.maintenance='1'");
    QCOMPARE(spy.count(), 0);

    // Positive control: clear the GUC, update again, NOTIFY must arrive.
    // (RESET in the scope guard supersedes this, but keep it for an explicit
    // in-test positive control of the GUC-off path.)
    QVERIFY(q.exec("SET dve.maintenance = '0'"));
    QVERIFY(q.exec(QString("UPDATE data_rows SET draw_pressure = 3.22 "
                           "WHERE id=%1").arg(m_dataRowId)));
    QVERIFY2(spy.wait(2000), "rowChanged NOTIFY should arrive with maintenance off");
}

void TstLiveSync::pingListen_detectsListenSocket()
{
    // Healthy connection: both the query socket and the LISTEN socket answer
    // SELECT 1. Before this fix only the query socket was probed, so a dead
    // LISTEN socket (the GFW half-open case) went undetected.
    QVERIFY(m_conn->ping());        // query socket
    QVERIFY(m_conn->pingListen());  // listen socket (NEW)
}

void TstLiveSync::resubscribeMissing_fillsDroppedChannel()
{
    PostgresConnection c;
    QVERIFY(c.open(pgConfig()));
    NotificationListener l(&c);
    QVERIFY(l.subscribe());
    QVERIFY(l.isSubscribedTo("dataviewer_changes"));

    // Simulate one channel dropped (test-only helper).
    l.unsubscribeFromChannelForTest("dataviewer_changes");
    QVERIFY(!l.isSubscribedTo("dataviewer_changes"));

    // resubscribeMissing() must refill exactly the dropped channel and report
    // success (all channels present again). The surviving channels are left
    // untouched (no double-subscribe).
    QVERIFY(l.resubscribeMissing());
    QVERIFY(l.isSubscribedTo("dataviewer_changes"));
    QVERIFY(l.isSubscribedTo("dataviewer_presence"));
    QVERIFY(l.isSubscribedTo("dataviewer_cell_focus"));

    // Idempotent: a second call with everything present is a no-op returning
    // true (every channel already subscribed -> size == channels().size()).
    QVERIFY(l.resubscribeMissing());
}

void TstLiveSync::offlineEnqueueFailure_bumpsUnsyncedExactlyOncePerEdit()
{
    // No live DB needed: this pins the COUNTER, not a write. We construct a
    // LiveSync with a NULL connection (so commitCell takes the offline-enqueue
    // path) and a real OfflineSnapshot whose pending-edits queue file is
    // MIP-encrypted junk (so enqueueCellEdit() returns false every time, just
    // like an undecodable queue on a live machine).
    QTemporaryDir tmp;
    QVERIFY(tmp.isValid());

    OfflineSnapshot snap;
    snap.setOverrideDirForTesting(tmp.path());

    // Plant a ciphertext-marker queue file so ensureQueueOpen() takes the
    // decrypt branch, fails to produce a valid SQLite db, and every
    // enqueueCellEdit() returns false.
    const QString qp = snap.queuePath();
    QDir().mkpath(QFileInfo(qp).absolutePath());
    {
        QByteArray bytes("%TSD-Header-###%");   // 16-byte MIP marker
        const char junk[] = { '\x01', '\x02', '\x00', ' ', 'n', 'o', 't',
                              ' ', 's', 'q', 'l', 'i', 't', 'e' };
        bytes.append(junk, int(sizeof(junk)));
        QFile f(qp);
        QVERIFY(f.open(QIODevice::WriteOnly));
        QVERIFY(f.write(bytes) == bytes.size());
        f.close();
    }
    QVERIFY(DVE::looksEncrypted(qp));
    QVERIFY(!snap.enqueueCellEdit("data_rows", 1, "notes", QVariant("probe")));

    IdentityManager id;
    id.setDisplayName("OfflineUser");
    id.setColor("#00aa00");

    // NULL connection -> commitCell hits the offline-enqueue branch.
    LiveSync sync(nullptr, &id);
    sync.setOfflineSnapshot(&snap);

    QSignalSpy unsyncedSpy(&sync, &LiveSync::unsyncedEditsChanged);
    QSignalSpy enqueueFailedSpy(&sync, &LiveSync::offlineEnqueueFailed);

    // Fire three distinct offline commits. Each fails to enqueue (degraded
    // queue) and must bump the tally EXACTLY once (not twice). commitCell still
    // returns true: the keystroke is preserved in the open session's dirty set.
    QVERIFY(sync.commitCell("data_rows", 1, "notes", QVariant("a")));
    QVERIFY(sync.commitCell("data_rows", 2, "notes", QVariant("b")));
    QVERIFY(sync.commitCell("data_rows", 3, "notes", QVariant("c")));

    // The crux: 3 edits -> count 3, NOT 6 (the old quadratic double-bump).
    QCOMPARE(sync.unsyncedEditCount(), 3);

    // Each failed enqueue surfaced the one-shot notification once per edit.
    QCOMPARE(enqueueFailedSpy.count(), 3);

    // unsyncedEditsChanged fired once per bump and the LAST emitted value is 3.
    QCOMPARE(unsyncedSpy.count(), 3);
    QCOMPARE(unsyncedSpy.last().at(0).toInt(), 3);

    // And the degraded flag is the channel carrying the "queue unreadable"
    // state (it is NOT folded into pendingCount(), which stays drainable-only).
    QVERIFY(sync.offlineQueueDegraded());
    QCOMPARE(sync.pendingCount(), 0);   // nothing drainable; degradation rides the flag
}

QTEST_MAIN(TstLiveSync)
#include "tst_livesync.moc"
