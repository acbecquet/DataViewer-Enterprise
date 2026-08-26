#include "LiveSync.h"
#include "LiveSyncWorker.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "NotificationListener.h"
#include "OfflineSnapshot.h"

#include <QHash>
#include <QSet>
#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QStringList>
#include <QRegularExpression>
#include <QUuid>
#include <QTimer>
#include <QThread>
#include <QEventLoop>

namespace DVE {

LiveSync::LiveSync(PostgresConnection* conn, IdentityManager* identity,
                   QObject* parent)
    : QObject(parent), m_conn(conn), m_identity(identity)
{
    m_throttleTimer = new QTimer(this);
    m_throttleTimer->setSingleShot(true);
    m_throttleTimer->setInterval(200);
    connect(m_throttleTimer, &QTimer::timeout, this, &LiveSync::onThrottleTick);
}

LiveSync::~LiveSync()
{
    // H1: flush the throttle queue synchronously so the last 200 ms of
    // edits don't vanish when the app closes. onThrottleTick() dispatches
    // each pending commit through the same path commitCell uses (worker
    // when wired, sync fallback otherwise).
    if (m_throttleTimer) m_throttleTimer->stop();
    onThrottleTick();

    // H9: only invoke the worker stop via BlockingQueuedConnection when
    // the worker's event loop is still running. Otherwise the blocking
    // call would deadlock waiting for a slot that never executes — a
    // scenario the v2.0.1 destructor could hit when the DB drops during
    // shutdown and the worker thread exited via QSqlDatabase error.
    if (m_workerThread) {
        if (m_worker && m_workerThread->isRunning()) {
            QMetaObject::invokeMethod(m_worker, "stop", Qt::BlockingQueuedConnection);
        }
        m_workerThread->quit();
        m_workerThread->wait(2000);
    }
}

void LiveSync::setWorkerConfig(const DbConfig& cfg)
{
    if (m_workerThread) return;
    m_workerThread = new QThread(this);
    m_worker       = new LiveSyncWorker(cfg);
    m_worker->moveToThread(m_workerThread);
    connect(m_workerThread, &QThread::finished, m_worker, &QObject::deleteLater);
    connect(m_worker, &LiveSyncWorker::commitFailed,
            this,     &LiveSync::onWorkerCommitFailed);
    connect(m_worker, &LiveSyncWorker::commitConflict,
            this,     &LiveSync::onWorkerCommitConflict);
    connect(m_worker, &LiveSyncWorker::measurementFailed,
            this,     &LiveSync::onWorkerMeasurementFailed);
    m_workerThread->start();
    QMetaObject::invokeMethod(m_worker, "start", Qt::QueuedConnection);
}

void LiveSync::setVersionLookup(VersionLookup lookup)
{
    m_versionLookup = std::move(lookup);
}

qint64 LiveSync::currentVersionFor(const QString& table, qint64 rowId) const
{
    if (!m_versionLookup) return -1;
    return m_versionLookup(table, rowId);
}

bool LiveSync::isLiveSyncTable(const QString& t)
{
    // v3 Phase 3d: the data_rows / samples / tests / files entries are GONE.
    // The 2026-07-30 audit found zero commitCell callers for tests and files
    // (dead allowlist entries), and the one data_rows caller
    // (MainWindow::onStoryCellEdited) now goes through commitMeasurement -
    // post-cutover data_rows is a read-only view and dve_commit_cell's UPDATE
    // on it can only fail. Sensory stays: json_path commits are its live path.
    return t == QLatin1String("sensory_sessions")
        || t == QLatin1String("detailed_sensory_sessions");
}

bool LiveSync::isLiveSyncColumn(const QString& table, const QString& column)
{
    static const QHash<QString, QSet<QString>> kAllowed = {
        { QStringLiteral("sensory_sessions"), {
            QStringLiteral("session_name"), QStringLiteral("tester_name"),
            QStringLiteral("assessor_name"), QStringLiteral("media"),
            QStringLiteral("puff_length"), QStringLiteral("date"),
            QStringLiteral("json_data")
        }},
        { QStringLiteral("detailed_sensory_sessions"), {
            QStringLiteral("session_name"), QStringLiteral("tester_name"),
            QStringLiteral("assessor_name"), QStringLiteral("media"),
            QStringLiteral("date"), QStringLiteral("json_data")
        }}
    };
    auto it = kAllowed.constFind(table);
    if (it == kAllowed.constEnd()) return false;
    return it.value().contains(column);
}

bool LiveSync::commitCell(const QString& table, qint64 rowId,
                          const QString& column, const QVariant& value,
                          bool allowQueue)
{
    if (!isLiveSyncTable(table)) {
        qWarning() << "LiveSync::commitCell unknown table" << table;
        return false;
    }
    const bool isJsonPath = column.startsWith(QLatin1String("json_path:"));
    const QString gateCol = isJsonPath ? QStringLiteral("json_data") : column;
    if (!isLiveSyncColumn(table, gateCol)) {
        qWarning() << "LiveSync::commitCell column not allowed for"
                   << table << ":" << column;
        return false;
    }

    if (!m_conn || !m_conn->isOpen()) {
        if (allowQueue && m_snapshot) {
            // v2.4.4 R5: do NOT swallow a failed offline enqueue. If the local
            // queue can't accept the edit (e.g. MIP-encrypted pending_edits),
            // warn loudly via handleEnqueueResult and count it as one unsynced
            // edit (it reached neither PG nor the local queue). The keystroke is
            // still preserved in the open session's dirty set, so we report
            // success to the caller (the UI doesn't reject the keystroke).
            const bool ok =
                m_snapshot->enqueueCellEdit(table, rowId, column, value);
            handleEnqueueResult(ok, table, rowId, column, value);
            if (!ok) bumpUnsynced();   // IMPORTANT 1: single-owner tally bump
            return true;
        }
        return false;
    }

    m_pendingCommits.insert({table, rowId, column, -1}, value);
    if (!m_throttleTimer->isActive()) m_throttleTimer->start();
    return true;
}

bool LiveSync::commitMeasurement(qint64 sampleId, const QString& key,
                                 int sortOrder, const QVariant& value,
                                 bool allowQueue)
{
    if (sortOrder < 0) {
        qWarning() << "LiveSync::commitMeasurement negative sortOrder" << sortOrder
                   << "for" << key;
        return false;
    }

    if (!m_conn || !m_conn->isOpen()) {
        if (allowQueue && m_snapshot) {
            // Same offline contract as commitCell: enqueue (schema_version=2),
            // surface a failed enqueue, count it once. The keystroke is still
            // in the open session's dirty set either way.
            const bool ok = m_snapshot->enqueueMeasurementEdit(
                sampleId, key, sortOrder, value);
            handleEnqueueResult(ok, QStringLiteral("measurements"), sampleId,
                                key, value);
            if (!ok) bumpUnsynced();
            return true;
        }
        return false;
    }

    m_pendingCommits.insert(
        {QStringLiteral("measurements"), sampleId, key, sortOrder}, value);
    if (!m_throttleTimer->isActive()) m_throttleTimer->start();
    return true;
}

void LiveSync::onThrottleTick()
{
    if (m_pendingCommits.isEmpty()) return;
    const auto drained = m_pendingCommits;
    m_pendingCommits.clear();
    for (auto it = drained.constBegin(); it != drained.constEnd(); ++it) {
        dispatchCommit(it.key().table, it.key().rowId,
                       it.key().column, it.value(), it.key().sortOrder);
    }
}

void LiveSync::dispatchCommit(const QString& table, qint64 rowId,
                              const QString& column, const QVariant& value,
                              int sortOrder)
{
    // v3 Phase 3d: a sortOrder >= 0 marks a measurement entry (see PendingKey).
    // rowId is the SAMPLE id and column is the metric key.
    const bool isMeasurement = sortOrder >= 0;
    const bool isJsonPath = column.startsWith(QLatin1String("json_path:"));
    if (!m_conn || !m_conn->isOpen()) {
        // v2.4.4 R5: surface a failed offline enqueue instead of discarding the
        // bool (a connection that dropped between commitCell and the throttle
        // tick lands here). On enqueue failure count one unsynced edit through
        // the single-owner tally (IMPORTANT 1).
        if (m_snapshot) {
            const bool ok = isMeasurement
                ? m_snapshot->enqueueMeasurementEdit(rowId, column, sortOrder, value)
                : m_snapshot->enqueueCellEdit(table, rowId, column, value);
            handleEnqueueResult(ok, table, rowId, column, value);
            if (!ok) bumpUnsynced();
        }
        return;
    }
    const QString uuid = m_identity
        ? m_identity->uuid().toString(QUuid::WithoutBraces)
        : QString();

    // v2.0.2: resolve the caller's expected version for this row. -1
    // (no lookup registered, or row not in cache) means "no OCC" and
    // the stored function falls back to its v2.0.1 behavior.
    const qint64 expectedVersion = currentVersionFor(table, rowId);

    if (m_worker) {
        if (isMeasurement) {
            QMetaObject::invokeMethod(m_worker, "commitMeasurement", Qt::QueuedConnection,
                Q_ARG(qint64, rowId),   Q_ARG(QString, column),
                Q_ARG(int, sortOrder),  Q_ARG(QVariant, value),
                Q_ARG(QString, uuid));
        } else if (isJsonPath) {
            const QString path = column.mid(QStringLiteral("json_path:").size());
            QMetaObject::invokeMethod(m_worker, "commitJson", Qt::QueuedConnection,
                Q_ARG(QString, table), Q_ARG(qint64, rowId),
                Q_ARG(QString, path),  Q_ARG(QVariant, value),
                Q_ARG(QString, uuid),  Q_ARG(qint64, expectedVersion));
        } else {
            QMetaObject::invokeMethod(m_worker, "commitScalar", Qt::QueuedConnection,
                Q_ARG(QString, table),  Q_ARG(qint64, rowId),
                Q_ARG(QString, column), Q_ARG(QVariant, value),
                Q_ARG(QString, uuid),   Q_ARG(qint64, expectedVersion));
        }
        return;
    }

    if (isMeasurement) {
        runMeasurementUpdateSync(rowId, column, sortOrder, value);
    } else if (isJsonPath) {
        runJsonPathUpdateSync(table, rowId,
            column.mid(QStringLiteral("json_path:").size()), value);
    } else {
        runScalarUpdateSync(table, rowId, column, value);
    }
}

void LiveSync::onWorkerCommitFailed(QString table, qint64 rowId,
                                    QString column, QVariant value)
{
    // Driver-level failure (network drop etc.) — preserve the user's
    // value for replay when we reconnect. OCC misses go through the
    // onWorkerCommitConflict path and DO NOT enqueue.
    //
    // v2.5.0 Task 4 (RC3): this fires only AFTER the worker's reconnect-and-
    // retry-once already failed, so the edit genuinely did not land. Surface
    // it (the dirty-cell merge still protects it on the next whole-save).
    if (m_snapshot) {
        // v2.4.4 R5: a driver-level failure means the edit didn't land remotely;
        // if it ALSO can't be queued locally, that's a double failure we must
        // not hide. handleEnqueueResult warns + emits the one-shot signal but
        // does NOT bump the tally (IMPORTANT 1) -- the bumpUnsynced() below
        // counts this failed commit exactly once regardless of the enqueue
        // outcome. The edit is still preserved in the open session's dirty set
        // and re-persisted by the next whole-session save (IMPORTANT 2).
        const bool ok =
            m_snapshot->enqueueCellEdit(table, rowId, column, value);
        handleEnqueueResult(ok, table, rowId, column, value);
    }
    bumpUnsynced();
}

void LiveSync::onWorkerMeasurementFailed(qint64 sampleId, QString key,
                                         int sortOrder, QVariant value)
{
    // Mirrors onWorkerCommitFailed: driver-level failure, so preserve the edit
    // for replay (schema_version=2 queue row) and count it once.
    if (m_snapshot) {
        const bool ok = m_snapshot->enqueueMeasurementEdit(
            sampleId, key, sortOrder, value);
        handleEnqueueResult(ok, QStringLiteral("measurements"), sampleId,
                            key, value);
    }
    bumpUnsynced();
}

void LiveSync::onWorkerCommitConflict(QString table, qint64 rowId,
                                      QString column, QVariant value,
                                      qint64 expectedVersion)
{
    qWarning() << "LiveSync OCC miss:" << table << column << "row" << rowId
               << "expected v" << expectedVersion;
    // v2.5.0 Task 4 (RC3): a conflict means the local per-cell write did NOT
    // apply (version moved / row gone). Count it as unsynced so the user sees
    // the cell isn't live; the next whole-session save reconciles it.
    bumpUnsynced();
    emit commitConflict(table, rowId, column, value, expectedVersion);
}

void LiveSync::bumpUnsynced()
{
    ++m_unsyncedEdits;
    emit unsyncedEditsChanged(m_unsyncedEdits);
}

void LiveSync::handleEnqueueResult(bool ok, const QString& table, qint64 rowId,
                                   const QString& column, const QVariant& value)
{
    if (ok) return;   // queued cleanly -- nothing to surface

    Q_UNUSED(value);

    // v2.4.4 R5: the offline queue refused the edit (most likely a
    // MIP-encrypted pending_edits.sqlite the python fallback couldn't decode).
    // Before this fix the bool was discarded at every call site and the edit
    // vanished silently.
    //
    // IMPORTANT 2: the edit is NOT lost and is NOT buffered here. The keystroke
    // that triggered this already recorded the cell in the open session's dirty
    // set, so the next online whole-session save re-persists it. We emit a
    // one-shot notification (no count) so MainWindow can warn ONCE; it must not
    // re-increment the counter on that signal.
    //
    // IMPORTANT 1: this helper does NOT bump the unsynced tally itself. The
    // unsynced TALLY has a single owner (m_unsyncedEdits via bumpUnsynced). Each
    // caller decides whether the failed enqueue is a NEW unsynced edit: the two
    // pure-offline sites bump once (the edit reached neither PG nor the local
    // queue); onWorkerCommitFailed already bumped for its commit failure and
    // must NOT bump again here. Centralising the bump in this helper would
    // double-count on the worker path.
    qWarning().nospace()
        << "LiveSync: offline enqueue FAILED for " << table << " row " << rowId
        << " column " << column << " -- the local pending-edits queue is "
           "unreadable/unwritable (MIP-encrypted?). The edit stays in the open "
           "session and will be written on the next online save (Ctrl+U).";
    emit offlineEnqueueFailed();
}

void LiveSync::setOfflineSnapshot(OfflineSnapshot* snap)
{
    m_snapshot = snap;
}

int LiveSync::flushPending()
{
    if (!m_snapshot) return 0;
    if (!m_conn || !m_conn->isOpen()) return 0;

    // v3 Phase 3d (hazard H9): three queue generations drain differently.
    //  - schema_version=2 measurement rows: replay through commitMeasurement.
    //  - v1 rows for data_rows: enqueued by a PRE-cutover build against real
    //    data_rows ids. Post-cutover, translate id -> (sample_id, sort_order)
    //    through the frozen pre-image data_rows_pre_v3, then commit as a
    //    measurement. A row whose id is not there (deleted before cutover)
    //    has no honest target: drop it with a warning and one unsynced bump -
    //    returning true so the drain clears it instead of retrying forever.
    //    Pre-cutover (no pre-image table), v1 data_rows rows replay through
    //    commitCell exactly as they always did... except the 3d allowlist no
    //    longer contains data_rows, so a direct dve_commit_cell replay is
    //    used - the queue row was validated against the OLD allowlist when it
    //    was enqueued.
    //  - v1 sensory rows: unchanged, commitCell path.
    const bool postCutover = [this]() {
        QSqlQuery p(m_conn->queryDb());
        return p.exec(QStringLiteral(
                   "SELECT 1 FROM pg_class c WHERE c.oid = to_regclass('data_rows_pre_v3')"))
            && p.next();
    }();

    return m_snapshot->drainPendingEdits(
        [this, postCutover](const QString& t, qint64 r, const QString& c,
                            const QVariant& v, int schemaVersion, int sortOrder) {
            if (schemaVersion >= 2)
                return this->commitMeasurement(r, c, sortOrder, v,
                                               /*allowQueue=*/false);
            if (t == QLatin1String("data_rows")) {
                if (!postCutover) {
                    // Pre-cutover replay: the legacy stored-function path,
                    // bypassing the (now sensory-only) allowlist gate.
                    return this->runScalarUpdateSync(t, r, c, v);
                }
                QSqlQuery look(m_conn->queryDb());
                look.prepare(QStringLiteral(
                    "SELECT sample_id, sort_order FROM data_rows_pre_v3 WHERE id = ?"));
                look.addBindValue(r);
                if (look.exec() && look.next()) {
                    return this->commitMeasurement(look.value(0).toLongLong(), c,
                                                   look.value(1).toInt(), v,
                                                   /*allowQueue=*/false);
                }
                qWarning() << "LiveSync::flushPending: v1 data_rows pending edit"
                           << "row" << r << "column" << c
                           << "has no data_rows_pre_v3 pre-image (deleted before"
                              " the cutover) -- dropping it";
                bumpUnsynced();
                return true;   // clear the queue row; it can never land
            }
            return this->commitCell(t, r, c, v, /*allowQueue=*/false);
        });
}

// C4: total queued work — throttle queue plus persistent snapshot queue.
// Callers (MainWindow::flushPendingEdits) gate file-level saveFile on this
// returning 0 so file-level UPDATEs don't race the per-cell drain. Returning
// the sum is intentional even though the throttle queue is in-process: a
// caller asserting "drain complete before save" needs both to be empty.
//
// v2.4.4 R5: this reports ONLY actionable/drainable work — in-memory pending
// commits plus the real queued-row count. A degraded (undecodable) queue does
// NOT contribute here: an undecodable queue can never drain, so folding it into
// the count would make pendingCount()>=1 forever and trip the
// Q_ASSERT(pendingCount()==0) invariant in MainWindow::flushPendingEdits on
// every flush. The "queue unreadable" state is signalled SOLELY via
// offlineQueueDegraded() (consumed by the sync indicator); pendingCount() stays
// a pure drainable-work counter so its ==0 invariant holds.
int LiveSync::pendingCount() const
{
    int n = m_pendingCommits.size();
    if (m_snapshot)
        n += m_snapshot->pendingEditCount();
    return n;
}

bool LiveSync::offlineQueueDegraded() const
{
    return m_snapshot && m_snapshot->queueDegraded();
}

bool LiveSync::flushNowAndWait(int timeoutMs)
{
    // Re-entrancy guard: a nested QEventLoop (below) keeps delivering posted
    // events, which could re-enter this method via another deferred persist
    // call. A nested flush would just wait on the same worker again, so make
    // it a no-op.
    if (m_flushing) return false;
    m_flushing = true;

    // Dispatch everything currently queued in the 200 ms throttle window to
    // the worker RIGHT NOW (don't wait for the timer). onThrottleTick() runs
    // the same dispatch path commitCell uses — worker (queued) when wired,
    // sync fallback otherwise.
    if (m_throttleTimer && m_throttleTimer->isActive())
        m_throttleTimer->stop();
    onThrottleTick();

    // No worker → the sync fallback inside onThrottleTick already executed the
    // UPDATEs on this thread, so the edits are durable; nothing to wait for.
    if (!m_worker || !m_workerThread) { m_flushing = false; return true; }

    // Wait, bounded by timeoutMs, for the worker to drain everything queued
    // above. sync() is posted via QueuedConnection so it runs only AFTER every
    // commitScalar/commitJson already in the worker's FIFO event queue; its
    // synced() reply therefore means "all preceding writes have landed". The
    // single-shot timer caps the wait so a stalled NAS can't freeze the UI.
    //
    // v2.5.0 Task 3 (RC2): a `synced` flag distinguishes a real drain from a
    // timeout so the bool return is meaningful to callers.
    // v2.5.0 Task 4 (RC3): snapshot the failure count BEFORE the barrier. The
    // worker thread processes its FIFO in order, so any commitFailed/
    // commitConflict for edits queued before this sync() is delivered to the UI
    // thread (and bumps m_unsyncedEdits) ahead of synced(). After a confirmed
    // drain we therefore clear exactly the failures that predated this drain —
    // failures that occurred DURING the drain are kept so they stay visible.
    const int preDrainUnsynced = m_unsyncedEdits;

    bool synced = false;
    QEventLoop loop;
    QTimer timeout;
    timeout.setSingleShot(true);
    connect(&timeout, &QTimer::timeout, &loop, &QEventLoop::quit);
    connect(m_worker, &LiveSyncWorker::synced, &loop, [&loop, &synced]() {
        synced = true;
        loop.quit();
    });
    QMetaObject::invokeMethod(m_worker, "sync", Qt::QueuedConnection);
    timeout.start(timeoutMs);
    loop.exec();

    // On a confirmed drain, subtract the pre-drain failures (now reconciled by
    // the whole-session save the caller is about to run). NB: this does NOT
    // recover lost edits — it clears the "failures pending from before this
    // drain" signal. The dirty-cell merge (Task 3) is what re-persists them.
    if (synced && preDrainUnsynced > 0) {
        const int remaining = m_unsyncedEdits - preDrainUnsynced;
        m_unsyncedEdits = remaining > 0 ? remaining : 0;
        emit unsyncedEditsChanged(m_unsyncedEdits);
    }

    m_flushing = false;
    return synced;
}

bool LiveSync::runScalarUpdateSync(const QString& table, qint64 rowId,
                                   const QString& column, const QVariant& value)
{
    const qint64 expectedVersion = currentVersionFor(table, rowId);
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("SELECT dve_commit_cell(?, ?, ?, ?, ?, ?)"));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(value.toString());
    q.addBindValue(m_identity
        ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString());
    if (expectedVersion < 0) {
        q.addBindValue(QVariant(QMetaType(QMetaType::Int)));
    } else {
        q.addBindValue(static_cast<int>(expectedVersion));
    }
    if (!q.exec()) {
        qWarning() << "LiveSync sync scalar failed:" << q.lastError().text();
        return false;
    }
    const bool ok = q.next() && q.value(0).toBool();
    if (!ok) {
        emit commitConflict(table, rowId, column, value, expectedVersion);
    }
    return ok;
}

bool LiveSync::runJsonPathUpdateSync(const QString& table, qint64 rowId,
                                     const QString& jsonPath, const QVariant& value)
{
    QStringList parts;
    static const QRegularExpression re(QStringLiteral(R"(([^.\[\]]+)|\[(\d+)\])"));
    auto it = re.globalMatch(jsonPath);
    while (it.hasNext()) {
        const auto m = it.next();
        parts << (m.captured(1).isEmpty() ? m.captured(2) : m.captured(1));
    }
    if (parts.isEmpty()) return false;
    const QString pgPath = QStringLiteral("{%1}").arg(parts.join(QLatin1Char(',')));

    const qint64 expectedVersion = currentVersionFor(table, rowId);
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral(
        "SELECT dve_commit_cell_json(?, ?, ?, ?::text[], ?, ?, ?)"));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(jsonPath);
    q.addBindValue(pgPath);
    q.addBindValue(value.toString());
    q.addBindValue(m_identity
        ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString());
    if (expectedVersion < 0) {
        q.addBindValue(QVariant(QMetaType(QMetaType::Int)));
    } else {
        q.addBindValue(static_cast<int>(expectedVersion));
    }
    if (!q.exec()) {
        qWarning() << "LiveSync sync jsonb failed:" << q.lastError().text();
        return false;
    }
    const bool ok = q.next() && q.value(0).toBool();
    if (!ok) {
        emit commitConflict(table, rowId,
                            QStringLiteral("json_path:") + jsonPath,
                            value, expectedVersion);
    }
    return ok;
}

// v3 Phase 3d: sync-fallback twin of the worker's commitMeasurement (used
// when no worker is wired - tests without a DbConfig). Same FALSE contract:
// unknown key is permanent, logged, never re-queued.
bool LiveSync::runMeasurementUpdateSync(qint64 sampleId, const QString& key,
                                        int sortOrder, const QVariant& value)
{
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("SELECT dve_commit_measurement(?, ?, ?, ?, ?)"));
    q.addBindValue(sampleId);
    q.addBindValue(key);
    q.addBindValue(sortOrder);
    q.addBindValue(value.toString());
    q.addBindValue(m_identity
        ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString());
    if (!q.exec()) {
        qWarning() << "LiveSync sync measurement failed:" << q.lastError().text();
        return false;
    }
    const bool ok = q.next() && q.value(0).toBool();
    if (!ok) {
        qWarning() << "LiveSync::runMeasurementUpdateSync: metric_defs has no"
                      " kind='metric' row for" << key << "-- edit not committed"
                      " per-cell (whole-file save still carries it)";
    }
    return ok;
}

bool LiveSync::focusCell(const QString& table, qint64 rowId, const QString& column)
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    const QString uuid = m_identity->uuid().toString(QUuid::WithoutBraces);
    if (m_worker) {
        QMetaObject::invokeMethod(m_worker, "focusCell", Qt::QueuedConnection,
            Q_ARG(QString, uuid),  Q_ARG(QString, table),
            Q_ARG(qint64,  rowId), Q_ARG(QString, column),
            Q_ARG(QString, m_identity->displayName()),
            Q_ARG(QString, m_identity->color()));
        return true;
    }
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)"));
    q.addBindValue(uuid);
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(m_identity->displayName());
    q.addBindValue(m_identity->color());
    return q.exec();
}

bool LiveSync::blurCell()
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    const QString uuid = m_identity->uuid().toString(QUuid::WithoutBraces);
    if (m_worker) {
        QMetaObject::invokeMethod(m_worker, "blurCell", Qt::QueuedConnection,
            Q_ARG(QString, uuid));
        return true;
    }
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("DELETE FROM cell_focus WHERE user_uuid = ?::uuid"));
    q.addBindValue(uuid);
    return q.exec();
}

void LiveSync::onRowChanged(const RowChange& c)
{
    if (c.column.isEmpty()) return;
    emit cellChanged(c.table, c.id, c.column, c.newValue);
}

void LiveSync::onCellFocusChanged(const CellFocusChange& f)
{
    if (f.op == QLatin1String("DELETE"))
        emit cellBlurred(f.tableName, f.rowId, f.columnName);
    else
        emit cellFocused(f.tableName, f.rowId, f.columnName,
                         f.userName, f.userColor);
}

} // namespace DVE
