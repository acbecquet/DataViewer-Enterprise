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
    return t == QLatin1String("data_rows")
        || t == QLatin1String("samples")
        || t == QLatin1String("tests")
        || t == QLatin1String("files")
        || t == QLatin1String("sensory_sessions")
        || t == QLatin1String("detailed_sensory_sessions");
}

bool LiveSync::isLiveSyncColumn(const QString& table, const QString& column)
{
    static const QHash<QString, QSet<QString>> kAllowed = {
        { QStringLiteral("data_rows"), {
            QStringLiteral("puffs"), QStringLiteral("before_weight"),
            QStringLiteral("after_weight"), QStringLiteral("draw_pressure"),
            QStringLiteral("resistance"), QStringLiteral("smell"),
            QStringLiteral("clog"), QStringLiteral("notes"),
            QStringLiteral("tpm"), QStringLiteral("tpm_power_density"),
            QStringLiteral("variation_tpm"), QStringLiteral("oil_consumed"),
            QStringLiteral("puffing_regime")
        }},
        { QStringLiteral("samples"), {
            QStringLiteral("sample_name"), QStringLiteral("sample_id"),
            QStringLiteral("date"), QStringLiteral("tester"),
            QStringLiteral("media"), QStringLiteral("viscosity"),
            QStringLiteral("resistance"), QStringLiteral("voltage"),
            QStringLiteral("power"), QStringLiteral("heating_technology"),
            QStringLiteral("puffing_regime"), QStringLiteral("initial_oil_mass"),
            QStringLiteral("burn_status"), QStringLiteral("clog_status"),
            QStringLiteral("leak_status")
        }},
        { QStringLiteral("tests"), {
            QStringLiteral("sheet_name"), QStringLiteral("template_version")
        }},
        { QStringLiteral("files"), {
            QStringLiteral("file_path"), QStringLiteral("file_name"),
            QStringLiteral("template_version")
        }},
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
            return m_snapshot->enqueueCellEdit(table, rowId, column, value);
        }
        return false;
    }

    m_pendingCommits.insert({table, rowId, column}, value);
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
                       it.key().column, it.value());
    }
}

void LiveSync::dispatchCommit(const QString& table, qint64 rowId,
                              const QString& column, const QVariant& value)
{
    const bool isJsonPath = column.startsWith(QLatin1String("json_path:"));
    if (!m_conn || !m_conn->isOpen()) {
        if (m_snapshot) m_snapshot->enqueueCellEdit(table, rowId, column, value);
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
        if (isJsonPath) {
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

    if (isJsonPath) {
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
        m_snapshot->enqueueCellEdit(table, rowId, column, value);
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

void LiveSync::setOfflineSnapshot(OfflineSnapshot* snap)
{
    m_snapshot = snap;
}

int LiveSync::flushPending()
{
    if (!m_snapshot) return 0;
    if (!m_conn || !m_conn->isOpen()) return 0;
    return m_snapshot->drainPendingEdits(
        [this](const QString& t, qint64 r, const QString& c, const QVariant& v) {
            return this->commitCell(t, r, c, v, /*allowQueue=*/false);
        });
}

// C4: total queued work — throttle queue plus persistent snapshot queue.
// Callers (MainWindow::flushPendingEdits) gate file-level saveFile on this
// returning 0 so file-level UPDATEs don't race the per-cell drain. Returning
// the sum is intentional even though the throttle queue is in-process: a
// caller asserting "drain complete before save" needs both to be empty.
int LiveSync::pendingCount() const
{
    int n = m_pendingCommits.size();
    if (m_snapshot) n += m_snapshot->pendingEditCount();
    return n;
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
