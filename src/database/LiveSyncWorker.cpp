#include "LiveSyncWorker.h"

#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QUuid>
#include <QRegularExpression>
#include <QStringList>
#include <QThread>

namespace DVE {

LiveSyncWorker::LiveSyncWorker(const DbConfig& cfg, QObject* parent)
    : QObject(parent), m_cfg(cfg)
{
    // Unique connection name per worker instance so multiple LiveSyncs (in
    // tests, or app + tests) don't collide on QSqlDatabase's global registry.
    m_connName = QStringLiteral("dve_livesync_worker_%1")
                     .arg(QUuid::createUuid().toString(QUuid::WithoutBraces));
}

LiveSyncWorker::~LiveSyncWorker()
{
    if (m_open) stop();
}

void LiveSyncWorker::start()
{
    if (m_open) return;
    m_open = openConnection();
    if (!m_open) {
        qWarning() << "LiveSyncWorker::start failed:" << m_db.lastError().text();
    }
}

void LiveSyncWorker::stop()
{
    if (m_db.isOpen()) m_db.close();
    if (QSqlDatabase::contains(m_connName)) {
        m_db = QSqlDatabase();
        QSqlDatabase::removeDatabase(m_connName);
    }
    m_open = false;
}

bool LiveSyncWorker::openConnection()
{
    if (QSqlDatabase::contains(m_connName)) {
        QSqlDatabase::removeDatabase(m_connName);
    }
    m_db = QSqlDatabase::addDatabase(QStringLiteral("QPSQL"), m_connName);
    m_db.setHostName(m_cfg.host);
    m_db.setPort(m_cfg.port);
    m_db.setDatabaseName(m_cfg.database);
    m_db.setUserName(m_cfg.user);
    m_db.setPassword(m_cfg.password);
    // Short connect timeout so a slow NAS doesn't hang worker startup; shared
    // options add application_name + keepalives + tcp_user_timeout (v2.4.2).
    m_db.setConnectOptions(QStringLiteral("connect_timeout=5;") + pgSharedConnectOptions());
    if (!m_db.open()) return false;
    applyPgSessionSettings(m_db);     // statement_timeout (best-effort)
    return true;
}

bool LiveSyncWorker::isConnectionError(const QSqlQuery& q, bool dbOpen)
{
    // Convenience overload: pull the query's error and delegate to the pure
    // classifier so both the live call site and the unit test share one body.
    return isConnectionError(q.lastError(), dbOpen);
}

bool LiveSyncWorker::isConnectionError(const QSqlError& err, bool dbOpen)
{
    // The handle reporting closed is the clearest signal the connection died.
    if (!dbOpen) return true;

    const QString code = err.nativeErrorCode();

    // Postgres SQLSTATEs: 26000 = invalid_sql_statement_name (the "unnamed
    // prepared statement does not exist" we see when libpq silently lost and
    // re-made the socket); 25P02 = in_failed_sql_transaction (a connection
    // wedged in an aborted transaction, e.g. after a statement_timeout cancel
    // mid-transaction); class 08 = connection_exception (08000/08003/08006/
    // 08001/08004/08007). All mean "re-establish the session".
    if (code == QLatin1String("26000")) return true;
    // 25P02 = in_failed_sql_transaction: a connection wedged in an aborted
    // transaction (e.g. after a statement_timeout cancel mid-transaction).
    // stop()/start() recreates the backend and clears the aborted state.
    if (code == QLatin1String("25P02")) return true;
    if (code.startsWith(QLatin1String("08"))) return true;

    // Fallback to libpq's human text for cases the driver doesn't surface a
    // SQLSTATE for (e.g. a hard socket close before any reply).
    const QString text = err.text().toLower();
    static const char* kNeedles[] = {
        "server closed", "unable to send query", "connection",
        "terminat", "broken pipe", "no connection", "could not send"
    };
    for (const char* needle : kNeedles) {
        if (text.contains(QLatin1String(needle))) return true;
    }
    return false;
}

bool LiveSyncWorker::execWithReconnect(QSqlQuery& q,
                                       const std::function<void(QSqlQuery&)>& build)
{
    // First attempt on the current connection.
    q = QSqlQuery(m_db);
    build(q);
    if (q.exec()) return true;

    // Only reconnect when the error looks like a dead connection, not a
    // statement-level error (syntax, constraint, OCC). A bad statement would
    // just fail again on a fresh connection and waste a teardown.
    if (!isConnectionError(q, m_db.isOpen())) return false;

    qWarning() << "LiveSyncWorker: connection-shaped error, reconnecting --"
               << q.lastError().nativeErrorCode() << q.lastError().text();

    // v2.5.0 Task 4 review: the replay below is at-least-once, not exactly-once.
    // If the FIRST exec actually reached the server and committed but the
    // connection died before libpq read the reply, the replay re-runs the same
    // statement on the fresh connection -- a second dve_commit_cell* call. For
    // our writes that is benign-but-not-free: each commit bumps the row version,
    // so a double-apply burns one extra version (and, against a now-newer row,
    // the replay can return FALSE -> a SPURIOUS commitConflict for an edit that
    // already landed). It is NEVER destructive -- the value written is identical
    // both times, sibling JSON paths are untouched (jsonb_set), and the merge
    // layer (Task 3 dirty cells) keeps the local edit authoritative regardless.
    // We accept the rare spurious conflict / extra version bump as the cost of
    // not silently dropping the edit, which was the v2.4.0 regression.
    stop();
    start();
    if (!m_open || !m_db.isOpen()) {
        qWarning() << "LiveSyncWorker: reconnect failed to re-open connection:"
                   << m_db.lastError().text();
        return false;
    }

    // Replay the SAME statement on the fresh connection, once.
    q = QSqlQuery(m_db);
    build(q);
    const bool ok = q.exec();
    if (ok) {
        qWarning() << "LiveSyncWorker: reconnect succeeded, statement replayed";
    } else {
        qWarning() << "LiveSyncWorker: statement still failing after reconnect:"
                   << q.lastError().text();
    }
    return ok;
}

QString LiveSyncWorker::parsePathToPgArray(const QString& jsonPath) const
{
    // "samples[2].scores.smoothness" -> "{samples,2,scores,smoothness}"
    QStringList parts;
    static const QRegularExpression re(QStringLiteral(R"(([^.\[\]]+)|\[(\d+)\])"));
    auto it = re.globalMatch(jsonPath);
    while (it.hasNext()) {
        const auto m = it.next();
        parts << (m.captured(1).isEmpty() ? m.captured(2) : m.captured(1));
    }
    return QStringLiteral("{%1}").arg(parts.join(QLatin1Char(',')));
}

void LiveSyncWorker::commitScalar(QString table, qint64 rowId,
                                  QString column, QVariant value, QString uuid,
                                  qint64 expectedVersion)
{
    if (!m_open || !m_db.isOpen()) {
        // Connection never opened (or torn down) -- try a reconnect before
        // giving up, so a transient startup race doesn't drop the edit.
        stop();
        start();
        if (!m_open || !m_db.isOpen()) {
            emit commitFailed(table, rowId, column, value);
            return;
        }
    }
    // SP4.5 Stage 2a: a row whose id isn't assigned yet (-1 during the
    // background persist-on-load window) can never match UPDATE ... WHERE id=-1.
    // Route the edit to the offline pending_edits queue via commitFailed so it
    // replays once the real id is stamped (and the session dirty-cell merge
    // re-persists it on the next whole-file save), instead of firing a no-op
    // UPDATE that the commitConflict path would discard as a spurious conflict.
    if (rowId <= 0) {
        qWarning() << "LiveSyncWorker::commitScalar: rowId<=0 (" << rowId
                   << ") for" << table << column << "-- deferring to offline queue";
        emit commitFailed(table, rowId, column, value);
        return;
    }
    QSqlQuery q;
    const bool ok = execWithReconnect(q, [&](QSqlQuery& qq) {
        qq.prepare(QStringLiteral("SELECT dve_commit_cell(?, ?, ?, ?, ?, ?)"));
        qq.addBindValue(table);
        qq.addBindValue(rowId);
        qq.addBindValue(column);
        qq.addBindValue(value.toString());
        qq.addBindValue(uuid);
        // expectedVersion < 0 -> bind NULL so the stored function takes its
        // pre-v2.0.2 no-OCC path (DEFAULT NULL). Otherwise we send the version
        // for the server-side AND version = $6 check.
        if (expectedVersion < 0) {
            qq.addBindValue(QVariant(QMetaType(QMetaType::Int)));
        } else {
            qq.addBindValue(static_cast<int>(expectedVersion));
        }
    });
    if (!ok) {
        qWarning() << "LiveSyncWorker::commitScalar failed:" << q.lastError().text()
                   << "for" << table << column << rowId;
        emit commitFailed(table, rowId, column, value);
        return;
    }
    // The stored function returns BOOLEAN. FALSE means the WHERE matched
    // zero rows: either expectedVersion didn't match OR the row was deleted.
    // The local edit did NOT land -- bubble it up as a conflict so LiveSync
    // does NOT replay-enqueue a stale value.
    if (!q.next() || !q.value(0).toBool()) {
        emit commitConflict(table, rowId, column, value, expectedVersion);
    }
}

void LiveSyncWorker::commitJson(QString table, qint64 rowId,
                                QString jsonPath, QVariant value, QString uuid,
                                qint64 expectedVersion)
{
    if (!m_open || !m_db.isOpen()) {
        stop();
        start();
        if (!m_open || !m_db.isOpen()) {
            emit commitFailed(table, rowId,
                              QStringLiteral("json_path:") + jsonPath, value);
            return;
        }
    }
    // SP4.5 Stage 2a: see commitScalar -- defer a pre-persist (id<=0) edit to the
    // offline queue rather than firing UPDATE ... WHERE id=-1.
    if (rowId <= 0) {
        qWarning() << "LiveSyncWorker::commitJson: rowId<=0 (" << rowId
                   << ") for" << table << jsonPath << "-- deferring to offline queue";
        emit commitFailed(table, rowId,
                          QStringLiteral("json_path:") + jsonPath, value);
        return;
    }
    const QString pgArr = parsePathToPgArray(jsonPath);
    QSqlQuery q;
    const bool ok = execWithReconnect(q, [&](QSqlQuery& qq) {
        qq.prepare(QStringLiteral(
            "SELECT dve_commit_cell_json(?, ?, ?, ?::text[], ?, ?, ?)"));
        qq.addBindValue(table);
        qq.addBindValue(rowId);
        qq.addBindValue(jsonPath);
        qq.addBindValue(pgArr);
        qq.addBindValue(value.toString());
        qq.addBindValue(uuid);
        if (expectedVersion < 0) {
            qq.addBindValue(QVariant(QMetaType(QMetaType::Int)));
        } else {
            qq.addBindValue(static_cast<int>(expectedVersion));
        }
    });
    if (!ok) {
        qWarning() << "LiveSyncWorker::commitJson failed:" << q.lastError().text()
                   << "for" << table << jsonPath << rowId;
        emit commitFailed(table, rowId,
                          QStringLiteral("json_path:") + jsonPath, value);
        return;
    }
    if (!q.next() || !q.value(0).toBool()) {
        emit commitConflict(table, rowId,
                            QStringLiteral("json_path:") + jsonPath,
                            value, expectedVersion);
    }
}

void LiveSyncWorker::focusCell(QString uuid, QString table, qint64 rowId,
                                QString column, QString userName, QString userColor)
{
    if (!m_open || !m_db.isOpen()) return;
    QSqlQuery q;
    const bool ok = execWithReconnect(q, [&](QSqlQuery& qq) {
        qq.prepare(QStringLiteral("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)"));
        qq.addBindValue(uuid);
        qq.addBindValue(table);
        qq.addBindValue(rowId);
        qq.addBindValue(column);
        qq.addBindValue(userName);
        qq.addBindValue(userColor);
    });
    if (!ok) {
        qWarning() << "LiveSyncWorker::focusCell failed:" << q.lastError().text();
    }
}

void LiveSyncWorker::blurCell(QString uuid)
{
    if (!m_open || !m_db.isOpen()) return;
    QSqlQuery q;
    const bool ok = execWithReconnect(q, [&](QSqlQuery& qq) {
        qq.prepare(QStringLiteral("DELETE FROM cell_focus WHERE user_uuid = ?::uuid"));
        qq.addBindValue(uuid);
    });
    if (!ok) {
        qWarning() << "LiveSyncWorker::blurCell failed:" << q.lastError().text();
    }
}

void LiveSyncWorker::sync()
{
    // DATAVIEWER-4 drain barrier. By the time this slot runs, the worker
    // thread has already processed every commit slot queued before it (Qt
    // event queues are FIFO), so all preceding writes have hit Postgres.
    // No DB work needed -- just signal completion back to the UI thread.
    emit synced();
}

} // namespace DVE
