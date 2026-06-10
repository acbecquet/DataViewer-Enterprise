#pragma once

#include <QObject>
#include <QSqlDatabase>
#include <QString>
#include <QVariant>

#include <functional>

#include "ConfigLoader.h"

class QSqlQuery;

namespace DVE {

// LiveSyncWorker runs on its own QThread. The UI thread fires fire-and-forget
// queued-connection signals; the worker owns its own QSqlDatabase connection
// so writes never block the UI on Postgres round trips.
//
// Lifecycle:
//   ctor builds the QObject (still on the caller's thread).
//   start() must be called AFTER moveToThread() — it runs on the worker
//   thread and opens the dedicated connection.
//   stop()  closes the connection on the worker thread.
//
// Signals back to the UI (commitFailed, commitSucceeded) are queued so the
// caller sees them on the UI thread and can update offline-snapshot state.
//
// v2.5.0 Task 4 (RC3): production logs showed the worker connection entering a
// permanently-broken state ("unnamed prepared statement does not exist
// (26000)" / "Unable to send query") that never recovered until app restart —
// every per-cell commit on that connection then died silently. The worker now
// detects connection-shaped errors, tears down and re-opens its connection,
// and replays the failed statement once (execWithReconnect). Only a failure
// AFTER the reconnect retry reaches the commitFailed/commitConflict paths.
class LiveSyncWorker : public QObject {
    Q_OBJECT
public:
    explicit LiveSyncWorker(const DbConfig& cfg, QObject* parent = nullptr);
    ~LiveSyncWorker() override;

public slots:
    // Open the dedicated connection on the worker thread.
    void start();
    // Close it on the worker thread.
    void stop();

    // Scalar UPDATE via dve_commit_cell stored function.
    // expectedVersion < 0 disables OCC (legacy callers/tests). When >= 0,
    // the stored function's WHERE clause matches AND version = expected;
    // a row with a different version triggers commitConflict, not commitFailed.
    void commitScalar(QString table, qint64 rowId,
                      QString column, QVariant value, QString uuid,
                      qint64 expectedVersion);

    // JSONB UPDATE via dve_commit_cell_json stored function. expectedVersion
    // semantics match commitScalar.
    // jsonPath is the unparsed path (e.g. "samples[2].name"); the worker
    // parses it into a TEXT[] array client-side and sends both forms.
    void commitJson(QString table, qint64 rowId,
                    QString jsonPath, QVariant value, QString uuid,
                    qint64 expectedVersion);

    void focusCell(QString uuid, QString table, qint64 rowId,
                   QString column, QString userName, QString userColor);
    void blurCell(QString uuid);

    // DATAVIEWER-4: drain barrier. Runs only after every commitJson()/
    // commitScalar() queued before it (the worker thread processes its
    // event queue in order), so emitting synced() signals to the UI
    // thread "everything queued so far has been written". LiveSync::
    // flushNowAndWait() posts this and blocks on synced() at deliberate
    // persist points so a whole-session write can't revert a just-typed cell.
    void sync();

signals:
    // Emitted on driver-level commit failure (network drop, syntax error,
    // permission denial). LiveSync (UI thread) enqueues the lost edit to
    // the offline snapshot for replay.
    void commitFailed(QString table, qint64 rowId,
                      QString column, QVariant value);

    // Emitted when the stored function returns FALSE — the caller passed
    // an expectedVersion that no longer matches the row's current version
    // (optimistic-concurrency miss) OR the row was deleted. Distinct from
    // commitFailed so LiveSync DOES NOT enqueue to the offline snapshot
    // (that would re-apply a stale write).
    void commitConflict(QString table, qint64 rowId,
                        QString column, QVariant attemptedValue,
                        qint64 expectedVersion);

    // DATAVIEWER-4: emitted at the end of sync() — see the sync() slot above.
    void synced();

private:
    bool        openConnection();
    QString     parsePathToPgArray(const QString& jsonPath) const;

    // v2.5.0 Task 4 (RC3): true when a failed q.exec() looks like the
    // connection (not the statement) is broken — Postgres 26000
    // "prepared statement does not exist", SQLSTATE class 08 (connection
    // exception), the db handle reporting closed, or libpq's "server
    // closed"/"Unable to send query"/"terminat"/"connection" text. Static so
    // it is unit-testable without a live connection.
    static bool isConnectionError(const QSqlQuery& q, bool dbOpen);

    // v2.5.0 Task 4 (RC3): prepare+bind via `build`, exec, and on a
    // connection-shaped failure stop()/start() the connection and replay
    // `build` once on the fresh handle. `q` is (re)assigned a QSqlQuery bound
    // to the live m_db. Returns true if exec ultimately succeeded; on a
    // false return q.lastError() reflects the final attempt. A non-connection
    // error (syntax, constraint) fails immediately with no reconnect.
    bool execWithReconnect(QSqlQuery& q,
                           const std::function<void(QSqlQuery&)>& build);

    DbConfig     m_cfg;
    QSqlDatabase m_db;
    QString      m_connName;
    bool         m_open = false;
};

} // namespace DVE
