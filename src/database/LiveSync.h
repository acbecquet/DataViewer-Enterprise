#pragma once

#include <QObject>
#include <QPointer>
#include <QString>
#include <QVariant>
#include <QHash>

#include "ConfigLoader.h"

class QThread;
class QTimer;

namespace DVE {

class PostgresConnection;
class IdentityManager;
class OfflineSnapshot;
class LiveSyncWorker;
struct RowChange;
struct CellFocusChange;

// LiveSync is the single chokepoint for per-cell writes and remote
// applies in v2.0.1. Every editable widget commits through
// commitCell(); every NOTIFY-driven cell change emits cellChanged().
//
// JSONB columns are addressed with column names prefixed "json_path:"
// e.g. column = "json_path:samples[2].scores.smoothness". LiveSync
// translates these into jsonb_set() UPDATEs at the database layer so
// concurrent edits to different paths in the same sample don't
// clobber each other.
//
// v2.0.1 polish-2: writes go to a background LiveSyncWorker on its
// own QThread so the UI thread never blocks on Postgres round trips.
// Rapid same-cell edits coalesce in a 200 ms throttle window — only
// the latest value per (table, rowId, column) is sent to the worker.
class LiveSync : public QObject {
    Q_OBJECT
public:
    LiveSync(PostgresConnection* conn, IdentityManager* identity,
             QObject* parent = nullptr);
    ~LiveSync() override;

    // Wire up the dedicated worker connection. Without this call,
    // commitCell falls back to the legacy synchronous path on the UI
    // thread (used by tests that don't carry a DbConfig). Call once
    // after construction; safe to call multiple times — only the
    // first wins.
    void setWorkerConfig(const DbConfig& cfg);

    // Scalar-column or JSONB UPDATE. Returns true if the edit was
    // accepted (dispatched to the worker or queued offline), false if
    // the table/column failed the allowlist or the snapshot was
    // missing during offline. This call never blocks on Postgres.
    //
    // allowQueue=false: caller is replaying from the offline queue;
    // skip re-enqueue so a mid-drain disconnect can't duplicate.
    bool commitCell(const QString& table, qint64 rowId,
                    const QString& column, const QVariant& value,
                    bool allowQueue = true);

    // Async cell-focus broadcast. UI thread debounces this upstream
    // (see MainWindow m_focusCommitTimer); the worker just runs the
    // DELETE+INSERT in a single stored-function call.
    bool focusCell(const QString& table, qint64 rowId, const QString& column);
    bool blurCell();

    // Offline replay hooks (unchanged contract).
    void setOfflineSnapshot(OfflineSnapshot* snap);
    int  flushPending();

signals:
    void cellChanged(const QString& table, qint64 rowId,
                     const QString& column, const QVariant& newValue);

    void cellFocused(const QString& table, qint64 rowId,
                     const QString& column,
                     const QString& userName, const QString& userColor);
    void cellBlurred(const QString& table, qint64 rowId,
                     const QString& column);

public slots:
    void onRowChanged(const RowChange& change);
    void onCellFocusChanged(const CellFocusChange& change);

private slots:
    // Throttle tick: drain m_pendingCommits, dispatch each to worker.
    void onThrottleTick();
    // Worker reported a commit failure — enqueue to snapshot for replay.
    void onWorkerCommitFailed(QString table, qint64 rowId,
                              QString column, QVariant value);

private:
    struct PendingKey {
        QString table;
        qint64  rowId;
        QString column;
        bool operator==(const PendingKey& o) const {
            return table == o.table && rowId == o.rowId && column == o.column;
        }
    };
    friend size_t qHash(const PendingKey& k, size_t seed) noexcept;

    // Allowlists (unchanged from v2.0.1; defined in .cpp).
    static bool isLiveSyncTable (const QString& table);
    static bool isLiveSyncColumn(const QString& table, const QString& column);

    // Legacy synchronous fallback (only used when no worker is set —
    // e.g. tests that don't pass a DbConfig). Kept so the test suite
    // doesn't need a worker thread spun up.
    bool runScalarUpdateSync(const QString& table, qint64 rowId,
                             const QString& column, const QVariant& value);
    bool runJsonPathUpdateSync(const QString& table, qint64 rowId,
                               const QString& jsonPath, const QVariant& value);

    // Dispatch helper — picks scalar vs json path and routes either to
    // the worker (async) or the sync fallback.
    void dispatchCommit(const QString& table, qint64 rowId,
                        const QString& column, const QVariant& value);

    QPointer<PostgresConnection> m_conn;
    QPointer<IdentityManager>    m_identity;
    QPointer<OfflineSnapshot>    m_snapshot;

    // Worker thread + worker object. Lazily constructed in setWorkerConfig.
    QThread*        m_workerThread = nullptr;
    LiveSyncWorker* m_worker       = nullptr;

    // 200 ms coalescing window for rapid same-cell edits. Single shared
    // tick timer so we don't allocate one timer per dirty cell.
    QTimer*                       m_throttleTimer = nullptr;
    QHash<PendingKey, QVariant>   m_pendingCommits;
};

inline size_t qHash(const LiveSync::PendingKey& k, size_t seed) noexcept
{
    // Qualify with ::qHash so the qulonglong/QString overloads from Qt's
    // global namespace are found instead of recursing into this overload.
    return ::qHash(k.table, seed)
         ^ ::qHash(static_cast<qulonglong>(k.rowId), seed)
         ^ ::qHash(k.column, seed);
}

} // namespace DVE
