#pragma once

#include <QObject>
#include <QPointer>
#include <QString>
#include <QVariant>
#include <QHash>

#include <functional>

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

    // C4: number of cells still queued for commit — sum of in-process throttle
    // queue and the persistent per-cell offline snapshot. Callers gate
    // file-level saveFile on this returning 0 so the saveFile path doesn't
    // race the per-cell drain.
    //
    // R5: this counts ONLY actionable/drainable work. A present-but-undecodable
    // (MIP-encrypted/corrupt) offline queue does NOT contribute here — it can
    // never drain, so counting it would keep pendingCount()>=1 forever and trip
    // MainWindow::flushPendingEdits' Q_ASSERT(pendingCount()==0) on every flush.
    // The unreadable-queue state is carried SOLELY by offlineQueueDegraded()
    // below, so pendingCount()==0 reliably means "no drainable work left".
    int  pendingCount() const;

    // R5: true iff the offline pending-edits queue exists but could not be
    // opened/decoded on the last access (so its row count can't be trusted).
    // This rides ALONGSIDE pendingCount() (which excludes the degraded queue),
    // not inside it: the sync indicator reads this flag to surface a distinct
    // non-blocking "queue unreadable" state. The stranded edits are not lost —
    // they remain in the open session's dirty set and re-persist on the next
    // online whole-session save (Ctrl+U).
    bool offlineQueueDegraded() const;

    // v2.5.0 Task 4 (RC3): count of per-cell edits the worker reported as
    // failed (driver-level) or conflicted (OCC miss / row gone) since the last
    // successful drain. Surfaced in the UI so silent per-cell loss can't hide:
    // production logs showed blurCell/commit failing 126× with NO visible cue.
    //
    // IMPORTANT: a reset to 0 (on the next successful flushNowAndWait drain)
    // does NOT mean the failed edits were recovered. It means "no NEW failures
    // are outstanding". The actual safety net is the panel dirty-cell merge
    // (Task 3), which keeps the local value authoritative on the next whole-
    // session save. This counter is a visibility signal, not a recovery state.
    int  unsyncedEditCount() const { return m_unsyncedEdits; }

    // DATAVIEWER-4: drain the throttle queue NOW and block (bounded by
    // timeoutMs) until the worker has written this client's pending per-cell
    // edits. Call at DELIBERATE persist points (Ctrl+U / Close / export /
    // program-close) before a whole-session read-merge-write so the merge
    // can't revert a just-typed score. No-op when nothing is pending or no
    // worker is wired (the sync fallback already wrote inside onThrottleTick).
    // Re-entrant calls are no-ops. The 5 s auto-save deliberately does NOT
    // call this (it stays async so typing never blocks on the NAS).
    //
    // v2.5.0 Task 3 (RC2): returns true when the worker confirmed the drain
    // (its synced() reply arrived within timeoutMs), false on timeout or when a
    // re-entrant call short-circuits. With the dirty-aware merge a timeout is no
    // longer a data-loss risk (the panel's dirtyCells keep the local edits
    // authoritative), so callers log a qWarning and proceed. A no-worker / sync-
    // fallback flush returns true (the edits are already durable on this thread).
    bool flushNowAndWait(int timeoutMs = 4000);

    // Register a callback that returns the caller's last-known row
    // version for (table, rowId), used by v2.0.2 optimistic-concurrency
    // checks in the stored functions. Return -1 to opt out of OCC for a
    // particular row (e.g. when the row isn't in the caller's cache
    // yet). If no callback is registered, all commits run with OCC
    // disabled — matching v2.0.1 behavior for tests/legacy callers.
    using VersionLookup = std::function<qint64(const QString& table, qint64 rowId)>;
    void setVersionLookup(VersionLookup lookup);

signals:
    void cellChanged(const QString& table, qint64 rowId,
                     const QString& column, const QVariant& newValue);

    void cellFocused(const QString& table, qint64 rowId,
                     const QString& column,
                     const QString& userName, const QString& userColor);
    void cellBlurred(const QString& table, qint64 rowId,
                     const QString& column);

    // Forwarded from the worker on an OCC miss. MainWindow (Phase 3
    // consumer) decides whether to show a stale-data dialog, refetch the
    // row, or yellow-flash the cell. v2.0.2 does NOT enqueue conflicted
    // writes to the offline snapshot.
    void commitConflict(const QString& table, qint64 rowId,
                        const QString& column, const QVariant& attemptedValue,
                        qint64 expectedVersion);

    // v2.5.0 Task 4 (RC3): emitted whenever unsyncedEditCount() changes — the
    // worker reported a new commitFailed/commitConflict (count goes up), or a
    // successful drain reset it (count goes to 0). MainWindow's DB sync
    // indicator listens and shows a warning state when non-zero.
    void unsyncedEditsChanged(int count);

    // v2.4.4 R5: emitted ONCE PER new offline-enqueue failure — an offline
    // per-cell edit that could not be persisted to the local pending_edits
    // queue (e.g. the queue file is MIP-encrypted and undecodable). Before this,
    // enqueueCellEdit's bool was discarded at every call site, so such an edit
    // vanished with no cue.
    //
    // R5 IMPORTANT 1: this is a one-shot NOTIFICATION, not a count. The unsynced
    // TALLY is owned solely by m_unsyncedEdits (bumped via bumpUnsynced() and
    // broadcast by unsyncedEditsChanged). handleEnqueueResult does NOT touch the
    // tally — it only warns and emits THIS signal; each CALLER bumps the tally
    // exactly once (the two pure-offline sites bump on a failed enqueue;
    // onWorkerCommitFailed bumps for its own commit failure regardless of the
    // enqueue outcome). MainWindow listens to THIS signal only to raise its
    // one-time "edits not saved locally" warning; it must NOT add to its
    // unsynced counter here (that would double-count and fight the
    // unsyncedEditsChanged handler — the old quadratic bug).
    //
    // The edit is NOT lost: the same keystroke already recorded the cell in the
    // open session's dirty set (SensoryPanel/DetailedSensoryPanel dirtyCells /
    // m_modifiedFilePaths), so the next online whole-session save re-persists it
    // via the dirty-aware merge. See handleEnqueueResult().
    void offlineEnqueueFailed();

public slots:
    void onRowChanged(const RowChange& change);
    void onCellFocusChanged(const CellFocusChange& change);

private slots:
    // Throttle tick: drain m_pendingCommits, dispatch each to worker.
    void onThrottleTick();
    // Worker reported a driver-level commit failure — enqueue to snapshot.
    void onWorkerCommitFailed(QString table, qint64 rowId,
                              QString column, QVariant value);
    // Worker reported an OCC miss — forward upstream without enqueuing.
    void onWorkerCommitConflict(QString table, qint64 rowId,
                                QString column, QVariant value,
                                qint64 expectedVersion);

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

    // DATAVIEWER-4: re-entrancy guard for flushNowAndWait. The nested
    // QEventLoop it spins could otherwise re-deliver a deferred persist
    // call (e.g. a queued close + export) and recurse into a second wait.
    bool                          m_flushing = false;

    // v2.5.0 Task 4 (RC3): running count of worker-reported per-cell failures
    // (commitFailed + commitConflict) since the last successful drain. See
    // unsyncedEditCount(). bumpUnsynced()/resetUnsynced() emit
    // unsyncedEditsChanged when the value actually moves.
    int                           m_unsyncedEdits = 0;
    void bumpUnsynced();

    // v2.4.4 R5 IMPORTANT 2: the single chokepoint the three offline-enqueue
    // sites route their enqueueCellEdit() result through. On failure it ONLY
    // warns and emits the one-shot offlineEnqueueFailed() notification — it does
    // NOT bump the unsynced tally. The tally has a single owner (m_unsyncedEdits
    // via bumpUnsynced), and each CALLER bumps it once: the two pure-offline
    // sites (commitCell, dispatchCommit) bump on a failed enqueue;
    // onWorkerCommitFailed bumps for its commit failure regardless of the
    // enqueue outcome. Bumping inside this helper would double-count on the
    // worker path (the old quadratic bug).
    //
    // There is deliberately NO in-memory fallback buffer here. An investigation
    // (R5) confirmed the edit is NOT lost when the enqueue fails: every offline
    // per-cell edit site records the touched cell in the open session's dirty
    // set BEFORE calling commitCell (SensoryPanel/DetailedSensoryPanel
    // dirtyCells; TPM m_modifiedFilePaths). The next online whole-session save
    // (Ctrl+U / close prompt / reconnect) re-persists it via the dirty-aware
    // merge. A write-only fallback vector that nothing drained was the dead code
    // this fix removed; the dialog wording now describes the real mechanism.
    void handleEnqueueResult(bool ok, const QString& table, qint64 rowId,
                             const QString& column, const QVariant& value);

    // v2.0.2: version resolver for optimistic-concurrency checks. -1
    // sentinel when unset → no OCC (matches v2.0.1 behavior).
    VersionLookup                 m_versionLookup;
    qint64 currentVersionFor(const QString& table, qint64 rowId) const;
};

inline size_t qHash(const LiveSync::PendingKey& k, size_t seed) noexcept
{
    // M7: chain seeds across the three fields so collisions in any one
    // component don't collapse the hash. The previous XOR-with-shared-seed
    // form caused O(n) lookup chains at high edit rates with narrow value
    // ranges.
    //
    // All inner calls qualify ::qHash to force the global-namespace
    // overloads (qulonglong/QString) instead of recursing into ourselves.
    size_t h = ::qHash(k.table, seed);
    h = ::qHash(static_cast<qulonglong>(k.rowId), h);
    h = ::qHash(k.column, h);
    return h;
}

} // namespace DVE
