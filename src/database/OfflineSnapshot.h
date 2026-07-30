#pragma once

#include <QObject>
#include <QString>
#include <QStringList>
#include <QSqlDatabase>
#include <QDateTime>
#include <atomic>
#include <QVariant>
#include <functional>
#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"
#include "DatabaseManager.h"  // for FileRecord / SensoryRecord / DetailedSensoryRecord

QT_FORWARD_DECLARE_CLASS(QSqlQuery)

namespace DVE {

class PostgresConnection;

// Local SQLite mirror of the Postgres database, used as the read-only data
// source when the NAS is unreachable. Path: %LOCALAPPDATA%/DataViewer/snapshot.sqlite.
//
// Lifecycle:
//   1. App boots, calls openReadOnly() -- fast.
//   2. On clean close (online), MainWindow calls regenerate(pgConn) which
//      dumps the live Postgres into a .tmp file and atomically replaces
//      the production snapshot.
//   3. Next time the app boots offline, openReadOnly() makes the snapshot
//      the data source. Saves are refused by DatabaseManager.
//
// Atomic write: regenerate() writes to snapshot.sqlite.tmp, then deletes
// snapshot.sqlite and renames .tmp over it. If anything fails mid-write,
// the previous snapshot.sqlite is intact.
class OfflineSnapshot : public QObject {
    Q_OBJECT
public:
    explicit OfflineSnapshot(QObject* parent = nullptr);
    ~OfflineSnapshot() override;

    // Resolves on first use. Creates parent directory if missing.
    QString path() const;

    // Regenerates the snapshot by reading every editable table from `live`
    // and writing the rows into a new SQLite file. Atomic: writes to .tmp,
    // then renames over the production path. Returns false on any error;
    // previous snapshot is preserved.
    //
    // Consistency contract: regenerate runs all Postgres reads inside a
    // REPEATABLE READ READ ONLY transaction so the snapshot is a consistent
    // point-in-time view of the source database (no torn snapshots if
    // another client commits between SELECTs).
    bool regenerate(PostgresConnection* live);

    // SP4.5 Stage 2b: regen progress callback (done, total, phase label). Null-safe.
    using RegenProgress = std::function<void(int done, int total, const QString& phase)>;

    // SP4.5 Stage 2b: per-run regen diagnostics (test/log visibility). wasIncremental
    // is true when the prior snapshot's blobs were reused; imageRowsPulledFromPg counts
    // blob rows actually fetched from PG (0 on a data-only change == the win).
    struct RegenStats {
        bool wasIncremental        = false;
        int  imageRowsPulledFromPg = 0;
    };

    // SP4.5 Stage 2b: same as regenerate() but drives a progress callback (used by
    // the close + manual-refresh progress dialog). The no-arg overload calls this
    // with an empty callback.
    bool regenerate(PostgresConnection* live, const RegenProgress& progress);

    // SP4.5 Stage 2a: thread-agnostic regen body. Runs on whatever thread the
    // caller's `live` connection belongs to (the background SnapshotRegenWorker
    // owns its own PostgresConnection). Writes destPath via the SAME atomic
    // .tmp -> prod promotion as regenerate(). `cancel` (nullable) is polled
    // between blob tables; when it flips true the function rolls back and
    // returns false. On success *outFingerprint / *outServerTimeUtc receive the
    // content fingerprint + PG-server stamp captured inside the txn.
    // INVARIANT: all PG reads run inside the one REPEATABLE READ READ ONLY txn.
    static bool regenToPath(PostgresConnection* live,
                            const QString&      destPath,
                            std::atomic<bool>*  cancel,
                            QString*            outFingerprint,
                            QDateTime*          outServerTimeUtc,
                            QString*            outError,
                            const RegenProgress& progress = {},
                            RegenStats*         outStats  = nullptr);

    // H14: runtime guard for the hand-maintained SELECT / INSERT / bind-loop
    // triples that regenToPath() copies each table with. The three counts must
    // agree or the snapshot silently mis-copies data. Returns false on a
    // mismatch after logging at CRITICAL severity (naming the table and all
    // three counts) and filling *outError; regenToPath then aborts, which
    // leaves the previous snapshot in place. ACTIVE IN RELEASE BUILDS -- it
    // replaced a Q_ASSERT_X that compiled out of the shipped binary. Public so
    // tst_offlinesnapshot can pin that contract directly.
    //
    // On `loopBound`, which today looks redundant and is not: every current
    // call site passes select.record().count(), the same value the guard reads
    // itself, so only the check with teeth (SELECT list vs the hand-maintained
    // INSERT literal) can currently fire. The parameter is kept because it
    // states the third leg of the contract explicitly at each site rather than
    // leaving it implied, and because a site that goes back to a literal bind
    // bound - the shape this guard exists to catch - re-arms it for free. The
    // unit test drives a genuine three-way mismatch directly, which is a case
    // production can no longer produce.
    static bool checkColumnArity(const QSqlQuery& select, int insertPlaceholders,
                                 int loopBound, const char* table,
                                 QString* outError = nullptr);

    // Opens the snapshot read-only (SQLite QSQLITE driver with
    // "QSQLITE_OPEN_READONLY" connect option). Subsequent listFiles/loadFile*/
    // listSensoryRecords/loadSensorySession/listDetailedSensoryRecords/
    // loadDetailedSensorySession go through this connection.
    bool openReadOnly();
    void close();
    bool isOpen() const { return m_open; }

    // Read-only accessors (mirror DatabaseManager's read methods).
    QVector<FileRecord>             listFiles() const;
    FileResult                      loadFileByPath(const QString& filePath) const;
    FileResult                      loadFile(int id) const;
    QVector<SensoryRecord>          listSensoryRecords() const;
    SensorySession                  loadSensorySession(int id) const;
    QVector<DetailedSensoryRecord>  listDetailedSensoryRecords() const;
    DetailedSensorySession          loadDetailedSensorySession(int id) const;

    // Settings table accessor — added in Plan C T3 so DatabaseManager can
    // route loadCumulativeLayout()/getSetting() to the snapshot when offline.
    // Returns defaultVal on missing key or any SQL error (with m_lastError set).
    QString getSetting(const QString& key, const QString& defaultVal = QString()) const;

    // When the snapshot was last regenerated (from _snapshot_meta table).
    // Returns invalid QDateTime if not yet regenerated.
    QDateTime snapshotTakenAt() const;

    // SP4.5 (perf): cheap change-detection so the close-time regenerate() can be
    // skipped when the live DB is unchanged since the last snapshot (the common
    // case — read-only sessions). The fingerprint is per-table COUNT + MAX
    // (updated_at); it is captured at regenerate() time into _snapshot_meta and
    // recomputed against the live DB on close.
    //   * liveContentFingerprint  — compute from the live PG ("" on error)
    //   * storedContentFingerprint — read what the open snapshot recorded ("" if absent)
    //   * isCurrentVsLive          — true iff both are present and equal (safe to skip
    //                                regen); ANY error/missing value returns false → regen.
    static QString liveContentFingerprint(PostgresConnection* live);
    QString        storedContentFingerprint() const;
    bool           isCurrentVsLive(PostgresConnection* live) const;

    // SP4.5 Stage 2b: the fingerprint table order (matches snapshotContentFingerprint).
    static const char* const kFingerprintTables[9];
    // Split a fingerprint into its `;`-separated segments. Empty list if the count
    // != 9 (malformed / schema drift).
    static QStringList fingerprintSegments(const QString& fp);
    // True if `table`'s segment differs between prior and live, OR either string is
    // unparseable, OR `table` is unknown (all safe defaults: force a refresh).
    static bool segmentChanged(const QString& priorFp, const QString& liveFp,
                               const char* table);

    // Test/diagnostic: the exact PG server timestamp (UTC) that the most
    // recent regenerate() captured and persisted as snapshot_taken_at. R7
    // sources the freshness stamp from the server clock rather than the
    // (possibly NTP-skewed) client clock; this getter lets a test prove the
    // persisted stamp equals what regenerate read from the server. Invalid
    // until a successful regenerate() runs.
    QDateTime lastRegenServerTimeUtc() const { return m_lastRegenServerTime; }

    // v2.0.1: persistent per-cell pending-edit queue. Lives in a separate
    // SQLite file (pending_edits.sqlite) alongside the read-only snapshot
    // so the snapshot itself stays QSQLITE_OPEN_READONLY. The queue table
    // survives across app restarts: edits captured while offline are
    // replayed by LiveSync::flushPending() once ConnectionMonitor reports
    // the connection came back online.
    //
    // column_name uses "json_path:..." for JSONB paths same as
    // LiveSync::commitCell, so the replay can dispatch by prefix.
    bool enqueueCellEdit(const QString& table, qint64 rowId,
                         const QString& column, const QVariant& value);

    // Replay queued cell edits via the supplied callback. Each entry for
    // which the callback returns true is DELETEd from the queue. Returns
    // the count of successfully-replayed entries. Failures stay queued
    // for the next flush.
    int drainPendingEdits(
        std::function<bool(const QString& table, qint64 rowId,
                            const QString& column,
                            const QVariant& value)> apply);

    // Test/diagnostic helper — returns the count of queued edits.
    //
    // R5 IMPORTANT 3: returns 0 when the queue file is present-but-undecodable
    // (MIP-encrypted/corrupt) just as it does for a genuinely-empty queue. A
    // caller MUST pair this with queueDegraded() before trusting "0 == drained":
    // a degraded queue reports 0 here while offline edits sit un-persisted on
    // disk. See queueDegraded().
    int pendingEditCount() const;

    // R5 IMPORTANT 3: true iff the LAST ensureQueueOpen() (driven by
    // enqueueCellEdit / drainPendingEdits / pendingEditCount) failed to open the
    // pending_edits queue because the file is present-but-unreadable — most
    // likely MIP-encrypted and undecodable by the python fallback. This lets the
    // save gate / sync indicator report "queue unreadable" rather than mistaking
    // a false pendingEditCount()==0 for "queue empty". Cleared on a successful
    // open. A genuinely-absent queue (first run, no offline edits yet) is NOT
    // degraded.
    bool queueDegraded() const { return m_queueDegraded; }

    QString lastError() const { return m_lastError; }

    // Test-only: lets tests override LOCALAPPDATA without relying on
    // QStandardPaths::setTestModeEnabled() (which gives a slightly different
    // path layout). Empty string clears the override and reverts to
    // QStandardPaths::AppLocalDataLocation. Has no production callers.
    void setOverrideDirForTesting(const QString& dir);

    // Filesystem path of the pending-edits queue SQLite file. Public so
    // tooling (the C5 drain-tests in tst_offlinesnapshot) and future
    // diagnostics can inspect the queue from outside.
    QString queuePath() const;

private:
    // Lazily opens (and creates if missing) the writable queue file at
    // <snapshot dir>/pending_edits.sqlite. Returns false on failure with
    // m_lastError populated. The queue connection name is separate from
    // the read-only snapshot's so they coexist cleanly.
    bool ensureQueueOpen() const;

    QSqlDatabase    m_db;
    QString         m_path;
    QString         m_connName;
    bool            m_open = false;
    mutable QString m_lastError;
    QString         m_overrideDir;  // test-only; see setOverrideDirForTesting()

    // R5 (MIP resilience): when the snapshot / queue file is MIP-encrypted at
    // rest, we decrypt it to a plaintext temp .sqlite via the bundled python
    // (MipFallback) and open THAT instead. The temp path is remembered so
    // close() can delete it. Empty when the real file opened directly.
    QString         m_decryptedSnapshotTmp;
    mutable QString m_decryptedQueueTmp;

    // R5: true iff the LAST openReadOnly() failed specifically because the
    // snapshot file was MIP-encrypted and could not be decoded (vs. mere
    // absence). Lets the MainWindow callers show a one-time LOUD warning on a
    // decode failure without nagging on a normal first-run absence.
    bool            m_lastOpenWasDecodeFailure = false;

public:
    // R5: see m_lastOpenWasDecodeFailure. Valid immediately after a false
    // return from openReadOnly().
    bool lastOpenWasDecodeFailure() const { return m_lastOpenWasDecodeFailure; }

private:

    // The server-clock timestamp captured by the last successful regenerate()
    // (R7). Exposed via lastRegenServerTimeUtc() for the server-clock test.
    QDateTime       m_lastRegenServerTime;

    // Writable queue connection — separate file from the read-only
    // snapshot. Lazily opened by ensureQueueOpen(). Mutable so
    // const accessors that diagnose the queue can populate it.
    mutable QSqlDatabase m_queueDb;
    mutable QString      m_queueConnName;
    mutable bool         m_queueOpen = false;

    // R5 IMPORTANT 3: set true by ensureQueueOpen() when the queue file EXISTS
    // but could not be opened (MIP-encrypted/undecodable, locked, corrupt);
    // cleared on a successful open. Surfaced via queueDegraded() so a false
    // pendingEditCount()==0 on an unreadable queue isn't mistaken for "empty".
    mutable bool         m_queueDegraded = false;
};

} // namespace DVE
