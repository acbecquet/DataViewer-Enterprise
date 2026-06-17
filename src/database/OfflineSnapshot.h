#pragma once

#include <QObject>
#include <QString>
#include <QSqlDatabase>
#include <QDateTime>
#include <QVariant>
#include <functional>
#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"
#include "DatabaseManager.h"  // for FileRecord / SensoryRecord / DetailedSensoryRecord

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
    int pendingEditCount() const;

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

    // The server-clock timestamp captured by the last successful regenerate()
    // (R7). Exposed via lastRegenServerTimeUtc() for the server-clock test.
    QDateTime       m_lastRegenServerTime;

    // Writable queue connection — separate file from the read-only
    // snapshot. Lazily opened by ensureQueueOpen(). Mutable so
    // const accessors that diagnose the queue can populate it.
    mutable QSqlDatabase m_queueDb;
    mutable QString      m_queueConnName;
    mutable bool         m_queueOpen = false;
};

} // namespace DVE
