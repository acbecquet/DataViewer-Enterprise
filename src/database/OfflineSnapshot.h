#pragma once

#include <QObject>
#include <QString>
#include <QSqlDatabase>
#include <QDateTime>
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

    QString lastError() const { return m_lastError; }

    // Test-only: lets tests override LOCALAPPDATA without relying on
    // QStandardPaths::setTestModeEnabled() (which gives a slightly different
    // path layout). Empty string clears the override and reverts to
    // QStandardPaths::AppLocalDataLocation. Has no production callers.
    void setOverrideDirForTesting(const QString& dir);

private:
    QSqlDatabase    m_db;
    QString         m_path;
    QString         m_connName;
    bool            m_open = false;
    mutable QString m_lastError;
    QString         m_overrideDir;  // test-only; see setOverrideDirForTesting()
};

} // namespace DVE
