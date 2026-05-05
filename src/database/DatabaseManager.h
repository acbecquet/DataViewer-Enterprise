#pragma once

#include <QObject>
#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariantMap>
#include <QSqlDatabase>
#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"

namespace DVE {

struct FileRecord {
    int    id;
    QString filePath;
    QString fileName;
    QString loadedAt;
    QString templateVersion;
    int     sheetCount;
    int     sampleCount;
};

struct SensoryRecord {
    int     id;
    QString sessionName;
    QString testTitle;
    QString assessorName;
    QString testerName;
    QString media;
    QString date;
    int     sampleCount;
};

struct DetailedSensoryRecord {
    int     id;
    QString sessionName;
    QString testTitle;
    QString assessorName;
    QString testerName;
    QString media;
    QString date;
    int     sampleCount;
};

class DatabaseManager : public QObject {
    Q_OBJECT

public:
    explicit DatabaseManager(QObject* parent = nullptr);
    ~DatabaseManager() override;

    bool open(const QString& dbPath);
    void close();
    bool isOpen() const;
    QString currentPath() const { return m_currentPath; }

    // ── Hierarchical file storage ────────────────────────────────────────────
    // Stores the full file hierarchy (tests, samples, data rows, images).
    // If the file already exists in DB (by file_path), replaces it entirely.
    bool saveFile(const FileResult& result);

    // Quick existence check
    bool hasFile(const QString& filePath) const;

    // Load a full FileResult from the database (inverse of saveFile).
    // Returns an empty FileResult (filePath.isEmpty()) on failure.
    FileResult loadFile(int id) const;
    FileResult loadFileByPath(const QString& filePath) const;

    // ── File listing / removal ───────────────────────────────────────────────
    QVector<FileRecord> listFiles() const;
    bool removeFile(int id);

    // Remove duplicates: keep only the N most recent entries per file_name.
    // Also deletes entries with "unknown" template version (corrupt/empty).
    int deduplicateFiles(int keepPerName = 3);

    // Recent files (last 20)
    QStringList recentFilePaths() const;

    // ── Sensory sessions ─────────────────────────────────────────────────────
    bool saveSensorySession(const SensorySession& s);
    QVector<SensorySession> loadSensorySessions() const;
    SensorySession loadSensorySession(int id) const;
    QVector<SensoryRecord> listSensoryRecords() const;
    bool removeSensorySession(int id);
    QString nextDefaultTestName() const;

    // ── Detailed sensory sessions ───────────────────────────────────────────
    bool saveDetailedSensorySession(const DetailedSensorySession& s);
    QVector<DetailedSensorySession> loadDetailedSensorySessions() const;
    DetailedSensorySession loadDetailedSensorySession(int id) const;
    QVector<DetailedSensoryRecord> listDetailedSensoryRecords() const;
    bool removeDetailedSensorySession(int id);

    // ── Settings key/value store ─────────────────────────────────────────────
    bool setSetting(const QString& key, const QString& value);
    QString getSetting(const QString& key, const QString& defaultVal = "") const;

    QString lastError() const { return m_lastError; }

    // ── Cross-machine lock (prevents concurrent open of a synced DB) ─────────
    // open() acquires an exclusive lock by writing <dbPath>.lock with this
    // host's name + PID + timestamp. If the lock is already held by *this*
    // machine, open() reuses it (treats prior-crash recovery as OK). If held
    // by another machine, open() returns false and lockHolder() returns the
    // foreign hostname so the UI can prompt the user.
    struct LockInfo {
        bool    present = false;
        bool    foreign = false;       // present and held by a DIFFERENT host
        QString host;
        QString pid;
        QString acquiredAt;            // ISO-8601 timestamp
    };
    LockInfo lockInfo() const { return m_lockInfo; }
    bool forceReleaseLock(const QString& dbPath);  // delete the .lock file unconditionally

private:
    QSqlDatabase m_db;
    QString      m_lastError;
    QString      m_currentPath;
    bool         m_open = false;
    LockInfo     m_lockInfo;
    QString      m_lockPath;     // <dbPath>.lock — empty if no lock taken

    bool initSchema();
    void logDebug(const QString& msg) const;

    // Lock helpers
    static QString thisHostName();
    static QString lockPathFor(const QString& dbPath);
    LockInfo readLockFile(const QString& path) const;
    bool writeLockFile(const QString& path);
    void removeLockFile();
    static bool pathLooksCloudSynced(const QString& dbPath);
};

} // namespace DVE
