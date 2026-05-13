#pragma once

#include <QObject>
#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariantMap>
#include "ConfigLoader.h"
#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"

namespace DVE {

class PostgresConnection;
class IdentityManager;

// ── WriteResult — outcome of a save method that participates in optimistic
//    concurrency control. Defined ahead of class DatabaseManager so callers
//    that include this header can switch on the enum.
//
//    Naming convention: the primary "save with rich result" API is
//    `tryWriteXxx(...) -> WriteResult`. The historical `saveXxx(...) -> bool`
//    shims forward to tryWriteXxx and return `result == Success`. Phase 7
//    (Task 21) will migrate MainWindow call sites onto the tryWrite* API
//    via ConflictResolver. Until then, the bool overloads keep working.
//
//    `OfflineReadOnly` is a forward declaration for Plan C — no method
//    currently returns it. It exists in the enum now so Plan C doesn't
//    have to re-ship this header.
enum class WriteResult {
    Success,
    VersionMismatch,    // expected version stale; server row has newer version
    RowDeleted,         // row no longer exists at this id
    UniqueViolation,    // hit a unique constraint (e.g., duplicate file_path)
    OfflineReadOnly,    // reserved for Plan C; saves blocked while offline
    OtherError          // generic SQL failure; details in lastError()
};

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

    bool open(const DbConfig& cfg, IdentityManager* identity);
    void close();
    bool isOpen() const;
    QString currentPath() const;

    // ── Hierarchical file storage ────────────────────────────────────────────
    // Stores the full file hierarchy (tests, samples, data rows, images).
    //
    // Optimistic concurrency: if result.id != -1 and result.version > 0,
    // the file row is UPDATEd with WHERE id = ? AND version = ?. A zero
    // rowcount triggers a follow-up SELECT to distinguish VersionMismatch
    // (row exists, version stale) from RowDeleted (row no longer there).
    // INSERT path returns UniqueViolation on SQLSTATE 23505 (duplicate
    // file_path). On Success, children (tests/samples/data_rows/images)
    // are wiped via DELETE WHERE file_id = ? and re-inserted.
    WriteResult tryWriteFile(const FileResult& result);
    // Mutable overload — identical to the const-ref version except that on
    // WriteResult::Success it writes the post-save id + version back into
    // `result`. Callers that need to round-trip the file across multiple
    // saves (e.g., MainWindow's recreate handler after a RowDeleted) MUST
    // use this overload; the const-ref version above is fire-and-forget.
    WriteResult tryWriteFile(FileResult& result);
    // Bool shim — returns true iff tryWriteFile returned Success. Kept for
    // existing MainWindow call sites until Phase 7 routes them through
    // ConflictResolver.
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
    //
    // Optimistic concurrency: identical scheme to tryWriteFile. The by-ref
    // overload also writes the post-save id and version back into the
    // struct so the next save round-trips them.
    WriteResult tryWriteSensorySession(const SensorySession& s);
    WriteResult tryWriteSensorySession(SensorySession& s);
    // Bool shims — return true iff Success. The by-ref overload still
    // populates s.id on Success (and now s.version too).
    bool saveSensorySession(const SensorySession& s);
    // Overload that populates s.id with the autoincrement id assigned by
    // Postgres on insert, plus s.version with the server-assigned version.
    // Post-condition: on true return, s.id is normally > 0 — but if a
    // concurrent writer deletes the row between the INSERT and the
    // lookup-by-natural-key, s.id stays at its prior value (-1 for a
    // fresh struct). Callers that depend on a valid id should verify
    // s.id > 0.
    bool saveSensorySession(SensorySession& s);
    QVector<SensorySession> loadSensorySessions() const;
    SensorySession loadSensorySession(int id) const;
    QVector<SensoryRecord> listSensoryRecords() const;
    bool removeSensorySession(int id);
    QString nextDefaultTestName() const;

    // ── Detailed sensory sessions ───────────────────────────────────────────
    WriteResult tryWriteDetailedSensorySession(const DetailedSensorySession& s);
    bool saveDetailedSensorySession(const DetailedSensorySession& s);
    QVector<DetailedSensorySession> loadDetailedSensorySessions() const;
    DetailedSensorySession loadDetailedSensorySession(int id) const;
    QVector<DetailedSensoryRecord> listDetailedSensoryRecords() const;
    bool removeDetailedSensorySession(int id);

    // ── Layout JSON persistence (sensory report preview) ─────────────────────
    QString loadSensoryLayout(int sessionId) const;
    bool    saveSensoryLayout(int sessionId, const QString& layoutJson);

    QString loadCumulativeLayout() const;
    bool    saveCumulativeLayout(const QString& layoutJson);

    // ── Settings key/value store ─────────────────────────────────────────────
    bool setSetting(const QString& key, const QString& value);
    QString getSetting(const QString& key, const QString& defaultVal = "") const;

    QString lastError() const { return m_lastError; }

private:
    PostgresConnection* m_pg = nullptr;
    IdentityManager*    m_identity = nullptr;
    // Read methods (loadFile, listFiles, getSetting, etc.) are const but
    // still need to clear/set m_lastError on entry/error. mutable lets them
    // do that without breaking the const contract; m_lastError is purely a
    // diagnostic side-channel.
    mutable QString     m_lastError;
    bool                m_open = false;

    void logDebug(const QString& msg) const;
};

} // namespace DVE
