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
    // Overload that populates s.id with the autoincrement id assigned by SQLite.
    // Use this when you need to persist the id (e.g., for layout JSON anchoring).
    // Post-condition: on true return, s.id is normally > 0 — but if a concurrent
    // writer deletes the row between the INSERT OR REPLACE and the lookup-by-
    // natural-key, s.id stays at its prior value (-1 for a fresh struct).
    // Callers that depend on a valid id should verify s.id > 0.
    bool saveSensorySession(SensorySession& s);
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
    QString             m_lastError;
    bool                m_open = false;

    void logDebug(const QString& msg) const;
};

} // namespace DVE
