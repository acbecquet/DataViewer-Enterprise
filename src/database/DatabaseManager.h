#pragma once

#include <QObject>
#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariantMap>
#include <QSqlDatabase>
#include "../pipeline/ReportData.h"

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

class DatabaseManager : public QObject {
    Q_OBJECT

public:
    explicit DatabaseManager(QObject* parent = nullptr);
    ~DatabaseManager() override;

    bool open(const QString& dbPath);
    void close();
    bool isOpen() const;

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

    // ── Settings key/value store ─────────────────────────────────────────────
    bool setSetting(const QString& key, const QString& value);
    QString getSetting(const QString& key, const QString& defaultVal = "") const;

    QString lastError() const { return m_lastError; }

private:
    QSqlDatabase m_db;
    QString      m_lastError;
    bool         m_open = false;

    bool initSchema();
    void logDebug(const QString& msg) const;
};

} // namespace DVE
