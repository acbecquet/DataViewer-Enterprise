#pragma once

#include <QObject>
#include <QString>
#include <QVector>
#include <QJsonObject>

namespace DVE {

// Plan C auto-recovery on-disk store.
//
// Every open TPM file / sensory session / detailed-sensory session is
// snapshotted as one JSON "blob" file under <root>/Recovery, with a small
// index.json listing them. Per-item blobs mean a single edit only rewrites
// that one item's blob, not the whole working set.
//
// This header carries ONLY the C3 store primitives (atomic blob + index
// write/read). Crash detection (move-to-_prev on startup, C4) and the
// debounced off-thread flush + state provider (C5) are layered on by later
// tasks onto this same class.
enum class RecoveryKind { Tpm, Sensory, Detailed };

struct RecoveryEntry {
    RecoveryKind kind = RecoveryKind::Tpm;
    QString id;            // stable per-item id (e.g. file path or session key)
    QString displayName;
    QString sourcePath;
    bool dirty = false;
    QString blobFile;      // basename of the blob, e.g. "tpm_<id>.json"
    QJsonObject payload;   // populated by readAll(); ignored by writeItem's index entry
};

class RecoveryManager : public QObject
{
    Q_OBJECT

public:
    // Default root is QStandardPaths::AppLocalDataLocation, i.e.
    // %LOCALAPPDATA%/SDR/DataViewer Enterprise, alongside the existing
    // dataviewer.log / snapshot.sqlite. Override with setDirOverride() in tests.
    explicit RecoveryManager(QObject* parent = nullptr);

    QString liveDir() const;   // <root>/Recovery
    QString prevDir() const;   // <root>/Recovery_prev

    // Write one item: serialize payload to its blob atomically (QSaveFile),
    // replace-or-append the matching index entry (by kind+id), rewrite the
    // index atomically. Returns false if any disk write fails.
    bool writeItem(const RecoveryEntry& e, const QJsonObject& payload);

    // Delete the blob for (kind,id), drop it from the index, rewrite the index.
    bool removeItem(RecoveryKind kind, const QString& id);

    // Pure read: parse dir/index.json into entries, then load each entry's blob
    // into its payload. Does not touch the live in-memory index mirror.
    QVector<RecoveryEntry> readAll(const QString& dir) const;

    // Test seam: override the root directory (replaces AppLocalDataLocation).
    void setDirOverride(const QString& dir);

private:
    QString m_root;
    QVector<RecoveryEntry> m_index;   // mirror of the live index.json

    bool writeIndex();
    static QString blobName(RecoveryKind kind, const QString& id);
};

} // namespace DVE
