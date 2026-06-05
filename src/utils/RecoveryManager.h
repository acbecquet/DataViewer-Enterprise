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
// This header carries the C3 store primitives (atomic blob + index write/read)
// and the C4 crash-detection lifecycle (move-to-_prev on startup + clear on
// clean exit). The debounced off-thread flush + state provider (C5) is layered
// on by the next task onto this same class.
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

    // --- C4: crash/clean detection lifecycle ---
    //
    // Called once at startup (before any writeItem). SingleInstance guarantees
    // a single running instance, so any leftover live Recovery/ dir means the
    // prior instance died uncleanly (crash, or the updater's std::_Exit). This
    // preserves that orphaned store as Recovery_prev/ for the reopen prompt and
    // Tools->Recover, then starts the new session with a fresh (empty) live
    // store. Only one generation is kept: a pre-existing _prev is discarded.
    //
    // Returns false if it could not cleanly promote the previous live store
    // (e.g. a locked file); callers (C6) MUST then skip wiring the rolling-flush
    // for this session, or the preserved crash data in liveDir() will be
    // overwritten.
    bool adoptPreviousSession();

    // True iff Recovery_prev holds a recoverable set (index.json present AND it
    // lists at least one entry). An empty store does not trigger the prompt.
    bool hasRecoverable() const;

    // The recoverable set from the previous session: readAll(prevDir()).
    QVector<RecoveryEntry> recoverableItems() const;

    // Clean exit: drop both the live and _prev stores and the in-memory mirror,
    // so the next startup finds nothing to recover.
    void clear();

    // Test seam: override the root directory (replaces AppLocalDataLocation).
    void setDirOverride(const QString& dir);

private:
    QString m_root;
    QVector<RecoveryEntry> m_index;   // mirror of the live index.json

    bool writeIndex();
    static QString blobName(RecoveryKind kind, const QString& id);
};

} // namespace DVE
