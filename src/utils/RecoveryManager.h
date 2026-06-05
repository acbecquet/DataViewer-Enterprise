#pragma once

#include <QObject>
#include <QString>
#include <QVector>
#include <QJsonObject>
#include <QTimer>
#include <QFuture>

#include <functional>

namespace DVE {

// Plan C auto-recovery on-disk store.
//
// Every open TPM file / sensory session / detailed-sensory session is
// snapshotted as one JSON "blob" file under <root>/Recovery, with a small
// index.json listing them. Per-item blobs mean a single edit only rewrites
// that one item's blob, not the whole working set.
//
// This header carries the C3 store primitives (atomic blob + index write/read),
// the C4 crash-detection lifecycle (move-to-_prev on startup + clear on clean
// exit), and the C5 debounced off-thread flush + state provider that keeps the
// rolling snapshot off the UI thread.
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

    // Forget ONLY the previous session: remove Recovery_prev/, leaving the live
    // Recovery/ store (this session's rolling snapshot) intact. Used when the
    // user declines the reopen prompt ("No, start fresh") so the offered session
    // is not re-offered next launch. Best-effort/logged on failure, like clear().
    void clearPrevious();

    // --- C5: debounced off-thread flush + state provider ---
    //
    // The whole point of the rolling snapshot is that it must NEVER block the
    // UI thread (the v2.0.6 "Not Responding" freezes came from synchronous disk
    // writes on the UI thread). C6 hands us a StateProvider that reads the live
    // UI model (m_loadedFiles + both sensory panels' sessions) and returns the
    // full current set as self-contained value copies (entries WITH payloads).

    // MainWindow supplies the current full state on demand (entries WITH payloads):
    using StateProvider = std::function<QVector<RecoveryEntry>()>;
    void setStateProvider(StateProvider p);

    // Arm the ~2 s debounce. Coalesces a burst of edits into a single flush.
    void noteDirty();

    // Capture via the provider, write every entry's blob, rewrite the index, and
    // prune any live blob the snapshot no longer references.
    //   synchronous == true  : run the disk I/O inline on the calling thread,
    //                          after first waiting out any in-flight async
    //                          worker so this is the dir's last writer. Used by
    //                          the updater's pre-std::_Exit hook (C7), where we
    //                          cannot wait for a worker via the event loop.
    //   synchronous == false : capture on the calling (UI) thread, then run the
    //                          disk I/O on a QtConcurrent worker so the UI never
    //                          blocks. Re-entrant calls coalesce via a pending
    //                          flag; only one worker flush runs at a time.
    void flushNow(bool synchronous);

    // Test seam: override the root directory (replaces AppLocalDataLocation).
    void setDirOverride(const QString& dir);

private:
    QString m_root;
    QVector<RecoveryEntry> m_index;   // mirror of the live index.json

    // C5 debounce/safety timers + provider. Both timers are parented to this and
    // fire on the thread that owns this object (the UI thread).
    QTimer* m_debounce = nullptr;     // single-shot ~2 s; coalesces edit bursts
    QTimer* m_safety   = nullptr;     // repeating ~30 s floor on snapshot age
    StateProvider m_provider;

    // C5 async re-entrancy state (touched only on the UI thread):
    bool m_flushInFlight = false;     // a worker flush is currently running
    bool m_flushPending  = false;     // another flush was requested mid-flight
    QFuture<void> m_flushFuture;      // the in-flight worker (for sync wait-out)

    // M2 efficiency gate: the ~30 s safety timer must not periodically
    // re-serialize + rewrite every open file when nothing has changed. This is
    // set true by noteDirty() and cleared whenever a flush is dispatched; the
    // safety timer skips its flush while it's false. The 2 s debounce already
    // implies a real edit, so it always flushes (and clears the flag). The sync
    // path (flushNow(true), e.g. the updater pre-exit hook) always flushes too.
    bool m_dirtySinceFlush = false;

    bool writeIndex();
    static QString blobName(RecoveryKind kind, const QString& id);

    // Dispatch the disk I/O for an already-captured, immutable snapshot. Updates
    // m_index to mirror the snapshot on the UI thread, then either runs the I/O
    // inline (synchronous) or on a worker via QtConcurrent (asynchronous).
    void dispatchFlush(const QVector<RecoveryEntry>& snapshot, bool synchronous);
};

} // namespace DVE
