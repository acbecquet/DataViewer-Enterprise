#include "utils/RecoveryManager.h"
#include "utils/MipFallback.h"

#include <QCryptographicHash>
#include <QDebug>
#include <QDir>
#include <QFile>
#include <QFutureWatcher>
#include <QJsonArray>
#include <QJsonDocument>
#include <QSaveFile>
#include <QSet>
#include <QStandardPaths>
#include <QtConcurrent>

namespace DVE {

namespace {

constexpr auto kIndexFile = "index.json";

// Short, stable string for a RecoveryKind. Used both as the blob-name prefix
// and as the on-disk index "kind" value, so detection survives schema reads.
QString kindToken(RecoveryKind kind)
{
    switch (kind) {
    case RecoveryKind::Tpm:      return QStringLiteral("tpm");
    case RecoveryKind::Sensory:  return QStringLiteral("sensory");
    case RecoveryKind::Detailed: return QStringLiteral("detailed");
    }
    return QStringLiteral("tpm");
}

RecoveryKind kindFromToken(const QString& token)
{
    if (token == QLatin1String("sensory"))  return RecoveryKind::Sensory;
    if (token == QLatin1String("detailed")) return RecoveryKind::Detailed;
    return RecoveryKind::Tpm;
}

// Replace anything that isn't filesystem-safe with '_'. Keeps the basename
// human-readable; uniqueness is guaranteed separately by a hash suffix (an id
// like "C:/a/b.xlsx" and "C:\a\b.xlsx" would otherwise collide on the slug).
QString slugify(const QString& id)
{
    QString out;
    out.reserve(id.size());
    for (const QChar c : id) {
        const bool safe = (c >= QLatin1Char('a') && c <= QLatin1Char('z'))
                       || (c >= QLatin1Char('A') && c <= QLatin1Char('Z'))
                       || (c >= QLatin1Char('0') && c <= QLatin1Char('9'))
                       || c == QLatin1Char('.') || c == QLatin1Char('-')
                       || c == QLatin1Char('_');
        out.append(safe ? c : QLatin1Char('_'));
    }
    return out;
}

// "<kind>_<slug>_<8-hex-of-sha1(id)>.json": readable prefix + collision-free
// suffix. Deterministic for a given (kind,id), unique across distinct ids.
// Shared by RecoveryManager::blobName (the static used by writeItem/removeItem)
// and by the off-thread flush worker below, so a blob's name is computed
// identically on the UI thread and on the worker.
QString blobNameFor(RecoveryKind kind, const QString& id)
{
    const QByteArray digest =
        QCryptographicHash::hash(id.toUtf8(), QCryptographicHash::Sha1);
    const QString hex = QString::fromLatin1(digest.toHex().left(8));
    return kindToken(kind) + QLatin1Char('_') + slugify(id)
         + QLatin1Char('_') + hex + QStringLiteral(".json");
}

// Atomic write via QSaveFile: writes to a temp sibling then renames into place
// on commit(), so a crash mid-write never leaves a half-written blob/index.
bool writeFileAtomic(const QString& path, const QByteArray& bytes)
{
    QSaveFile f(path);
    if (!f.open(QIODevice::WriteOnly))
        return false;
    if (f.write(bytes) != bytes.size()) {
        f.cancelWriting();
        return false;
    }
    return f.commit();
}

// Serialize a snapshot's index entries to index.json bytes. Mirrors
// RecoveryManager::writeIndex's schema exactly so a flushed store reads back
// identically to a writeItem-built one. The blob payload lives in the blob, not
// the index, so payloads are intentionally not emitted here.
QByteArray indexBytesFor(const QVector<RecoveryEntry>& entries)
{
    QJsonArray arr;
    for (const RecoveryEntry& e : entries) {
        QJsonObject o;
        o[QStringLiteral("kind")]        = kindToken(e.kind);
        o[QStringLiteral("id")]          = e.id;
        o[QStringLiteral("displayName")] = e.displayName;
        o[QStringLiteral("sourcePath")]  = e.sourcePath;
        o[QStringLiteral("dirty")]       = e.dirty;
        o[QStringLiteral("blobFile")]    = blobNameFor(e.kind, e.id);
        arr.append(o);
    }
    return QJsonDocument(arr).toJson(QJsonDocument::Indented);
}

// The off-thread flush worker (C5). MUST be self-contained: it takes an
// immutable snapshot value + the live dir path and touches NOTHING on the
// RecoveryManager (no m_index, no signals, no other QObject state), so it is
// safe to run on a QtConcurrent thread. The re-entrancy guard in dispatchFlush
// ensures only ONE writer touches the live dir at a time (one async worker, or
// a sync flush that first waits any worker out), so the per-file atomic writes
// stay race-free.
//
// Steps:
//   1. mkpath the live dir;
//   2. write every snapshot entry's blob atomically;
//   3. rewrite index.json atomically from the snapshot;
//   4. prune: list *.json in the live dir, delete any whose name is not one of
//      the snapshot's expected blob names (computed here from the snapshot, not
//      from m_index, so the worker is fully self-contained). index.json itself
//      is never a blob name, so it is never pruned.
void runFlush(const QVector<RecoveryEntry>& snapshot, const QString& liveDirPath)
{
    if (!QDir().mkpath(liveDirPath))
        return;

    // (2) blobs, and collect the set of names this snapshot legitimately owns.
    QSet<QString> keep;
    keep.insert(QString::fromLatin1(kIndexFile));   // never prune the index
    for (const RecoveryEntry& e : snapshot) {
        const QString blob = blobNameFor(e.kind, e.id);
        keep.insert(blob);
        const QByteArray bytes =
            QJsonDocument(e.payload).toJson(QJsonDocument::Compact);
        writeFileAtomic(liveDirPath + QLatin1Char('/') + blob, bytes);
    }

    // (3) fresh index from the snapshot.
    writeFileAtomic(liveDirPath + QLatin1Char('/') + QLatin1String(kIndexFile),
                    indexBytesFor(snapshot));

    // (4) prune stale blobs the snapshot no longer references.
    const QStringList stray =
        QDir(liveDirPath).entryList(QStringList{ QStringLiteral("*.json") },
                                    QDir::Files);
    for (const QString& name : stray) {
        if (!keep.contains(name))
            QFile::remove(liveDirPath + QLatin1Char('/') + name);
    }
}

} // namespace

RecoveryManager::RecoveryManager(QObject* parent)
    : QObject(parent)
    , m_root(QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation))
{
    // C5 debounce: a burst of edits restarts the 2 s window, so we flush once
    // ~2 s after the last edit rather than on every keystroke.
    m_debounce = new QTimer(this);
    m_debounce->setSingleShot(true);
    m_debounce->setInterval(2000);
    connect(m_debounce, &QTimer::timeout, this, [this]() { flushNow(false); });

    // C5 safety net: a floor on snapshot staleness for the (rare) case where a
    // noteDirty() is somehow missed -- guarantees the rolling snapshot is never
    // older than ~30 s while the app runs. Repeating; started immediately.
    //
    // M2 efficiency gate: only actually flush when something changed since the
    // last flush. Without this, the timer would re-serialize + rewrite every
    // open file's blob every 30 s even on a perfectly idle session (the real
    // provider, wired by C6, reads the whole open set on each capture). The 2 s
    // debounce path still always flushes -- it only fires right after an edit.
    m_safety = new QTimer(this);
    m_safety->setSingleShot(false);
    m_safety->setInterval(30000);
    connect(m_safety, &QTimer::timeout, this, [this]() {
        if (m_dirtySinceFlush)
            flushNow(false);
    });
    m_safety->start();
}

QString RecoveryManager::liveDir() const
{
    return m_root + QStringLiteral("/Recovery");
}

QString RecoveryManager::prevDir() const
{
    return m_root + QStringLiteral("/Recovery_prev");
}

void RecoveryManager::setDirOverride(const QString& dir)
{
    m_root = dir;
}

QString RecoveryManager::blobName(RecoveryKind kind, const QString& id)
{
    return blobNameFor(kind, id);
}

bool RecoveryManager::writeItem(const RecoveryEntry& e, const QJsonObject& payload)
{
    if (!QDir().mkpath(liveDir()))
        return false;

    const QString blob = blobName(e.kind, e.id);
    const QByteArray bytes = QJsonDocument(payload).toJson(QJsonDocument::Compact);
    if (!writeFileAtomic(liveDir() + QLatin1Char('/') + blob, bytes))
        return false;

    // Replace-or-append the index entry, matched by (kind,id). Keep insertion
    // order so the live set is stable across rewrites.
    RecoveryEntry stored = e;
    stored.blobFile = blob;
    stored.payload  = QJsonObject();   // the payload lives in the blob, not the index

    bool replaced = false;
    for (RecoveryEntry& existing : m_index) {
        if (existing.kind == stored.kind && existing.id == stored.id) {
            existing = stored;
            replaced = true;
            break;
        }
    }
    if (!replaced)
        m_index.append(stored);

    return writeIndex();
}

bool RecoveryManager::removeItem(RecoveryKind kind, const QString& id)
{
    const QString blob = blobName(kind, id);
    QFile::remove(liveDir() + QLatin1Char('/') + blob);

    for (int i = 0; i < m_index.size(); ++i) {
        if (m_index[i].kind == kind && m_index[i].id == id) {
            m_index.removeAt(i);
            break;
        }
    }
    return writeIndex();
}

bool RecoveryManager::adoptPreviousSession()
{
    // Discard any older recoverable generation: we keep only one session back.
    // If the stale _prev can't be removed (a locked file), do NOT proceed: the
    // rename below would fail too and we'd risk clobbering the live crash data.
    if (QDir(prevDir()).exists()) {
        if (!QDir(prevDir()).removeRecursively()) {
            qWarning() << "RecoveryManager: failed to remove stale prev dir" << prevDir()
                       << "- preserving live store, recovery prompt will be skipped this launch";
            return false;   // do NOT risk clobbering live
        }
    }

    // Promote the orphaned live store (if any) to _prev. A leftover live dir
    // here means the previous instance never reached its clean clear(), so its
    // contents are exactly what the reopen prompt should offer to recover. If
    // the promote fails, the crash data is still in liveDir() and the caller
    // must not wire the rolling-flush (which would overwrite it) -- signal false.
    if (QDir(liveDir()).exists()) {
        if (!QDir().rename(liveDir(), prevDir())) {
            qWarning() << "RecoveryManager: failed to promote live ->" << prevDir()
                       << "- crash data remains in" << liveDir() << "and must not be overwritten";
            return false;   // live still holds the crash data
        }
    }

    // Start the new session with an empty in-memory mirror. The live dir is now
    // gone; the first writeItem() recreates it via mkpath.
    m_index.clear();
    return true;
}

bool RecoveryManager::hasRecoverable() const
{
    // An index file alone isn't enough: an empty store (index present, no
    // entries) must not trigger the recovery prompt.
    if (!QFile::exists(prevDir() + QLatin1Char('/') + QLatin1String(kIndexFile)))
        return false;
    // R5: a present-but-undecodable store (e.g. MIP-encrypted index even the
    // python fallback couldn't read) yields an empty list here, but it is NOT
    // "nothing to recover" -- it is "couldn't read what's there". Treat that as
    // recoverable so the caller surfaces a warning rather than silently
    // dropping the user's crashed session. lastReadFailed() lets the caller
    // tell the two apart.
    const bool nonEmpty = !readAll(prevDir()).isEmpty();
    return nonEmpty || m_lastReadFailed;
}

QVector<RecoveryEntry> RecoveryManager::recoverableItems() const
{
    return readAll(prevDir());
}

void RecoveryManager::clear()
{
    if (QDir(liveDir()).exists() && !QDir(liveDir()).removeRecursively())
        qWarning() << "RecoveryManager: failed to remove live dir" << liveDir();
    if (QDir(prevDir()).exists() && !QDir(prevDir()).removeRecursively())
        qWarning() << "RecoveryManager: failed to remove prev dir" << prevDir();
    m_index.clear();
}

// -- C5: debounced off-thread flush + state provider -------------------------

void RecoveryManager::setStateProvider(StateProvider p)
{
    m_provider = std::move(p);
}

void RecoveryManager::noteDirty()
{
    // (Re)start the single-shot window: a burst of edits collapses to one flush
    // ~2 s after the final edit. The flush itself runs off the UI thread.
    // Mark the session dirty so the 30 s safety timer knows a flush is actually
    // warranted (M2 gate); the flag is cleared when the next flush is dispatched.
    m_dirtySinceFlush = true;
    m_debounce->start();
}

void RecoveryManager::flushNow(bool synchronous)
{
    // Capture the snapshot on the CALLING thread. For the debounce/safety timers
    // and the C6/C7 hooks, that is the UI thread -- and the provider reads live
    // UI data (m_loadedFiles / panel sessions), so it MUST run there. The result
    // is a self-contained value copy (entries + payloads) safe to hand to a
    // worker.
    if (!m_provider)
        return;
    // A flush is now happening (sync inline, or captured here + dispatched to a
    // worker), so the snapshot is about to reflect current state: clear the M2
    // dirty flag. Any edit landing after this point re-sets it via noteDirty().
    m_dirtySinceFlush = false;
    const QVector<RecoveryEntry> snapshot = m_provider();
    dispatchFlush(snapshot, synchronous);
}

void RecoveryManager::dispatchFlush(const QVector<RecoveryEntry>& snapshot,
                                    bool synchronous)
{
    // Mirror the snapshot into m_index on the UI thread (NEVER from the worker),
    // so hasRecoverable()/readAll consumers and the next writeItem see the same
    // live set the worker is about to persist. blobFile is filled in to match
    // what runFlush() will write; payloads are dropped from the mirror (they
    // live in the blobs).
    m_index.clear();
    m_index.reserve(snapshot.size());
    for (const RecoveryEntry& e : snapshot) {
        RecoveryEntry stored = e;
        stored.blobFile = blobName(e.kind, e.id);
        stored.payload  = QJsonObject();
        m_index.append(stored);
    }

    const QString liveDirPath = liveDir();

    if (synchronous) {
        // Inline path: used by the updater's pre-std::_Exit hook (C7), where we
        // cannot wait for a worker via the event loop. Block on any in-flight
        // async worker FIRST so this inline pass is the dir's final writer --
        // otherwise a still-running worker could prune/overwrite from a stale
        // snapshot after we exit. Blocks the calling thread, but the process is
        // about to terminate anyway.
        if (m_flushInFlight && m_flushFuture.isValid())
            m_flushFuture.waitForFinished();
        runFlush(snapshot, liveDirPath);
        return;
    }

    // Async path: keep the UI responsive. If a worker flush is already running,
    // record that another is wanted and bail -- the in-flight worker's finished
    // handler will re-fire flushNow(false) with fresh state. This bounds us to a
    // single concurrent writer of the live dir, so atomic per-file writes never
    // race each other.
    if (m_flushInFlight) {
        m_flushPending = true;
        return;
    }

    m_flushInFlight = true;
    auto* watcher = new QFutureWatcher<void>(this);
    connect(watcher, &QFutureWatcher<void>::finished, this, [this, watcher]() {
        watcher->deleteLater();
        m_flushInFlight = false;
        m_flushFuture = QFuture<void>();   // drop the finished future
        // Coalesced request(s) arrived mid-flight: run one more pass with the
        // latest state (re-captures via the provider on the UI thread).
        if (m_flushPending) {
            m_flushPending = false;
            flushNow(false);
        }
    });
    // Capture the snapshot + path by value so the worker owns immutable copies.
    m_flushFuture = QtConcurrent::run([snapshot, liveDirPath]() {
        runFlush(snapshot, liveDirPath);
    });
    watcher->setFuture(m_flushFuture);
}

bool RecoveryManager::writeIndex()
{
    QJsonArray arr;
    for (const RecoveryEntry& e : m_index) {
        QJsonObject o;
        o[QStringLiteral("kind")]        = kindToken(e.kind);
        o[QStringLiteral("id")]          = e.id;
        o[QStringLiteral("displayName")] = e.displayName;
        o[QStringLiteral("sourcePath")]  = e.sourcePath;
        o[QStringLiteral("dirty")]       = e.dirty;
        o[QStringLiteral("blobFile")]    = e.blobFile;
        arr.append(o);
    }
    const QByteArray bytes =
        QJsonDocument(arr).toJson(QJsonDocument::Indented);
    return writeFileAtomic(liveDir() + QLatin1Char('/') + QLatin1String(kIndexFile),
                           bytes);
}

bool RecoveryManager::readBytesMipAware(const QString& filePath,
                                        QByteArray& out) const
{
    out.clear();

    // Genuinely absent file: not a failure, not encrypted -- the caller decides
    // whether absence matters (e.g. a missing index.json simply means no store).
    if (!QFile::exists(filePath))
        return false;

    // Fast path: a readable, plaintext file.
    if (!looksEncrypted(filePath)) {
        QFile f(filePath);
        if (f.open(QIODevice::ReadOnly)) {
            out = f.readAll();
            f.close();
            return true;
        }
        // Present but unreadable for a non-MIP reason (locked / perms). LOUD.
        m_lastReadFailed = true;
        m_lastError = QStringLiteral("RecoveryManager: cannot open recovery file ")
                      + filePath + QStringLiteral(": ") + f.errorString();
        qWarning().noquote() << m_lastError;
        return false;
    }

    // MIP-encrypted at rest: route through the bundled MIP-allowlisted python
    // to a plaintext temp copy, then read THAT. Mirrors ExcelReader's approach
    // (openpyxl runs inside the authenticated user session, sees plaintext).
    const QString python = resolveBundledPython();
    QString err;
    const QString tmp = decryptToTempViaPython(filePath, python, err);
    if (tmp.isEmpty()) {
        m_lastReadFailed = true;
        m_lastError = QStringLiteral("RecoveryManager: recovery file is "
                                     "MIP-encrypted and could not be decrypted (")
                      + filePath + QStringLiteral("): ") + err;
        qWarning().noquote() << m_lastError;
        return false;
    }
    QFile f(tmp);
    const bool ok = f.open(QIODevice::ReadOnly);
    if (ok) {
        out = f.readAll();
        f.close();
    } else {
        m_lastReadFailed = true;
        m_lastError = QStringLiteral("RecoveryManager: decrypted temp copy "
                                     "unreadable for ") + filePath;
        qWarning().noquote() << m_lastError;
    }
    QFile::remove(tmp);   // we hold the bytes; the temp is no longer needed
    return ok;
}

QVector<RecoveryEntry> RecoveryManager::readAll(const QString& dir) const
{
    QVector<RecoveryEntry> out;
    // Reset the per-call read-failure state; readBytesMipAware sets it on a
    // present-but-undecodable file.
    m_lastReadFailed = false;
    m_lastError.clear();

    const QString idxPath = dir + QLatin1Char('/') + QLatin1String(kIndexFile);

    // Absent index = no store at all (a true first run / cleared session).
    // readBytesMipAware returns false WITHOUT setting the failure flag for an
    // absent file, so this stays the silent "nothing to recover" case.
    if (!QFile::exists(idxPath))
        return out;

    QByteArray idxBytes;
    if (!readBytesMipAware(idxPath, idxBytes)) {
        // The index EXISTS (checked above) but we couldn't read/decode it.
        // m_lastReadFailed + m_lastError are already set + logged. Returning
        // empty here would look identical to "nothing to recover" -- the
        // failure flag is exactly how the caller tells them apart.
        return out;
    }

    const QJsonDocument idxDoc = QJsonDocument::fromJson(idxBytes);
    if (!idxDoc.isArray()) {
        m_lastReadFailed = true;
        m_lastError = QStringLiteral("RecoveryManager: recovery index is not "
                                     "valid JSON (") + idxPath + QLatin1Char(')');
        qWarning().noquote() << m_lastError;
        return out;
    }

    const QJsonArray arr = idxDoc.array();
    for (const QJsonValue& v : arr) {
        const QJsonObject o = v.toObject();
        RecoveryEntry e;
        e.kind        = kindFromToken(o[QStringLiteral("kind")].toString());
        e.id          = o[QStringLiteral("id")].toString();
        e.displayName = o[QStringLiteral("displayName")].toString();
        e.sourcePath  = o[QStringLiteral("sourcePath")].toString();
        e.dirty       = o[QStringLiteral("dirty")].toBool();
        e.blobFile    = o[QStringLiteral("blobFile")].toString();

        // Load the blob payload, MIP-aware. A missing/corrupt/undecodable blob
        // leaves payload empty but still surfaces the index entry (so the user
        // sees the item exists). An undecodable blob also trips the loud
        // read-failure flag via readBytesMipAware so the caller can warn.
        if (!e.blobFile.isEmpty()) {
            const QString blobPath = dir + QLatin1Char('/') + e.blobFile;
            QByteArray blobBytes;
            if (readBytesMipAware(blobPath, blobBytes)) {
                const QJsonDocument blobDoc = QJsonDocument::fromJson(blobBytes);
                if (blobDoc.isObject())
                    e.payload = blobDoc.object();
                else {
                    m_lastReadFailed = true;
                    m_lastError = QStringLiteral("RecoveryManager: recovery blob "
                                                 "is not valid JSON (")
                                  + blobPath + QLatin1Char(')');
                    qWarning().noquote() << m_lastError;
                }
            }
            // readBytesMipAware already logged + flagged a present-but-
            // undecodable blob; a genuinely-absent blob is left silent.
        }
        out.append(e);
    }
    return out;
}

} // namespace DVE
