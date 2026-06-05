#include "utils/RecoveryManager.h"

#include <QCryptographicHash>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QSaveFile>
#include <QStandardPaths>

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

} // namespace

RecoveryManager::RecoveryManager(QObject* parent)
    : QObject(parent)
    , m_root(QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation))
{
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
    // "<kind>_<slug>_<8-hex-of-sha1(id)>.json": readable prefix + collision-free
    // suffix. Deterministic for a given (kind,id), unique across distinct ids.
    const QByteArray digest =
        QCryptographicHash::hash(id.toUtf8(), QCryptographicHash::Sha1);
    const QString hex = QString::fromLatin1(digest.toHex().left(8));
    return kindToken(kind) + QLatin1Char('_') + slugify(id)
         + QLatin1Char('_') + hex + QStringLiteral(".json");
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

QVector<RecoveryEntry> RecoveryManager::readAll(const QString& dir) const
{
    QVector<RecoveryEntry> out;

    QFile idx(dir + QLatin1Char('/') + QLatin1String(kIndexFile));
    if (!idx.open(QIODevice::ReadOnly))
        return out;
    const QByteArray idxBytes = idx.readAll();
    idx.close();

    const QJsonDocument idxDoc = QJsonDocument::fromJson(idxBytes);
    if (!idxDoc.isArray())
        return out;

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

        // Load the blob payload. A missing/corrupt blob leaves payload empty
        // but still surfaces the index entry, so callers can report it.
        if (!e.blobFile.isEmpty()) {
            QFile blob(dir + QLatin1Char('/') + e.blobFile);
            if (blob.open(QIODevice::ReadOnly)) {
                const QByteArray blobBytes = blob.readAll();
                blob.close();
                const QJsonDocument blobDoc = QJsonDocument::fromJson(blobBytes);
                if (blobDoc.isObject())
                    e.payload = blobDoc.object();
            }
        }
        out.append(e);
    }
    return out;
}

} // namespace DVE
