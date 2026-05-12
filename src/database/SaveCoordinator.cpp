#include "SaveCoordinator.h"

#include <QDebug>
#include <QDir>
#include <QFileInfo>
#include <QMessageBox>

#include "DatabaseManager.h"
#include "ConflictResolver.h"

namespace DVE {

namespace {

// -- Conflict-diff field builders --------------------------------------------
// Each builds a QVariantMap of user-visible scalar fields. The
// VersionMismatchDialog renders these side-by-side so the user can pick
// which value to keep per field. Hierarchical data (sheets, samples,
// scores) is intentionally NOT included -- the dialog only supports
// scalar merges; deep merging is deferred.

QVariantMap diffFieldsForFile(const FileResult& f) {
    QVariantMap m;
    m["filePath"]        = f.filePath;
    m["fileName"]        = f.fileName;
    m["templateVersion"] = f.templateVersion;
    m["sheetCount"]      = f.sheets.size();
    int sampleCount = 0;
    for (const auto& sh : f.sheets) sampleCount += sh.samples.size();
    m["sampleCount"]     = sampleCount;
    return m;
}

QVariantMap diffFieldsForSensory(const SensorySession& s) {
    QVariantMap m;
    m["sessionName"]  = s.sessionName;
    m["testerName"]   = s.testerName;
    m["assessorName"] = s.assessorName;
    m["date"]         = s.date;
    m["media"]        = s.media;
    m["samples"]      = s.samples.size();
    return m;
}

QVariantMap diffFieldsForDetailed(const DetailedSensorySession& s) {
    QVariantMap m;
    m["sessionName"]  = s.sessionName;
    m["testerName"]   = s.testerName;
    m["assessorName"] = s.assessorName;
    m["date"]         = s.date;
    m["media"]        = s.media;
    return m;
}

// -- Apply user-selected scalars from `merged` back onto the local struct ----
// Only the few user-visible scalar fields are mutated; vectors/maps stay put.

void applyMergedToFile(FileResult& f, const QVariantMap& merged) {
    if (merged.contains("filePath"))        f.filePath        = merged["filePath"].toString();
    if (merged.contains("fileName"))        f.fileName        = merged["fileName"].toString();
    if (merged.contains("templateVersion")) f.templateVersion = merged["templateVersion"].toString();
    // sheetCount / sampleCount are read-only diagnostics; cannot meaningfully
    // be applied back to the struct since the underlying vectors are richer.
}

void applyMergedToSensory(SensorySession& s, const QVariantMap& merged) {
    if (merged.contains("sessionName"))  s.sessionName  = merged["sessionName"].toString();
    if (merged.contains("testerName"))   s.testerName   = merged["testerName"].toString();
    if (merged.contains("assessorName")) s.assessorName = merged["assessorName"].toString();
    if (merged.contains("date"))         s.date         = merged["date"].toString();
    if (merged.contains("media"))        s.media        = merged["media"].toString();
    // samples count is informational only.
}

void applyMergedToDetailed(DetailedSensorySession& s, const QVariantMap& merged) {
    if (merged.contains("sessionName"))  s.sessionName  = merged["sessionName"].toString();
    if (merged.contains("testerName"))   s.testerName   = merged["testerName"].toString();
    if (merged.contains("assessorName")) s.assessorName = merged["assessorName"].toString();
    if (merged.contains("date"))         s.date         = merged["date"].toString();
    if (merged.contains("media"))        s.media        = merged["media"].toString();
}

// -- Rename suggestion helpers -----------------------------------------------
QString suggestRename(const QString& name) {
    if (name.isEmpty()) return QStringLiteral("Untitled (2)");
    // If ends with " (N)" bump it; otherwise append " (2)".
    const int open  = name.lastIndexOf(QLatin1Char('('));
    const int close = name.lastIndexOf(QLatin1Char(')'));
    if (open >= 0 && close == name.size() - 1 && close > open + 1) {
        bool ok = false;
        const int n = name.mid(open + 1, close - open - 1).trimmed().toInt(&ok);
        if (ok) {
            return name.left(open).trimmed() + QStringLiteral(" (") +
                   QString::number(n + 1) + QStringLiteral(")");
        }
    }
    return name + QStringLiteral(" (2)");
}

// Same logic for a file name -- bump the basename, keep the extension.
QString suggestRenameFile(const QString& fileName) {
    if (fileName.isEmpty()) return QStringLiteral("Untitled (2).xlsx");
    const QFileInfo fi(fileName);
    const QString base    = fi.completeBaseName();
    const QString suffix  = fi.suffix();
    const QString newBase = suggestRename(base);
    return suffix.isEmpty() ? newBase
                            : newBase + QLatin1Char('.') + suffix;
}

} // anonymous namespace

// ----------------------------------------------------------------------------
SaveCoordinator::SaveCoordinator(DatabaseManager* db,
                                 ConflictResolver* resolver,
                                 QObject* parent)
    : QObject(parent), m_db(db), m_resolver(resolver) {}

// --- saveFile ---------------------------------------------------------------
SaveCoordinator::Outcome SaveCoordinator::saveFile(FileResult& result,
                                                    QWidget* parent) {
    if (!m_db) return Failed;

    auto warn = [&](const QString& msg) {
        QMessageBox::warning(parent, QStringLiteral("Database Save Failed"), msg);
    };

    const WriteResult wr = m_db->tryWriteFile(result);
    switch (wr) {
    case WriteResult::Success:
        return Saved;

    case WriteResult::VersionMismatch: {
        if (!m_resolver) { warn(QStringLiteral("Version mismatch (no resolver wired)")); return Failed; }
        const FileResult fresh = m_db->loadFileByPath(result.filePath);
        const QVariantMap mine   = diffFieldsForFile(result);
        const QVariantMap theirs = diffFieldsForFile(fresh);
        QVariantMap merged;
        const auto choice = m_resolver->resolveVersionMismatch(
            QStringLiteral("File"),
            QString(), QString(),
            mine, theirs, &merged, parent);

        switch (choice) {
        case ConflictResolver::KeepMine:
            result.version = fresh.version;
            if (result.id <= 0) result.id = fresh.id;
            return (m_db->tryWriteFile(result) == WriteResult::Success) ? Saved : Failed;
        case ConflictResolver::TakeTheirs:
            emit localEditsDiscarded(QStringLiteral("files"), fresh.id);
            return UserCancelled;
        case ConflictResolver::ApplySelection:
            applyMergedToFile(result, merged);
            result.version = fresh.version;
            if (result.id <= 0) result.id = fresh.id;
            return (m_db->tryWriteFile(result) == WriteResult::Success) ? Saved : Failed;
        case ConflictResolver::CancelVersion:
        default:
            return UserCancelled;
        }
    }

    case WriteResult::RowDeleted: {
        if (!m_resolver) { warn(QStringLiteral("Row deleted (no resolver wired)")); return Failed; }
        const QVariantMap myEdits = diffFieldsForFile(result);
        const auto choice = m_resolver->resolveDeleted(
            QStringLiteral("File"), QString(), myEdits, parent);

        if (choice == ConflictResolver::Recreate) {
            result.id = -1;
            result.version = 0;
            const WriteResult retry = m_db->tryWriteFile(result);
            if (retry == WriteResult::Success) return Saved;
            warn(QStringLiteral("Could not recreate row: ") + m_db->lastError());
            return Failed;
        }
        // Discard
        emit localEditsDiscarded(QStringLiteral("files"), result.id);
        return UserCancelled;
    }

    case WriteResult::UniqueViolation: {
        if (!m_resolver) { warn(QStringLiteral("Unique violation (no resolver wired)")); return Failed; }
        const QString desc = QStringLiteral(
            "A file with this name or path already exists in the database.");
        const QString suggested = suggestRenameFile(result.fileName);
        QString renamed;
        const auto choice = m_resolver->resolveUniqueViolation(
            desc, suggested, &renamed, parent);

        switch (choice) {
        case ConflictResolver::RenameMine:
            if (!renamed.isEmpty()) {
                result.fileName = renamed;
                const QFileInfo fi(result.filePath);
                if (!fi.absoluteDir().path().isEmpty()) {
                    result.filePath = fi.absoluteDir().filePath(renamed);
                }
            }
            result.id = -1;
            result.version = 0;
            return (m_db->tryWriteFile(result) == WriteResult::Success) ? Saved : Failed;
        case ConflictResolver::OpenExisting:
            emit openExistingRequested(QStringLiteral("files"), result.filePath);
            return UserCancelled;
        case ConflictResolver::CancelUnique:
        default:
            return UserCancelled;
        }
    }

    case WriteResult::OfflineReadOnly:
        warn(QStringLiteral("Saves are disabled while offline."));
        return Failed;

    case WriteResult::OtherError:
    default:
        warn(QStringLiteral("Save failed: ") + m_db->lastError());
        return Failed;
    }
}

// --- saveSensorySession -----------------------------------------------------
SaveCoordinator::Outcome SaveCoordinator::saveSensorySession(SensorySession& s,
                                                              QWidget* parent) {
    if (!m_db) return Failed;

    auto warn = [&](const QString& msg) {
        QMessageBox::warning(parent, QStringLiteral("Database Save Failed"), msg);
    };

    const WriteResult wr = m_db->tryWriteSensorySession(s);
    switch (wr) {
    case WriteResult::Success:
        return Saved;

    case WriteResult::VersionMismatch: {
        if (!m_resolver) { warn(QStringLiteral("Version mismatch (no resolver wired)")); return Failed; }
        const SensorySession fresh = (s.id > 0) ? m_db->loadSensorySession(s.id)
                                                : SensorySession{};
        const QVariantMap mine   = diffFieldsForSensory(s);
        const QVariantMap theirs = diffFieldsForSensory(fresh);
        QVariantMap merged;
        const auto choice = m_resolver->resolveVersionMismatch(
            QStringLiteral("Sensory Session"),
            QString(), QString(),
            mine, theirs, &merged, parent);

        switch (choice) {
        case ConflictResolver::KeepMine:
            s.version = fresh.version;
            if (s.id <= 0) s.id = fresh.id;
            return (m_db->tryWriteSensorySession(s) == WriteResult::Success) ? Saved : Failed;
        case ConflictResolver::TakeTheirs:
            emit localEditsDiscarded(QStringLiteral("sensory_sessions"), fresh.id);
            return UserCancelled;
        case ConflictResolver::ApplySelection:
            applyMergedToSensory(s, merged);
            s.version = fresh.version;
            if (s.id <= 0) s.id = fresh.id;
            return (m_db->tryWriteSensorySession(s) == WriteResult::Success) ? Saved : Failed;
        case ConflictResolver::CancelVersion:
        default:
            return UserCancelled;
        }
    }

    case WriteResult::RowDeleted: {
        if (!m_resolver) { warn(QStringLiteral("Row deleted (no resolver wired)")); return Failed; }
        const QVariantMap myEdits = diffFieldsForSensory(s);
        const auto choice = m_resolver->resolveDeleted(
            QStringLiteral("Sensory Session"), QString(), myEdits, parent);

        if (choice == ConflictResolver::Recreate) {
            s.id = -1;
            s.version = 0;
            const WriteResult retry = m_db->tryWriteSensorySession(s);
            if (retry == WriteResult::Success) return Saved;
            warn(QStringLiteral("Could not recreate row: ") + m_db->lastError());
            return Failed;
        }
        emit localEditsDiscarded(QStringLiteral("sensory_sessions"), s.id);
        return UserCancelled;
    }

    case WriteResult::UniqueViolation: {
        if (!m_resolver) { warn(QStringLiteral("Unique violation (no resolver wired)")); return Failed; }
        const QString desc = QStringLiteral(
            "A sensory session with the same name, tester, and date already exists.");
        const QString suggested = suggestRename(s.sessionName);
        QString renamed;
        const auto choice = m_resolver->resolveUniqueViolation(
            desc, suggested, &renamed, parent);

        switch (choice) {
        case ConflictResolver::RenameMine:
            if (!renamed.isEmpty()) s.sessionName = renamed;
            s.id = -1;
            s.version = 0;
            return (m_db->tryWriteSensorySession(s) == WriteResult::Success) ? Saved : Failed;
        case ConflictResolver::OpenExisting:
            emit openExistingRequested(QStringLiteral("sensory_sessions"), s.sessionName);
            return UserCancelled;
        case ConflictResolver::CancelUnique:
        default:
            return UserCancelled;
        }
    }

    case WriteResult::OfflineReadOnly:
        warn(QStringLiteral("Saves are disabled while offline."));
        return Failed;

    case WriteResult::OtherError:
    default:
        warn(QStringLiteral("Save failed: ") + m_db->lastError());
        return Failed;
    }
}

// --- saveDetailedSensorySession ---------------------------------------------
SaveCoordinator::Outcome SaveCoordinator::saveDetailedSensorySession(
    DetailedSensorySession& s, QWidget* parent) {
    if (!m_db) return Failed;

    auto warn = [&](const QString& msg) {
        QMessageBox::warning(parent, QStringLiteral("Database Save Failed"), msg);
    };

    const WriteResult wr = m_db->tryWriteDetailedSensorySession(s);
    switch (wr) {
    case WriteResult::Success:
        return Saved;

    case WriteResult::VersionMismatch: {
        if (!m_resolver) { warn(QStringLiteral("Version mismatch (no resolver wired)")); return Failed; }
        const DetailedSensorySession fresh = (s.id > 0)
            ? m_db->loadDetailedSensorySession(s.id) : DetailedSensorySession{};
        const QVariantMap mine   = diffFieldsForDetailed(s);
        const QVariantMap theirs = diffFieldsForDetailed(fresh);
        QVariantMap merged;
        const auto choice = m_resolver->resolveVersionMismatch(
            QStringLiteral("Detailed Sensory Session"),
            QString(), QString(),
            mine, theirs, &merged, parent);

        switch (choice) {
        case ConflictResolver::KeepMine:
            s.version = fresh.version;
            if (s.id <= 0) s.id = fresh.id;
            return (m_db->tryWriteDetailedSensorySession(s) == WriteResult::Success)
                ? Saved : Failed;
        case ConflictResolver::TakeTheirs:
            emit localEditsDiscarded(QStringLiteral("detailed_sensory_sessions"), fresh.id);
            return UserCancelled;
        case ConflictResolver::ApplySelection:
            applyMergedToDetailed(s, merged);
            s.version = fresh.version;
            if (s.id <= 0) s.id = fresh.id;
            return (m_db->tryWriteDetailedSensorySession(s) == WriteResult::Success)
                ? Saved : Failed;
        case ConflictResolver::CancelVersion:
        default:
            return UserCancelled;
        }
    }

    case WriteResult::RowDeleted: {
        if (!m_resolver) { warn(QStringLiteral("Row deleted (no resolver wired)")); return Failed; }
        const QVariantMap myEdits = diffFieldsForDetailed(s);
        const auto choice = m_resolver->resolveDeleted(
            QStringLiteral("Detailed Sensory Session"), QString(), myEdits, parent);

        if (choice == ConflictResolver::Recreate) {
            s.id = -1;
            s.version = 0;
            const WriteResult retry = m_db->tryWriteDetailedSensorySession(s);
            if (retry == WriteResult::Success) return Saved;
            warn(QStringLiteral("Could not recreate row: ") + m_db->lastError());
            return Failed;
        }
        emit localEditsDiscarded(QStringLiteral("detailed_sensory_sessions"), s.id);
        return UserCancelled;
    }

    case WriteResult::UniqueViolation: {
        if (!m_resolver) { warn(QStringLiteral("Unique violation (no resolver wired)")); return Failed; }
        const QString desc = QStringLiteral(
            "A detailed sensory session with the same name, tester, and date already exists.");
        const QString suggested = suggestRename(s.sessionName);
        QString renamed;
        const auto choice = m_resolver->resolveUniqueViolation(
            desc, suggested, &renamed, parent);

        switch (choice) {
        case ConflictResolver::RenameMine:
            if (!renamed.isEmpty()) s.sessionName = renamed;
            s.id = -1;
            s.version = 0;
            return (m_db->tryWriteDetailedSensorySession(s) == WriteResult::Success)
                ? Saved : Failed;
        case ConflictResolver::OpenExisting:
            emit openExistingRequested(QStringLiteral("detailed_sensory_sessions"), s.sessionName);
            return UserCancelled;
        case ConflictResolver::CancelUnique:
        default:
            return UserCancelled;
        }
    }

    case WriteResult::OfflineReadOnly:
        warn(QStringLiteral("Saves are disabled while offline."));
        return Failed;

    case WriteResult::OtherError:
    default:
        warn(QStringLiteral("Save failed: ") + m_db->lastError());
        return Failed;
    }
}

} // namespace DVE
