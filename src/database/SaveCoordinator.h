#pragma once

#include <QObject>
#include <QPointer>
#include <QString>

#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"

class QWidget;

namespace DVE {

class DatabaseManager;
class ConflictResolver;

// ── SaveCoordinator ──────────────────────────────────────────────────────────
// Routes every save through ConflictResolver. Each save method calls the
// matching `tryWriteXxx` on DatabaseManager and dispatches on the returned
// WriteResult. Conflicts (VersionMismatch / RowDeleted / UniqueViolation)
// pop the appropriate dialog; the user's choice is then applied to the
// in-place struct and the save is retried where the choice implies a retry.
//
// Phase 7 Task 21 hoists this out of MainWindow so 13 save call sites
// (7 in MainWindow, 3 each in SensoryPanel / DetailedSensoryPanel) share
// one implementation. Before 7b, every save used the bool shim that
// swallowed conflicts as `false` — the safety-critical regression this
// class fixes for the 10-user dropday goal.
class SaveCoordinator : public QObject {
    Q_OBJECT

public:
    enum Outcome { Saved, UserCancelled, Failed };

    SaveCoordinator(DatabaseManager* db,
                    ConflictResolver* resolver,
                    QObject* parent = nullptr);

    // The struct is mutated in-place on conflict resolution:
    //  - KeepMine bumps `version` to the server's fresh version then retries.
    //  - ApplySelection overlays the user-selected scalar fields, bumps
    //    version, then retries.
    //  - RowDeleted+Recreate clears id (-1) / version (0) so the retry
    //    becomes an INSERT.
    //  - UniqueViolation+RenameMine renames the natural-key field,
    //    clears id/version, then retries the INSERT.
    Outcome saveFile(FileResult& result, QWidget* parent = nullptr);
    Outcome saveSensorySession(SensorySession& s, QWidget* parent = nullptr);
    Outcome saveDetailedSensorySession(DetailedSensorySession& s,
                                       QWidget* parent = nullptr);

signals:
    // Emitted when a save discards the user's local edits (TakeTheirs on
    // VersionMismatch, or Discard on RowDeleted). MainWindow is expected
    // to reload the affected resource so the UI matches what's now in
    // the database.
    void localEditsDiscarded(const QString& table, qint64 id);

    // Emitted when the user picks OpenExisting on a UniqueViolation
    // dialog. The full natural key is in `naturalKey` so MainWindow can
    // find and switch to the existing row. Full reload is deferred to
    // a follow-up commit — for now the save returns UserCancelled.
    void openExistingRequested(const QString& table,
                               const QString& naturalKey);

private:
    QPointer<DatabaseManager>  m_db;
    QPointer<ConflictResolver> m_resolver;
};

} // namespace DVE
