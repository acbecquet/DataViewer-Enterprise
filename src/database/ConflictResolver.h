#pragma once

#include <QObject>
#include <QString>
#include <QVariantMap>

class QWidget;

namespace DVE {

class ConflictResolver : public QObject {
    Q_OBJECT
public:
    enum VersionChoice { KeepMine, TakeTheirs, ApplySelection, CancelVersion };
    enum DeletedChoice { Recreate, Discard };
    enum UniqueChoice  { OpenExisting, RenameMine, CancelUnique };

    explicit ConflictResolver(QObject* parent = nullptr);

    // Pop a version-mismatch dialog showing per-field comparison. The user
    // picks per-field choices and clicks Apply selection (or KeepMine/TakeTheirs
    // bulk). On Apply selection, the caller's `merged` map is populated with
    // the field-by-field choices.
    VersionChoice resolveVersionMismatch(const QString& tableLabel,
                                          const QString& whoChanged,
                                          const QString& howLongAgo,
                                          const QVariantMap& mine,
                                          const QVariantMap& theirs,
                                          QVariantMap* merged,
                                          QWidget* parent = nullptr);

    // Pop a row-deleted dialog. Shows what the user has unsaved.
    DeletedChoice resolveDeleted(const QString& tableLabel,
                                  const QString& whoDeleted,
                                  const QVariantMap& myUnsavedEdits,
                                  QWidget* parent = nullptr);

    // Pop a unique-violation dialog. Suggests a renamed name (caller can
    // accept via RenameMine which sets *renamedName).
    UniqueChoice resolveUniqueViolation(const QString& description,
                                         const QString& suggestedRename,
                                         QString* renamedName,
                                         QWidget* parent = nullptr);
};

} // namespace DVE
