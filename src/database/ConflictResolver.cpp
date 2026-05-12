#include "ConflictResolver.h"
#include "VersionMismatchDialog.h"
#include "RowDeletedDialog.h"
#include "UniqueViolationDialog.h"

namespace DVE {

ConflictResolver::ConflictResolver(QObject* parent) : QObject(parent) {}

ConflictResolver::VersionChoice ConflictResolver::resolveVersionMismatch(
    const QString& tableLabel, const QString& whoChanged, const QString& howLongAgo,
    const QVariantMap& mine, const QVariantMap& theirs,
    QVariantMap* merged, QWidget* parent) {

    VersionMismatchDialog dlg(tableLabel, whoChanged, howLongAgo, mine, theirs, parent);
    const int rc = dlg.exec();
    if (rc != QDialog::Accepted) return CancelVersion;

    switch (dlg.bulkChoice()) {
        case VersionMismatchDialog::Mine:
            return KeepMine;
        case VersionMismatchDialog::Theirs:
            return TakeTheirs;
        case VersionMismatchDialog::Selection:
            if (merged) *merged = dlg.selectedValues();
            return ApplySelection;
    }
    return CancelVersion;
}

ConflictResolver::DeletedChoice ConflictResolver::resolveDeleted(
    const QString& tableLabel, const QString& whoDeleted,
    const QVariantMap& myUnsavedEdits, QWidget* parent) {

    RowDeletedDialog dlg(tableLabel, whoDeleted, myUnsavedEdits, parent);
    return (dlg.exec() == QDialog::Accepted && dlg.choice() == RowDeletedDialog::Recreate)
        ? Recreate : Discard;
}

ConflictResolver::UniqueChoice ConflictResolver::resolveUniqueViolation(
    const QString& description, const QString& suggestedRename,
    QString* renamedName, QWidget* parent) {

    UniqueViolationDialog dlg(description, suggestedRename, parent);
    const int rc = dlg.exec();
    if (rc != QDialog::Accepted) return CancelUnique;
    if (dlg.choice() == UniqueViolationDialog::OpenExisting) return OpenExisting;
    if (renamedName) *renamedName = dlg.renamedName();
    return RenameMine;
}

} // namespace DVE
