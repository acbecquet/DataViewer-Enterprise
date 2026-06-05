#pragma once

#include <QDialog>
#include <QVector>

#include "../utils/RecoveryManager.h"

class QListWidget;

namespace DVE {

// ─── RecoverDialog ────────────────────────────────────────────────────────────
// Plan C C9: selective reload of work recovered from a previous session.
//
// The C8 startup prompt is all-or-nothing; this dialog (driven from
// Tools -> Recover) lets the user tick exactly which previous-session items to
// reload. It is pure selection UI — it owns no recovery logic. The caller passes
// the recoverable set (RecoveryManager::recoverableItems()); after exec() ==
// QDialog::Accepted the caller reads selected() and hands the subset to
// MainWindow::restoreItems().
//
// One checkable QListWidget row per entry, all checked by default. Labels read
// like "[TPM] Foo.xlsx" / "[Sensory] <name>" / "[Detailed] <name>", with a
// " — unsaved" suffix when the entry was dirty.
// ─────────────────────────────────────────────────────────────────────────────
class RecoverDialog : public QDialog
{
    Q_OBJECT

public:
    explicit RecoverDialog(const QVector<RecoveryEntry>& items,
                           QWidget* parent = nullptr);

    // Call after exec() == QDialog::Accepted. Returns the entries whose row is
    // checked, in their original order.
    QVector<RecoveryEntry> selected() const;

private:
    // The original entries, indexed 1:1 with the list rows (row i ↔ m_items[i]).
    QVector<RecoveryEntry> m_items;
    QListWidget*           m_list = nullptr;
};

} // namespace DVE
