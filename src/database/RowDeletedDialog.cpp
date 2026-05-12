#include "RowDeletedDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QListWidget>
#include <QPushButton>

namespace DVE {

RowDeletedDialog::RowDeletedDialog(const QString& tableLabel,
                                    const QString& whoDeleted,
                                    const QVariantMap& myUnsavedEdits,
                                    QWidget* parent)
    : QDialog(parent) {
    setWindowTitle(tr("Row deleted — %1").arg(tableLabel));
    resize(420, 260);

    auto* root = new QVBoxLayout(this);
    root->addWidget(new QLabel(
        tr("This %1 was deleted by %2 while you were editing.")
            .arg(tableLabel, whoDeleted)));

    if (!myUnsavedEdits.isEmpty()) {
        root->addWidget(new QLabel(tr("Your unsaved changes:")));
        auto* list = new QListWidget(this);
        for (auto it = myUnsavedEdits.constBegin(); it != myUnsavedEdits.constEnd(); ++it) {
            list->addItem(QString("%1: %2").arg(it.key(), it.value().toString()));
        }
        list->setMaximumHeight(120);
        root->addWidget(list);
    }

    auto* btns = new QHBoxLayout;
    btns->addStretch();
    auto* btnDiscard  = new QPushButton(tr("Discard my changes"));
    auto* btnRecreate = new QPushButton(tr("Recreate as new"));
    btnRecreate->setDefault(true);
    btns->addWidget(btnDiscard);
    btns->addWidget(btnRecreate);
    root->addLayout(btns);

    connect(btnDiscard,  &QPushButton::clicked, this, &RowDeletedDialog::onDiscard);
    connect(btnRecreate, &QPushButton::clicked, this, &RowDeletedDialog::onRecreate);
}

} // namespace DVE
