#include "UniqueViolationDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>

namespace DVE {

UniqueViolationDialog::UniqueViolationDialog(const QString& description,
                                              const QString& suggestedRename,
                                              QWidget* parent)
    : QDialog(parent) {
    setWindowTitle(tr("Already exists"));
    resize(420, 200);

    auto* root = new QVBoxLayout(this);
    root->addWidget(new QLabel(description));

    auto* renameRow = new QHBoxLayout;
    renameRow->addWidget(new QLabel(tr("Rename to:")));
    m_renameEdit = new QLineEdit(suggestedRename);
    renameRow->addWidget(m_renameEdit);
    root->addLayout(renameRow);

    auto* btns = new QHBoxLayout;
    auto* btnCancel = new QPushButton(tr("Cancel"));
    auto* btnOpen   = new QPushButton(tr("Open existing"));
    auto* btnRename = new QPushButton(tr("Rename"));
    btnRename->setDefault(true);
    btns->addWidget(btnCancel);
    btns->addStretch();
    btns->addWidget(btnOpen);
    btns->addWidget(btnRename);
    root->addLayout(btns);

    connect(btnCancel, &QPushButton::clicked, this, &UniqueViolationDialog::onCancel);
    connect(btnOpen,   &QPushButton::clicked, this, &UniqueViolationDialog::onOpenExisting);
    connect(btnRename, &QPushButton::clicked, this, &UniqueViolationDialog::onRename);
}

void UniqueViolationDialog::onRename() {
    m_choice = RenameMine;
    accept();
}

QString UniqueViolationDialog::renamedName() const {
    return m_renameEdit->text().trimmed();
}

} // namespace DVE
