#include "VersionMismatchDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QLabel>
#include <QTableWidget>
#include <QTableWidgetItem>
#include <QHeaderView>
#include <QRadioButton>
#include <QButtonGroup>
#include <QPushButton>
#include <QSet>

namespace DVE {

VersionMismatchDialog::VersionMismatchDialog(const QString& tableLabel,
                                              const QString& whoChanged,
                                              const QString& howLongAgo,
                                              const QVariantMap& mine,
                                              const QVariantMap& theirs,
                                              QWidget* parent)
    : QDialog(parent), m_mine(mine), m_theirs(theirs) {
    setWindowTitle(tr("Conflict -- %1").arg(tableLabel));
    resize(640, 400);

    auto* root = new QVBoxLayout(this);
    root->addWidget(new QLabel(tr("\"%1\" was changed by %2 (%3).").arg(tableLabel, whoChanged, howLongAgo)));

    m_table = new QTableWidget(this);
    m_table->setColumnCount(4);
    m_table->setHorizontalHeaderLabels({tr("Field"), tr("Their value"),
                                         tr("Your value"), tr("Use which?")});
    m_table->horizontalHeader()->setStretchLastSection(true);
    m_table->verticalHeader()->setVisible(false);
    m_table->setSelectionMode(QAbstractItemView::NoSelection);
    m_table->setEditTriggers(QAbstractItemView::NoEditTriggers);

    QSet<QString> keys = QSet<QString>(mine.keyBegin(), mine.keyEnd());
    for (auto it = theirs.constBegin(); it != theirs.constEnd(); ++it)
        keys.insert(it.key());

    QStringList sortedKeys(keys.constBegin(), keys.constEnd());
    sortedKeys.sort();
    m_table->setRowCount(sortedKeys.size());

    for (int row = 0; row < sortedKeys.size(); ++row) {
        const QString& key = sortedKeys[row];
        const QVariant theirVal = theirs.value(key);
        const QVariant myVal    = mine.value(key);
        const bool same = (theirVal == myVal);

        m_table->setItem(row, 0, new QTableWidgetItem(key));
        m_table->setItem(row, 1, new QTableWidgetItem(theirVal.toString()));
        m_table->setItem(row, 2, new QTableWidgetItem(myVal.toString()));

        if (same) {
            auto* label = new QTableWidgetItem(tr("same"));
            label->setFlags(Qt::ItemIsEnabled);
            m_table->setItem(row, 3, label);
        } else {
            auto* container = new QWidget;
            auto* hl = new QHBoxLayout(container);
            hl->setContentsMargins(2, 2, 2, 2);
            auto* btnTheirs = new QRadioButton(tr("theirs"));
            auto* btnMine   = new QRadioButton(tr("yours"));
            btnMine->setChecked(true);  // default to user's value
            auto* group = new QButtonGroup(container);
            group->addButton(btnTheirs, 1);
            group->addButton(btnMine,   0);
            hl->addWidget(btnTheirs);
            hl->addWidget(btnMine);
            hl->addStretch();
            m_table->setCellWidget(row, 3, container);
            m_radios.insert(key, group);
        }
    }
    m_table->resizeColumnsToContents();
    root->addWidget(m_table);

    auto* btns = new QHBoxLayout;
    auto* btnAllMine   = new QPushButton(tr("Keep all mine"));
    auto* btnAllTheirs = new QPushButton(tr("Take all theirs"));
    auto* btnApply     = new QPushButton(tr("Apply selection"));
    btnApply->setDefault(true);
    btns->addWidget(btnAllMine);
    btns->addWidget(btnAllTheirs);
    btns->addStretch();
    btns->addWidget(btnApply);
    root->addLayout(btns);

    connect(btnAllMine,   &QPushButton::clicked, this, &VersionMismatchDialog::onKeepMine);
    connect(btnAllTheirs, &QPushButton::clicked, this, &VersionMismatchDialog::onTakeTheirs);
    connect(btnApply,     &QPushButton::clicked, this, &VersionMismatchDialog::onApplySelection);
}

void VersionMismatchDialog::onKeepMine()        { m_choice = Mine;      accept(); }
void VersionMismatchDialog::onTakeTheirs()      { m_choice = Theirs;    accept(); }
void VersionMismatchDialog::onApplySelection()  { m_choice = Selection; accept(); }

QVariantMap VersionMismatchDialog::selectedValues() const {
    QVariantMap out;
    QSet<QString> keys = QSet<QString>(m_mine.keyBegin(), m_mine.keyEnd());
    for (auto it = m_theirs.constBegin(); it != m_theirs.constEnd(); ++it)
        keys.insert(it.key());
    for (const QString& key : keys) {
        const QVariant theirVal = m_theirs.value(key);
        const QVariant myVal    = m_mine.value(key);
        if (theirVal == myVal) { out.insert(key, myVal); continue; }
        auto it = m_radios.find(key);
        if (it == m_radios.end()) { out.insert(key, myVal); continue; }
        const int picked = (*it)->checkedId();   // 0=mine, 1=theirs
        out.insert(key, (picked == 1) ? theirVal : myVal);
    }
    return out;
}

} // namespace DVE
