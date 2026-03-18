#include "NewFileDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QPushButton>
#include <QLabel>
#include <QDialogButtonBox>
#include <QScrollArea>

namespace DVE {

NewFileDialog::NewFileDialog(const QStringList& availableTests, QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle("Create New Test File");
    setMinimumWidth(380);
    setMinimumHeight(360);

    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(8);

    mainLayout->addWidget(new QLabel("Select tests to include:"));

    m_list = new QListWidget(this);
    m_list->setSelectionMode(QAbstractItemView::NoSelection);
    for (const QString& test : availableTests) {
        QListWidgetItem* item = new QListWidgetItem(test, m_list);
        item->setFlags(item->flags() | Qt::ItemIsUserCheckable);
        item->setCheckState(Qt::Checked);
    }
    mainLayout->addWidget(m_list);

    QHBoxLayout* btnRow = new QHBoxLayout;
    QPushButton* selAll   = new QPushButton("Select All",   this);
    QPushButton* deselAll = new QPushButton("Deselect All", this);
    connect(selAll,   &QPushButton::clicked, this, &NewFileDialog::onSelectAll);
    connect(deselAll, &QPushButton::clicked, this, &NewFileDialog::onDeselectAll);
    btnRow->addWidget(selAll);
    btnRow->addWidget(deselAll);
    btnRow->addStretch();
    mainLayout->addLayout(btnRow);

    QDialogButtonBox* box = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    connect(box, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(box, &QDialogButtonBox::rejected, this, &QDialog::reject);
    mainLayout->addWidget(box);
}

QStringList NewFileDialog::selectedTests() const
{
    QStringList result;
    for (int i = 0; i < m_list->count(); ++i) {
        QListWidgetItem* item = m_list->item(i);
        if (item->checkState() == Qt::Checked)
            result.append(item->text());
    }
    return result;
}

void NewFileDialog::onSelectAll()
{
    for (int i = 0; i < m_list->count(); ++i)
        m_list->item(i)->setCheckState(Qt::Checked);
}

void NewFileDialog::onDeselectAll()
{
    for (int i = 0; i < m_list->count(); ++i)
        m_list->item(i)->setCheckState(Qt::Unchecked);
}

} // namespace DVE
