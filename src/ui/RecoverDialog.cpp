#include "RecoverDialog.h"

#include <QListWidget>
#include <QListWidgetItem>
#include <QDialogButtonBox>
#include <QVBoxLayout>
#include <QLabel>
#include <QPushButton>

namespace DVE {

namespace {

// Short human tag for the entry's kind, as it appears in brackets in the label.
QString kindTag(RecoveryKind kind)
{
    switch (kind) {
    case RecoveryKind::Tpm:      return QStringLiteral("TPM");
    case RecoveryKind::Sensory:  return QStringLiteral("Sensory");
    case RecoveryKind::Detailed: return QStringLiteral("Detailed");
    }
    return QStringLiteral("Item");
}

} // namespace

RecoverDialog::RecoverDialog(const QVector<RecoveryEntry>& items, QWidget* parent)
    : QDialog(parent)
    , m_items(items)
{
    setWindowTitle(tr("Recover Previous Session"));
    setMinimumWidth(420);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(16, 16, 16, 16);
    mainLayout->setSpacing(10);

    auto* intro = new QLabel(
        tr("Select the items to reload from the previous session:"), this);
    intro->setWordWrap(true);
    mainLayout->addWidget(intro);

    m_list = new QListWidget(this);
    // Row i maps 1:1 to m_items[i]; default every row to checked.
    for (const RecoveryEntry& e : m_items) {
        QString label = QStringLiteral("[%1] %2")
                            .arg(kindTag(e.kind), e.displayName);
        if (e.dirty)
            label += tr(" — unsaved");

        auto* item = new QListWidgetItem(label, m_list);
        item->setFlags(item->flags() | Qt::ItemIsUserCheckable);
        item->setCheckState(Qt::Checked);
    }
    mainLayout->addWidget(m_list);

    auto* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    if (auto* ok = buttons->button(QDialogButtonBox::Ok))
        ok->setText(tr("Recover"));
    mainLayout->addWidget(buttons);

    connect(buttons, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, this, &QDialog::reject);
}

QVector<RecoveryEntry> RecoverDialog::selected() const
{
    QVector<RecoveryEntry> chosen;
    // m_list rows are 1:1 with m_items (same order, same count). Guard the count
    // defensively anyway so a row/entry mismatch can never index out of range.
    const int n = qMin(m_list->count(), m_items.size());
    for (int i = 0; i < n; ++i) {
        if (m_list->item(i)->checkState() == Qt::Checked)
            chosen.append(m_items[i]);
    }
    return chosen;
}

} // namespace DVE
