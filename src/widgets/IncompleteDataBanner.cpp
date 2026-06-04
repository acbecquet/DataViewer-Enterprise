#include "IncompleteDataBanner.h"

#include <QHBoxLayout>
#include <QLabel>
#include <QToolButton>

namespace DVE {

IncompleteDataBanner::IncompleteDataBanner(QWidget* parent)
    : QFrame(parent)
{
    setObjectName(QStringLiteral("incompleteDataBanner"));
    // Amber palette, same family as OfflineBanner (warning, not destructive).
    setStyleSheet(QStringLiteral(
        "QFrame#incompleteDataBanner {"
        "  background-color: #fef3c7;"
        "  border-bottom: 1px solid #f59e0b;"
        "  padding: 6px 12px;"
        "}"
        "QLabel { color: #92400e; }"
        "QToolButton {"
        "  background: transparent;"
        "  color: #92400e;"
        "  border: none;"
        "  font-weight: 600;"
        "  padding: 2px 4px;"
        "}"
        "QToolButton:hover { color: #78350f; }"
    ));

    auto* layout = new QHBoxLayout(this);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setSpacing(8);

    m_label = new QLabel(this);
    m_label->setTextInteractionFlags(Qt::TextSelectableByMouse);
    m_label->setWordWrap(true);
    layout->addWidget(m_label, 1);

    m_closeBtn = new QToolButton(this);
    m_closeBtn->setText(QStringLiteral("X"));
    m_closeBtn->setToolTip(tr("Dismiss"));
    connect(m_closeBtn, &QToolButton::clicked, this, [this]() {
        dismiss();
        emit dismissed();
    });
    layout->addWidget(m_closeBtn);

    setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);
    hide();
}

void IncompleteDataBanner::showForFile(const QString& fileName)
{
    const QString name = fileName.isEmpty() ? tr("this file") : fileName;
    // Use tr() so the string is translatable; the warning glyph is plain unicode.
    m_label->setText(
        tr("⚠  Some data in “%1” isn’t stored in the database — the file was saved "
           "before per-row data syncing was introduced. "
           "Open the source .xlsx to restore it automatically, or ask an admin "
           "to run the database repair tool.")
            .arg(name));
    show();
}

void IncompleteDataBanner::dismiss()
{
    hide();
    m_label->setText(QString());
}

} // namespace DVE
