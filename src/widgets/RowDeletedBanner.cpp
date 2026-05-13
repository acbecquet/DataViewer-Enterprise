#include "RowDeletedBanner.h"

#include <QHBoxLayout>
#include <QLabel>
#include <QMouseEvent>
#include <QToolButton>

namespace DVE {

RowDeletedBanner::RowDeletedBanner(QWidget* parent)
    : QFrame(parent)
{
    setFrameShape(QFrame::StyledPanel);
    // Red-ish palette to flag the destructive nature of the event, but not
    // alarming — the user is offered a clear path forward (recreate).
    setStyleSheet(
        "DVE--RowDeletedBanner {"
        "  background: #fee2e2;"
        "  border: 1px solid #fca5a5;"
        "}"
        "QLabel {"
        "  color: #991b1b;"
        "  font-weight: 600;"
        "}"
        "QToolButton {"
        "  background: transparent;"
        "  color: #991b1b;"
        "  border: 1px solid transparent;"
        "  padding: 2px 8px;"
        "  font-weight: 600;"
        "}"
        "QToolButton:hover {"
        "  background: #fecaca;"
        "  border: 1px solid #991b1b;"
        "  border-radius: 3px;"
        "}"
    );

    auto* layout = new QHBoxLayout(this);
    layout->setContentsMargins(8, 6, 8, 6);
    layout->setSpacing(8);

    m_messageLabel = new QLabel(this);
    m_messageLabel->setTextInteractionFlags(Qt::TextSelectableByMouse);
    m_messageLabel->setText(QString());
    layout->addWidget(m_messageLabel, 1);

    m_recreateBtn = new QToolButton(this);
    m_recreateBtn->setText(tr("Recreate"));
    m_recreateBtn->setToolTip(tr("Save the current local state as a new row"));
    connect(m_recreateBtn, &QToolButton::clicked, this, [this]() {
        emit recreateRequested();
    });
    layout->addWidget(m_recreateBtn);

    m_closeBtn = new QToolButton(this);
    m_closeBtn->setText(QStringLiteral("X"));
    m_closeBtn->setToolTip(tr("Dismiss without recreating"));
    connect(m_closeBtn, &QToolButton::clicked, this, [this]() {
        dismiss();
        emit dismissed();
    });
    layout->addWidget(m_closeBtn);

    setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);
    hide();
}

void RowDeletedBanner::showFor(const QString& resourceLabel,
                               const QString& userName)
{
    const QString who = userName.isEmpty() ? tr("another user") : userName;
    const QString label = resourceLabel.isEmpty() ? tr("resource")
                                                   : resourceLabel;
    m_messageLabel->setText(
        tr("This %1 was deleted by %2. Click Recreate to save your local copy as a new row.")
            .arg(label, who));
    show();
}

void RowDeletedBanner::dismiss()
{
    hide();
    m_messageLabel->setText(QString());
}

} // namespace DVE
