#include "OfflineBanner.h"

#include <QDateTime>
#include <QHBoxLayout>
#include <QLabel>
#include <QPushButton>
#include <QVBoxLayout>

namespace DVE {

OfflineBanner::OfflineBanner(QWidget* parent)
    : QFrame(parent)
{
    setObjectName(QStringLiteral("offlineBanner"));
    // Amber palette signals "warning, but not destructive". Background and
    // border match the design spec; padding gives the text breathing room.
    setStyleSheet(QStringLiteral(
        "QFrame#offlineBanner {"
        "  background-color: #fef3c7;"
        "  border-bottom: 1px solid #f59e0b;"
        "  padding: 6px 12px;"
        "}"
    ));

    auto* outer = new QHBoxLayout(this);
    outer->setContentsMargins(0, 0, 0, 0);
    outer->setSpacing(8);

    auto* textColumn = new QVBoxLayout();
    textColumn->setContentsMargins(0, 0, 0, 0);
    textColumn->setSpacing(2);

    m_primaryLabel = new QLabel(this);
    m_primaryLabel->setTextInteractionFlags(Qt::TextSelectableByMouse);
    textColumn->addWidget(m_primaryLabel);

    m_secondaryLabel = new QLabel(this);
    m_secondaryLabel->setTextInteractionFlags(Qt::TextSelectableByMouse);
    m_secondaryLabel->setVisible(false);
    textColumn->addWidget(m_secondaryLabel);

    // H7: tertiary line for retry-failure callout, styled red to distinguish
    // from the "queued" secondary line. Hidden until setPendingFailureCount
    // sets a non-zero value.
    m_tertiaryLabel = new QLabel(this);
    m_tertiaryLabel->setTextInteractionFlags(Qt::TextSelectableByMouse);
    m_tertiaryLabel->setStyleSheet(QStringLiteral("color: #b91c1c;"));
    m_tertiaryLabel->setVisible(false);
    textColumn->addWidget(m_tertiaryLabel);

    outer->addLayout(textColumn, 1);
    outer->addStretch(0);

    m_retryButton = new QPushButton(tr("Retry"), this);
    m_retryButton->setCursor(Qt::PointingHandCursor);
    m_retryButton->setStyleSheet(QStringLiteral(
        "QPushButton {"
        "  background-color: #f59e0b;"
        "  color: white;"
        "  border: none;"
        "  padding: 4px 12px;"
        "  border-radius: 3px;"
        "}"
        "QPushButton:hover {"
        "  background-color: #d97706;"
        "}"
    ));
    connect(m_retryButton, &QPushButton::clicked,
            this, &OfflineBanner::retryClicked);
    outer->addWidget(m_retryButton);

    setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);

    // Default state: invalid sync timestamp, no pending edits, hidden.
    refreshPrimaryText();
    refreshSecondaryText();
    refreshTertiaryText();
    setVisible(false);
}

OfflineBanner::~OfflineBanner() = default;

void OfflineBanner::setLastSync(const QDateTime& whenLocal)
{
    m_lastSync = whenLocal;
    refreshPrimaryText();
}

void OfflineBanner::setPendingCount(int n)
{
    m_pendingCount = n;
    refreshSecondaryText();
}

void OfflineBanner::setPendingFailureCount(int n)
{
    m_pendingFailureCount = n;
    refreshTertiaryText();
}

void OfflineBanner::refreshPrimaryText()
{
    if (m_lastSync.isValid()) {
        m_primaryLabel->setText(
            tr("⚠  Working offline — read-only. Last synced %1.")
                .arg(humanizeRelativeTime(m_lastSync)));
    } else {
        m_primaryLabel->setText(
            tr("⚠  Working offline — read-only. (No previous snapshot.)"));
    }
}

void OfflineBanner::refreshSecondaryText()
{
    if (m_pendingCount <= 0) {
        m_secondaryLabel->setText(QString());
        m_secondaryLabel->setVisible(false);
    } else if (m_pendingCount == 1) {
        m_secondaryLabel->setText(
            tr("1 unsaved change. Will retry when reconnected."));
        m_secondaryLabel->setVisible(true);
    } else {
        m_secondaryLabel->setText(
            tr("%1 unsaved changes. Will retry when reconnected.")
                .arg(m_pendingCount));
        m_secondaryLabel->setVisible(true);
    }
}

void OfflineBanner::refreshTertiaryText()
{
    if (m_pendingFailureCount <= 0) {
        m_tertiaryLabel->setText(QString());
        m_tertiaryLabel->setVisible(false);
        return;
    }
    if (m_pendingFailureCount == 1) {
        m_tertiaryLabel->setText(
            tr("⚠  1 edit failed to save. Click Retry."));
    } else {
        m_tertiaryLabel->setText(
            tr("⚠  %1 edits failed to save. Click Retry.")
                .arg(m_pendingFailureCount));
    }
    m_tertiaryLabel->setVisible(true);
}

QString OfflineBanner::humanizeRelativeTime(const QDateTime& whenLocal)
{
    if (!whenLocal.isValid()) {
        return QString();
    }

    const QDateTime now = QDateTime::currentDateTime();
    const qint64 secs = whenLocal.secsTo(now);

    // Future timestamps (clock skew, etc.) collapse to "just now" rather
    // than emitting nonsensical "-3 minutes ago".
    if (secs < 60) {
        return QObject::tr("just now");
    }

    const qint64 mins = secs / 60;
    if (mins < 60) {
        return QObject::tr("%1 minutes ago").arg(mins);
    }

    const qint64 hours = mins / 60;
    if (hours < 24) {
        return QObject::tr("%1 hours ago").arg(hours);
    }

    if (hours < 48) {
        return QObject::tr("yesterday at %1").arg(whenLocal.toString("h:mm AP"));
    }

    return whenLocal.toString("MMM d, yyyy");
}

} // namespace DVE
