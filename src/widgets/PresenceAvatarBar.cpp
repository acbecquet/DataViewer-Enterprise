#include "PresenceAvatarBar.h"

#include <QColor>
#include <QEvent>
#include <QFont>
#include <QFontMetrics>
#include <QHelpEvent>
#include <QPainter>
#include <QPen>
#include <QToolTip>

namespace DVE {

// Layout constants — the bar lives above (or beside) the central editor so
// the avatars need to be small but legible. 28 px feels right next to the
// 22 px ribbon labels without dominating the title area.
namespace {
constexpr int kAvatarSize      = 28;  // diameter in px
constexpr int kAvatarSpacing   = 4;   // px between avatars
constexpr int kSelfRingPx      = 2;   // white inner ring for self-user
constexpr int kHMargin         = 6;   // px outer margin (each side)
constexpr int kVMargin         = 2;   // px outer margin (top + bottom)

QString initialsFor(const QString& name) {
    // First letter of first word; first letter of second word if present.
    const QString trimmed = name.trimmed();
    if (trimmed.isEmpty()) return QStringLiteral("?");
    const QStringList parts = trimmed.split(QLatin1Char(' '), Qt::SkipEmptyParts);
    QString out;
    out += parts.value(0, QString()).left(1).toUpper();
    if (parts.size() >= 2)
        out += parts.value(1).left(1).toUpper();
    if (out.isEmpty()) return QStringLiteral("?");
    return out;
}
} // namespace

PresenceAvatarBar::PresenceAvatarBar(QWidget* parent)
    : QWidget(parent)
{
    setAttribute(Qt::WA_TransparentForMouseEvents, false);
    setMouseTracking(true);
    setMinimumHeight(kAvatarSize + 2 * kVMargin);
    setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Fixed);
    // Start hidden — we show the bar only when at least one remote user is
    // present so no whitespace gap appears above the editor when working alone.
    setVisible(false);
}

void PresenceAvatarBar::setPresence(const QVector<PresenceRow>& rows,
                                    const QUuid& selfUuid)
{
    m_rows     = rows;
    m_selfUuid = selfUuid;
    setVisible(!m_rows.isEmpty());
    updateGeometry();
    update();
}

void PresenceAvatarBar::clear()
{
    m_rows.clear();
    setVisible(false);
    update();
}

QSize PresenceAvatarBar::sizeHint() const
{
    const int n = m_rows.size();
    const int width = 2 * kHMargin +
                      (n > 0 ? n * kAvatarSize + (n - 1) * kAvatarSpacing : 0);
    return QSize(width, kAvatarSize + 2 * kVMargin);
}

QSize PresenceAvatarBar::minimumSizeHint() const
{
    // Same as sizeHint: the bar must keep its full content width even when it
    // lives in a tight horizontal layout (the plot control row), so the avatars
    // are never clipped to zero width and hidden.
    return sizeHint();
}

QRect PresenceAvatarBar::avatarRect(int index) const
{
    if (index < 0 || index >= m_rows.size()) return {};
    // Right-aligned layout: first avatar (index 0) lives at the right edge,
    // additional avatars stack to the left. This keeps the bar from
    // shifting horizontally as users come and go.
    const int rightEdge = width() - kHMargin;
    const int x = rightEdge - (index + 1) * kAvatarSize - index * kAvatarSpacing;
    const int y = kVMargin;
    return QRect(x, y, kAvatarSize, kAvatarSize);
}

void PresenceAvatarBar::paintEvent(QPaintEvent* /*event*/)
{
    if (m_rows.isEmpty()) return;

    QPainter painter(this);
    painter.setRenderHint(QPainter::Antialiasing, true);

    QFont font = this->font();
    font.setBold(true);
    font.setPointSize(qMax(7, font.pointSize() - 1));
    painter.setFont(font);
    const QFontMetrics fm(font);

    for (int i = 0; i < m_rows.size(); ++i) {
        const QRect r = avatarRect(i);
        if (!r.isValid()) continue;

        QColor fill(m_rows[i].userColor);
        if (!fill.isValid()) fill = QColor("#888888");

        painter.setBrush(fill);
        painter.setPen(Qt::NoPen);
        painter.drawEllipse(r);

        // Self-user gets a 2 px white inner ring.
        if (!m_selfUuid.isNull() && m_rows[i].userUuid == m_selfUuid) {
            QPen ringPen(Qt::white);
            ringPen.setWidth(kSelfRingPx);
            painter.setBrush(Qt::NoBrush);
            painter.setPen(ringPen);
            const int inset = kSelfRingPx + 1;
            painter.drawEllipse(r.adjusted(inset, inset, -inset, -inset));
        }

        // Initials, white text centered.
        painter.setBrush(Qt::NoBrush);
        painter.setPen(Qt::white);
        const QString initials = initialsFor(m_rows[i].userName);
        painter.drawText(r, Qt::AlignCenter, initials);
        Q_UNUSED(fm);
    }
}

bool PresenceAvatarBar::event(QEvent* event)
{
    if (event->type() == QEvent::ToolTip) {
        auto* helpEvent = static_cast<QHelpEvent*>(event);
        for (int i = 0; i < m_rows.size(); ++i) {
            const QRect r = avatarRect(i);
            if (!r.contains(helpEvent->pos())) continue;
            const QString tip = QStringLiteral("%1 (%2)")
                                    .arg(m_rows[i].userName,
                                         m_rows[i].intent);
            QToolTip::showText(helpEvent->globalPos(), tip, this);
            return true;
        }
        QToolTip::hideText();
        return true;
    }
    return QWidget::event(event);
}

} // namespace DVE
