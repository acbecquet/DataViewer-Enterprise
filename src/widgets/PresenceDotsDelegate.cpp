#include "PresenceDotsDelegate.h"

#include <algorithm>

#include <QApplication>
#include <QColor>
#include <QFontMetrics>
#include <QModelIndex>
#include <QPainter>
#include <QStringList>
#include <QStyleOptionViewItem>

namespace DVE {

// Layout constants. Sizing is generous enough to be readable but not so big
// it crowds the file/sheet labels in the navigator.
namespace {
constexpr int kDotDiameter   = 8;   // px diameter of each presence circle
constexpr int kDotSpacing    = 4;   // px between adjacent dots
constexpr int kTextPadding   = 6;   // px between item text and first dot
constexpr int kEditingRingPx = 1;   // white inner ring thickness for "editing"
} // namespace

PresenceDotsDelegate::PresenceDotsDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {}

void PresenceDotsDelegate::paint(QPainter* painter,
                                 const QStyleOptionViewItem& option,
                                 const QModelIndex& index) const
{
    const QStringList colors  = index.data(kColorsRole).toStringList();
    const QStringList intents = index.data(kIntentsRole).toStringList();

    if (colors.isEmpty()) {
        QStyledItemDelegate::paint(painter, option, index);
        return;
    }

    // Reserve a right-edge band wide enough for all dots + padding. The base
    // delegate paints into the remaining rect, so its text-elision logic kicks
    // in on narrow columns and the dots stay visible flush with the right edge.
    const int dotsTotal = colors.size() * kDotDiameter
                        + (colors.size() - 1) * kDotSpacing;
    const int dotsRegionWidth = dotsTotal + kTextPadding;

    QStyleOptionViewItem opt = option;
    opt.rect = QRect(option.rect.left(), option.rect.top(),
                     std::max(0, option.rect.width() - dotsRegionWidth),
                     option.rect.height());

    QStyledItemDelegate::paint(painter, opt, index);

    int dotX = option.rect.right() - dotsTotal;
    const int dotY = option.rect.center().y() - kDotDiameter / 2;

    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, true);

    for (int i = 0; i < colors.size(); ++i) {
        QColor fill(colors[i]);
        if (!fill.isValid()) fill = QColor("#888888");

        painter->setBrush(fill);
        painter->setPen(Qt::NoPen);
        painter->drawEllipse(dotX, dotY, kDotDiameter, kDotDiameter);

        // Editing intent: small white inner ring so the dot reads as "active".
        const QString intent = i < intents.size() ? intents[i] : QString();
        if (intent.compare("editing", Qt::CaseInsensitive) == 0) {
            painter->setBrush(Qt::NoBrush);
            QPen ringPen(Qt::white);
            ringPen.setWidth(kEditingRingPx);
            painter->setPen(ringPen);
            const int inset = kEditingRingPx + 1;
            painter->drawEllipse(dotX + inset, dotY + inset,
                                 kDotDiameter - 2 * inset,
                                 kDotDiameter - 2 * inset);
        }

        dotX += kDotDiameter + kDotSpacing;
    }

    painter->restore();
}

QSize PresenceDotsDelegate::sizeHint(const QStyleOptionViewItem& option,
                                     const QModelIndex& index) const
{
    QSize base = QStyledItemDelegate::sizeHint(option, index);
    const QStringList colors = index.data(kColorsRole).toStringList();
    if (colors.isEmpty()) return base;

    const int dotsTotal = colors.size() * kDotDiameter
                        + (colors.size() - 1) * kDotSpacing;
    base.rwidth() += kTextPadding + dotsTotal;
    return base;
}

} // namespace DVE
