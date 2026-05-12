#include "PresenceDotsDelegate.h"

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
    // Render the normal item first so we get the right text, icon, selection
    // background, alternating row color, etc. Then we paint the dots on top
    // in the remaining horizontal space.
    QStyledItemDelegate::paint(painter, option, index);

    const QStringList colors  = index.data(kColorsRole).toStringList();
    const QStringList intents = index.data(kIntentsRole).toStringList();
    if (colors.isEmpty()) return;

    // Measure the item's text to find where the dots should start. We use
    // the option's font (respects per-item font overrides) and the displayed
    // text from the model.
    const QFontMetrics fm(option.font);
    const QString text = index.data(Qt::DisplayRole).toString();
    const int textWidth = fm.horizontalAdvance(text);

    // For tree items the option.rect is the row rect; the actual text starts
    // after the icon + indentation. We approximate the icon offset from the
    // decoration size, falling back to 0 when no icon is present.
    int xOffset = option.rect.left();
    if (!option.icon.isNull()) {
        xOffset += option.decorationSize.width() + 4;
    }
    // Padding inside the option rect (consistent with default Qt item delegates).
    xOffset += 4;

    int dotX = xOffset + textWidth + kTextPadding;
    const int dotY = option.rect.center().y() - kDotDiameter / 2;

    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, true);

    for (int i = 0; i < colors.size(); ++i) {
        // Stop painting if we run out of horizontal room — better to truncate
        // than scribble past the column edge.
        if (dotX + kDotDiameter > option.rect.right()) break;

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

    // Reserve extra horizontal room for the dots so item resize-to-contents
    // (where used) doesn't clip them.
    const int dotsWidth =
        kTextPadding + colors.size() * (kDotDiameter + kDotSpacing);
    base.rwidth() += dotsWidth;
    return base;
}

} // namespace DVE
