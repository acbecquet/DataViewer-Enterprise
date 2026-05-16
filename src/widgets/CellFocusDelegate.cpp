#include "CellFocusDelegate.h"

#include <QPainter>
#include <QPainterPath>
#include <QFontMetrics>
#include <QColor>

namespace DVE {

namespace {
constexpr int kBorderThickness = 2;
constexpr int kFlagHeight      = 14;
constexpr int kFlagPaddingX    = 4;
constexpr int kFlagFontPt      = 8;
constexpr int kFlagMaxChars    = 14;
} // namespace

CellFocusDelegate::CellFocusDelegate(QObject* parent)
    : QStyledItemDelegate(parent) {}

void CellFocusDelegate::paint(QPainter* painter,
                              const QStyleOptionViewItem& option,
                              const QModelIndex& index) const
{
    // Default render first.
    QStyledItemDelegate::paint(painter, option, index);

    // Flash overlay -- yellow wash on top of the painted cell.
    if (index.data(kFlashRole).toBool()) {
        painter->save();
        painter->fillRect(option.rect, QColor(255, 240, 130, 160));
        painter->restore();
    }

    const QString colorHex = index.data(kFocusColorRole).toString();
    if (colorHex.isEmpty()) return;
    const QColor color(colorHex);
    if (!color.isValid()) return;

    QString name = index.data(kFocusNameRole).toString();
    if (name.size() > kFlagMaxChars) name = name.left(kFlagMaxChars - 1) + QChar(0x2026);

    painter->save();
    painter->setRenderHint(QPainter::Antialiasing, true);

    // Border inside the cell.
    QPen pen(color);
    pen.setWidth(kBorderThickness);
    pen.setJoinStyle(Qt::MiterJoin);
    painter->setPen(pen);
    painter->setBrush(Qt::NoBrush);
    QRect inner = option.rect.adjusted(kBorderThickness / 2,
                                       kBorderThickness / 2,
                                       -kBorderThickness / 2,
                                       -kBorderThickness / 2);
    painter->drawRect(inner);

    // Name flag, rounded top, positioned above the cell.
    if (!name.isEmpty() && option.rect.top() >= kFlagHeight) {
        QFont f = painter->font();
        f.setPointSize(kFlagFontPt);
        f.setBold(true);
        painter->setFont(f);
        const QFontMetrics fm(f);
        const int textW = fm.horizontalAdvance(name);
        const int flagW = textW + 2 * kFlagPaddingX;
        const QRect flagRect(option.rect.left(),
                             option.rect.top() - kFlagHeight,
                             qMin(flagW, option.rect.width()),
                             kFlagHeight);
        QPainterPath path;
        path.addRoundedRect(flagRect, 3, 3);
        path.addRect(flagRect.left(), flagRect.bottom() - 2,
                     flagRect.width(), 3);
        painter->setPen(Qt::NoPen);
        painter->setBrush(color);
        painter->drawPath(path);

        painter->setPen(Qt::white);
        painter->drawText(flagRect.adjusted(kFlagPaddingX, 0, -kFlagPaddingX, 0),
                          Qt::AlignVCenter | Qt::AlignLeft, name);
    }

    painter->restore();
}

} // namespace DVE
