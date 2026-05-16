#pragma once

#include <QStyledItemDelegate>

namespace DVE {

// Paints a "someone else is editing this cell" decoration on TPM cells:
// a 2px colored border around the cell + a small name flag positioned
// above. The cell still shows its underlying value beneath.
//
// Roles consumed (set per-item on the model):
//   kFocusColorRole : hex color string of the remote user (e.g. "#16a34a").
//   kFocusNameRole  : display name string ("Tina"). Truncated to ~14 chars.
//   kFlashRole      : bool; when true the cell is rendered with a yellow
//                     tint on top of the normal text for the next paint
//                     cycle. The caller is responsible for clearing the
//                     role after the flash interval.
class CellFocusDelegate : public QStyledItemDelegate {
    Q_OBJECT
public:
    explicit CellFocusDelegate(QObject* parent = nullptr);

    void paint(QPainter* painter, const QStyleOptionViewItem& option,
               const QModelIndex& index) const override;

    static constexpr int kFocusColorRole = Qt::UserRole + 10;
    static constexpr int kFocusNameRole  = Qt::UserRole + 11;
    static constexpr int kFlashRole      = Qt::UserRole + 12;
};

} // namespace DVE
