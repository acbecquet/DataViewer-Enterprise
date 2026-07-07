#pragma once

#include <QStyledItemDelegate>

namespace DVE {

// Custom Qt item-view delegate that paints small colored circles after the
// item text. Colors come from a QStringList stored at Qt::UserRole+1 (one
// hex string per active user); intent ("viewing" | "editing") comes from a
// parallel QStringList at Qt::UserRole+2. Absent data renders normally
// without dots. Editing-intent dots get a small white inner ring so the
// user can distinguish active editors from passive viewers at a glance.
class PresenceDotsDelegate : public QStyledItemDelegate {
    Q_OBJECT
public:
    explicit PresenceDotsDelegate(QObject* parent = nullptr);

    void paint(QPainter* painter, const QStyleOptionViewItem& option,
               const QModelIndex& index) const override;

    QSize sizeHint(const QStyleOptionViewItem& option,
                   const QModelIndex& index) const override;

    // Roles used to attach presence data to view items. Callers populate
    // these on each QTreeWidgetItem / QListWidgetItem; the delegate reads
    // them at paint time. Colors and intents must be parallel lists.
    static constexpr int kColorsRole  = Qt::UserRole + 1;
    static constexpr int kIntentsRole = Qt::UserRole + 2;
};

// PresenceDotsDelegate that additionally WRAPS long item text to the view's
// current viewport width (v2.7.0 owner directive: no cut-off text at any
// window size). Needed for QTreeView, whose setWordWrap() does not actually
// produce wrapped rows on this Qt/style combination - the delegate must
// report a wrapped-height sizeHint itself. The view must still have
// wordWrap(true) so the style paints the text with Qt::TextWordWrap.
class WrappingPresenceDelegate : public PresenceDotsDelegate {
    Q_OBJECT
public:
    explicit WrappingPresenceDelegate(QObject* parent = nullptr);

    QSize sizeHint(const QStyleOptionViewItem& option,
                   const QModelIndex& index) const override;
};

} // namespace DVE
