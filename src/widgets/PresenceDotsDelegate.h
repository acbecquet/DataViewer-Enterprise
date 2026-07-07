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

// NOTE (v2.7.0): a WrappingPresenceDelegate subclass (viewport-width-aware
// wrapped size hints, for word-wrapped tree rows) was tried and REVERTED —
// size hints that depend on the live viewport width create relayout feedback
// that can freeze the UI under live-update churn. Long names are handled by
// content-sized columns + full horizontal scroll + tooltips instead. See
// tasks/lessons.md before reintroducing wrapping in item views.

} // namespace DVE
