#pragma once

// FlowLayout — cards/chips wrap left-to-right, then down. Adapted from Qt's
// "flowlayout" Widgets example. Implements hasHeightForWidth()/
// heightForWidth()/doLayout() so it cooperates with a width-driven QScrollArea
// (vertical scroll only). Shared by SensoryPanel (card grid, centre-justified)
// and NotesStoryPanel (chips + editor controls, left-justified). Rows are
// centre-justified by default; call setRowAlignment(Qt::AlignLeft) to
// left-justify. Default spacing hSpacing 6 / vSpacing 4, pass explicit values
// via the constructor.

#include <QLayout>
#include <QRect>
#include <QList>

namespace DVE {

class FlowLayout : public QLayout
{
public:
    explicit FlowLayout(QWidget* parent = nullptr, int margin = -1,
                        int hSpacing = 6, int vSpacing = 4);
    ~FlowLayout() override;

    void addItem(QLayoutItem* item) override;
    int  horizontalSpacing() const;
    int  verticalSpacing()  const;
    Qt::Orientations expandingDirections() const override;
    bool hasHeightForWidth() const override;
    int  heightForWidth(int) const override;
    int  count() const override;
    QLayoutItem* itemAt(int index) const override;
    QSize minimumSize() const override;
    void  setGeometry(const QRect& rect) override;
    QSize sizeHint() const override;
    QLayoutItem* takeAt(int index) override;

    // Horizontal justification of each wrapped row. Only the horizontal flag is
    // honoured: Qt::AlignLeft, Qt::AlignRight, or Qt::AlignHCenter (default).
    void setRowAlignment(Qt::Alignment a) { m_rowAlign = a; invalidate(); }

private:
    int doLayout(const QRect& rect, bool testOnly) const;

    QList<QLayoutItem*> m_items;
    int m_hSpace;
    int m_vSpace;
    Qt::Alignment m_rowAlign = Qt::AlignHCenter;
};

} // namespace DVE
