#pragma once

// FlowLayout — cards/chips wrap left-to-right, then down. Adapted from Qt's
// "flowlayout" Widgets example (rows are centre-justified). Implements
// hasHeightForWidth()/heightForWidth()/doLayout() so it cooperates with a
// width-driven QScrollArea (vertical scroll only). Shared by SensoryPanel
// (card grid) and NotesStoryPanel (chips + editor controls); default spacing
// hSpacing 6 / vSpacing 4, pass explicit values via the constructor.

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

private:
    int doLayout(const QRect& rect, bool testOnly) const;

    QList<QLayoutItem*> m_items;
    int m_hSpace;
    int m_vSpace;
};

} // namespace DVE
