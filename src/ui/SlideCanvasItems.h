#pragma once
#include <QGraphicsObject>
#include <QPixmap>

namespace DVE {

class ResizableSlideItem : public QGraphicsObject {
    Q_OBJECT
public:
    ResizableSlideItem(const QString& elementId,
                        bool aspectLocked, QGraphicsItem* parent = nullptr);

    QString elementId() const { return m_elementId; }
    QRectF  itemRectInches() const;            // converts back to inches

    void setSelectedItem(bool s);
    bool isSelectedItem() const { return m_selected; }

signals:
    void itemClicked(ResizableSlideItem*);
    void rectChanged(const QRectF& newRectInches);

protected:
    void   mousePressEvent(QGraphicsSceneMouseEvent*) override;
    void   mouseMoveEvent(QGraphicsSceneMouseEvent*) override;
    void   mouseReleaseEvent(QGraphicsSceneMouseEvent*) override;
    QRectF boundingRect() const override;

    // Subclass paints content; base paints handles when selected
    virtual void paintContent(QPainter*) = 0;
    void paint(QPainter*, const QStyleOptionGraphicsItem*, QWidget*) override;

    static constexpr double kPxPerInch = 60.0;   // matches ImageViewDialog
    QString m_elementId;
    double  m_w = 200, m_h = 100;
    bool    m_aspectLocked;
    bool    m_selected = false;
    bool    m_resizing = false, m_moving = false;
    QPointF m_pressScenePos, m_pressItemPos;
    double  m_pressW = 0, m_pressH = 0;
    QRectF  handleRect() const;     // bottom-right resize handle
};

class PlotItem : public ResizableSlideItem {
    Q_OBJECT
public:
    PlotItem(const QString& elementId, const QPixmap& pixmap,
              QGraphicsItem* parent = nullptr);
    void setPixmap(const QPixmap&);
protected:
    void paintContent(QPainter*) override;
private:
    QPixmap m_pixmap;
};

} // namespace DVE
