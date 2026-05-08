#pragma once
#include <QGraphicsObject>
#include <QPixmap>
#include <QStringList>
#include <QVector>

namespace DVE {

class ResizableSlideItem : public QGraphicsObject {
    Q_OBJECT
public:
    // Eight resize-handle positions plus None (interior = move). Ordering is
    // arbitrary but stable for paint/hit-test iteration.
    enum class Handle {
        None, TopLeft, Top, TopRight, Right, BottomRight, Bottom, BottomLeft, Left
    };

    ResizableSlideItem(const QString& elementId,
                        bool aspectLocked, QGraphicsItem* parent = nullptr);

    QString elementId() const { return m_elementId; }
    QRectF  itemRectInches() const;            // converts back to inches

    // Set both position and size from a layout rect (inches). No-op for null
    // rects so callers can pass through default-constructed ContentSlideLayout
    // fields without guarding.
    void setRectInches(const QRectF& rectInches);

    void setSelectedItem(bool s);
    bool isSelectedItem() const { return m_selected; }

signals:
    void itemClicked(ResizableSlideItem*);
    void rectChanged(const QRectF& newRectInches);

protected:
    void   mousePressEvent(QGraphicsSceneMouseEvent*) override;
    void   mouseMoveEvent(QGraphicsSceneMouseEvent*) override;
    void   mouseReleaseEvent(QGraphicsSceneMouseEvent*) override;
    void   hoverMoveEvent(QGraphicsSceneHoverEvent*) override;
    void   hoverLeaveEvent(QGraphicsSceneHoverEvent*) override;
    QRectF boundingRect() const override;

    // Subclass paints content; base paints handles when selected
    virtual void paintContent(QPainter*) = 0;
    void paint(QPainter*, const QStyleOptionGraphicsItem*, QWidget*) override;

    // Hit-test: returns the handle under `pos` (item coords), or Handle::None
    // if pos is over the body or outside the item entirely.
    Handle handleAt(const QPointF& pos) const;
    QRectF handleRectFor(Handle h) const;

    static constexpr double kPxPerInch = 60.0;   // matches ImageViewDialog
    QString m_elementId;
    double  m_w = 200, m_h = 100;
    bool    m_aspectLocked;
    bool    m_selected = false;
    bool    m_resizing = false, m_moving = false;
    Handle  m_grabbedHandle = Handle::None;
    QPointF m_pressScenePos, m_pressItemPos;
    double  m_pressW = 0, m_pressH = 0;
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

class TableItem : public ResizableSlideItem {
    Q_OBJECT
public:
    TableItem(const QString& elementId, QGraphicsItem* parent = nullptr);
    void setHeaders(const QStringList&);
    void setRows(const QVector<QStringList>&);
    void setSort(const QString& column, Qt::SortOrder);
    void setFontPointSize(int pt);     // clamped to >= 1
    int  fontPointSize() const { return m_fontPt; }
signals:
    void columnHeaderClicked(const QString& column);
protected:
    void paintContent(QPainter*) override;
    void mousePressEvent(QGraphicsSceneMouseEvent*) override;
private:
    QStringList m_headers;
    QVector<QStringList> m_rows;
    QString m_sortColumn;
    Qt::SortOrder m_sortOrder = Qt::DescendingOrder;
    int     m_fontPt = 11;            // default body font size
    QRectF headerRectFor(int colIdx) const;
    double  rowHeight() const;        // computed from m_fontPt + line padding
};

// TextItem renders editable text on the slide canvas. Double-click pops a
// modal QInputDialog. By design, `textCommitted` fires ONLY on user-confirmed
// edits — programmatic `setText(...)` is silent. Consumers (e.g. the report
// preview dialog) that need to observe every state change should track
// programmatic edits in their own model and use `textCommitted` purely to
// record user intent (undo/redo, dirty-state tracking).
class TextItem : public ResizableSlideItem {
    Q_OBJECT
public:
    TextItem(const QString& elementId, QGraphicsItem* parent = nullptr);
    void setText(const QString&);
    QString text() const { return m_text; }
    void setFontPointSize(int pt);     // clamped to >= 1
    int  fontPointSize() const { return m_fontPt; }
signals:
    void textCommitted(const QString&);
protected:
    void paintContent(QPainter*) override;
    void mouseDoubleClickEvent(QGraphicsSceneMouseEvent*) override;
private:
    QString m_text;
    int     m_fontPt = 14;
};

} // namespace DVE
