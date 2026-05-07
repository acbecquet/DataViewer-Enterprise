#pragma once
#include <QGraphicsObject>
#include <QPixmap>
#include <QStringList>
#include <QVector>

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

class TableItem : public ResizableSlideItem {
    Q_OBJECT
public:
    TableItem(const QString& elementId, QGraphicsItem* parent = nullptr);
    void setHeaders(const QStringList&);
    void setRows(const QVector<QStringList>&);
    void setSort(const QString& column, Qt::SortOrder);
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
    QRectF headerRectFor(int colIdx) const;
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
