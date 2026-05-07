#include "SlideCanvasItems.h"
#include <QPainter>
#include <QGraphicsSceneMouseEvent>

namespace DVE {

ResizableSlideItem::ResizableSlideItem(const QString& id, bool aspect, QGraphicsItem* p)
    : QGraphicsObject(p), m_elementId(id), m_aspectLocked(aspect) {
    setFlags(ItemIsSelectable | ItemIsFocusable);
    setAcceptHoverEvents(true);
}

QRectF ResizableSlideItem::boundingRect() const {
    return QRectF(0, 0, m_w, m_h).adjusted(-1, -1, 7, 7);
}

QRectF ResizableSlideItem::handleRect() const {
    constexpr double k = 12;
    return QRectF(m_w - k/2, m_h - k/2, k, k);
}

QRectF ResizableSlideItem::itemRectInches() const {
    return QRectF(x() / kPxPerInch, y() / kPxPerInch,
                   m_w / kPxPerInch, m_h / kPxPerInch);
}

void ResizableSlideItem::setSelectedItem(bool s) {
    if (m_selected != s) { m_selected = s; update(); }
}

void ResizableSlideItem::mousePressEvent(QGraphicsSceneMouseEvent* e) {
    m_pressScenePos = e->scenePos();
    m_pressItemPos  = pos();
    m_pressW = m_w; m_pressH = m_h;
    m_resizing = handleRect().contains(e->pos());
    m_moving   = !m_resizing;
    emit itemClicked(this);
    e->accept();
}

void ResizableSlideItem::mouseMoveEvent(QGraphicsSceneMouseEvent* e) {
    const QPointF d = e->scenePos() - m_pressScenePos;
    if (m_moving) {
        setPos(m_pressItemPos + d);
    } else if (m_resizing) {
        double newW = qMax(20.0, m_pressW + d.x());
        double newH = qMax(20.0, m_pressH + d.y());
        if (m_aspectLocked) {
            const double aspect = m_pressW / qMax(1.0, m_pressH);
            if (newW / newH > aspect) newW = newH * aspect;
            else                       newH = newW / aspect;
        }
        prepareGeometryChange();
        m_w = newW; m_h = newH;
    }
    update();
}

void ResizableSlideItem::mouseReleaseEvent(QGraphicsSceneMouseEvent*) {
    if (m_moving || m_resizing) emit rectChanged(itemRectInches());
    m_moving = m_resizing = false;
}

void ResizableSlideItem::paint(QPainter* p, const QStyleOptionGraphicsItem*, QWidget*) {
    paintContent(p);
    if (m_selected) {
        p->setPen(QPen(QColor(30, 130, 230), 1.5, Qt::DashLine));
        p->setBrush(Qt::NoBrush);
        p->drawRect(QRectF(0, 0, m_w, m_h));
        p->setPen(Qt::NoPen);
        p->setBrush(QColor(30, 130, 230));
        p->drawRect(handleRect());
    }
}

PlotItem::PlotItem(const QString& id, const QPixmap& pix, QGraphicsItem* p)
    : ResizableSlideItem(id, /*aspectLocked=*/true, p), m_pixmap(pix) {
    if (!pix.isNull()) {
        m_w = pix.width()  * kPxPerInch / 96.0;     // assume 96 dpi source
        m_h = pix.height() * kPxPerInch / 96.0;
    }
}

void PlotItem::setPixmap(const QPixmap& pix) {
    m_pixmap = pix; update();
}

void PlotItem::paintContent(QPainter* p) {
    if (!m_pixmap.isNull())
        p->drawPixmap(QRectF(0, 0, m_w, m_h), m_pixmap, m_pixmap.rect());
    else {
        p->fillRect(QRectF(0, 0, m_w, m_h), QColor(240, 240, 240));
        p->setPen(QColor(120, 120, 120));
        p->drawText(QRectF(0, 0, m_w, m_h), Qt::AlignCenter, "(plot)");
    }
}

} // namespace DVE
