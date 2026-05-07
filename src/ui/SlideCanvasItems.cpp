#include "SlideCanvasItems.h"
#include <QPainter>
#include <QGraphicsSceneMouseEvent>
#include <QInputDialog>

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

TableItem::TableItem(const QString& id, QGraphicsItem* p)
    : ResizableSlideItem(id, /*aspectLocked=*/false, p) {
    m_w = 600; m_h = 100;
}
void TableItem::setHeaders(const QStringList& h) { m_headers = h; update(); }
void TableItem::setRows(const QVector<QStringList>& r) { m_rows = r; update(); }
void TableItem::setSort(const QString& c, Qt::SortOrder o) {
    m_sortColumn = c; m_sortOrder = o; update();
}

QRectF TableItem::headerRectFor(int colIdx) const {
    const int n = m_headers.size();
    if (n == 0) return {};
    const double cw = m_w / n;
    return QRectF(colIdx * cw, 0, cw, 24);
}

void TableItem::paintContent(QPainter* p) {
    p->setRenderHint(QPainter::Antialiasing);
    p->fillRect(QRectF(0, 0, m_w, m_h), Qt::white);
    p->setPen(QColor(80, 80, 80));

    const int n = m_headers.size();
    if (n == 0) return;
    const double cw = m_w / n;
    const double rowH = qMax(18.0, (m_h - 24) / qMax(1, m_rows.size()));

    // Header row
    p->fillRect(QRectF(0, 0, m_w, 24), QColor(0x1F, 0x4E, 0x79));
    p->setPen(Qt::white);
    for (int c = 0; c < n; ++c) {
        QString h = m_headers[c];
        if (h == m_sortColumn) h += (m_sortOrder == Qt::AscendingOrder ? " \xe2\x96\xb2" : " \xe2\x96\xbc");
        p->drawText(QRectF(c * cw + 4, 4, cw - 8, 16), Qt::AlignVCenter, h);
    }

    // Rows
    p->setPen(QColor(40, 40, 40));
    for (int r = 0; r < m_rows.size(); ++r) {
        if (r % 2) p->fillRect(QRectF(0, 24 + r*rowH, m_w, rowH),
                                QColor(245, 248, 252));
        for (int c = 0; c < n && c < m_rows[r].size(); ++c)
            p->drawText(QRectF(c*cw + 4, 24 + r*rowH, cw - 8, rowH),
                        Qt::AlignVCenter, m_rows[r][c]);
    }

    // Borders
    p->setPen(QColor(0xCC, 0xCC, 0xCC));
    p->drawRect(QRectF(0, 0, m_w, m_h));
}

void TableItem::mousePressEvent(QGraphicsSceneMouseEvent* e) {
    if (e->pos().y() < 24) {
        // Header click - find column
        const int n = m_headers.size();
        if (n > 0) {
            // qBound guards against negative-x events that scene routing can deliver
            // because boundingRect() extends 1 px past the item's left edge.
            const int col = qBound(0, int(e->pos().x() / (m_w / n)), n - 1);
            emit columnHeaderClicked(m_headers[col]);
            e->accept(); return;
        }
    }
    ResizableSlideItem::mousePressEvent(e);
}

TextItem::TextItem(const QString& id, QGraphicsItem* p)
    : ResizableSlideItem(id, /*aspectLocked=*/false, p) {
    m_w = 300; m_h = 40;
}
void TextItem::setText(const QString& t) { m_text = t; update(); }
void TextItem::setFontPointSize(int pt) { m_fontPt = pt; update(); }

void TextItem::paintContent(QPainter* p) {
    p->setRenderHint(QPainter::TextAntialiasing);
    QFont f = p->font(); f.setPointSize(m_fontPt); p->setFont(f);
    p->setPen(QColor(0x33, 0x33, 0x33));
    p->drawText(QRectF(4, 4, m_w - 8, m_h - 8),
                Qt::AlignTop | Qt::AlignLeft | Qt::TextWordWrap, m_text);
}

void TextItem::mouseDoubleClickEvent(QGraphicsSceneMouseEvent*) {
    bool ok = false;
    const QString next = QInputDialog::getMultiLineText(
        nullptr, "Edit Text", "Text:", m_text, &ok);
    if (ok && next != m_text) {
        m_text = next;
        update();
        emit textCommitted(m_text);
    }
}

} // namespace DVE
