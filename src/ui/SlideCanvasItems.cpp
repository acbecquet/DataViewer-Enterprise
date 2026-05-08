#include "SlideCanvasItems.h"
#include <QPainter>
#include <QFont>
#include <QFontMetrics>
#include <QGraphicsSceneMouseEvent>
#include <QGraphicsSceneHoverEvent>
#include <QInputDialog>
#include <array>

namespace DVE {

ResizableSlideItem::ResizableSlideItem(const QString& id, bool aspect, QGraphicsItem* p)
    : QGraphicsObject(p), m_elementId(id), m_aspectLocked(aspect) {
    setFlags(ItemIsSelectable | ItemIsFocusable);
    setAcceptHoverEvents(true);
}

QRectF ResizableSlideItem::boundingRect() const {
    // Symmetric expansion to cover all 8 handles (each 12px square centered on
    // an edge midpoint or corner; +1 buffer for the 1.5px dashed-pen overhang).
    return QRectF(0, 0, m_w, m_h).adjusted(-7, -7, 7, 7);
}

QRectF ResizableSlideItem::handleRectFor(Handle h) const {
    constexpr double k = 12;
    constexpr double half = k / 2.0;
    switch (h) {
    case Handle::TopLeft:     return QRectF(-half,         -half,         k, k);
    case Handle::Top:         return QRectF(m_w/2 - half,  -half,         k, k);
    case Handle::TopRight:    return QRectF(m_w  - half,   -half,         k, k);
    case Handle::Right:       return QRectF(m_w  - half,   m_h/2 - half,  k, k);
    case Handle::BottomRight: return QRectF(m_w  - half,   m_h   - half,  k, k);
    case Handle::Bottom:      return QRectF(m_w/2 - half,  m_h   - half,  k, k);
    case Handle::BottomLeft:  return QRectF(-half,         m_h   - half,  k, k);
    case Handle::Left:        return QRectF(-half,         m_h/2 - half,  k, k);
    case Handle::None: break;
    }
    return QRectF();
}

ResizableSlideItem::Handle ResizableSlideItem::handleAt(const QPointF& pos) const {
    // Iterate handles in standard order. Aspect-locked items skip edge midpoints
    // (only corners produce useful resizes when aspect is fixed).
    static constexpr std::array<Handle, 8> kAll = {
        Handle::TopLeft, Handle::Top, Handle::TopRight,
        Handle::Right, Handle::BottomRight,
        Handle::Bottom, Handle::BottomLeft, Handle::Left
    };
    for (Handle h : kAll) {
        if (m_aspectLocked && (h == Handle::Top || h == Handle::Right
                                || h == Handle::Bottom || h == Handle::Left))
            continue;
        if (handleRectFor(h).contains(pos)) return h;
    }
    return Handle::None;
}

QRectF ResizableSlideItem::itemRectInches() const {
    return QRectF(x() / kPxPerInch, y() / kPxPerInch,
                   m_w / kPxPerInch, m_h / kPxPerInch);
}

void ResizableSlideItem::setRectInches(const QRectF& r) {
    if (r.isNull()) return;             // null rect = no override; keep ctor defaults
    setPos(r.x() * kPxPerInch, r.y() * kPxPerInch);
    prepareGeometryChange();
    m_w = r.width()  * kPxPerInch;
    m_h = r.height() * kPxPerInch;
    update();
}

void ResizableSlideItem::setSelectedItem(bool s) {
    if (m_selected != s) { m_selected = s; update(); }
}

void ResizableSlideItem::mousePressEvent(QGraphicsSceneMouseEvent* e) {
    m_pressScenePos = e->scenePos();
    m_pressItemPos  = pos();
    m_pressW = m_w; m_pressH = m_h;
    m_grabbedHandle = handleAt(e->pos());
    m_resizing = (m_grabbedHandle != Handle::None);
    m_moving   = !m_resizing;
    emit itemClicked(this);
    e->accept();
}

void ResizableSlideItem::mouseMoveEvent(QGraphicsSceneMouseEvent* e) {
    const QPointF d = e->scenePos() - m_pressScenePos;
    if (m_moving) {
        setPos(m_pressItemPos + d);
        update();
        return;
    }
    if (!m_resizing) return;

    // Compute new geometry per handle. Each handle moves a specific edge or
    // corner; the others stay anchored.
    double newX = m_pressItemPos.x();
    double newY = m_pressItemPos.y();
    double newW = m_pressW;
    double newH = m_pressH;
    switch (m_grabbedHandle) {
    case Handle::TopLeft:
        newX = m_pressItemPos.x() + d.x(); newY = m_pressItemPos.y() + d.y();
        newW = m_pressW - d.x();           newH = m_pressH - d.y();
        break;
    case Handle::Top:
        newY = m_pressItemPos.y() + d.y(); newH = m_pressH - d.y();
        break;
    case Handle::TopRight:
        newY = m_pressItemPos.y() + d.y();
        newW = m_pressW + d.x();           newH = m_pressH - d.y();
        break;
    case Handle::Right:
        newW = m_pressW + d.x();
        break;
    case Handle::BottomRight:
        newW = m_pressW + d.x();           newH = m_pressH + d.y();
        break;
    case Handle::Bottom:
        newH = m_pressH + d.y();
        break;
    case Handle::BottomLeft:
        newX = m_pressItemPos.x() + d.x();
        newW = m_pressW - d.x();           newH = m_pressH + d.y();
        break;
    case Handle::Left:
        newX = m_pressItemPos.x() + d.x(); newW = m_pressW - d.x();
        break;
    case Handle::None: break;
    }

    // Clamp to minimum 20x20. When clamping a left/top-anchored edge, also
    // clamp the position so the right/bottom edge stays put.
    constexpr double kMin = 20.0;
    if (newW < kMin) {
        if (m_grabbedHandle == Handle::TopLeft || m_grabbedHandle == Handle::BottomLeft
            || m_grabbedHandle == Handle::Left)
            newX = m_pressItemPos.x() + (m_pressW - kMin);
        newW = kMin;
    }
    if (newH < kMin) {
        if (m_grabbedHandle == Handle::TopLeft || m_grabbedHandle == Handle::TopRight
            || m_grabbedHandle == Handle::Top)
            newY = m_pressItemPos.y() + (m_pressH - kMin);
        newH = kMin;
    }

    // Aspect lock: corners only (edge handles are filtered out in handleAt for
    // aspect-locked items, so we only get here from corners). Maintain aspect
    // and re-anchor the opposite corner.
    if (m_aspectLocked) {
        const double aspect = m_pressW / qMax(1.0, m_pressH);
        if (newW / newH > aspect) newW = newH * aspect;
        else                       newH = newW / aspect;
        if (m_grabbedHandle == Handle::TopLeft) {
            newX = m_pressItemPos.x() + (m_pressW - newW);
            newY = m_pressItemPos.y() + (m_pressH - newH);
        } else if (m_grabbedHandle == Handle::TopRight) {
            newY = m_pressItemPos.y() + (m_pressH - newH);
        } else if (m_grabbedHandle == Handle::BottomLeft) {
            newX = m_pressItemPos.x() + (m_pressW - newW);
        }
    }

    if (QPointF(newX, newY) != pos()) setPos(newX, newY);
    if (newW != m_w || newH != m_h) {
        prepareGeometryChange();
        m_w = newW; m_h = newH;
    }
    update();
}

void ResizableSlideItem::mouseReleaseEvent(QGraphicsSceneMouseEvent*) {
    if (m_moving || m_resizing) emit rectChanged(itemRectInches());
    m_moving = m_resizing = false;
    m_grabbedHandle = Handle::None;
}

void ResizableSlideItem::hoverMoveEvent(QGraphicsSceneHoverEvent* e) {
    // Cursor affordance: handle direction vs. body-move.
    Qt::CursorShape c = m_selected ? Qt::SizeAllCursor : Qt::ArrowCursor;
    if (m_selected) {
        switch (handleAt(e->pos())) {
        case Handle::TopLeft:    case Handle::BottomRight:  c = Qt::SizeFDiagCursor; break;
        case Handle::TopRight:   case Handle::BottomLeft:   c = Qt::SizeBDiagCursor; break;
        case Handle::Top:        case Handle::Bottom:       c = Qt::SizeVerCursor;   break;
        case Handle::Left:       case Handle::Right:        c = Qt::SizeHorCursor;   break;
        case Handle::None: break;
        }
    }
    setCursor(c);
}

void ResizableSlideItem::hoverLeaveEvent(QGraphicsSceneHoverEvent*) {
    unsetCursor();
}

void ResizableSlideItem::paint(QPainter* p, const QStyleOptionGraphicsItem*, QWidget*) {
    paintContent(p);
    if (m_selected) {
        // Dashed selection rectangle around the body.
        p->setPen(QPen(QColor(30, 130, 230), 1.5, Qt::DashLine));
        p->setBrush(Qt::NoBrush);
        p->drawRect(QRectF(0, 0, m_w, m_h));

        // Eight handles (corners + edge midpoints; aspect-locked items skip the
        // edge midpoints so edge-drags don't break the locked ratio).
        p->setPen(QPen(QColor(30, 130, 230), 1));
        p->setBrush(Qt::white);
        static constexpr std::array<Handle, 8> kAll = {
            Handle::TopLeft, Handle::Top, Handle::TopRight,
            Handle::Right, Handle::BottomRight,
            Handle::Bottom, Handle::BottomLeft, Handle::Left
        };
        for (Handle h : kAll) {
            if (m_aspectLocked && (h == Handle::Top || h == Handle::Right
                                     || h == Handle::Bottom || h == Handle::Left))
                continue;
            p->drawRect(handleRectFor(h));
        }
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
void TableItem::setRows(const QVector<QStringList>& r) {
    m_rows = r;
    // Auto-grow m_h so the last row never gets clipped. paintContent uses a
    // minimum 18 px row height; grow to header(24) + N*18 if the layout-set
    // height is smaller. Never shrinks.
    constexpr double headerH = 24.0;
    constexpr double minRowH = 18.0;
    const double required = headerH + r.size() * minRowH + 4.0;   // +4 padding
    if (required > m_h) {
        prepareGeometryChange();
        m_h = required;
    }
    update();
}
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
void TextItem::setText(const QString& t) {
    m_text = t;
    // Auto-grow m_h so wrapped text never gets clipped. We never SHRINK — the
    // user's layout-set size is treated as a minimum height. This is what makes
    // the propertiesBox fit "Highest Rated:" / "Lowest Rated:" lines without
    // manual resize when the box is set to the default 2" height.
    if (m_w > 8 && !m_text.isEmpty()) {
        QFont f;
        f.setFamilies({QStringLiteral("Montserrat"), QStringLiteral("Calibri")});
        f.setPointSize(m_fontPt);
        const QFontMetrics fm(f);
        const QRect bounded = fm.boundingRect(
            QRect(0, 0, int(m_w - 8), 100000),
            Qt::AlignTop | Qt::AlignLeft | Qt::TextWordWrap, m_text);
        const double required = bounded.height() + 8;
        if (required > m_h) {
            prepareGeometryChange();
            m_h = required;
        }
    }
    update();
}
void TextItem::setFontPointSize(int pt) {
    // Clamp against negative or zero values that QFont accepts but renders as
    // garbage. Defends against corrupted layout JSON.
    m_fontPt = qMax(1, pt);
    update();
}

void TextItem::paintContent(QPainter* p) {
    p->setRenderHint(QPainter::TextAntialiasing);
    // Match production PPTX font (Montserrat for slide titles); Qt falls back
    // to system default sans-serif if Montserrat isn't installed.
    QFont f = p->font();
    f.setFamilies({QStringLiteral("Montserrat"), QStringLiteral("Calibri")});
    f.setPointSize(m_fontPt);
    p->setFont(f);
    p->setPen(QColor(0x33, 0x33, 0x33));
    if (m_text.isEmpty()) {
        // Placeholder so an empty TextItem is discoverable on canvas.
        p->setPen(QColor(0xAA, 0xAA, 0xAA));
        p->drawText(QRectF(4, 4, m_w - 8, m_h - 8),
                    Qt::AlignTop | Qt::AlignLeft, QStringLiteral("(text)"));
        return;
    }
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
