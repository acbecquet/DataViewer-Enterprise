#include "ImageViewDialog.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QDialogButtonBox>
#include <QPainter>
#include <QPainterPath>
#include <QGraphicsSceneMouseEvent>
#include <QGraphicsRectItem>
#include <QPushButton>
#include <QStyle>
#include <QGraphicsDropShadowEffect>
#include "../utils/AppTheme.h"

namespace DVE {

// ═══════════════════════════════════════════════════════════════════════════════
// ResizableImageItem
// ═══════════════════════════════════════════════════════════════════════════════

ResizableImageItem::ResizableImageItem(const QPixmap& pixmap,
                                       const QString& path,
                                       double w, double h,
                                       QGraphicsItem* parent)
    : QGraphicsObject(parent)
    , m_pixmap(pixmap)
    , m_path(path)
    , m_w(w)
    , m_h(h)
{
    setCursor(Qt::SizeAllCursor);
    setZValue(1);
}

QRectF ResizableImageItem::boundingRect() const
{
    return QRectF(0, 0, m_w, m_h);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

QRectF ResizableImageItem::handleRect() const
{
    return QRectF(m_w - kHandle, m_h - kHandle, kHandle, kHandle);
}

QRectF ResizableImageItem::cropHandleRect(int idx) const
{
    const double hw = kCropHandle / 2.0;
    QRectF ci(m_cropRect.x() * m_w, m_cropRect.y() * m_h,
              m_cropRect.width() * m_w, m_cropRect.height() * m_h);
    switch (idx) {
        case 0: return QRectF(ci.left()  - hw, ci.top()    - hw, kCropHandle, kCropHandle); // TL
        case 1: return QRectF(ci.right() - hw, ci.top()    - hw, kCropHandle, kCropHandle); // TR
        case 2: return QRectF(ci.right() - hw, ci.bottom() - hw, kCropHandle, kCropHandle); // BR
        case 3: return QRectF(ci.left()  - hw, ci.bottom() - hw, kCropHandle, kCropHandle); // BL
        default: return QRectF();
    }
}

int ResizableImageItem::cropHandleAt(const QPointF& p) const
{
    for (int i = 0; i < 4; ++i)
        if (cropHandleRect(i).contains(p)) return i;
    return -1;
}

// ── Public setters ────────────────────────────────────────────────────────────

void ResizableImageItem::setItemSelected(bool sel)
{
    if (m_selected == sel) return;
    m_selected = sel;
    if (!sel && m_cropMode) setCropMode(false);
    update();
}

void ResizableImageItem::setCropMode(bool on)
{
    m_cropMode   = on;
    m_cropHandle = -1;
    setCursor(on ? Qt::CrossCursor : Qt::SizeAllCursor);
    update();
}

void ResizableImageItem::setCropRect(const QRectF& r)
{
    m_effectiveCrop = r;
    m_cropRect = QRectF(0.0, 0.0, 1.0, 1.0);
    // Crop the displayed pixmap so no stretching occurs
    if (!m_pixmap.isNull() && !(r.x() == 0.0 && r.y() == 0.0 &&
                                 r.width() == 1.0 && r.height() == 1.0)) {
        QRect src(qRound(r.x() * m_pixmap.width()),
                  qRound(r.y() * m_pixmap.height()),
                  qMax(1, qRound(r.width()  * m_pixmap.width())),
                  qMax(1, qRound(r.height() * m_pixmap.height())));
        m_pixmap = m_pixmap.copy(src);
        // Item is already the right size (layout was saved post-crop)
    }
    update();
}

void ResizableImageItem::commitCrop()
{
    if (m_cropRect.x() == 0.0 && m_cropRect.y() == 0.0 &&
        m_cropRect.width() == 1.0 && m_cropRect.height() == 1.0)
        return; // nothing to commit

    // Compose with the cumulative effective crop (both relative to original image)
    m_effectiveCrop = QRectF(
        m_effectiveCrop.x() + m_cropRect.x() * m_effectiveCrop.width(),
        m_effectiveCrop.y() + m_cropRect.y() * m_effectiveCrop.height(),
        m_cropRect.width()  * m_effectiveCrop.width(),
        m_cropRect.height() * m_effectiveCrop.height());

    // Crop the displayed pixmap (m_pixmap is already the prev-crop version)
    if (!m_pixmap.isNull()) {
        QRect src(qRound(m_cropRect.x() * m_pixmap.width()),
                  qRound(m_cropRect.y() * m_pixmap.height()),
                  qMax(1, qRound(m_cropRect.width()  * m_pixmap.width())),
                  qMax(1, qRound(m_cropRect.height() * m_pixmap.height())));
        m_pixmap = m_pixmap.copy(src);
    }

    // Resize item so cropped portion displays at same pixel scale (no stretch)
    prepareGeometryChange();
    m_w = qMax(40.0, m_w * m_cropRect.width());
    m_h = qMax(30.0, m_h * m_cropRect.height());

    // Reset to full for next crop session
    m_cropRect = QRectF(0.0, 0.0, 1.0, 1.0);
    update();
}

// ── Paint ─────────────────────────────────────────────────────────────────────

void ResizableImageItem::paint(QPainter* p,
                               const QStyleOptionGraphicsItem*,
                               QWidget*)
{
    // Draw the full pixmap — m_pixmap is always pre-cropped after commitCrop/setCropRect.
    // m_cropRect is the overlay selection only; it must NOT be used as a source rect here.
    if (!m_pixmap.isNull()) {
        p->drawPixmap(QRectF(0, 0, m_w, m_h), m_pixmap,
                      QRectF(0, 0, m_pixmap.width(), m_pixmap.height()));
    } else {
        p->fillRect(QRectF(0, 0, m_w, m_h), QColor(0xE0, 0xE0, 0xE0));
    }

    // Selection / normal border
    if (m_selected && !m_cropMode) {
        p->setPen(QPen(QColor(0x00, 0x66, 0xCC), 2));
        p->setBrush(Qt::NoBrush);
        p->drawRect(QRectF(1, 1, m_w - 2, m_h - 2));
    } else if (!m_cropMode) {
        p->setPen(QPen(QColor(0x80, 0x80, 0x80), 1));
        p->setBrush(Qt::NoBrush);
        p->drawRect(QRectF(0, 0, m_w, m_h));
    }

    // Resize corner handle (only when selected and not in crop mode)
    if (m_selected && !m_cropMode) {
        p->setPen(Qt::NoPen);
        p->setBrush(QColor(0x00, 0x66, 0xCC));
        p->drawRect(handleRect());
    }

    // Crop mode overlay
    if (m_cropMode) {
        QRectF ci(m_cropRect.x() * m_w, m_cropRect.y() * m_h,
                  m_cropRect.width() * m_w, m_cropRect.height() * m_h);

        // Darken outside the crop area
        QPainterPath outside;
        outside.addRect(QRectF(0, 0, m_w, m_h));
        outside.addRect(ci);
        p->fillPath(outside, QColor(0, 0, 0, 120));

        // Crop boundary
        p->setPen(QPen(Qt::white, 1.5, Qt::DashLine));
        p->setBrush(Qt::NoBrush);
        p->drawRect(ci);

        // Crop edge handles
        for (int i = 0; i < 4; ++i) {
            QRectF hr = cropHandleRect(i);
            p->setPen(QPen(QColor(0x00, 0x44, 0xAA), 1));
            p->setBrush(Qt::white);
            p->drawRect(hr);
        }
    }
}

// ── Mouse events ─────────────────────────────────────────────────────────────

void ResizableImageItem::mousePressEvent(QGraphicsSceneMouseEvent* e)
{
    emit itemClicked(this);

    if (m_cropMode) {
        m_pressLocalPos  = e->pos();
        m_pressCropRect  = m_cropRect;
        m_cropHandle     = cropHandleAt(e->pos());
        e->accept();
        return;
    }

    m_pressScenePos = e->scenePos();
    m_pressItemPos  = pos();
    m_pressW        = m_w;
    m_pressH        = m_h;

    if (handleRect().contains(e->pos())) {
        m_resizing = true;
        m_moving   = false;
        setCursor(Qt::SizeFDiagCursor);
    } else {
        m_moving   = true;
        m_resizing = false;
        setCursor(Qt::ClosedHandCursor);
    }
    e->accept();
}

void ResizableImageItem::mouseMoveEvent(QGraphicsSceneMouseEvent* e)
{
    // ── Crop mode ────────────────────────────────────────────────────────────
    if (m_cropMode) {
        if (m_cropHandle < 0) { e->accept(); return; }
        QPointF delta = e->pos() - m_pressLocalPos;
        double  dx    = delta.x() / m_w;
        double  dy    = delta.y() / m_h;
        const double kMin = 10.0 / qMax(m_w, m_h);
        QRectF  cr    = m_pressCropRect;
        switch (m_cropHandle) {
            case 0: // TL
                cr.setLeft  (qBound(0.0,            cr.left()   + dx, cr.right()  - kMin));
                cr.setTop   (qBound(0.0,            cr.top()    + dy, cr.bottom() - kMin));
                break;
            case 1: // TR
                cr.setRight (qBound(cr.left()+kMin, cr.right()  + dx, 1.0));
                cr.setTop   (qBound(0.0,            cr.top()    + dy, cr.bottom() - kMin));
                break;
            case 2: // BR
                cr.setRight (qBound(cr.left()+kMin, cr.right()  + dx, 1.0));
                cr.setBottom(qBound(cr.top()+kMin,  cr.bottom() + dy, 1.0));
                break;
            case 3: // BL
                cr.setLeft  (qBound(0.0,            cr.left()   + dx, cr.right()  - kMin));
                cr.setBottom(qBound(cr.top()+kMin,  cr.bottom() + dy, 1.0));
                break;
        }
        m_cropRect = cr;
        update();
        e->accept();
        return;
    }

    QPointF delta = e->scenePos() - m_pressScenePos;

    // ── Resize (free, no aspect lock) ────────────────────────────────────────
    if (m_resizing) {
        double newW = qMax(60.0, m_pressW + delta.x());
        double newH = qMax(40.0, m_pressH + delta.y());
        prepareGeometryChange();
        m_w = newW;
        m_h = newH;
        update();
    }

    // ── Move (clamped to scene rect) ─────────────────────────────────────────
    else if (m_moving) {
        QPointF newPos = m_pressItemPos + delta;
        if (scene()) {
            QRectF sr = scene()->sceneRect();
            newPos.setX(qBound(sr.left(), newPos.x(), sr.right()  - m_w));
            newPos.setY(qBound(sr.top(),  newPos.y(), sr.bottom() - m_h));
        }
        setPos(newPos);
    }

    e->accept();
}

void ResizableImageItem::mouseReleaseEvent(QGraphicsSceneMouseEvent* e)
{
    m_resizing   = false;
    m_moving     = false;
    m_cropHandle = -1;
    setCursor(m_cropMode ? Qt::CrossCursor : Qt::SizeAllCursor);
    e->accept();
}

// ═══════════════════════════════════════════════════════════════════════════════
// ImageViewDialog
// ═══════════════════════════════════════════════════════════════════════════════

ImageViewDialog::ImageViewDialog(const QStringList&     paths,
                                 const QVector<QRectF>& layouts,
                                 const QVector<QRectF>& crops,
                                 const QString&         sampleName,
                                 QWidget*               parent)
    : QDialog(parent)
    , m_paths(paths)
    , m_initLayouts(layouts)
    , m_initCrops(crops)
{
    setWindowTitle(QString("Images \u2014 %1").arg(sampleName));
    setMinimumSize(560, 400);
    resize(920, 620);

    auto* shadow = new QGraphicsDropShadowEffect(this);
    shadow->setBlurRadius(16);
    shadow->setOffset(0, 4);
    shadow->setColor(QColor(0, 0, 0, 30));
    setGraphicsEffect(shadow);

    // ── Header label ──────────────────────────────────────────────────────────
    m_hintLabel = new QLabel(
        QString("<b>%1</b> &nbsp;\u2014 drag to reposition &nbsp;\u00B7&nbsp; "
                "drag <b>blue corner</b> to resize freely &nbsp;\u00B7&nbsp; "
                "use <b>Crop Image</b> then drag <b>corner handles</b> to trim")
            .arg(sampleName.toHtmlEscaped()),
        this);
    m_hintLabel->setWordWrap(true);
    m_hintLabel->setStyleSheet(
        "padding: 6px 8px; background: #E8F4FD; border-bottom: 1px solid #BCBCBC;");

    // ── Canvas ────────────────────────────────────────────────────────────────
    m_scene = new QGraphicsScene(0, 0, kSceneW, kSceneH, this);

    m_view = new QGraphicsView(m_scene, this);
    m_view->setRenderHint(QPainter::Antialiasing);
    m_view->setDragMode(QGraphicsView::NoDrag);
    m_view->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_view->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_view->setAlignment(Qt::AlignCenter);
    m_view->setBackgroundBrush(QColor(0xA8, 0xA8, 0xA8));  // gray outside slide
    m_view->setMinimumSize(480, 320);

    auto makeBtn = [this](const QString& text) {
        auto* b = new QPushButton(text, this);
        b->setFixedHeight(28);
        return b;
    };

    m_removeBtn = makeBtn("\u2715  Remove Selected");
    m_removeBtn->setProperty("destructive", true);
    m_removeBtn->style()->unpolish(m_removeBtn);
    m_removeBtn->style()->polish(m_removeBtn);

    m_cropBtn  = makeBtn("Crop Image");
    m_resetBtn = makeBtn("\u21BA  Reset All");

    m_removeBtn->setEnabled(false);
    m_cropBtn->setEnabled(false);

    // ── Dialog buttons ────────────────────────────────────────────────────────
    auto* bbox = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, this);
    connect(bbox, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(bbox, &QDialogButtonBox::rejected, this, &QDialog::reject);
    if (auto* okBtn = bbox->button(QDialogButtonBox::Ok)) {
        okBtn->setText("Save Layout");
        okBtn->setProperty("primary", true);
        okBtn->style()->unpolish(okBtn);
        okBtn->style()->polish(okBtn);
    }

    // ── Bottom bar ────────────────────────────────────────────────────────────
    auto* bottomBar = new QHBoxLayout();
    bottomBar->setSpacing(6);
    bottomBar->addWidget(m_removeBtn);
    bottomBar->addWidget(m_cropBtn);
    bottomBar->addWidget(m_resetBtn);
    bottomBar->addStretch(1);
    bottomBar->addWidget(bbox);

    // ── Main layout ───────────────────────────────────────────────────────────
    auto* main = new QVBoxLayout(this);
    main->setContentsMargins(0, 0, 0, 16);
    main->setSpacing(0);
    main->addWidget(m_hintLabel);
    {
        auto* inner = new QWidget(this);
        auto* il    = new QVBoxLayout(inner);
        il->setContentsMargins(16, 8, 16, 0);
        il->addWidget(m_view, 1);
        il->addSpacing(6);
        il->addLayout(bottomBar);
        main->addWidget(inner, 1);
    }

    connect(m_removeBtn, &QPushButton::clicked, this, &ImageViewDialog::onRemove);
    connect(m_cropBtn,   &QPushButton::clicked, this, &ImageViewDialog::onCropToggle);
    connect(m_resetBtn,  &QPushButton::clicked, this, &ImageViewDialog::onReset);

    buildScene();
}

// ─────────────────────────────────────────────────────────────────────────────

QStringList ImageViewDialog::imagePaths() const
{
    QStringList result;
    for (auto* item : m_items)
        result.append(item->path());
    return result;
}

QVector<QRectF> ImageViewDialog::imageLayouts() const
{
    QVector<QRectF> result;
    for (auto* item : m_items)
        result.append(QRectF(item->pos().x() / kPxPerInch,
                             item->pos().y() / kPxPerInch,
                             item->itemW()   / kPxPerInch,
                             item->itemH()   / kPxPerInch));
    return result;
}

QVector<QRectF> ImageViewDialog::imageCrops() const
{
    QVector<QRectF> result;
    for (auto* item : m_items)
        result.append(item->cropRect());
    return result;
}

// ─────────────────────────────────────────────────────────────────────────────

void ImageViewDialog::buildScene()
{
    // Exit crop mode cleanly
    if (m_inCropMode && m_selected)
        m_selected->setCropMode(false);
    m_inCropMode = false;
    if (m_cropBtn) {
        m_cropBtn->setText("Crop Image");
        m_cropBtn->setEnabled(false);
    }

    m_scene->clear();
    m_items.clear();
    m_selected = nullptr;
    if (m_removeBtn) m_removeBtn->setEnabled(false);

    // Slide boundary — white rect marks the usable area
    auto* slideBg = new QGraphicsRectItem(0, 0, kSceneW, kSceneH);
    slideBg->setPen(QPen(QColor(0x50, 0x50, 0x50), 1));
    slideBg->setBrush(Qt::white);
    slideBg->setZValue(0);
    slideBg->setFlag(QGraphicsItem::ItemIsSelectable, false);
    m_scene->addItem(slideBg);

    const int total = m_paths.size();
    for (int i = 0; i < total; ++i) {
        QPixmap pm(m_paths[i]);
        if (pm.isNull()) {
            pm = QPixmap(static_cast<int>(kSceneW / 4),
                         static_cast<int>(kSceneH / 3));
            pm.fill(QColor(0xE0, 0xE0, 0xE0));
        }

        // Position / size — skip zero-size saved layouts (images added before dialog was opened)
        QRectF r = defaultItemRect(i);
        if (i < m_initLayouts.size()) {
            const QRectF& lay = m_initLayouts[i];
            if (lay.width() > 0.0 && lay.height() > 0.0)
                r = QRectF(lay.x()     * kPxPerInch, lay.y()      * kPxPerInch,
                           lay.width() * kPxPerInch, lay.height() * kPxPerInch);
        }

        auto* item = new ResizableImageItem(pm, m_paths[i], r.width(), r.height());
        item->setPos(r.topLeft());

        // Restore crop
        if (i < m_initCrops.size())
            item->setCropRect(m_initCrops[i]);

        connect(item, &ResizableImageItem::itemClicked,
                this, &ImageViewDialog::onItemClicked);
        m_scene->addItem(item);
        m_items.append(item);
    }
}

QRectF ImageViewDialog::defaultItemRect(int index) const
{
    // Matches the PptxWriter auto-grid:
    //   cellW=4.2", cellH=3.0", startX=0.3", startY=0.7", gapX=0.15", gapY=0.1"
    const double cellW  = 4.2  * kPxPerInch;
    const double cellH  = 3.0  * kPxPerInch;
    const double startX = 0.3  * kPxPerInch;
    const double startY = 0.7  * kPxPerInch;
    const double gapX   = 0.15 * kPxPerInch;
    const double gapY   = 0.1  * kPxPerInch;
    int col = index % 2;
    int row = index / 2;
    return QRectF(startX + col * (cellW + gapX),
                  startY + row * (cellH + gapY),
                  cellW, cellH);
}

// ── Slots ─────────────────────────────────────────────────────────────────────

void ImageViewDialog::onItemClicked(ResizableImageItem* item)
{
    // Deselect previous (and exit crop mode if needed)
    if (m_selected && m_selected != item) {
        m_selected->setItemSelected(false);
        if (m_inCropMode) {
            m_selected->setCropMode(false);
            m_inCropMode = false;
            m_cropBtn->setText("Crop Image");
        }
    }
    m_selected = item;
    item->setItemSelected(true);
    m_removeBtn->setEnabled(!m_inCropMode);
    m_cropBtn->setEnabled(true);
}

void ImageViewDialog::onRemove()
{
    if (!m_selected) return;
    m_scene->removeItem(m_selected);
    m_items.removeOne(m_selected);
    delete m_selected;
    m_selected   = nullptr;
    m_inCropMode = false;
    m_cropBtn->setText("Crop Image");
    m_removeBtn->setEnabled(false);
    m_cropBtn->setEnabled(false);
}

void ImageViewDialog::onCropToggle()
{
    if (!m_selected) return;
    m_inCropMode = !m_inCropMode;
    m_selected->setCropMode(m_inCropMode);
    m_cropBtn->setText(m_inCropMode ? "Done Cropping" : "Crop Image");
    m_removeBtn->setEnabled(!m_inCropMode);
    if (!m_inCropMode)
        m_selected->commitCrop();  // resize item + apply crop, no stretch
}

void ImageViewDialog::onReset()
{
    // Forget any saved customisations and rebuild with defaults
    m_initLayouts.clear();
    m_initCrops.clear();
    buildScene();
}

} // namespace DVE
