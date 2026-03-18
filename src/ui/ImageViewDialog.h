#pragma once
#include <QDialog>
#include <QStringList>
#include <QVector>
#include <QRectF>
#include <QGraphicsView>
#include <QGraphicsScene>
#include <QGraphicsObject>
#include <QPushButton>
#include <QLabel>
#include <QPixmap>

namespace DVE {

// ─── ResizableImageItem ───────────────────────────────────────────────────────
// Draggable, aspect-ratio-locked resizable image item.
// In crop mode the four edge handles adjust the visible crop region.
class ResizableImageItem : public QGraphicsObject {
    Q_OBJECT
public:
    ResizableImageItem(const QPixmap& pixmap, const QString& path,
                       double w, double h, QGraphicsItem* parent = nullptr);

    QString path()     const { return m_path; }
    double  itemW()    const { return m_w; }
    double  itemH()    const { return m_h; }
    QRectF  cropRect() const { return m_cropRect; }   // normalized [0,1]

    void setItemSelected(bool sel);
    bool isItemSelected() const { return m_selected; }

    void setCropMode(bool on);
    bool inCropMode()  const { return m_cropMode; }
    void setCropRect(const QRectF& r);

    QRectF boundingRect() const override;
    void   paint(QPainter* painter,
                 const QStyleOptionGraphicsItem* option,
                 QWidget* widget) override;

signals:
    void itemClicked(ResizableImageItem* self);

protected:
    void mousePressEvent(QGraphicsSceneMouseEvent*   e) override;
    void mouseMoveEvent(QGraphicsSceneMouseEvent*    e) override;
    void mouseReleaseEvent(QGraphicsSceneMouseEvent* e) override;

private:
    QRectF handleRect()               const;  // resize corner handle
    QRectF cropHandleRect(int idx)    const;  // 0=top 1=right 2=bottom 3=left
    int    cropHandleAt(const QPointF& p) const;

    QPixmap m_pixmap;
    QString m_path;
    double  m_w, m_h;
    double  m_origAspect = 1.0;

    // Selection / resize state
    bool    m_selected  = false;
    bool    m_resizing  = false;
    bool    m_moving    = false;
    QPointF m_pressScenePos;
    QPointF m_pressItemPos;
    double  m_pressW    = 0.0;
    double  m_pressH    = 0.0;

    // Crop state
    bool    m_cropMode  = false;
    QRectF  m_cropRect  = QRectF(0.0, 0.0, 1.0, 1.0);  // normalized
    int     m_cropHandle = -1;
    QPointF m_pressLocalPos;
    QRectF  m_pressCropRect;

    static constexpr double kHandle     = 12.0;
    static constexpr double kCropHandle = 10.0;
};

// ─── ImageViewDialog ──────────────────────────────────────────────────────────
// Canvas (800×450 px = slide proportions) where images can be dragged,
// resized (aspect-ratio locked), and cropped.  Custom layout is reproduced in
// the generated report.
class ImageViewDialog : public QDialog {
    Q_OBJECT
public:
    ImageViewDialog(const QStringList&     paths,
                    const QVector<QRectF>& layouts,   // positions in inches
                    const QVector<QRectF>& crops,     // normalized [0,1] crop per image
                    const QString&         sampleName,
                    QWidget*               parent = nullptr);

    QStringList     imagePaths()   const;
    QVector<QRectF> imageLayouts() const;  // positions in inches
    QVector<QRectF> imageCrops()   const;  // normalized [0,1] crop rects

private slots:
    void onRemove();
    void onCropToggle();
    void onReset();
    void onItemClicked(ResizableImageItem* item);

private:
    void   buildScene();
    QRectF defaultItemRect(int index) const;  // scene pixels

    QGraphicsView*  m_view;
    QGraphicsScene* m_scene;
    QPushButton*    m_removeBtn;
    QPushButton*    m_cropBtn;
    QPushButton*    m_resetBtn;
    QLabel*         m_hintLabel;

    QStringList                m_paths;
    QVector<QRectF>            m_initLayouts;   // incoming, inches
    QVector<QRectF>            m_initCrops;     // incoming, normalized
    QList<ResizableImageItem*> m_items;
    ResizableImageItem*        m_selected  = nullptr;
    bool                       m_inCropMode = false;

    // Canvas = 800×450 px  ≡  13.333" × 7.5"  (60 px/inch)
    static constexpr double kSceneW    = 800.0;
    static constexpr double kSceneH    = 450.0;
    static constexpr double kPxPerInch = 60.0;
};

} // namespace DVE
