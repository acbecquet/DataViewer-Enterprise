#include <QtTest>
#include <QGraphicsScene>
#include "SlideCanvasItems.h"

class tst_SlideCanvasItems : public QObject {
    Q_OBJECT
private slots:
    void testPlotItemAspectLockedDuringResize() {
        QGraphicsScene scene;
        QPixmap pix(96, 96); pix.fill(Qt::red);
        DVE::PlotItem* item = new DVE::PlotItem("radar", pix);
        scene.addItem(item);

        // 96 px source @ 96 dpi = 1 inch; constructor scales to 60 px/inch canvas
        // m_w = 96 * 60 / 96 = 60 scene px; itemRectInches = 60 / 60 = 1.0 inch
        const QRectF r = item->itemRectInches();
        QCOMPARE(r.width(), r.height());      // square (aspect-locked)
        QCOMPARE(r.width(), 1.0);             // 96 px @ 96 dpi = 1 inch
    }
};

QTEST_MAIN(tst_SlideCanvasItems)
#include "tst_slidecanvasitems.moc"
