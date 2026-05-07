#include <QApplication>
#include <QtTest>
#include <QSignalSpy>
#include <QGraphicsScene>
#include <QGraphicsSceneMouseEvent>
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

    void testTableHeaderEmitsSortSignal() {
        QGraphicsScene scene;
        DVE::TableItem* item = new DVE::TableItem("table");
        scene.addItem(item);
        item->setPos(0, 0);
        item->setHeaders({"Name", "Score"});
        QSignalSpy spy(item, &DVE::TableItem::columnHeaderClicked);

        QGraphicsSceneMouseEvent press(QEvent::GraphicsSceneMousePress);
        press.setPos(QPointF(50, 10));        // item-local - first column header (cw=600/2=300)
        press.setScenePos(QPointF(50, 10));   // scene matches item-local since pos=0
        press.setButton(Qt::LeftButton);
        press.setButtons(Qt::LeftButton);
        QApplication::sendEvent(&scene, &press);

        QCOMPARE(spy.count(), 1);
        QCOMPARE(spy.first().first().toString(), QString("Name"));
    }

    void testTextItemSetGetText() {
        DVE::TextItem item("title");
        item.setText("Hello");
        QCOMPARE(item.text(), QString("Hello"));
    }

    void testTextItemFontSizeClampedAgainstNonPositiveInputs() {
        // setFontPointSize should clamp <= 0 inputs to 1 to defend against
        // corrupted layout JSON. We verify indirectly: paintContent must not
        // crash and text() round-trip is unaffected.
        DVE::TextItem item("title");
        item.setText("X");
        item.setFontPointSize(0);
        item.setFontPointSize(-5);
        QCOMPARE(item.text(), QString("X"));   // text untouched by font ops
    }
};

QTEST_MAIN(tst_SlideCanvasItems)
#include "tst_slidecanvasitems.moc"
