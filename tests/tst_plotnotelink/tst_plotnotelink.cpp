#include <QtTest>
#include "plotting/PlotHitTest.h"
#include "plotting/PlotEngine.h"

using namespace DVE;

class TstPlotNoteLink : public QObject
{
    Q_OBJECT
private slots:
    void nearestNotePoint_picksClosestWithinTolerance()
    {
        PlotTransform tf;
        tf.plotRect = QRectF(70, 50, 700, 400);   // left,top,w,h
        tf.xMin = 0; tf.xMax = 10; tf.yMin = 0; tf.yMax = 5;
        tf.valid = true;
        QVector<NotePoint> pts = {
            {0, 2, 1.0, 1.0},   // sample 0, row 2, puff 1, tpm 1.0
            {1, 5, 6.0, 3.0},   // sample 1, row 5, puff 6, tpm 3.0
        };
        // Click near the second point'''s pixel.
        const QPointF near2 = tf.dataToPixel(6.0, 3.0);
        int outSample = -1, outRow = -1;
        QVERIFY(nearestNotePoint(near2.toPoint(), tf, pts, 18, outSample, outRow));
        QCOMPARE(outSample, 1);
        QCOMPARE(outRow, 5);
    }

    void nearestNotePoint_missesWhenFartherThanTolerance()
    {
        PlotTransform tf;
        tf.plotRect = QRectF(70, 50, 700, 400);
        tf.xMin = 0; tf.xMax = 10; tf.yMin = 0; tf.yMax = 5; tf.valid = true;
        QVector<NotePoint> pts = { {0, 2, 1.0, 1.0} };
        int s=-9, r=-9;
        QPointF far = tf.dataToPixel(1.0, 1.0) + QPointF(100, 100);
        QVERIFY(!nearestNotePoint(far.toPoint(), tf, pts, 18, s, r));
        QCOMPARE(s, -1); QCOMPARE(r, -1);
    }

    void nearestNotePoint_invalidTransformReturnsFalse()
    {
        PlotTransform tf;   // default-constructed: valid == false
        QVector<NotePoint> pts = { {0, 2, 1.0, 1.0} };
        int s=-9, r=-9;
        QVERIFY(!nearestNotePoint(QPoint(100, 100), tf, pts, 18, s, r));
        QCOMPARE(s, -1); QCOMPARE(r, -1);
    }

    void nearestNotePoint_emptyPointsReturnsFalse()
    {
        PlotTransform tf;
        tf.plotRect = QRectF(70, 50, 700, 400);
        tf.xMin = 0; tf.xMax = 10; tf.yMin = 0; tf.yMax = 5; tf.valid = true;
        QVector<NotePoint> pts;   // no note-bearing points at all
        int s=-9, r=-9;
        QVERIFY(!nearestNotePoint(QPoint(100, 100), tf, pts, 18, s, r));
        QCOMPARE(s, -1); QCOMPARE(r, -1);
    }
};

QTEST_MAIN(TstPlotNoteLink)
#include "tst_plotnotelink.moc"
