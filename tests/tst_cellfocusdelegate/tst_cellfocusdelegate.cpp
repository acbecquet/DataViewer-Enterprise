#include <QApplication>
#include <QtTest>
#include <QStandardItemModel>
#include <QStyleOptionViewItem>
#include <QPixmap>
#include <QPainter>

#include "widgets/CellFocusDelegate.h"

using namespace DVE;

class TstCellFocusDelegate : public QObject
{
    Q_OBJECT
private slots:
    void noFocusRolesRendersDefault();
    void focusRolesPaintBorderAndFlag();
};

void TstCellFocusDelegate::noFocusRolesRendersDefault()
{
    QStandardItemModel m;
    m.appendRow(new QStandardItem("0.522"));

    CellFocusDelegate d;
    QStyleOptionViewItem opt;
    opt.rect = QRect(0, 0, 80, 24);

    QPixmap pm(80, 24); pm.fill(Qt::white);
    QPainter p(&pm);
    d.paint(&p, opt, m.index(0, 0));
    p.end();

    // No green / orange pixels in a no-roles render.
    const QImage img = pm.toImage();
    bool sawColor = false;
    for (int y = 0; y < img.height() && !sawColor; ++y)
        for (int x = 0; x < img.width(); ++x) {
            const QRgb px = img.pixel(x, y);
            if ((qRed(px) < 50 && qGreen(px) > 200) ||
                (qRed(px) > 200 && qGreen(px) > 100 && qBlue(px) < 50)) {
                sawColor = true; break;
            }
        }
    QVERIFY(!sawColor);
}

void TstCellFocusDelegate::focusRolesPaintBorderAndFlag()
{
    QStandardItemModel m;
    auto* item = new QStandardItem("0.522");
    item->setData(QString("#16a34a"), CellFocusDelegate::kFocusColorRole);
    item->setData(QString("Tina"),    CellFocusDelegate::kFocusNameRole);
    m.appendRow(item);

    CellFocusDelegate d;
    QStyleOptionViewItem opt;
    opt.rect = QRect(0, 16, 80, 24);   // leave room above for the flag

    QPixmap pm(80, 50); pm.fill(Qt::white);
    QPainter p(&pm);
    d.paint(&p, opt, m.index(0, 0));
    p.end();

    const QImage img = pm.toImage();

    // Border check: at least one solid-green pixel along each edge of opt.rect.
    auto greenAt = [&](int x, int y) {
        const QRgb px = img.pixel(x, y);
        return qRed(px) < 60 && qGreen(px) > 140 && qBlue(px) < 100;
    };
    QVERIFY(greenAt(0, 16) || greenAt(0, 17));            // left edge
    QVERIFY(greenAt(79, 16) || greenAt(79, 17));          // right edge
    QVERIFY(greenAt(0, 16) || greenAt(1, 16));            // top edge
    QVERIFY(greenAt(0, 39) || greenAt(1, 39));            // bottom edge

    // Flag check: in the region above opt.rect (y < 16) there should be
    // green-filled pixels and "Tina" text (white-ish pixels on green).
    bool sawFlagGreen = false;
    for (int y = 0; y < 16 && !sawFlagGreen; ++y)
        for (int x = 0; x < 60; ++x)
            if (greenAt(x, y)) { sawFlagGreen = true; break; }
    QVERIFY(sawFlagGreen);
}

QTEST_MAIN(TstCellFocusDelegate)
#include "tst_cellfocusdelegate.moc"
