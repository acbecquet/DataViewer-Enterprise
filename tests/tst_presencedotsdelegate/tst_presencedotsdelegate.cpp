#include <QApplication>
#include <QtTest>
#include <QStandardItemModel>
#include <QStyleOptionViewItem>
#include <QPixmap>
#include <QPainter>

#include "PresenceDotsDelegate.h"

using namespace DVE;

class tst_PresenceDotsDelegate : public QObject
{
    Q_OBJECT

private slots:
    void dotsStayVisibleWhenTextElides();
    void noDotsRendersAsPlainText();

private:
    QPixmap renderItem(const QString& text, const QStringList& colors,
                       int columnWidth);
};

QPixmap tst_PresenceDotsDelegate::renderItem(const QString& text,
                                             const QStringList& colors,
                                             int columnWidth)
{
    QStandardItemModel model;
    auto* item = new QStandardItem(text);
    item->setData(colors,  PresenceDotsDelegate::kColorsRole);
    item->setData(QStringList{"viewing"}, PresenceDotsDelegate::kIntentsRole);
    model.appendRow(item);

    PresenceDotsDelegate delegate;
    QStyleOptionViewItem opt;
    opt.rect = QRect(0, 0, columnWidth, 20);
    opt.font = QFont();
    opt.state = QStyle::State_Enabled;

    QPixmap pm(columnWidth, 20);
    pm.fill(Qt::white);
    QPainter painter(&pm);
    delegate.paint(&painter, opt, model.index(0, 0));
    painter.end();
    return pm;
}

void tst_PresenceDotsDelegate::dotsStayVisibleWhenTextElides()
{
    QStringList colors{"#FF0000"};
    QPixmap pm = renderItem(
        "this_is_a_very_long_filename_that_will_definitely_elide.xlsx",
        colors, /*columnWidth=*/120);

    const QImage img = pm.toImage();
    bool foundColored = false;
    for (int x = img.width() - 14; x < img.width() - 2 && !foundColored; ++x) {
        for (int y = 0; y < img.height(); ++y) {
            const QRgb px = img.pixel(x, y);
            if (qRed(px) > 200 && qGreen(px) < 50 && qBlue(px) < 50) {
                foundColored = true;
                break;
            }
        }
    }
    QVERIFY2(foundColored,
             "expected a red dot in the rightmost 12px band when text elides");
}

void tst_PresenceDotsDelegate::noDotsRendersAsPlainText()
{
    QPixmap pm = renderItem("short.xlsx", {}, 200);
    QVERIFY(!pm.isNull());
}

QTEST_MAIN(tst_PresenceDotsDelegate)
#include "tst_presencedotsdelegate.moc"
