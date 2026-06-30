#include <QtTest/QtTest>
#include <QApplication>
#include <QToolButton>
#include <QFont>
#include <QFontMetrics>
#include <QStringList>
#include <QIcon>
#include "RibbonWidget.h"
#include "ScrollHost.h"

// Count how many visual lines QToolButton's word-wrap would produce for a
// given (already newline-split) label inside a text rectangle of `textW`
// pixels wide, at font `f`. Mirrors QToolButton's Qt::TextWordWrap behavior
// closely enough to detect a clipped 3rd row.
static int wrappedLineCount(const QString& label, int textW, const QFont& f)
{
    const QFontMetrics fm(f);
    int lines = 0;
    for (const QString& hardLine : label.split('\n')) {
        const QRect br = fm.boundingRect(QRect(0, 0, textW, 100000),
                                         Qt::TextWordWrap, hardLine);
        const int h = qMax(1, fm.lineSpacing());
        lines += qMax(1, (br.height() + h - 1) / h);
    }
    return lines;
}

class TstRibbonLayout : public QObject
{
    Q_OBJECT
private slots:
    void wrapSplitsViewRawDataIntoTwoFittingLines() {
        const QFont f = RibbonGroup::largeButtonFont();
        const QString wrapped = RibbonGroup::wrapLabelText("View Raw Data", f);
        const QStringList parts = wrapped.split('\n');
        QCOMPARE(parts.size(), 2);
        const int textW = RibbonGroup::largeButtonTextWidth();
        const QFontMetrics fm(f);
        QVERIFY2(fm.horizontalAdvance(parts[0]) <= textW,
                 qPrintable(QString("line1 '%1' overflows %2px").arg(parts[0]).arg(textW)));
        QVERIFY2(fm.horizontalAdvance(parts[1]) <= textW,
                 qPrintable(QString("line2 '%1' overflows %2px").arg(parts[1]).arg(textW)));
    }

    void renderedButtonNeverExceedsTwoLines_standardScale() {
        const QFont f = RibbonGroup::largeButtonFont();
        const QString wrapped = RibbonGroup::wrapLabelText("View Raw Data", f);
        QVERIFY(wrappedLineCount(wrapped, RibbonGroup::largeButtonTextWidth(), f) <= 2);
    }

    void renderedButtonNeverExceedsTwoLines_at150Percent() {
        QFont f = RibbonGroup::largeButtonFont();
        f.setPointSizeF(f.pointSizeF() * 1.5);
        const QString wrapped = RibbonGroup::wrapLabelText("View Raw Data", f);
        const int textW = RibbonGroup::largeButtonTextWidth(f);
        QVERIFY(wrappedLineCount(wrapped, textW, f) <= 2);
    }

    void liveButtonFitsWithinItsBounds() {
        RibbonGroup grp("Data");
        QToolButton* b = grp.addLargeButton("View Raw Data", QIcon());
        QVERIFY(b != nullptr);
        b->resize(b->minimumSize());
        const QFont f = b->font();
        const int twoLineH = RibbonGroup::largeButtonHeight(f);
        QVERIFY2(b->minimumHeight() >= twoLineH,
                 qPrintable(QString("button min height %1 < required %2")
                                .arg(b->minimumHeight()).arg(twoLineH)));
        QVERIFY(b->text().contains('\n'));
    }
};

QTEST_MAIN(TstRibbonLayout)
#include "tst_ribbonlayout.moc"
