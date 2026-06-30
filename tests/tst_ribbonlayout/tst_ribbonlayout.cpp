#include <QtTest/QtTest>
#include <QApplication>
#include <QToolButton>
#include <QFont>
#include <QFontMetrics>
#include <QStringList>
#include <QIcon>
#include <QTextLayout>
#include <QTextOption>
#include "RibbonWidget.h"
#include "ScrollHost.h"

// Count how many visual lines QToolButton's word-wrap would produce for a
// given (already newline-split) label inside a text rectangle of `textW`
// pixels wide, at font `f`. Lays each hard line out with QTextLayout in the
// available width and counts the wrapped lines it actually produces -- this is
// how Qt itself breaks the text, so a half that doesn't fit genuinely yields a
// clipped extra row (the regression this test guards). Avoids the brittle
// (boundingRect.height()+lineSpacing-1)/lineSpacing ceiling form, which
// over-counts a single line to 2 at fractional point sizes where
// boundingRect().height() exceeds lineSpacing() by a pixel.
static int wrappedLineCount(const QString& label, int textW, const QFont& f)
{
    int lines = 0;
    QTextOption opt;
    opt.setWrapMode(QTextOption::WordWrap);
    for (const QString& hardLine : label.split('\n')) {
        QTextLayout layout(hardLine, f);
        layout.setTextOption(opt);
        layout.beginLayout();
        int hardLineLines = 0;
        for (;;) {
            QTextLine line = layout.createLine();
            if (!line.isValid())
                break;
            line.setLineWidth(textW);
            ++hardLineLines;
        }
        layout.endLayout();
        lines += qMax(1, hardLineLines);
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

    void groupAndRibbonHeightsGrowWithFont() {
        const QFont base = RibbonGroup::largeButtonFont();
        QFont scaled = base;
        scaled.setPointSizeF(base.pointSizeF() * 1.5);
        QVERIFY(RibbonGroup::groupMinimumHeight(base) >= 90);
        QVERIFY(RibbonGroup::groupMinimumHeight(base) <= 104);
        QVERIFY(RibbonGroup::groupMinimumHeight(scaled) >
                RibbonGroup::groupMinimumHeight(base));
        QVERIFY(RibbonWidget::ribbonMinimumHeight() >=
                RibbonGroup::groupMinimumHeight(base));
    }

    void groupRowScrollsHorizontallyWhenNarrow() {
        RibbonWidget ribbon;
        RibbonTab* tab = ribbon.addTab("Home");
        RibbonGroup* g = tab->addGroup("Data");
        for (int i = 0; i < 8; ++i)
            g->addLargeButton(QString("Button %1").arg(i), QIcon());
        ribbon.resize(200, 120);
        ribbon.show();
        QVERIFY(QTest::qWaitForWindowExposed(&ribbon));
        QWidget* page = ribbon.tabWidget()->widget(0);
        DVE::ScrollHost* host = qobject_cast<DVE::ScrollHost*>(page);
        QVERIFY2(host != nullptr, "ribbon tab page is not wrapped in a ScrollHost");
        QCOMPARE(host->verticalScrollBarPolicy(), Qt::ScrollBarAlwaysOff);
    }
};

QTEST_MAIN(TstRibbonLayout)
#include "tst_ribbonlayout.moc"
