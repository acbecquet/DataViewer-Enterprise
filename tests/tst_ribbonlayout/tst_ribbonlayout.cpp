#include <QtTest/QtTest>
#include <QApplication>
#include <QTabBar>
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
        // The group is now SHORTER: the bottom group-title row (1px separator +
        // one text line) was removed. groupMinimumHeight = top+bottom margin +
        // large-button band (32px icon + 2 label lines + chrome). At 8pt Segoe
        // UI this measures 82 here (was ~94 with the title row); band ~78-86.
        const int gh = RibbonGroup::groupMinimumHeight(base);
        QVERIFY2(gh >= 78, qPrintable(QString("groupMinimumHeight %1 < 78").arg(gh)));
        QVERIFY2(gh <= 86, qPrintable(QString("groupMinimumHeight %1 > 86").arg(gh)));
        // Still fits the two-line large button (invariant unchanged).
        QVERIFY(gh >= RibbonGroup::largeButtonHeight(base));
        QVERIFY(RibbonGroup::groupMinimumHeight(scaled) >
                RibbonGroup::groupMinimumHeight(base));
        // A populated ribbon is at least tab bar + group band tall.
        RibbonWidget ribbon;
        ribbon.addTab("Home")->addGroup("G")->addLargeButton("X", QIcon());
        QVERIFY(ribbon.minimumSizeHint().height() >
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
        // And it actually SCROLLS: even icons-only (8 x 36px) overflows 200px,
        // so the horizontal scrollbar must be engaged, not just permitted.
        QCoreApplication::processEvents();
        QVERIFY2(host->scrollbarActive(Qt::Horizontal),
                 "overflowing ribbon row did not engage its horizontal scrollbar");
    }

    // Compact mode is ONE global flag driven by the WIDEST tab, so the icon
    // style stays identical no matter which tab is showing. Owner bug
    // 2026-07-08: at a narrow width the wide Home tab compacted to icons but
    // the narrower Reports/Tools/Settings tabs kept their labels -- switching
    // tabs must not change the toolbar's appearance.
    void compactModeIsConsistentAcrossTabs() {
        RibbonWidget ribbon;
        RibbonTab* narrow = ribbon.addTab("Narrow");
        narrow->addGroup("A")->addLargeButton("One", QIcon());
        RibbonTab* wide = ribbon.addTab("Wide");
        RibbonGroup* g = wide->addGroup("B");
        for (int i = 0; i < 8; ++i)
            g->addLargeButton(QString("Button %1").arg(i), QIcon());

        // 400px fits the narrow tab alone but not the wide one. The wide tab's
        // need governs, so every tab is icons-only regardless of which shows.
        ribbon.setCurrentTab(0);
        ribbon.resize(400, 140);
        ribbon.show();
        QVERIFY(QTest::qWaitForWindowExposed(&ribbon));
        QCoreApplication::processEvents();
        QVERIFY2(ribbon.isCompactMode(),
                 "wide tab overflows 400px: all tabs must be icons-only");
        ribbon.setCurrentTab(1);
        QCoreApplication::processEvents();
        QVERIFY2(ribbon.isCompactMode(), "wide tab stays icons-only");
        ribbon.setCurrentTab(0);
        QCoreApplication::processEvents();
        QVERIFY2(ribbon.isCompactMode(),
                 "narrow tab stays icons-only too -- style is consistent");

        // Past the widest tab's need every tab shows labels, again uniformly.
        const int need = ribbon.fullModeNeededWidth();
        ribbon.resize(need + 60, 140);
        QCoreApplication::processEvents();
        QVERIFY2(!ribbon.isCompactMode(), "all tabs fit: labels return everywhere");
        ribbon.setCurrentTab(1);
        QCoreApplication::processEvents();
        QVERIFY2(!ribbon.isCompactMode(), "wide tab is labeled at full width too");
    }

    // ── Owner feedback 2026-07-01: tight 5px band + content-driven collapse ──

    // Full mode: exactly 5px between the tab page's top edge and the button,
    // and 5px between the button's bottom and the tab page's bottom edge --
    // the visual band the owner sees (tab margins + group margins combined).
    void fullModeBandPaddingIsFivePx() {
        RibbonWidget ribbon;
        RibbonTab* tab = ribbon.addTab("Home");
        RibbonGroup* g = tab->addGroup("File");
        QToolButton* btn = g->addLargeButton("New File", QIcon());
        // Height = the ribbon's own hint, exactly like MainWindow's layout
        // (vertical policy Fixed). A taller forced height would just center
        // the fixed-height group in the slack and hide real padding bugs.
        ribbon.resize(600, ribbon.sizeHint().height());
        ribbon.show();
        QVERIFY(QTest::qWaitForWindowExposed(&ribbon));
        QCoreApplication::processEvents();
        QVERIFY(!ribbon.isCompactMode());
        const QPoint p = btn->mapTo(tab, QPoint(0, 0));
        QCOMPARE(p.y(), RibbonGroup::kBandPadding);
        QCOMPARE(g->height() - (btn->mapTo(g, QPoint(0, 0)).y() + btn->height()),
                 RibbonGroup::kBandPadding);
    }

    // Compact mode must LOWER the group's minimum height to the tight icon
    // band (5 + 36 + 5 = 46) instead of keeping the full-mode ~82px minimum
    // that left the icons floating in dead space.
    void compactModeTightensGroupBand() {
        RibbonWidget ribbon;
        RibbonTab* tab = ribbon.addTab("Home");
        RibbonGroup* g = tab->addGroup("File");
        g->addLargeButton("New File", QIcon());
        const int fullH = ribbon.minimumSizeHint().height();
        // No processEvents on purpose: mode flips must refresh the hint chain
        // in the SAME tick (setCompactMode invalidates the group's layouts).
        ribbon.setCompactMode(true);
        QCOMPARE(RibbonGroup::compactGroupHeight(),
                 RibbonGroup::kBandPadding + RibbonGroup::kCompactButtonSide
                     + RibbonGroup::kBandPadding);
        QCOMPARE(g->minimumHeight(), RibbonGroup::compactGroupHeight());
        QVERIFY2(ribbon.minimumSizeHint().height() < fullH,
                 qPrintable(QString("compact ribbon hint %1 not below full %2")
                                .arg(ribbon.minimumSizeHint().height()).arg(fullH)));
        ribbon.setCompactMode(false);
        QCOMPARE(g->minimumHeight(), RibbonGroup::groupMinimumHeight());
        QCOMPARE(ribbon.minimumSizeHint().height(), fullH);
    }

    // And the RENDERED compact group is actually tight (the live bug: the
    // stale minimum kept the rendered band ~82px tall around 36px icons).
    void compactRibbonRendersTight() {
        RibbonWidget ribbon;
        RibbonTab* tab = ribbon.addTab("Home");
        RibbonGroup* g = tab->addGroup("File");
        for (int i = 0; i < 12; ++i)
            g->addLargeButton(QString("Button %1").arg(i), QIcon());
        ribbon.resize(420, 200);
        ribbon.show();
        QVERIFY(QTest::qWaitForWindowExposed(&ribbon));
        QCoreApplication::processEvents();
        QVERIFY2(ribbon.isCompactMode(),
                 "12 labeled buttons cannot fit 420px; icons-only must engage");
        QVERIFY2(g->height() <= RibbonGroup::compactGroupHeight() + 2,
                 qPrintable(QString("compact group renders %1px tall, want <= %2")
                                .arg(g->height())
                                .arg(RibbonGroup::compactGroupHeight() + 2)));
    }

    // Labels persist until the labeled buttons genuinely no longer fit the
    // ribbon's width; only then does icons-only engage (and it recovers).
    void labelsPersistUntilSqueezedThenIconsOnly() {
        RibbonWidget ribbon;
        RibbonTab* tab = ribbon.addTab("Home");
        RibbonGroup* g = tab->addGroup("File");
        for (int i = 0; i < 8; ++i)
            g->addLargeButton(QString("Button %1").arg(i), QIcon());
        const int need = ribbon.fullModeNeededWidth();
        QVERIFY2(need > 500 && need < 780,
                 qPrintable(QString("unexpected full-mode need %1").arg(need)));

        ribbon.resize(need + 60, 140);
        ribbon.show();
        QVERIFY(QTest::qWaitForWindowExposed(&ribbon));
        QCoreApplication::processEvents();
        QVERIFY2(!ribbon.isCompactMode(), "labels must persist while they fit");

        ribbon.resize(need - 60, 140);
        QCoreApplication::processEvents();
        QVERIFY2(ribbon.isCompactMode(),
                 "icons-only must engage once labels no longer fit");

        ribbon.resize(need + 60, 140);
        QCoreApplication::processEvents();
        QVERIFY2(!ribbon.isCompactMode(),
                 "labels must come back when there is room again");
    }

    // Mode-swapped (hidden) buttons must not inflate the width need, or TPM
    // mode would collapse because of hidden sensory buttons.
    void hiddenButtonsDontCountTowardNeededWidth() {
        RibbonWidget ribbon;
        RibbonTab* tab = ribbon.addTab("Home");
        RibbonGroup* g = tab->addGroup("File");
        QToolButton* a = g->addLargeButton("Alpha", QIcon());
        g->addLargeButton("Beta", QIcon());
        const int both = ribbon.fullModeNeededWidth();
        a->setVisible(false);
        const int one = ribbon.fullModeNeededWidth();
        QVERIFY2(one < both,
                 qPrintable(QString("need %1 not reduced from %2 by hiding")
                                .arg(one).arg(both)));
        QVERIFY(one >= RibbonGroup::kLargeButtonWidth);
    }
};

QTEST_MAIN(TstRibbonLayout)
#include "tst_ribbonlayout.moc"
