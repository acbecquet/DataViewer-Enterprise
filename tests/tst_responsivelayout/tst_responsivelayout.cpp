#include <QtTest/QtTest>
#include <QWidget>
#include <QSignalSpy>
#include "utils/ResponsiveLayout.h"

using DVE::ResponsiveLayout;

class TstResponsiveLayout : public QObject
{
    Q_OBJECT
private slots:
    void cleanup() {
        // Detach singleton from any widget destroyed at end of each test.
        ResponsiveLayout::instance().stopTracking();
    }

    void thresholdBoundary() {
        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(1200, 800);
        w.show();
        rl.beginTracking(&w);
        QTest::qWait(80);  // debounce window is 50 ms
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Standard);
        QVERIFY(!rl.isCompact());

        QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);
        w.resize(1000, 800);
        QTest::qWait(80);
        QCOMPARE(bpSpy.count(), 1);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Compact);
        QVERIFY(rl.isCompact());

        w.resize(1200, 800);
        QTest::qWait(80);
        QCOMPARE(bpSpy.count(), 2);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Standard);
    }

    void debounceCoalescesResizes() {
        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(1500, 800);
        w.show();
        rl.beginTracking(&w);
        QTest::qWait(80);

        QSignalSpy widthSpy(&rl, &ResponsiveLayout::widthChanged);
        // Rapid-fire 5 resize events within the debounce window
        for (int i = 0; i < 5; ++i) {
            w.resize(1500 - i * 10, 800);
            QTest::qWait(5);
        }
        QTest::qWait(80);
        // Should coalesce to 1 emission of the final size
        QCOMPARE(widthSpy.count(), 1);
        QCOMPARE(widthSpy.first().at(0).toInt(), 1460);
    }
};

QTEST_MAIN(TstResponsiveLayout)
#include "tst_responsivelayout.moc"
