#include <QtTest/QtTest>
#include <QCoreApplication>
#include <QEvent>
#include <QWidget>
#include <QSignalSpy>
#include "utils/ResponsiveLayout.h"

using DVE::ResponsiveLayout;

class TstResponsiveLayout : public QObject
{
    Q_OBJECT
private slots:
    void cleanup() {
        ResponsiveLayout& rl = ResponsiveLayout::instance();
        rl.stopTracking();
        // Drain any debounce timeout queued before stopTracking().
        QCoreApplication::processEvents();
        QCoreApplication::sendPostedEvents(nullptr, QEvent::Timer);
        // Re-seed the singleton to a known Standard baseline so the next test does
        // not inherit this test's m_lastWidth / m_breakpoint.
        QWidget seed;
        seed.resize(1280, 800);
        rl.beginTracking(&seed);              // recompute(1280) -> Standard
        rl.stopTracking();
        QVERIFY(!rl.isCompact());
    }

    void thresholdBoundary() {
        // Use QTRY_* macros so assertions poll until the debounce timer fires
        // rather than relying on a fixed wall-clock delay that can expire before
        // the QTimer::timeout() slot runs on a loaded machine.
        static constexpr int kWaitMs = ResponsiveLayout::kDebounceIntervalMs * 5;

        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(1200, 800);
        w.show();
        rl.beginTracking(&w);
        QTRY_COMPARE_WITH_TIMEOUT(rl.currentBreakpoint(), ResponsiveLayout::Standard, kWaitMs);
        QVERIFY(!rl.isCompact());

        QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);
        w.resize(1000, 800);
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 1, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Compact);
        QVERIFY(rl.isCompact());

        w.resize(1200, 800);
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 2, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Standard);
    }

    void veryNarrowBoundary() {
        static constexpr int kWaitMs = ResponsiveLayout::kDebounceIntervalMs * 5;

        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(900, 800);                 // Compact band (760..1100)
        w.show();
        rl.beginTracking(&w);
        QTRY_COMPARE_WITH_TIMEOUT(rl.currentBreakpoint(), ResponsiveLayout::Compact, kWaitMs);
        QVERIFY(rl.isCompact());
        QVERIFY(!rl.isVeryNarrow());

        QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);

        w.resize(700, 800);                 // Compact -> VeryNarrow
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 1, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::VeryNarrow);
        QVERIFY(rl.isVeryNarrow());
        QVERIFY(!rl.isCompact());

        w.resize(900, 800);                 // VeryNarrow -> Compact
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 2, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Compact);
        QVERIFY(rl.isCompact());
        QVERIFY(!rl.isVeryNarrow());
    }

    void veryNarrowExactThreshold() {
        static constexpr int kWaitMs = ResponsiveLayout::kDebounceIntervalMs * 5;

        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(760, 800);                 // exactly at the threshold
        w.show();
        rl.beginTracking(&w);
        QTRY_COMPARE_WITH_TIMEOUT(rl.currentBreakpoint(), ResponsiveLayout::Compact, kWaitMs);

        QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);
        w.resize(759, 800);
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 1, kWaitMs);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::VeryNarrow);
    }

    void debounceCoalescesResizes() {
        ResponsiveLayout& rl = ResponsiveLayout::instance();
        QWidget w;
        w.resize(1500, 800);
        w.show();
        rl.beginTracking(&w);
        QTest::qWait(ResponsiveLayout::kDebounceIntervalMs + 30);

        QSignalSpy widthSpy(&rl, &ResponsiveLayout::widthChanged);
        // Rapid-fire 5 resize events within the debounce window.
        // processEvents() delivers each resize to the event filter without
        // yielding enough wall-clock time for the 50 ms debounce to fire.
        for (int i = 0; i < 5; ++i) {
            w.resize(1500 - i * 10, 800);
            QCoreApplication::processEvents();
        }
        // Poll until the single coalesced emission arrives (or time out).
        QTRY_COMPARE_WITH_TIMEOUT(widthSpy.count(), 1,
                                  ResponsiveLayout::kDebounceIntervalMs * 5);
        QCOMPARE(widthSpy.first().at(0).toInt(), 1460);
    }

    void singletonStateResetsBetweenWindows() {
        ResponsiveLayout& rl = ResponsiveLayout::instance();

        QWidget narrow;
        narrow.resize(900, 700);
        narrow.show();
        rl.beginTracking(&narrow);
        QTRY_COMPARE_WITH_TIMEOUT(rl.currentBreakpoint(),
            ResponsiveLayout::Compact,
            ResponsiveLayout::kDebounceIntervalMs * 5);

        rl.stopTracking();
        QCoreApplication::processEvents();
        QWidget seed; seed.resize(1280, 800);
        rl.beginTracking(&seed);
        rl.stopTracking();
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Standard);

        QWidget again; again.resize(900, 700); again.show();
        QSignalSpy bpSpy(&rl, &ResponsiveLayout::breakpointChanged);
        rl.beginTracking(&again);
        QTRY_COMPARE_WITH_TIMEOUT(bpSpy.count(), 1,
            ResponsiveLayout::kDebounceIntervalMs * 5);
        QCOMPARE(rl.currentBreakpoint(), ResponsiveLayout::Compact);
    }
};

QTEST_MAIN(TstResponsiveLayout)
#include "tst_responsivelayout.moc"
