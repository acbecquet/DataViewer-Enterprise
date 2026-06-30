#include <QApplication>
#include <QScrollBar>
#include <QWidget>
#include <QtTest>

#include "ScrollHost.h"

using DVE::ScrollHost;

namespace {

// A widget that reports a fixed, oversized sizeHint in the requested
// direction(s) so the host's ScrollBarAsNeeded policy must engage when the
// viewport is smaller. minimumSizeHint mirrors sizeHint so widgetResizable
// can never shrink it below the overflow point.
class FixedHintWidget : public QWidget {
public:
    explicit FixedHintWidget(QSize hint, QWidget* parent = nullptr)
        : QWidget(parent), m_hint(hint) {}
    QSize sizeHint() const override { return m_hint; }
    QSize minimumSizeHint() const override { return m_hint; }
private:
    QSize m_hint;
};

// Lay the host out at a known small viewport and let the as-needed policy
// evaluate. A top-level show + event flush is required for the scrollbars to
// re-range against the viewport.
void settle(ScrollHost& host, QSize viewport) {
    host.resize(viewport);
    host.show();
    QApplication::processEvents();
    QTest::qWait(50);
    QApplication::processEvents();
}

} // namespace

class tst_ScrollHost : public QObject {
    Q_OBJECT
private slots:
    // The factory yields a configured, non-null host that adopts the content.
    void testWrapAdoptsContent() {
        auto* content = new QWidget;
        ScrollHost* host = ScrollHost::wrap(content);
        QVERIFY(host != nullptr);
        QCOMPARE(host->widget(), content);
        QVERIFY(host->widgetResizable());
        QCOMPARE(host->frameShape(), QFrame::NoFrame);
        QCOMPARE(host->horizontalScrollBarPolicy(), Qt::ScrollBarAsNeeded);
        QCOMPARE(host->verticalScrollBarPolicy(), Qt::ScrollBarAsNeeded);
        delete host;
    }

    // Content taller than the viewport -> vertical scrollbar engages.
    void testVerticalScrollWhenContentTaller() {
        ScrollHost host;
        host.setWidget(new FixedHintWidget(QSize(120, 4000)));
        settle(host, QSize(300, 200));
        QVERIFY2(host.verticalScrollBar()->maximum() > 0,
                 "vertical content overflow must produce a scrollable range");
        QVERIFY(host.scrollbarActive(Qt::Vertical));
        QVERIFY(host.contentOverflows());
    }

    // Content wider than the viewport -> horizontal scrollbar engages.
    void testHorizontalScrollWhenContentWider() {
        ScrollHost host;
        host.setWidget(new FixedHintWidget(QSize(4000, 120)));
        settle(host, QSize(300, 200));
        QVERIFY2(host.horizontalScrollBar()->maximum() > 0,
                 "horizontal content overflow must produce a scrollable range");
        QVERIFY(host.scrollbarActive(Qt::Horizontal));
    }

    // Content that fits -> no scroll range in either direction (invisible no-op).
    void testNoScrollWhenContentFits() {
        ScrollHost host;
        host.setWidget(new FixedHintWidget(QSize(100, 100)));
        settle(host, QSize(600, 600));
        QCOMPARE(host.verticalScrollBar()->maximum(), 0);
        QCOMPARE(host.horizontalScrollBar()->maximum(), 0);
        QVERIFY(!host.contentOverflows());
    }

    // Horizontal-only wrap forces the vertical axis off (ribbon group row).
    void testHorizontalOnlyWrapDisablesVerticalAxis() {
        ScrollHost* host = ScrollHost::wrap(new FixedHintWidget(QSize(4000, 4000)),
                                            Qt::Horizontal);
        QCOMPARE(host->verticalScrollBarPolicy(), Qt::ScrollBarAlwaysOff);
        QCOMPARE(host->horizontalScrollBarPolicy(), Qt::ScrollBarAsNeeded);
        delete host;
    }
};

QTEST_MAIN(tst_ScrollHost)
#include "tst_scrollhost.moc"
