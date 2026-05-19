#include <QApplication>
#include <QDateTime>
#include <QLabel>
#include <QPushButton>
#include <QSignalSpy>
#include <QtTest>

#include "OfflineBanner.h"

using DVE::OfflineBanner;

namespace {

// Helper: locate the labels by walking findChildren in document order. The
// constructor adds m_primaryLabel before m_secondaryLabel, so the first
// QLabel encountered is the primary.
QLabel* primaryLabel(OfflineBanner& b) {
    const auto labels = b.findChildren<QLabel*>();
    return labels.isEmpty() ? nullptr : labels.at(0);
}
QLabel* secondaryLabel(OfflineBanner& b) {
    const auto labels = b.findChildren<QLabel*>();
    return labels.size() < 2 ? nullptr : labels.at(1);
}
QLabel* tertiaryLabel(OfflineBanner& b) {
    const auto labels = b.findChildren<QLabel*>();
    return labels.size() < 3 ? nullptr : labels.at(2);
}
QPushButton* retryButton(OfflineBanner& b) {
    return b.findChild<QPushButton*>();
}

} // namespace

class tst_OfflineBanner : public QObject {
    Q_OBJECT
private slots:
    void testDefaultStateHidden() {
        OfflineBanner b;
        QVERIFY(!b.isVisible());
    }

    void testPrimaryTextWithLastSync() {
        OfflineBanner b;
        b.setLastSync(QDateTime::currentDateTime().addSecs(-30));
        QLabel* prim = primaryLabel(b);
        QVERIFY(prim != nullptr);
        const QString txt = prim->text();
        QVERIFY2(txt.contains("just now"), qPrintable(txt));
        QVERIFY2(txt.contains("Working offline"), qPrintable(txt));
    }

    void testPrimaryTextWithoutLastSync() {
        OfflineBanner b;
        b.setLastSync(QDateTime()); // invalid
        QLabel* prim = primaryLabel(b);
        QVERIFY(prim != nullptr);
        const QString txt = prim->text();
        QVERIFY2(txt.contains("No previous snapshot"), qPrintable(txt));
    }

    void testHumanizeRelativeTimeBoundaries() {
        const QDateTime now = QDateTime::currentDateTime();

        // Invalid input -> empty string
        QCOMPARE(OfflineBanner::humanizeRelativeTime(QDateTime()), QString());

        // 30s ago -> "just now"
        QCOMPARE(OfflineBanner::humanizeRelativeTime(now.addSecs(-30)),
                 QString("just now"));

        // 5 min ago -> "5 minutes ago"
        QCOMPARE(OfflineBanner::humanizeRelativeTime(now.addSecs(-5 * 60)),
                 QString("5 minutes ago"));

        // 2 hours ago -> "2 hours ago"
        QCOMPARE(OfflineBanner::humanizeRelativeTime(now.addSecs(-2 * 3600)),
                 QString("2 hours ago"));

        // 30 hours ago -> "yesterday at ..."
        const QString yesterday =
            OfflineBanner::humanizeRelativeTime(now.addSecs(-30 * 3600));
        QVERIFY2(yesterday.startsWith("yesterday at "),
                 qPrintable(yesterday));

        // 5 days ago -> "MMM d, yyyy" (e.g., "May 8, 2026")
        const QDateTime fiveDaysAgo = now.addDays(-5);
        const QString older =
            OfflineBanner::humanizeRelativeTime(fiveDaysAgo);
        const QString expected = fiveDaysAgo.toString("MMM d, yyyy");
        QCOMPARE(older, expected);
    }

    void testPendingCountZeroHidesSecondary() {
        OfflineBanner b;
        b.setPendingCount(0);
        QLabel* sec = secondaryLabel(b);
        QVERIFY(sec != nullptr);
        QVERIFY(!sec->isVisible() || sec->text().isEmpty());
        // The widget itself is hidden by default, so we cannot rely on
        // isVisible() returning true for the secondary. We assert via
        // visibility hint: the secondary's QWidget::isHidden() flag.
        QVERIFY(sec->isHidden());
    }

    void testPendingCountOne() {
        OfflineBanner b;
        b.setPendingCount(1);
        QLabel* sec = secondaryLabel(b);
        QVERIFY(sec != nullptr);
        QCOMPARE(sec->text(),
                 QString("1 unsaved change. Will retry when reconnected."));
        // The setVisible(true) call inside refreshSecondaryText flips the
        // hidden flag back off, even while the parent banner is hidden.
        QVERIFY(!sec->isHidden());
    }

    void testPendingCountMany() {
        OfflineBanner b;
        b.setPendingCount(7);
        QLabel* sec = secondaryLabel(b);
        QVERIFY(sec != nullptr);
        QCOMPARE(sec->text(),
                 QString("7 unsaved changes. Will retry when reconnected."));
        QVERIFY(!sec->isHidden());
    }

    // -- H7: setPendingFailureCount surfaces the residual count of edits that
    //        a recent flush could NOT push to Postgres. Distinct from the
    //        secondary "N unsaved changes" line so the user sees the issue
    //        is not "queued" but "failed to send".
    void testPendingFailureCountZeroHidesTertiary() {
        OfflineBanner b;
        b.setPendingFailureCount(0);
        QLabel* ter = tertiaryLabel(b);
        QVERIFY(ter != nullptr);
        QVERIFY(ter->isHidden());
    }

    void testPendingFailureCountOne() {
        OfflineBanner b;
        b.setPendingFailureCount(1);
        QLabel* ter = tertiaryLabel(b);
        QVERIFY(ter != nullptr);
        QVERIFY(!ter->isHidden());
        QVERIFY2(ter->text().contains("1"), qPrintable(ter->text()));
        QVERIFY2(ter->text().contains("failed"), qPrintable(ter->text()));
    }

    void testPendingFailureCountMany() {
        OfflineBanner b;
        b.setPendingFailureCount(5);
        QLabel* ter = tertiaryLabel(b);
        QVERIFY(ter != nullptr);
        QVERIFY(!ter->isHidden());
        QVERIFY2(ter->text().contains("5"), qPrintable(ter->text()));
        QVERIFY2(ter->text().contains("failed"), qPrintable(ter->text()));
    }

    void testRetryButtonEmitsSignal() {
        OfflineBanner b;
        QSignalSpy spy(&b, &OfflineBanner::retryClicked);
        QPushButton* btn = retryButton(b);
        QVERIFY(btn != nullptr);
        btn->click();
        QCOMPARE(spy.count(), 1);
    }
};

QTEST_MAIN(tst_OfflineBanner)
#include "tst_offlinebanner.moc"
