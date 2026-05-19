#pragma once

#include <QFrame>
#include <QDateTime>

class QLabel;
class QPushButton;

namespace DVE {

// Yellow "Working offline" banner shown at the top of MainWindow's central
// area when the live Postgres connection is unreachable. Designed to mount
// above the central widget stack via QVBoxLayout::insertWidget(0, ...).
//
// Visual:
//   +------------------------------------------------------------+
//   | (warn) Working offline - read-only. Last synced 2 hours ago|
//   |        1 unsaved change. Will retry when reconnected.      |
//   |                                                  [Retry]   |
//   +------------------------------------------------------------+
//
// Styling: background #fef3c7, bottom border 1px #f59e0b, padding 6px.
// The icon is plain unicode for portability; no QIcon needed.
//
// Public API:
//   - setVisible(bool) - toggles the banner; not auto-managed by the widget.
//   - setLastSync(QDateTime) - formats the timestamp into a human-readable
//     relative string ("2 hours ago", "just now", "yesterday at 4:12 PM").
//     Pass an invalid QDateTime to suppress the "Last synced..." clause.
//   - setPendingCount(int) - sets the second-line copy. 0 hides the second
//     line entirely. 1 -> "1 unsaved change. Will retry when reconnected."
//     2+ -> "N unsaved changes. Will retry when reconnected."
//   - retryClicked signal - emitted when user clicks Retry.
//
// C5 wires this:
//   - connect(monitor, &ConnectionMonitor::wentOffline, banner, [=]() {
//         banner->setLastSync(m_snapshot->snapshotTakenAt());
//         banner->setVisible(true);
//     });
//   - connect(monitor, &ConnectionMonitor::cameOnline, banner, [=]() {
//         banner->setVisible(false);
//     });
//   - connect(banner, &OfflineBanner::retryClicked, this, &MainWindow::onRetryClicked);
class OfflineBanner : public QFrame {
    Q_OBJECT
public:
    explicit OfflineBanner(QWidget* parent = nullptr);
    ~OfflineBanner() override;

    void setLastSync(const QDateTime& whenLocal);
    void setPendingCount(int n);
    // H7: surfaces the count of edits that a recent flush attempt couldn't
    // push to Postgres (as opposed to queued-but-not-yet-tried edits, which
    // setPendingCount tracks). 0 hides the tertiary line.
    void setPendingFailureCount(int n);

    // Exposed for unit testing of the boundary cases (just-now, minutes,
    // hours, yesterday, older). Pure function with no side effects.
    static QString humanizeRelativeTime(const QDateTime& whenLocal);

signals:
    void retryClicked();

private:
    QLabel*      m_primaryLabel        = nullptr;
    QLabel*      m_secondaryLabel      = nullptr;
    QLabel*      m_tertiaryLabel       = nullptr;
    QPushButton* m_retryButton         = nullptr;
    QDateTime    m_lastSync;
    int          m_pendingCount        = 0;
    int          m_pendingFailureCount = 0;

    void refreshPrimaryText();
    void refreshSecondaryText();
    void refreshTertiaryText();
};

} // namespace DVE
