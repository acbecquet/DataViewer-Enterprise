#pragma once

#include <QFrame>

class QLabel;
class QToolButton;

namespace DVE {

// Amber banner shown above the central editor area when the active TPM
// FileResult has at least one sheet with dbDataIncomplete == true --
// meaning it was loaded from a legacy DB record that had no per-row /
// raw-grid data persisted.
//
// Non-blocking, informational only.  No buttons to trigger repair --
// repair is either:
//   (a) automatic self-heal when the source .xlsx is present on disk
//       (the Load-from-Database path re-parses and calls saveFile), or
//   (b) the headless DataViewer.exe --repair-db admin tool.
//
// The user can dismiss the banner via the X button; it reappears the
// next time a file with incomplete data is loaded.
//
// Pattern mirrors OfflineBanner / RowDeletedBanner: constructed by
// MainWindow::setupCentralWidget(), inserted into the central VBoxLayout
// just below m_offlineBanner, hidden by default.  MainWindow calls
// setVisible(true/false) from updateIncompleteDataBanner().
class IncompleteDataBanner : public QFrame {
    Q_OBJECT
public:
    explicit IncompleteDataBanner(QWidget* parent = nullptr);

    // Show the banner for the given filename.  Idempotent -- recalling
    // with a different name just updates the text.
    void showForFile(const QString& fileName);

    // Hide and reset state.
    void dismiss();

signals:
    void dismissed();

private:
    QLabel*      m_label   = nullptr;
    QToolButton* m_closeBtn = nullptr;
};

} // namespace DVE
