#pragma once

#include <QFrame>
#include <QString>

class QLabel;
class QToolButton;

namespace DVE {

// Non-modal banner that drops down over the central editor when a row
// currently open in the UI is deleted by another user. Shows the resource
// label (e.g. "file 'foo.xlsx'") and the remote user's name; the body acts
// as a clickable button that emits recreateRequested for the caller to push
// the local state back into the DB as a fresh INSERT. A trailing "x"
// dismisses the banner without action (emits dismissed).
//
// Lives at the top of the central editor area between PresenceAvatarBar
// and the central stack. Hidden by default; MainWindow toggles visibility
// based on the DELETE NOTIFY for the currently-open resource.
class RowDeletedBanner : public QFrame {
    Q_OBJECT
public:
    explicit RowDeletedBanner(QWidget* parent = nullptr);

    // Populate and show the banner. resourceLabel is presented verbatim
    // (e.g. "file 'foo.xlsx'"), userName is the remote actor displayed
    // in the message. Idempotent — recalling overrides the previous text.
    void showFor(const QString& resourceLabel, const QString& userName);

    // Hide and clear state. Called either by the user dismissing or by
    // MainWindow after the recreate succeeds.
    void dismiss();

signals:
    // The user clicked the banner body — caller should push local state
    // as a fresh INSERT (clear id/version, then save through coordinator).
    void recreateRequested();

    // The user clicked the dismiss button — caller takes no action.
    void dismissed();

private:
    QLabel*      m_messageLabel = nullptr;
    QToolButton* m_recreateBtn  = nullptr;
    QToolButton* m_closeBtn     = nullptr;
};

} // namespace DVE
