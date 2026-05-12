#pragma once

#include <QWidget>
#include <QVector>
#include <QUuid>

#include "../database/PresenceManager.h"  // for PresenceRow

namespace DVE {

// Horizontal strip of colored avatar circles, one per active user on the
// currently-open resource. The self-user gets a 2 px white inner ring so the
// current user can spot themselves in the bar at a glance. Hovering an
// avatar shows a tooltip with the user's full name and intent.
//
// Placed at the top of the central editor area by MainWindow; updated by
// MainWindow::refreshPresenceFor when the active resource's presence
// changes. clear() hides every avatar and is used when no resource is
// active.
class PresenceAvatarBar : public QWidget {
    Q_OBJECT
public:
    explicit PresenceAvatarBar(QWidget* parent = nullptr);

    void setPresence(const QVector<PresenceRow>& rows, const QUuid& selfUuid);
    void clear();

    QSize sizeHint() const override;

protected:
    void paintEvent(QPaintEvent* event) override;
    bool event(QEvent* event) override;  // tooltips

private:
    QVector<PresenceRow> m_rows;
    QUuid                m_selfUuid;

    // Returns the per-avatar rect inside this widget for hit-testing.
    QRect avatarRect(int index) const;
};

} // namespace DVE
