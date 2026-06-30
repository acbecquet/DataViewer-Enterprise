#include "ScrollHost.h"

#include <QPalette>
#include <QScrollBar>

namespace DVE {

ScrollHost::ScrollHost(QWidget* parent)
    : QScrollArea(parent)
{
    setObjectName(QStringLiteral("scrollHost"));
    // Content fills the viewport when there is room; scrolls when there is not.
    setWidgetResizable(true);
    // Scrollbars appear only on overflow, independently per direction.
    setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    // Visual no-op until it scrolls: no frame, transparent viewport, no inset.
    setFrameShape(QFrame::NoFrame);
    setLineWidth(0);
    setViewportMargins(0, 0, 0, 0);
    setContentsMargins(0, 0, 0, 0);
    // Let the wrapped region's own background show through; do NOT paint our
    // own surface (the scroll area's viewport autofills by default).
    setBackgroundRole(QPalette::NoRole);
    viewport()->setAutoFillBackground(false);
    setAttribute(Qt::WA_TranslucentBackground, false);
}

ScrollHost* ScrollHost::wrap(QWidget* content, Qt::Orientations scroll)
{
    auto* host = new ScrollHost();
    if (!(scroll & Qt::Horizontal))
        host->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    if (!(scroll & Qt::Vertical))
        host->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    host->setWidget(content);   // ScrollHost takes ownership of content
    return host;
}

bool ScrollHost::scrollbarActive(Qt::Orientation o) const
{
    const QScrollBar* bar =
        (o == Qt::Horizontal) ? horizontalScrollBar() : verticalScrollBar();
    return bar && bar->isVisible() && bar->maximum() > 0;
}

bool ScrollHost::contentOverflows() const
{
    return (horizontalScrollBar() && horizontalScrollBar()->maximum() > 0) ||
           (verticalScrollBar()   && verticalScrollBar()->maximum()   > 0);
}

} // namespace DVE
