#include "ResponsiveLayout.h"
#include <QWidget>
#include <QTimer>
#include <QResizeEvent>

namespace DVE {

ResponsiveLayout::ResponsiveLayout() {
    m_debounce = new QTimer(this);
    m_debounce->setSingleShot(true);
    m_debounce->setInterval(50);
    connect(m_debounce, &QTimer::timeout, this, [this]() {
        recompute(m_pendingWidth);
    });
}

ResponsiveLayout& ResponsiveLayout::instance() {
    // Heap-allocate to avoid static destructor running after QApplication
    // teardown (which would double-delete QTimer children via ~QObject).
    static ResponsiveLayout* s_instance = new ResponsiveLayout();
    return *s_instance;
}

void ResponsiveLayout::beginTracking(QWidget* window) {
    if (m_window == window) return;
    if (m_window) m_window->removeEventFilter(this);
    m_window = window;
    if (m_window) {
        m_window->installEventFilter(this);
        recompute(m_window->width());
    }
}

void ResponsiveLayout::stopTracking() {
    if (m_window) {
        m_window->removeEventFilter(this);
        m_window = nullptr;
    }
    m_debounce->stop();
}

bool ResponsiveLayout::eventFilter(QObject* watched, QEvent* event) {
    if (watched == m_window && event->type() == QEvent::Resize) {
        auto* re = static_cast<QResizeEvent*>(event);
        m_pendingWidth = re->size().width();
        m_debounce->start();
    }
    return QObject::eventFilter(watched, event);
}

void ResponsiveLayout::recompute(int width) {
    const Breakpoint newBp = (width < kCompactThreshold) ? Compact : Standard;
    const bool widthChangedFlag = (width != m_lastWidth);
    const bool bpChanged = (newBp != m_breakpoint);

    m_lastWidth = width;
    m_breakpoint = newBp;

    if (widthChangedFlag) emit widthChanged(width);
    if (bpChanged)        emit breakpointChanged(newBp, width);
}

} // namespace DVE
