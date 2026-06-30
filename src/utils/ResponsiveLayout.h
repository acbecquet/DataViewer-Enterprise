#pragma once
#include <QObject>

class QWidget;
class QTimer;

namespace DVE {

// Singleton breakpoint detector + signal for responsive UI rules.
//
// THREAD SAFETY: Must be used on the main (GUI) thread only. The singleton
// holds a QTimer and installs a QObject event filter; neither is safe to
// call from worker threads.
//
// Usage:
//   ResponsiveLayout::instance().beginTracking(mainWindow);
//   connect(&ResponsiveLayout::instance(),
//           &ResponsiveLayout::breakpointChanged,
//           panel, &SensoryPanel::onBreakpointChanged);
//   if (ResponsiveLayout::instance().isCompact()) { ... }
class ResponsiveLayout : public QObject
{
    Q_OBJECT
public:
    enum Breakpoint {
        Standard,    // >= kCompactThreshold (default 1100 px)
        Compact,     // [kVeryNarrowThreshold, kCompactThreshold) -> icons-only ribbon + 32px sidebar strip
        VeryNarrow   // < kVeryNarrowThreshold (default 760 px) -> also auto-collapse both side docks
    };
    Q_ENUM(Breakpoint)

    static constexpr int kCompactThreshold = 1100;
    static constexpr int kVeryNarrowThreshold = 760;          // < 760 -> auto-collapse both side docks
    static constexpr int kSensoryNarrowThreshold = 700;        // < 700 -> 1-up cards
    static constexpr int kDetailedNarrowThreshold = 800;       // < 800 -> 1-col form
    static constexpr int kDetailedStackChartsThreshold = 1000; // < 1000 -> stack radars
    static constexpr int kDebounceIntervalMs = 50;             // resize-event debounce window

    static ResponsiveLayout& instance();

    // Hook the singleton up to a window. Installs an event filter that
    // listens for resize events on the window. Breakpoint computation is
    // debounced (50 ms) so dragging does not thrash. Safe to call multiple
    // times - re-arms on the new window.
    void beginTracking(QWidget* window);

    // Detach from the currently tracked window. Called automatically by
    // beginTracking() when switching windows, but may also be called
    // explicitly (e.g. in tests between test functions).
    void stopTracking();

    int  currentWidth() const { return m_lastWidth; }
    Breakpoint currentBreakpoint() const { return m_breakpoint; }
    bool isCompact() const { return m_breakpoint == Compact; }
    bool isVeryNarrow() const { return m_breakpoint == VeryNarrow; }

signals:
    // Emitted only when the breakpoint actually changes.
    void breakpointChanged(Breakpoint newBp, int newWidth);

    // Emitted on every (debounced) resize, even within the same breakpoint.
    // Useful for sub-breakpoint rules (e.g. sensory 3-up/2-up/1-up).
    void widthChanged(int newWidth);

protected:
    bool eventFilter(QObject* watched, QEvent* event) override;

private:
    ResponsiveLayout();
    void recompute(int width);

    QWidget*   m_window = nullptr;
    QTimer*    m_debounce = nullptr;
    int        m_lastWidth = 0;
    int        m_pendingWidth = 0;
    Breakpoint m_breakpoint = Standard;
};

} // namespace DVE
