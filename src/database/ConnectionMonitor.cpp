#include "ConnectionMonitor.h"

#include <QRandomGenerator>

#include "PostgresConnection.h"

namespace DVE {

ConnectionMonitor::ConnectionMonitor(PostgresConnection* conn,
                                     const DbConfig& cfg,
                                     QObject* parent)
    : QObject(parent),
      m_conn(conn),
      m_cfg(cfg),
      m_pingTimer(this),
      m_reconnectTimer(this) {
    // Repeating ping; the timeout cadence is fixed in online mode.
    m_pingTimer.setSingleShot(false);

    // Single-shot reconnect: re-armed in onReconnectAttempt() with a
    // freshly-jittered interval after each failed attempt.
    m_reconnectTimer.setSingleShot(true);

    connect(&m_pingTimer,      &QTimer::timeout,
            this, &ConnectionMonitor::onPing);
    connect(&m_reconnectTimer, &QTimer::timeout,
            this, &ConnectionMonitor::onReconnectAttempt);
}

ConnectionMonitor::~ConnectionMonitor() = default;

void ConnectionMonitor::start() {
    // Caller chose to begin monitoring -- no signal emitted, the world
    // outside is already assumed to be in "online" mode at start time.
    switchToOnlineMode();
}

void ConnectionMonitor::stop() {
    m_pingTimer.stop();
    m_reconnectTimer.stop();
}

void ConnectionMonitor::forceReconnect() {
    // Cancel whatever's pending — we're running the attempt synchronously
    // on this call instead of letting the timer fire.
    m_pingTimer.stop();
    m_reconnectTimer.stop();
    // onReconnectAttempt() emits cameOnline() on success or re-arms the
    // reconnect timer with a fresh jittered delay on failure.
    onReconnectAttempt();
}

void ConnectionMonitor::setIntervalsForTesting(int pingMs,
                                               int reconnectBaseMs,
                                               int reconnectJitterCapMs) {
    m_pingMs               = pingMs;
    m_reconnectBaseMs      = reconnectBaseMs;
    m_reconnectJitterCapMs = reconnectJitterCapMs;

    // If the ping timer is already running, restart it with the new
    // cadence so the change takes effect immediately.
    if (m_pingTimer.isActive()) {
        m_pingTimer.start(m_pingMs);
    }
    // No need to restart the reconnect timer here -- its single-shot
    // re-arm in onReconnectAttempt() will pick up the new values on the
    // next iteration.
}

void ConnectionMonitor::onPing() {
    if (!m_conn) return;
    // v2.4.2 R4b: probe BOTH sockets. A half-open NOTIFY socket (the LISTEN
    // socket is dead but the query socket still answers SELECT 1) silently
    // stopped live updates on flaky / GFW networks because we only pinged the
    // query socket. Requiring both healthy means a dead listen socket also
    // flips offline -> the reconnect path reopens both sockets, re-subscribes,
    // and runs catch-up.
    if (m_conn->ping() && m_conn->pingListen()) {
        // Still online. Nothing to do.
        return;
    }
    // Ping failed -> drop to offline mode and notify listeners.
    switchToOfflineMode();
    emit wentOffline();
}

void ConnectionMonitor::onReconnectAttempt() {
    if (!m_conn) {
        // Re-arm with a fresh delay anyway so we keep trying when conn
        // becomes valid again. In practice this path shouldn't hit
        // unless the caller is misusing the class.
        m_reconnectTimer.start(nextReconnectDelayMs());
        return;
    }

    if (m_conn->tryOpenWithRetry(m_cfg, 3000)) {
        switchToOnlineMode();
        emit cameOnline();
        return;
    }

    // Still offline -- re-arm with a freshly-jittered delay.
    m_reconnectTimer.start(nextReconnectDelayMs());
}

void ConnectionMonitor::switchToOnlineMode() {
    m_reconnectTimer.stop();
    m_pingTimer.start(m_pingMs);
    m_online = true;
}

void ConnectionMonitor::switchToOfflineMode() {
    m_pingTimer.stop();
    m_reconnectTimer.start(nextReconnectDelayMs());
    m_online = false;
}

int ConnectionMonitor::nextReconnectDelayMs() {
    const int jitter = (m_reconnectJitterCapMs > 0)
        ? QRandomGenerator::global()->bounded(m_reconnectJitterCapMs)
        : 0;
    m_lastReconnectDelayMs = m_reconnectBaseMs + jitter;
    return m_lastReconnectDelayMs;
}

} // namespace DVE
