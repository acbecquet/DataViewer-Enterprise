#pragma once

#include <QObject>
#include <QTimer>
#include "ConfigLoader.h"

namespace DVE {

class PostgresConnection;

// Watches a PostgresConnection for disconnects + reconnects.
// Owned by MainWindow; instantiated with the SEPARATE PostgresConnection
// MainWindow uses for NOTIFY/heartbeat (m_pgConn), NOT DatabaseManager's
// internal one. This keeps the failure-detection surface aligned with
// where the live-updates traffic actually flows.
//
// Behavior:
//   start() launches the 30s ping timer (online mode). On ping failure
//   the monitor flips to offline, emits wentOffline(), and switches to
//   the reconnect timer (~30s + 0-5s jitter). On a successful
//   tryOpenWithRetry() it emits cameOnline() and switches back to the
//   30s ping.
//
// Threading: lives on the GUI thread (timers are QObject children).
//
// Disconnect detection is one of three signals per the spec (section
// "Mid-session disconnect detection"):
//   1. NOTIFY socket dies -- owned by NotificationListener; not this class.
//   2. Query returns 08006 / 57P01 -- bubbled by DatabaseManager.
//   3. 30s SELECT 1 ping fails -- THIS class.
class ConnectionMonitor : public QObject {
    Q_OBJECT
public:
    ConnectionMonitor(PostgresConnection* conn,
                      const DbConfig& cfg,
                      QObject* parent = nullptr);
    ~ConnectionMonitor() override;

    void start();
    void stop();

    // v2.1.0: user-initiated retry. Cancels the pending reconnect/ping
    // timer and runs one reconnect attempt immediately. Used by the
    // offline-banner Retry button so clicking it produces a visible
    // attempt right now instead of waiting up to 30 s for the next timer.
    void forceReconnect();

    bool isOnline() const { return m_online; }

    // For testing -- overrides the 30s ping and ~30s reconnect cadences.
    // pingMs is the period of the online ping timer.
    // reconnectBaseMs is the base period of the offline reconnect timer
    // (0-jitter cap is applied on top -- see reconnectJitterCapMs).
    // reconnectJitterCapMs is the random extra delay added to each
    // reconnect attempt (uniform in [0, cap)).
    // Default: 30000 / 30000 / 5000 -- matches spec.
    void setIntervalsForTesting(int pingMs,
                                int reconnectBaseMs,
                                int reconnectJitterCapMs);

    // Test-only: returns the most recent value computed by
    // nextReconnectDelayMs(). Lets tests assert that jitter is bounded
    // ([base, base+cap)) without friending or subclassing. Returns -1
    // before the first reconnect attempt is scheduled.
    int debugLastReconnectDelay() const { return m_lastReconnectDelayMs; }

signals:
    void wentOffline();
    void cameOnline();

private slots:
    void onPing();
    void onReconnectAttempt();

protected:
    // Protected (not private) so tests can subclass for observation.
    int nextReconnectDelayMs();

private:
    void switchToOnlineMode();
    void switchToOfflineMode();

    PostgresConnection* m_conn;     // not owned
    DbConfig            m_cfg;
    QTimer              m_pingTimer;
    QTimer              m_reconnectTimer;
    bool                m_online = true;

    int m_pingMs               = 30000;
    int m_reconnectBaseMs      = 30000;
    int m_reconnectJitterCapMs = 5000;

    int m_lastReconnectDelayMs = -1;  // for debugLastReconnectDelay()
};

} // namespace DVE
