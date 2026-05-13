// -- DataViewer Enterprise -- ConnectionMonitor Integration Tests --
//
// Verifies the 30s-ping + jittered-reconnect connection watchdog used by
// MainWindow to drive the online/offline mode switch. Tests open a real
// PostgresConnection against the ephemeral test container (DVE_TEST_PG_CONN)
// and tune the timers via setIntervalsForTesting() so the full suite runs
// in well under 60s.
//
// To simulate a disconnect we call PostgresConnection::close() directly:
// the next ping() returns false (m_open is false), the monitor flips to
// offline, and the reconnect path then runs tryOpenWithRetry() against
// the still-running test container. This is more deterministic than
// `docker stop dve-test-pg` and keeps the suite hermetic to docker state.
//
// REQUIRES: DVE_TEST_PG_CONN set (the same ephemeral test Postgres used
// by tst_postgresconnection / tst_databasemanager). Skipped otherwise.

#include <QtTest>
#include <QSignalSpy>
#include <QCoreApplication>

#include "ConnectionMonitor.h"
#include "PostgresConnection.h"
#include "ConfigLoader.h"

namespace {

DVE::DbConfig configFromEnv() {
    DVE::DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    if (env.isEmpty()) return c;
    for (const QString& part : env.split(' ', Qt::SkipEmptyParts)) {
        const QStringList kv = part.split('=');
        if (kv.size() != 2) continue;
        const QString k = kv[0], v = kv[1];
        if      (k == "host")     c.host = v;
        else if (k == "port")     c.port = v.toInt();
        else if (k == "dbname")   c.database = v;
        else if (k == "user")     c.user = v;
        else if (k == "password") c.password = v;
    }
    return c;
}

} // anonymous

class TstConnectionMonitor : public QObject {
    Q_OBJECT

private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) {
            QSKIP("DVE_TEST_PG_CONN not set; ConnectionMonitor integration tests skipped");
        }
    }

    // -- T1: start() puts the monitor in online mode with PG reachable -----
    void testStartReportsOnline() {
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));

        DVE::ConnectionMonitor mon(&pg, configFromEnv());
        // Lower intervals so the test runs fast. Ping cadence 500ms is
        // enough to fire a few times in the wait below.
        mon.setIntervalsForTesting(500, 500, 200);

        QSignalSpy wentOff(&mon, &DVE::ConnectionMonitor::wentOffline);
        QSignalSpy cameOn(&mon, &DVE::ConnectionMonitor::cameOnline);

        mon.start();
        QVERIFY(mon.isOnline());

        // Let the ping timer fire at least once. With PG reachable this
        // should NOT produce a wentOffline signal.
        QTest::qWait(800);

        QCOMPARE(wentOff.count(), 0);
        QCOMPARE(cameOn.count(),  0);
        QVERIFY(mon.isOnline());

        mon.stop();
        pg.close();
    }

    // -- T2: a failed ping flips the monitor to offline + emits signal ----
    void testPingDetectsOffline() {
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));

        DVE::ConnectionMonitor mon(&pg, configFromEnv());
        // Fast ping; reconnect base set unreachably high so the reconnect
        // timer cannot accidentally fire during this test and bring us
        // back online (which would race the QSignalSpy assertion below).
        mon.setIntervalsForTesting(200, 60000, 100);

        QSignalSpy wentOff(&mon, &DVE::ConnectionMonitor::wentOffline);

        mon.start();
        QVERIFY(mon.isOnline());

        // Slam the connection shut so the next ping() returns false.
        pg.close();

        // Within a few ping cycles the monitor must drop to offline.
        QVERIFY(QTest::qWaitFor([&]() { return wentOff.count() >= 1; }, 5000));
        QCOMPARE(wentOff.count(), 1);
        QVERIFY(!mon.isOnline());

        mon.stop();
    }

    // -- T3: after a failed ping, a successful reconnect flips back online -
    void testReconnectComesBackOnline() {
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));

        DVE::ConnectionMonitor mon(&pg, configFromEnv());
        // Short ping + short reconnect base so the offline -> online round
        // trip completes well within the 60s suite budget.
        mon.setIntervalsForTesting(200, 500, 200);

        QSignalSpy wentOff(&mon, &DVE::ConnectionMonitor::wentOffline);
        QSignalSpy cameOn(&mon, &DVE::ConnectionMonitor::cameOnline);

        mon.start();
        QVERIFY(mon.isOnline());

        // Trigger offline.
        pg.close();
        QVERIFY(QTest::qWaitFor([&]() { return wentOff.count() >= 1; }, 5000));
        QVERIFY(!mon.isOnline());

        // Postgres is still running on port 5433. The reconnect timer will
        // call tryOpenWithRetry() which opens a fresh QSqlDatabase against
        // the live server -- success.
        QVERIFY(QTest::qWaitFor([&]() { return cameOn.count() >= 1; }, 10000));
        QCOMPARE(cameOn.count(), 1);
        QVERIFY(mon.isOnline());

        mon.stop();
        pg.close();
    }

    // -- T4: stop() halts both timers; no further signals fire ------------
    void testStopHaltsTimers() {
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));

        DVE::ConnectionMonitor mon(&pg, configFromEnv());
        mon.setIntervalsForTesting(100, 100, 50);

        QSignalSpy wentOff(&mon, &DVE::ConnectionMonitor::wentOffline);
        QSignalSpy cameOn(&mon, &DVE::ConnectionMonitor::cameOnline);

        mon.start();
        mon.stop();

        // Now slam PG shut so a ping WOULD have detected offline if the
        // timer were still running.
        pg.close();

        QTest::qWait(800);  // multiple would-be ping cycles
        QCOMPARE(wentOff.count(), 0);
        QCOMPARE(cameOn.count(),  0);
    }

    // -- T5: nextReconnectDelayMs() is bounded to [base, base+cap) --------
    //
    // Drives the monitor offline so the reconnect path runs; captures the
    // delay each iteration via debugLastReconnectDelay() and asserts the
    // bound. The reconnect attempts will all fail because we point at a
    // bad port -- this lets us collect several jitter samples without
    // bringing the monitor back online mid-test.
    void testJitterIsBounded() {
        // Open then close so isOpen() is false (so the first ping fails).
        // The bad-port config below makes tryOpenWithRetry() fail with a
        // short timeout, leaving the monitor in offline mode.
        DVE::PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));
        pg.close();

        DVE::DbConfig badCfg = configFromEnv();
        badCfg.port = 1;   // nothing listens here; tryOpenWithRetry() fails fast

        DVE::ConnectionMonitor mon(&pg, badCfg);

        // Short ping (fails on first tick since pg is closed -> wentOffline
        // -> reconnect timer schedules with jitter).
        // Reconnect base 300ms, jitter cap 200ms -> [300, 500).
        constexpr int kBase = 300;
        constexpr int kCap  = 200;
        mon.setIntervalsForTesting(100, kBase, kCap);

        mon.start();

        // Sample debugLastReconnectDelay() periodically and assert the bound.
        // Each reconnect attempt blocks for ~3000ms in tryOpenWithRetry's
        // internal retry loop, so we may only get one or two distinct samples
        // in the test's wall time -- but each one tests the spec.
        QVector<int> samples;
        for (int i = 0; i < 20; ++i) {
            QTest::qWait(200);
            const int d = mon.debugLastReconnectDelay();
            if (d >= 0 && (samples.isEmpty() || samples.last() != d)) {
                samples.append(d);
                if (samples.size() >= 3) break;
            }
        }

        // At minimum the first scheduled reconnect must have computed a
        // delay (called inside switchToOfflineMode()).
        QVERIFY2(!samples.isEmpty(),
                 "debugLastReconnectDelay() never produced a value; reconnect "
                 "scheduling did not fire");

        for (int d : samples) {
            QVERIFY2(d >= kBase && d < kBase + kCap,
                     qPrintable(QString("reconnect delay %1 out of bounds [%2, %3)")
                                    .arg(d).arg(kBase).arg(kBase + kCap)));
        }

        mon.stop();
    }
};

QTEST_MAIN(TstConnectionMonitor)
#include "tst_connectionmonitor.moc"
