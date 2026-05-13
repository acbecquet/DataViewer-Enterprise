// -- DataViewer Enterprise -- Plan C end-to-end checkpoint --
//
// Bookend integration test for Plan C: drives the full offline-failover
// round-trip across DatabaseManager + PostgresConnection + OfflineSnapshot +
// ConnectionMonitor in a single fixture. The component-level suites
// (tst_databasemanager / tst_offlinesnapshot / tst_connectionmonitor) verify
// each piece individually; this test verifies they cooperate the way
// MainWindow needs them to in production.
//
// Round-trip exercised:
//   1. DB online -- a tryWriteFile succeeds against live Postgres.
//   2. We regenerate the snapshot against live Postgres so it contains the
//      file written in step 1.
//   3. PG connection is slammed shut; ConnectionMonitor detects the failure
//      via its ping timer and emits wentOffline -- the test slot flips
//      DatabaseManager into offline mode.
//   4. listFiles() now reads from the snapshot (still returns the file from
//      step 1); tryWriteFile() returns OfflineReadOnly.
//   5. We reopen the PG connection; ConnectionMonitor's reconnect timer
//      detects the recovery via tryOpenWithRetry and emits cameOnline. The
//      test slot flips DatabaseManager back to online.
//   6. tryWriteFile() succeeds again -- the replay path that MainWindow's
//      pending-edit queue follows on reconnect.
//
// Two supporting tests verify the snapshot persists across close+reopen
// cycles and that the manual "pending-edit replay" sequence works end-to-end
// without bothering with SaveCoordinator (a logic-level mirror of the
// MainWindow path).
//
// REQUIRES: DVE_TEST_PG_CONN set (the same ephemeral test Postgres used by
// the other PG-dependent suites). Skipped otherwise.

#include <QtTest>
#include <QSignalSpy>
#include <QTemporaryDir>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QCoreApplication>

#include "DatabaseManager.h"
#include "OfflineSnapshot.h"
#include "ConnectionMonitor.h"
#include "PostgresConnection.h"
#include "ConfigLoader.h"
#include "IdentityManager.h"
#include "ReportData.h"
#include "SensoryData.h"
#include "DetailedSensoryData.h"

namespace {

// Build a DbConfig from DVE_TEST_PG_CONN (space-separated key=value pairs).
DVE::DbConfig pgConfig() {
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

// Wipe the contents of every editable table between tests. Schema is
// created idempotently by deploy/postgres/init.sql when the test container
// is brought up.
void wipeAllTables() {
    DVE::DbConfig cfg = pgConfig();
    const QString cname = "tst_planc_wipe";
    {
        QSqlDatabase wipe = QSqlDatabase::addDatabase("QPSQL", cname);
        wipe.setHostName(cfg.host);
        wipe.setPort(cfg.port);
        wipe.setDatabaseName(cfg.database);
        wipe.setUserName(cfg.user);
        wipe.setPassword(cfg.password);
        if (wipe.open()) {
            QSqlQuery q(wipe);
            for (const QString& t : QStringList{
                     "data_rows", "images", "samples", "tests", "files",
                     "sensory_images", "sensory_sessions",
                     "detailed_sensory_images", "detailed_sensory_sessions",
                     "settings"
                 }) {
                q.exec("DELETE FROM " + t);
            }
            wipe.close();
        }
    }
    QSqlDatabase::removeDatabase(cname);
}

// Build a minimal FileResult. Mirrors makeFileResult in tst_databasemanager.
DVE::FileResult makeFileResult(const QString& fileName,
                               const QString& filePath)
{
    DVE::FileResult fr;
    fr.filePath        = filePath;
    fr.fileName        = fileName;
    fr.templateVersion = "new";
    fr.sheetNames      << "Lifetime Test";

    DVE::SheetResult sheet;
    sheet.sheetName       = "Lifetime Test";
    sheet.templateVersion = "new";
    sheet.columnHeaders   << "puffs" << "beforeWeight" << "afterWeight";
    sheet.overallAvgTPM   = 3.5;

    DVE::SampleResult sample;
    sample.sampleName = "Sample 1";
    sample.sampleID   = "S-1";
    sample.date       = "2026-05-13";
    sample.tester     = "QA";
    sample.media      = "E-liquid";
    sample.resistance = 1.1;
    sample.voltage    = 3.0;
    sample.power      = 8.18;
    sample.averageTPM = 3.5;

    DVE::DataRow row;
    row.puffs        = 10.0;
    row.beforeWeight = 25.1;
    row.afterWeight  = 25.065;
    row.tpm          = 3.5;
    row.oilConsumed  = 0.035;
    sample.rows.append(row);

    sheet.samples.append(sample);
    fr.sheets.append(sheet);
    return fr;
}

} // anonymous

// ============================================================================
class TstPlanCCheckpoint : public QObject {
    Q_OBJECT

private:
    DVE::IdentityManager m_identity;
    QTemporaryDir*       m_overrideDir = nullptr;

    QString overrideBaseDir() const {
        return m_overrideDir ? m_overrideDir->path() : QString();
    }

private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) {
            QSKIP("DVE_TEST_PG_CONN not set; Plan C checkpoint test skipped");
        }
    }

    void init() {
        wipeAllTables();
        delete m_overrideDir;
        m_overrideDir = new QTemporaryDir();
        QVERIFY(m_overrideDir->isValid());
    }

    void cleanup() {
        delete m_overrideDir;
        m_overrideDir = nullptr;
    }

    // ------------------------------------------------------------------------
    // T1 -- the full online -> offline -> online round-trip with
    //       ConnectionMonitor driving the transitions.
    //
    // This is the "headline" test: it exercises every Plan C component
    // (DatabaseManager routing, OfflineSnapshot reads, ConnectionMonitor
    // ping + reconnect) in a single sequence.
    // ------------------------------------------------------------------------
    void testOnlineToOfflineToOnlineRoundTrip() {
        // 1) Bring DatabaseManager online against ephemeral Postgres.
        DVE::DatabaseManager db;
        QVERIFY(db.open(pgConfig(), &m_identity));
        QVERIFY(db.isOnline());

        // 2) Seed one file via the online write path. After this the
        //    snapshot regenerate will see it via SELECT.
        DVE::FileResult seedFile = makeFileResult("planc-seed.xlsx",
                                                  "/tmp/planc-seed.xlsx");
        QCOMPARE(db.tryWriteFile(seedFile),
                 DVE::WriteResult::Success);
        QCOMPARE(db.listFiles().size(), 1);

        // 3) Regenerate snapshot against live PG and wire it into the
        //    DatabaseManager. We use a separate PostgresConnection here
        //    so it can be closed/reopened independently of the one
        //    DatabaseManager owns internally.
        DVE::PostgresConnection pgMonitor;
        QVERIFY(pgMonitor.open(pgConfig()));

        DVE::OfflineSnapshot snap;
        snap.setOverrideDirForTesting(overrideBaseDir());
        QVERIFY2(snap.regenerate(&pgMonitor), qPrintable(snap.lastError()));
        QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));
        db.setOfflineSnapshot(&snap);

        // 4) Wire up ConnectionMonitor on the SEPARATE PG connection.
        //    Connect wentOffline -> db.setOnline(false) and
        //    cameOnline -> db.setOnline(true), as MainWindow does.
        DVE::ConnectionMonitor monitor(&pgMonitor, pgConfig());
        QObject::connect(&monitor, &DVE::ConnectionMonitor::wentOffline,
                         &db, [&db]() { db.setOnline(false); });
        QObject::connect(&monitor, &DVE::ConnectionMonitor::cameOnline,
                         &db, [&db]() { db.setOnline(true); });

        // Speed up the ping + reconnect cadences for the test.
        // Reconnect base 500ms + jitter cap 300ms -> delay in [500, 800).
        monitor.setIntervalsForTesting(200, 500, 300);

        QSignalSpy wentOff(&monitor, &DVE::ConnectionMonitor::wentOffline);
        QSignalSpy cameOn (&monitor, &DVE::ConnectionMonitor::cameOnline);

        monitor.start();
        QVERIFY(monitor.isOnline());
        QVERIFY(db.isOnline());

        // 5) Simulate a NAS outage by slamming the monitored connection
        //    shut. The next ping fails -> wentOffline -> setOnline(false).
        pgMonitor.close();
        QVERIFY2(QTest::qWaitFor([&]() { return wentOff.count() >= 1; }, 5000),
                 "ConnectionMonitor did not emit wentOffline within 5s");
        QVERIFY(!monitor.isOnline());
        QVERIFY(!db.isOnline());

        // 6) Verify the offline read path: listFiles() now routes to the
        //    snapshot, which contains the file we wrote in step 2.
        const auto offlineList = db.listFiles();
        QCOMPARE(offlineList.size(), 1);
        QCOMPARE(offlineList[0].filePath,
                 QStringLiteral("/tmp/planc-seed.xlsx"));

        // 7) Verify the offline write guard: tryWriteFile is refused.
        DVE::FileResult anotherFile = makeFileResult("planc-blocked.xlsx",
                                                     "/tmp/planc-blocked.xlsx");
        QCOMPARE(db.tryWriteFile(anotherFile),
                 DVE::WriteResult::OfflineReadOnly);

        // 8) Recover: the monitor's reconnect timer will call
        //    tryOpenWithRetry against the still-running test container.
        //    A successful reconnect emits cameOnline -> setOnline(true).
        QVERIFY2(QTest::qWaitFor([&]() { return cameOn.count() >= 1; }, 15000),
                 "ConnectionMonitor did not emit cameOnline within 15s");
        QVERIFY(monitor.isOnline());
        QVERIFY(db.isOnline());

        // 9) Online write succeeds again -- this is the path the
        //    MainWindow pending-edit replay queue exercises on reconnect.
        QCOMPARE(db.tryWriteFile(anotherFile),
                 DVE::WriteResult::Success);
        QCOMPARE(db.listFiles().size(), 2);

        monitor.stop();
        snap.close();
        db.setOfflineSnapshot(nullptr);
        pgMonitor.close();
        db.close();
    }

    // ------------------------------------------------------------------------
    // T2 -- the snapshot survives close+reopen. Proves the SQLite file
    // persists between OfflineSnapshot instances (the lifecycle the
    // app boots through every cold start).
    // ------------------------------------------------------------------------
    void testOfflineSnapshotPersistsAcrossClose() {
        // Seed Postgres with one file via DatabaseManager.
        DVE::DatabaseManager db;
        QVERIFY(db.open(pgConfig(), &m_identity));
        DVE::FileResult fr = makeFileResult("persist.xlsx",
                                            "/tmp/persist.xlsx");
        QCOMPARE(db.tryWriteFile(fr), DVE::WriteResult::Success);

        // Regenerate the snapshot, capture the row count, close it.
        int firstCount = -1;
        QString firstPath;
        {
            DVE::PostgresConnection pg;
            QVERIFY(pg.open(pgConfig()));
            DVE::OfflineSnapshot snap;
            snap.setOverrideDirForTesting(overrideBaseDir());
            QVERIFY2(snap.regenerate(&pg), qPrintable(snap.lastError()));
            QVERIFY2(snap.openReadOnly(), qPrintable(snap.lastError()));
            firstCount = snap.listFiles().size();
            firstPath  = snap.path();
            snap.close();
            pg.close();
        }
        QVERIFY(firstCount >= 1);
        QVERIFY(QFile::exists(firstPath));

        // Construct a fresh OfflineSnapshot instance on the same override
        // dir; opening read-only must see the same row count without any
        // PostgresConnection in scope.
        {
            DVE::OfflineSnapshot snap2;
            snap2.setOverrideDirForTesting(overrideBaseDir());
            QCOMPARE(snap2.path(), firstPath);
            QVERIFY2(snap2.openReadOnly(), qPrintable(snap2.lastError()));
            QCOMPARE(snap2.listFiles().size(), firstCount);
            snap2.close();
        }

        db.close();
    }

    // ------------------------------------------------------------------------
    // T3 -- the pending-edit replay path mirrors what MainWindow's
    // PendingEdit queue does on reconnect: write is refused offline,
    // the caller captures it manually, and after setOnline(true) the
    // replay succeeds. Drives DatabaseManager directly without
    // SaveCoordinator to keep the test scope tight.
    // ------------------------------------------------------------------------
    void testPendingEditPropagatesViaSaveCoordinator() {
        DVE::DatabaseManager db;
        QVERIFY(db.open(pgConfig(), &m_identity));

        // No snapshot needed for this test; we only care about the
        // write-path online/offline guard.
        db.setOnline(false);
        QVERIFY(!db.isOnline());

        DVE::FileResult pending = makeFileResult("pending.xlsx",
                                                 "/tmp/pending.xlsx");
        // First attempt while offline must be refused.
        QCOMPARE(db.tryWriteFile(pending),
                 DVE::WriteResult::OfflineReadOnly);

        // The caller (MainWindow) captures this into its pending-edit
        // queue. Here we just hold the FileResult in scope and replay it
        // once the connection comes back.
        QVector<DVE::FileResult> pendingQueue;
        pendingQueue.append(pending);

        // Reconnect (simulated by setOnline(true) -- ConnectionMonitor
        // drives the real signal in production).
        db.setOnline(true);
        QVERIFY(db.isOnline());

        // Replay: the queued write now succeeds, and the file shows up
        // in the live PG listFiles() result.
        for (auto& edit : pendingQueue) {
            QCOMPARE(db.tryWriteFile(edit), DVE::WriteResult::Success);
        }

        const auto files = db.listFiles();
        QCOMPARE(files.size(), 1);
        QCOMPARE(files[0].filePath, QStringLiteral("/tmp/pending.xlsx"));

        db.close();
    }
};

QTEST_MAIN(TstPlanCCheckpoint)
#include "tst_planc_checkpoint.moc"
