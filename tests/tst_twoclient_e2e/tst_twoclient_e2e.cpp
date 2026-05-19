// -- DataViewer Enterprise -- Two-Client E2E Tests (Postgres) --
//
// Phase 8 Task 23 of the Plan B Postgres multi-user migration. This suite
// spins up TWO PostgresConnections (each with its own DatabaseManager,
// IdentityManager, and NotificationListener) inside a SINGLE process, all
// pointed at the same ephemeral Postgres reachable via DVE_TEST_PG_CONN.
// It exercises the full Plan B stack end-to-end:
//
//   * NOTIFY round-trip for INSERT/UPDATE/DELETE on `files`
//     and `sensory_sessions`.
//   * NOTIFY round-trip for INSERT/DELETE on `presence`.
//   * Optimistic-concurrency producing WriteResult::VersionMismatch when
//     two clients race on the same row.
//   * WriteResult::RowDeleted when one client removes a row the other
//     client still has a stale id for.
//   * WriteResult::UniqueViolation on duplicate file_path / duplicate
//     sensory natural-key (session_name, tester_name, date).
//   * The own-UUID filter is APP-level (in MainWindow); the
//     NotificationListener layer surfaces every NOTIFY regardless of who
//     wrote. This suite asserts that, so future maintainers don't move
//     the filter into the listener.
//
// IDENTITY-SETUP WORKAROUND
// =========================
// IdentityManager reads its UUID from `QSettings()` (no-arg), which is
// keyed by QCoreApplication::organizationName/applicationName. To get
// two distinct identities in one process we toggle the application name
// to "ClientA" / "ClientB" before each IdentityManager construction or
// DB write. The helper class `ClientScope` is an RAII wrapper for that
// toggle: while a ClientScope is alive, QCoreApplication's app name
// matches the scope's client, so any IdentityManager::uuid() call (and
// therefore any DatabaseManager::tryWrite*) reads the right UUID.
//
// REQUIRES DVE_TEST_PG_CONN, otherwise every slot is skipped.

#include <QtTest/QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QSignalSpy>
#include <QSettings>
#include <QCoreApplication>
#include <QUuid>

#include "DatabaseManager.h"
#include "PostgresConnection.h"
#include "NotificationListener.h"
#include "PresenceManager.h"
#include "IdentityManager.h"
#include "ConfigLoader.h"
#include "LiveSync.h"
#include "ReportData.h"
#include "SensoryData.h"

#include <QElapsedTimer>

using namespace DVE;

namespace {

// Parse DVE_TEST_PG_CONN (space-separated key=value pairs) into a DbConfig.
DbConfig pgConfig() {
    DbConfig c;
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

// RAII helper: while alive, QCoreApplication::applicationName() is set to
// `app`, so any IdentityManager constructed (or any settings read by an
// existing manager) inside the scope sees the right UUID storage.
//
// Save the previous name on entry, restore on destruction.
class ClientScope {
public:
    explicit ClientScope(const QString& app) {
        m_prev = QCoreApplication::applicationName();
        QCoreApplication::setApplicationName(app);
    }
    ~ClientScope() { QCoreApplication::setApplicationName(m_prev); }
private:
    QString m_prev;
};

// Minimal FileResult builder reused across tests.
FileResult makeFileResult(const QString& fileName = "test.xlsx",
                          const QString& filePath = "/tmp/test.xlsx") {
    FileResult fr;
    fr.filePath        = filePath;
    fr.fileName        = fileName;
    fr.templateVersion = "new";

    SheetResult sheet;
    sheet.sheetName       = "Lifetime Test";
    sheet.templateVersion = "new";
    sheet.columnHeaders   << "puffs" << "beforeWeight" << "afterWeight";
    sheet.overallAvgTPM   = 3.5;

    SampleResult sample;
    sample.sampleName = "Sample 1";
    sample.sampleID   = "S-1";
    sample.date       = "2026-01-01";
    sample.tester     = "QA";
    sample.media      = "E-liquid";
    sample.resistance = 1.1;
    sample.voltage    = 3.0;
    sample.power      = 8.18;
    sample.averageTPM = 3.5;

    DataRow row;
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

// Minimal SensorySession builder.
SensorySession makeSensorySession(const QString& sessionName = "S1",
                                  const QString& tester = "T1",
                                  const QString& date = "2026-01-01") {
    SensorySession s;
    s.sessionName  = sessionName;
    s.testTitle    = "E2E Test";
    s.testerName   = tester;
    s.assessorName = tester;
    s.media        = "Sample A";
    s.date         = date;
    s.timestamp    = date + "T10:00:00Z";

    SensorySample sample;
    sample.name = "Sample 1";
    for (const QString& m : kSensoryMetrics) sample.scores[m] = 5.0;
    s.samples.append(sample);
    return s;
}

} // anonymous namespace

class TstTwoClientE2E : public QObject {
    Q_OBJECT

private:
    // Per-test handles. Each test constructs its own pair of clients via
    // setupClients() to keep slots independent and avoid state leakage.
    struct Client {
        PostgresConnection*  pg = nullptr;
        DatabaseManager*     db = nullptr;
        IdentityManager*     id = nullptr;
        NotificationListener* nl = nullptr;
        LiveSync*            ls = nullptr;
        QSignalSpy*          cellChangedSpy = nullptr;
        QString              uuid;   // expected UUID string (no braces)
        QString              appName;
    };

    // Build a fresh client. The app-name toggle ensures IdentityManager
    // reads its UUID from the right QSettings storage.
    Client buildClient(const QString& appName, const QString& displayName,
                       const QString& color, bool withListener) {
        Client c;
        c.appName = appName;
        ClientScope scope(appName);

        // Clear any stale settings under this client and pre-seed an
        // identity so we KNOW which UUID it will return.
        {
            QSettings s;
            s.clear();
        }

        c.pg = new PostgresConnection;
        if (!c.pg->open(pgConfig())) {
            qWarning() << "buildClient open failed:" << c.pg->lastError();
            return c;
        }
        c.id = new IdentityManager;
        c.id->setDisplayName(displayName);
        c.id->setColor(color);
        // QSettings creation order: ensureUuid() in IdentityManager ctor
        // already wrote a fresh UUID under this app's storage; cache it.
        c.uuid = c.id->uuid().toString(QUuid::WithoutBraces);

        c.db = new DatabaseManager;
        if (!c.db->open(pgConfig(), c.id)) {
            qWarning() << "buildClient db open failed:" << c.db->lastError();
        }

        if (withListener) {
            c.nl = new NotificationListener(c.pg);
            if (!c.nl->subscribe()) {
                qWarning() << "buildClient subscribe failed";
            }
        }
        return c;
    }

    // Attach a LiveSync (and optional cellChanged spy) to a client built by
    // buildClient(). Wires the listener's rowChanged signal into LiveSync's
    // onRowChanged slot so remote NOTIFYs surface as cellChanged emissions.
    // The own-UUID echo filter is APP-level (MainWindow); we don't filter
    // here so the live-sync test sees the full unfiltered stream.
    void attachLiveSync(Client& c, bool withSpy) {
        ClientScope scope(c.appName);
        c.ls = new LiveSync(c.pg, c.id, /*parent=*/nullptr);
        if (c.nl) {
            QObject::connect(c.nl, &NotificationListener::rowChanged,
                             c.ls, &LiveSync::onRowChanged);
        }
        if (withSpy) {
            c.cellChangedSpy = new QSignalSpy(c.ls, &LiveSync::cellChanged);
        }
    }

    void destroyClient(Client& c) {
        delete c.cellChangedSpy; c.cellChangedSpy = nullptr;
        delete c.ls;  c.ls = nullptr;
        delete c.nl;  c.nl = nullptr;
        if (c.db) { c.db->close(); delete c.db; c.db = nullptr; }
        delete c.id; c.id = nullptr;
        if (c.pg) { c.pg->close(); delete c.pg; c.pg = nullptr; }
    }

private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
            QSKIP("DVE_TEST_PG_CONN not set; Postgres e2e tests skipped");
        QCoreApplication::setOrganizationName("DataViewerTest");
        QCoreApplication::setApplicationName("tst_twoclient_e2e");
    }

    // Per-slot setup: wipe every table the suite touches on a dedicated
    // cleanup connection so each test starts from a clean schema.
    void init() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) return;
        const QString cname = "tst_e2e_wipe";
        {
            QSqlDatabase wipe = QSqlDatabase::addDatabase("QPSQL", cname);
            const DbConfig cfg = pgConfig();
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
                         "presence", "settings"
                     }) {
                    q.exec("DELETE FROM " + t);
                }
                wipe.close();
            }
        }
        QSqlDatabase::removeDatabase(cname);
    }

    // ── 1. NOTIFY round-trip: INSERT on sensory_sessions reaches the other
    //       client's NotificationListener within 2s, with correct payload.
    void testNotifyRoundTrip_insert() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", true);
        QVERIFY(A.db && A.db->isOpen());
        QVERIFY(B.nl && B.nl->isSubscribed());

        QSignalSpy spy(B.nl, &NotificationListener::rowChanged);

        // Client A writes a sensory session (UUID-tagged via IdentityManager).
        {
            ClientScope scope(A.appName);
            SensorySession s = makeSensorySession("notify-ins", "T1", "2026-01-01");
            QCOMPARE(A.db->tryWriteSensorySession(s), WriteResult::Success);
        }

        QVERIFY(spy.wait(3000));
        QVERIFY(spy.count() >= 1);

        // Find the sensory_sessions INSERT among any emitted signals.
        bool found = false;
        for (int i = 0; i < spy.count(); ++i) {
            const RowChange c = spy.at(i).at(0).value<RowChange>();
            if (c.table == "sensory_sessions" && c.op == "INSERT") {
                QCOMPARE(c.updatedBy, A.uuid);
                QVERIFY(c.id > 0);
                found = true; break;
            }
        }
        QVERIFY2(found, "No sensory_sessions INSERT NOTIFY received");

        destroyClient(A);
        destroyClient(B);
    }

    // ── 2. NOTIFY round-trip: INSERT then UPDATE on files reaches B with
    //       two distinct ops. Verifies the bump_version trigger doesn't
    //       break NOTIFY (the AFTER ROW trigger fires post bump_version).
    void testNotifyRoundTrip_update() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", true);

        QSignalSpy spy(B.nl, &NotificationListener::rowChanged);

        // Step 1: A inserts a file.
        FileResult fr = makeFileResult("upd.xlsx", "/tmp/upd.xlsx");
        {
            ClientScope scope(A.appName);
            QCOMPARE(A.db->tryWriteFile(fr), WriteResult::Success);
        }

        QVERIFY(spy.wait(3000));
        bool sawInsert = false;
        for (int i = 0; i < spy.count(); ++i) {
            const RowChange c = spy.at(i).at(0).value<RowChange>();
            if (c.table == "files" && c.op == "INSERT") {
                sawInsert = true; break;
            }
        }
        QVERIFY2(sawInsert, "No files INSERT NOTIFY received");

        spy.clear();

        // Step 2: A reloads (to get id+version) then re-saves -> UPDATE.
        FileResult loaded;
        {
            ClientScope scope(A.appName);
            loaded = A.db->loadFileByPath("/tmp/upd.xlsx");
            QVERIFY(loaded.id > 0 && loaded.version > 0);
            loaded.sheets[0].samples[0].sampleName = "updated";
            QCOMPARE(A.db->tryWriteFile(loaded), WriteResult::Success);
        }

        QVERIFY(spy.wait(3000));
        bool sawUpdate = false;
        for (int i = 0; i < spy.count(); ++i) {
            const RowChange c = spy.at(i).at(0).value<RowChange>();
            if (c.table == "files" && c.op == "UPDATE" && c.id == loaded.id) {
                sawUpdate = true; break;
            }
        }
        QVERIFY2(sawUpdate, "No files UPDATE NOTIFY received");

        destroyClient(A);
        destroyClient(B);
    }

    // ── 3. NOTIFY round-trip: DELETE on files reaches B.
    void testNotifyRoundTrip_delete() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", true);

        // Step 1: A inserts a file.
        FileResult fr = makeFileResult("del.xlsx", "/tmp/del.xlsx");
        int fileId = -1;
        {
            ClientScope scope(A.appName);
            QCOMPARE(A.db->tryWriteFile(fr), WriteResult::Success);
            FileResult loaded = A.db->loadFileByPath("/tmp/del.xlsx");
            fileId = loaded.id;
            QVERIFY(fileId > 0);
        }

        QSignalSpy spy(B.nl, &NotificationListener::rowChanged);

        // Step 2: A deletes.
        {
            ClientScope scope(A.appName);
            QVERIFY(A.db->removeFile(fileId));
        }

        QVERIFY(spy.wait(3000));
        bool sawDelete = false;
        for (int i = 0; i < spy.count(); ++i) {
            const RowChange c = spy.at(i).at(0).value<RowChange>();
            if (c.table == "files" && c.op == "DELETE" && c.id == fileId) {
                sawDelete = true; break;
            }
        }
        QVERIFY2(sawDelete, "No files DELETE NOTIFY received");

        destroyClient(A);
        destroyClient(B);
    }

    // ── 4. Documentation test: confirm the NotificationListener layer does
    //       NOT filter out self-writes. The own-UUID filter is APP-level in
    //       MainWindow. If a future change moves filtering into the listener,
    //       this test breaks first.
    void testOwnUuidEcho_notFiltered_atListener() {
        // Same connection writes and listens. The listener uses listenDb()
        // (a separate QSqlDatabase from queryDb()) so the NOTIFY round-trips
        // through the Postgres server even within one process.
        Client A = buildClient("ClientA", "Alice", "#ef4444", true);

        QSignalSpy spy(A.nl, &NotificationListener::rowChanged);

        {
            ClientScope scope(A.appName);
            FileResult fr = makeFileResult("echo.xlsx", "/tmp/echo.xlsx");
            QCOMPARE(A.db->tryWriteFile(fr), WriteResult::Success);
        }

        QVERIFY(spy.wait(3000));
        bool sawOwnInsert = false;
        for (int i = 0; i < spy.count(); ++i) {
            const RowChange c = spy.at(i).at(0).value<RowChange>();
            if (c.table == "files" && c.op == "INSERT"
                && c.updatedBy == A.uuid) {
                sawOwnInsert = true; break;
            }
        }
        QVERIFY2(sawOwnInsert,
                 "NotificationListener should surface own-UUID writes "
                 "(filter is app-level, not listener-level)");

        destroyClient(A);
    }

    // ── 5. Optimistic conflict: VersionMismatch when two clients race on
    //       the same file row.
    void testOptimisticConflict_versionMismatch() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", false);

        // Seed: A inserts a file, both reload to get matching id+version.
        {
            ClientScope scope(A.appName);
            FileResult fr = makeFileResult("conflict.xlsx", "/tmp/conflict.xlsx");
            QCOMPARE(A.db->tryWriteFile(fr), WriteResult::Success);
        }

        FileResult frA, frB;
        {
            ClientScope scope(A.appName);
            frA = A.db->loadFileByPath("/tmp/conflict.xlsx");
            QVERIFY(frA.id > 0 && frA.version > 0);
        }
        {
            ClientScope scope(B.appName);
            frB = B.db->loadFileByPath("/tmp/conflict.xlsx");
            QVERIFY(frB.id == frA.id);
            QCOMPARE(frB.version, frA.version);
        }

        // Capture the shared pre-save version BEFORE A writes — tryWriteFile's
        // mutable overload now writes the post-save version back into frA,
        // so frA.version is no longer a stable baseline after the first save.
        const int sharedPreVer = frA.version;

        // A modifies and saves -> Success, server version bumps.
        {
            ClientScope scope(A.appName);
            frA.sheets[0].samples[0].sampleName = "A wins";
            QCOMPARE(A.db->tryWriteFile(frA), WriteResult::Success);
        }

        // B modifies and saves with stale version -> VersionMismatch.
        {
            ClientScope scope(B.appName);
            frB.sheets[0].samples[0].sampleName = "B loses";
            QCOMPARE(B.db->tryWriteFile(frB), WriteResult::VersionMismatch);
        }

        // B re-loads, retries -> Success.
        {
            ClientScope scope(B.appName);
            frB = B.db->loadFileByPath("/tmp/conflict.xlsx");
            QVERIFY(frB.version > sharedPreVer);
            frB.sheets[0].samples[0].sampleName = "B retries";
            QCOMPARE(B.db->tryWriteFile(frB), WriteResult::Success);
        }

        destroyClient(A);
        destroyClient(B);
    }

    // ── 6. Optimistic conflict: RowDeleted when one client removes while
    //       the other holds a stale id.
    void testOptimisticConflict_rowDeleted() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", false);

        // A inserts a sensory session. Both clients load it.
        SensorySession seedA;
        {
            ClientScope scope(A.appName);
            seedA = makeSensorySession("rd-sess", "T1", "2026-02-01");
            QVERIFY(A.db->saveSensorySession(seedA));
            QVERIFY(seedA.id > 0);
        }
        SensorySession sessB;
        {
            ClientScope scope(B.appName);
            sessB = B.db->loadSensorySession(seedA.id);
            QCOMPARE(sessB.id, seedA.id);
        }

        // B removes the session.
        {
            ClientScope scope(B.appName);
            QVERIFY(B.db->removeSensorySession(seedA.id));
        }

        // A modifies its stale-id copy and saves -> RowDeleted.
        {
            ClientScope scope(A.appName);
            seedA.samples[0].scores["Overall Liking"] = 9.0;
            QCOMPARE(A.db->tryWriteSensorySession(seedA),
                     WriteResult::RowDeleted);
        }

        destroyClient(A);
        destroyClient(B);
    }

    // ── 7. UniqueViolation on duplicate file_path across clients.
    void testOptimisticConflict_uniqueViolation_files() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", false);

        {
            ClientScope scope(A.appName);
            FileResult frA = makeFileResult("uv.xlsx", "/tmp/uv.xlsx");
            QCOMPARE(A.db->tryWriteFile(frA), WriteResult::Success);
        }

        {
            ClientScope scope(B.appName);
            // Fresh struct (id=-1) with same file_path -> INSERT collides.
            FileResult frB = makeFileResult("uv-other.xlsx", "/tmp/uv.xlsx");
            QCOMPARE(B.db->tryWriteFile(frB), WriteResult::UniqueViolation);
        }

        destroyClient(A);
        destroyClient(B);
    }

    // ── 8. UniqueViolation on duplicate sensory natural key across clients.
    void testOptimisticConflict_uniqueViolation_sensory() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", false);

        {
            ClientScope scope(A.appName);
            SensorySession sA = makeSensorySession("S1", "T1", "2026-03-01");
            QCOMPARE(A.db->tryWriteSensorySession(sA), WriteResult::Success);
        }

        {
            ClientScope scope(B.appName);
            // Fresh struct with same (session_name, tester_name, date).
            SensorySession sB = makeSensorySession("S1", "T1", "2026-03-01");
            QCOMPARE(B.db->tryWriteSensorySession(sB),
                     WriteResult::UniqueViolation);
        }

        destroyClient(A);
        destroyClient(B);
    }

    // ── 9. Presence INSERT/DELETE round-trip via NotificationListener.
    //       Verifies the dataviewer_presence channel propagates between
    //       clients and that PresenceChange payload fields are correct.
    void testPresenceRoundTrip() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", true);

        QSignalSpy insertSpy(B.nl, &NotificationListener::presenceChanged);

        // A activates presence on resource (file, 42).
        QUuid aUuid;
        {
            ClientScope scope(A.appName);
            aUuid = A.id->uuid();
            PresenceManager pmA(A.pg, A.id);
            QVERIFY(pmA.activate("file", 42, "viewing"));

            // Verify B sees the INSERT, then deactivate inside the same
            // scope so pmA destructor doesn't fire deactivate() too late.
            QVERIFY(insertSpy.wait(3000));
            bool sawInsert = false;
            for (int i = 0; i < insertSpy.count(); ++i) {
                const PresenceChange p = insertSpy.at(i).at(0)
                                                  .value<PresenceChange>();
                if (p.op == "INSERT" && p.resourceType == "file"
                    && p.resourceId == 42 && p.userUuid == aUuid) {
                    QCOMPARE(p.intent, QString("viewing"));
                    sawInsert = true; break;
                }
            }
            QVERIFY2(sawInsert, "No presence INSERT NOTIFY received");

            QSignalSpy deleteSpy(B.nl, &NotificationListener::presenceChanged);
            QVERIFY(pmA.deactivate());

            QVERIFY(deleteSpy.wait(3000));
            bool sawDelete = false;
            for (int i = 0; i < deleteSpy.count(); ++i) {
                const PresenceChange p = deleteSpy.at(i).at(0)
                                                  .value<PresenceChange>();
                if (p.op == "DELETE" && p.resourceType == "file"
                    && p.resourceId == 42 && p.userUuid == aUuid) {
                    sawDelete = true; break;
                }
            }
            QVERIFY2(sawDelete, "No presence DELETE NOTIFY received");
        }

        destroyClient(A);
        destroyClient(B);
    }

    // ── 10. Sanity: two ClientScopes produce DIFFERENT UUIDs. If this ever
    //        regresses (e.g., a future refactor caches IdentityManager
    //        state across scopes), the rest of the suite would silently
    //        give both clients the same UUID and the conflict tests would
    //        produce false-positive passes.
    void testIdentityWorkaround_yieldsDistinctUuids() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", false);

        QVERIFY(!A.uuid.isEmpty());
        QVERIFY(!B.uuid.isEmpty());
        QVERIFY2(A.uuid != B.uuid,
                 qPrintable(QString("Identity workaround broken: both "
                                    "clients got UUID %1").arg(A.uuid)));

        destroyClient(A);
        destroyClient(B);
    }

    // ── 11. C3: a second tryWriteFile call after a single-cell in-memory edit
    //       must NOT delete-and-reinsert the untouched rows. With the v2.0.1
    //       DELETE-cascade-rebuild code this slot fails because the entire
    //       tests subtree is wiped on every save, so data_rows.id values
    //       change on every save — and any concurrent client referencing the
    //       old ids is now pointing at dead rows. The fix is the id-aware
    //       upsert in C3: UPDATE existing children by id+version, INSERT new
    //       children, post-prune rows present pre-save but absent post-save.
    void tryWriteFile_concurrentUsersDoNotWipeEachOther() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        QVERIFY(A.db && A.db->isOpen());

        // Step 1: seed a file with 1 sheet / 1 sample / 2 data_rows via the
        // production save path. Tag each row's notes so we can re-identify
        // them by content.
        {
            ClientScope scope(A.appName);
            FileResult fr = makeFileResult("c3.xlsx", "/tmp/c3.xlsx");
            fr.sheets[0].samples[0].rows[0].notes = "ROW1-ORIGINAL";
            DataRow r2;
            r2.puffs = 20.0;
            r2.beforeWeight = 30.0;
            r2.afterWeight  = 29.9;
            r2.tpm          = 1.0;
            r2.oilConsumed  = 0.1;
            r2.notes        = "ROW2-ORIGINAL";
            fr.sheets[0].samples[0].rows.append(r2);
            QCOMPARE(A.db->tryWriteFile(fr), WriteResult::Success);
        }

        // Step 2: read back the pre-image row ids directly.
        qint64 fileId = -1, sampleId = -1, r1Id = -1, r2Id = -1;
        {
            QSqlQuery q(A.pg->queryDb());
            QVERIFY(q.exec(
                "SELECT id FROM files WHERE file_path='/tmp/c3.xlsx'"));
            QVERIFY(q.next());
            fileId = q.value(0).toLongLong();
            QVERIFY(q.exec(QString(
                "SELECT s.id FROM samples s JOIN tests t ON s.test_id=t.id "
                "WHERE t.file_id=%1 ORDER BY s.sort_order LIMIT 1").arg(fileId)));
            QVERIFY(q.next());
            sampleId = q.value(0).toLongLong();
            QVERIFY(q.exec(QString(
                "SELECT id FROM data_rows WHERE sample_id=%1 "
                "ORDER BY sort_order").arg(sampleId)));
            QVERIFY(q.next()); r1Id = q.value(0).toLongLong();
            QVERIFY(q.next()); r2Id = q.value(0).toLongLong();
            QVERIFY(r1Id > 0 && r2Id > 0 && r1Id != r2Id);
        }

        // Step 3: A loads the file (gets id+version), edits row 1's notes
        // only, leaves row 2 untouched, and re-saves.
        {
            ClientScope scope(A.appName);
            FileResult loaded = A.db->loadFile(static_cast<int>(fileId));
            QCOMPARE(loaded.id, static_cast<int>(fileId));
            QCOMPARE(loaded.sheets.size(), 1);
            QCOMPARE(loaded.sheets[0].samples.size(), 1);
            QCOMPARE(loaded.sheets[0].samples[0].rows.size(), 2);
            loaded.sheets[0].samples[0].rows[0].notes = "ROW1-EDITED";
            QCOMPARE(A.db->tryWriteFile(loaded), WriteResult::Success);
        }

        // Step 4: assert ids are stable across the save (i.e., the rows were
        // UPDATEd in place, not DELETEd and re-INSERTed), and that row 1's
        // edit landed while row 2's original value is intact.
        {
            QSqlQuery q(A.pg->queryDb());
            QVERIFY(q.exec(QString(
                "SELECT id, notes FROM data_rows WHERE sample_id=%1 "
                "ORDER BY sort_order").arg(sampleId)));
            QVERIFY(q.next());
            const qint64 newR1Id    = q.value(0).toLongLong();
            const QString newR1Note = q.value(1).toString();
            QVERIFY(q.next());
            const qint64 newR2Id    = q.value(0).toLongLong();
            const QString newR2Note = q.value(1).toString();

            QCOMPARE(newR1Id, r1Id);
            QCOMPARE(newR2Id, r2Id);
            QCOMPARE(newR1Note, QStringLiteral("ROW1-EDITED"));
            QCOMPARE(newR2Note, QStringLiteral("ROW2-ORIGINAL"));
        }

        destroyClient(A);
    }

    // ── 12. Per-cell live sync round-trip (v2.0.1 end-to-end contract).
    //       Client A commits a cell via LiveSync; client B observes the
    //       value land in the DB and the cellChanged signal fire within
    //       1.5 s.
    void cellEditOnClientAReachesClientBLive() {
        Client A = buildClient("ClientA", "Alice", "#ef4444", false);
        Client B = buildClient("ClientB", "Bob",   "#3b82f6", true);
        QVERIFY(A.db && A.db->isOpen());
        QVERIFY(B.nl && B.nl->isSubscribed());

        attachLiveSync(A, /*withSpy=*/false);
        attachLiveSync(B, /*withSpy=*/true);

        // Seed a data_rows row directly (mirrors tst_livesync fixture).
        qint64 dataRowId = -1;
        {
            ClientScope scope(A.appName);
            QSqlQuery q(A.pg->queryDb());
            QVERIFY(q.exec(
                "INSERT INTO files(file_path, file_name, loaded_at) "
                "VALUES('/tmp/live-e2e.xlsx', 'live-e2e.xlsx', '2026-05-16') "
                "RETURNING id"));
            QVERIFY(q.next());
            const qint64 fileId = q.value(0).toLongLong();

            QVERIFY(q.exec(QString(
                "INSERT INTO tests(file_id, sheet_name) "
                "VALUES(%1, 'S1') RETURNING id").arg(fileId)));
            QVERIFY(q.next());
            const qint64 testId = q.value(0).toLongLong();

            QVERIFY(q.exec(QString(
                "INSERT INTO samples(test_id, sample_name) "
                "VALUES(%1, 'A') RETURNING id").arg(testId)));
            QVERIFY(q.next());
            const qint64 sampleId = q.value(0).toLongLong();

            QVERIFY(q.exec(QString(
                "INSERT INTO data_rows(sample_id, sort_order, draw_pressure) "
                "VALUES(%1, 0, 0.0) RETURNING id").arg(sampleId)));
            QVERIFY(q.next());
            dataRowId = q.value(0).toLongLong();
            QVERIFY(dataRowId > 0);
        }

        // Drain any NOTIFYs generated by the seed inserts so the spy only
        // captures the upcoming commitCell event.
        QCoreApplication::processEvents(QEventLoop::AllEvents, 200);
        B.cellChangedSpy->clear();

        // Client A: commit a cell via LiveSync.
        {
            ClientScope scope(A.appName);
            QVERIFY(A.ls->commitCell("data_rows", dataRowId,
                                     "draw_pressure", 3.14));
        }

        // Client B: poll up to 1.5 s for the value to land.
        QElapsedTimer timer; timer.start();
        bool seen = false;
        while (timer.elapsed() < 1500) {
            QCoreApplication::processEvents(QEventLoop::AllEvents, 50);
            QSqlQuery q(B.pg->queryDb());
            if (q.exec(QString("SELECT draw_pressure FROM data_rows WHERE id=%1")
                       .arg(dataRowId)) && q.next()
                && qAbs(q.value(0).toDouble() - 3.14) < 1e-6) {
                seen = true; break;
            }
        }
        QVERIFY2(seen, "Client B did not see updated value within 1.5s");

        // Drain any pending NOTIFYs and confirm the cellChanged signal fired.
        if (B.cellChangedSpy->count() == 0) {
            B.cellChangedSpy->wait(1500);
        }
        QVERIFY2(B.cellChangedSpy->count() >= 1,
                 "Client B did not receive cellChanged signal");

        destroyClient(A);
        destroyClient(B);
    }
};

QTEST_MAIN(TstTwoClientE2E)
#include "tst_twoclient_e2e.moc"
