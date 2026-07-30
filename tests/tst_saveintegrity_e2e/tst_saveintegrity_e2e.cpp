// -- DataViewer Enterprise -- Save-Integrity E2E Harness (Postgres) --
//
// v2.5.0 Task 8 (F7). The user asked for "your own test harness to test and
// validate before handing over". This suite chains the REAL components the way
// the v2.4.0 production failures chained them -- DatabaseManager whole-session/
// file writes + the live LiveSync per-cell worker against the ephemeral
// `dve-test-pg` container -- and asserts the v2.5.0 fixes hold end to end.
//
// Each scenario reproduces a PRODUCTION failure documented in
// docs/superpowers/specs/2026-06-10-v24-save-sync-regressions-evidence.md and
// cites the matching root cause (RCn / Fn) in a comment. Unlike the focused
// unit suites (tst_databasemanager, tst_livesync, tst_twoclient_e2e), this one
// deliberately mixes the two write paths in a single session so a regression in
// the merge/dirty-cell arbitration -- the exact thing that silently reverted
// the user's edits "the next day" -- surfaces here even if every component
// passes in isolation.
//
// REQUIRES DVE_TEST_PG_CONN, otherwise every slot is skipped.

#include <QtTest/QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>
#include <QSignalSpy>
#include <QSettings>
#include <QCoreApplication>
#include <QSet>
#include <QString>

#include "DatabaseManager.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "ConfigLoader.h"
#include "LiveSync.h"
#include "ReportData.h"
#include "SensoryData.h"
#include "DetailedSensoryData.h"
#include "OutputPaths.h"

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

// A sensory session with all five metric scores defaulted to 5.0.
SensorySession makeSensorySession(const QString& sessionName,
                                  const QString& tester,
                                  const QString& date,
                                  const QString& title = QString()) {
    SensorySession s;
    s.sessionName  = sessionName;
    s.testTitle    = title.isEmpty() ? sessionName : title;
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

// Minimal one-sheet/one-sample/one-row FileResult. `marker` lands in the sample
// name so callers can tell which version a loadFileByPath returned.
FileResult makeFileResult(const QString& filePath, const QString& marker) {
    FileResult fr;
    fr.filePath        = filePath;
    fr.fileName        = "tpm.xlsx";
    fr.templateVersion = "new";

    SheetResult sheet;
    sheet.sheetName       = "Lifetime Test";
    sheet.templateVersion = "new";
    sheet.columnHeaders   << "puffs" << "beforeWeight" << "afterWeight";
    sheet.overallAvgTPM   = 3.5;

    SampleResult sample;
    sample.sampleName = marker;
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

// A detailed-sensory session with all 11 metric scores defaulted to 5.0 and the
// title/tester set so isDetailedSessionSavable() passes. Twin of
// makeSensorySession, for the DV-25 detailed repro + flip-guard scenarios.
DetailedSensorySession makeDetailedSession(const QString& sessionName,
                                           const QString& tester,
                                           const QString& date,
                                           const QString& title = QString()) {
    DetailedSensorySession s;
    s.sessionName  = sessionName;
    s.testTitle    = title.isEmpty() ? sessionName : title;
    s.testerName   = tester;
    s.assessorName = tester;
    s.media        = "Sample A";
    s.date         = date;
    s.timestamp    = date + "T10:00:00Z";

    DetailedSensorySample sample;
    sample.name = "Sample 1";
    for (const QString& m : kDetailedAllMetrics) sample.scores[m] = 5.0;
    s.samples.append(sample);
    return s;
}

} // anonymous namespace

class TstSaveIntegrityE2E : public QObject {
    Q_OBJECT

private:
    PostgresConnection* m_pg       = nullptr;   // shared connection (DB + readback)
    DatabaseManager*    m_db       = nullptr;
    IdentityManager*    m_identity = nullptr;

    // Wipe every table this suite writes so each slot starts clean.
    void wipe() {
        QSqlQuery q(m_pg->queryDb());
        for (const QString& t : QStringList{
                 "data_rows", "images", "samples", "tests", "files",
                 "sensory_images", "sensory_sessions",
                 "detailed_sensory_images", "detailed_sensory_sessions",
                 "presence", "cell_focus" }) {
            q.exec("DELETE FROM " + t);
        }
    }

    // Read one metric score straight from the sensory_sessions JSONB blob via
    // the shared connection -- the ground truth the merge must honor.
    double dbScore(int sessionId, int sampleIdx, const QString& metric) {
        QSqlQuery q(m_pg->queryDb());
        q.prepare(QString(
            "SELECT (json_data->'samples'->%1->>'%2')::float8 "
            "FROM sensory_sessions WHERE id = ?")
            .arg(sampleIdx).arg(metric));
        q.addBindValue(sessionId);
        if (!q.exec() || !q.next()) return -999.0;
        return q.value(0).toDouble();
    }

    // DV-25: read a per-sample TEXT field straight from the DB blob (the ground
    // truth a rename must land in). ->> yields the JSON string value untyped.
    QString dbText(int id, int sampleIdx, const QString& key) {
        QSqlQuery q(m_pg->queryDb());
        q.prepare(QStringLiteral(
            "SELECT json_data->'samples'->%1->>'%2' "
            "FROM sensory_sessions WHERE id = ?").arg(sampleIdx).arg(key));
        q.addBindValue(id);
        if (!q.exec() || !q.next()) return QString();
        return q.value(0).toString();
    }

    // DV-25: dbText twin against the detailed-sensory blob.
    QString dbTextDetailed(int id, int sampleIdx, const QString& key) {
        QSqlQuery q(m_pg->queryDb());
        q.prepare(QStringLiteral(
            "SELECT json_data->'samples'->%1->>'%2' "
            "FROM detailed_sensory_sessions WHERE id = ?").arg(sampleIdx).arg(key));
        q.addBindValue(id);
        if (!q.exec() || !q.next()) return QString();
        return q.value(0).toString();
    }

    // DV-28: a row's ctid is a transactionally-exact "was this tuple physically
    // UPDATEd" probe. Postgres MVCC writes a new tuple version (new ctid) on every
    // UPDATE, even one the DV-23 trigger suppresses the version bump for; a gated
    // (skipped) save issues no UPDATE, so the ctid is stable. pg_stat counters lag
    // and can't distinguish a no-op UPDATE, so we use the ctid.
    QString rowCtid(int id) {
        QSqlQuery q(m_pg->queryDb());
        q.prepare("SELECT ctid::text FROM sensory_sessions WHERE id = ?");
        q.addBindValue(id);
        if (!q.exec() || !q.next()) return QString();
        return q.value(0).toString();
    }

    int sensoryRowCount(const QString& tester, const QString& date) {
        QSqlQuery q(m_pg->queryDb());
        q.prepare("SELECT count(*) FROM sensory_sessions "
                  "WHERE tester_name = ? AND date = ?");
        q.addBindValue(tester);
        q.addBindValue(date);
        if (!q.exec() || !q.next()) return -1;
        return q.value(0).toInt();
    }

    int sensoryRowCountNamed(const QString& name, const QString& tester,
                             const QString& date) {
        QSqlQuery q(m_pg->queryDb());
        q.prepare("SELECT count(*) FROM sensory_sessions "
                  "WHERE session_name = ? AND tester_name = ? AND date = ?");
        q.addBindValue(name);
        q.addBindValue(tester);
        q.addBindValue(date);
        if (!q.exec() || !q.next()) return -1;
        return q.value(0).toInt();
    }

    int filesRowCount(const QString& path) {
        QSqlQuery q(m_pg->queryDb());
        q.prepare("SELECT count(*) FROM files WHERE file_path = ?");
        q.addBindValue(path);
        if (!q.exec() || !q.next()) return -1;
        return q.value(0).toInt();
    }

    // 3a/H7: the id of the ONLY row in a child table, or -1 when the table does
    // not hold exactly one row. init() wipes every table, so inside a slot that
    // persists a single one-sheet/one-sample/one-row FileResult this is the
    // churn probe: an UPDATE keeps the id, an INSERT-plus-prune replaces it, and
    // a failed prune leaves two rows. `table` is a test-side literal, never
    // user input, so interpolating it is safe.
    qint64 soleId(const QString& table) {
        QSqlQuery q(m_pg->queryDb());
        if (!q.exec("SELECT id FROM " + table)) return -1;
        if (!q.next()) return -1;
        const qint64 id = q.value(0).toLongLong();
        return q.next() ? -1 : id;
    }

private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty())
            QSKIP("DVE_TEST_PG_CONN not set; save-integrity e2e tests skipped");
        QCoreApplication::setOrganizationName("DataViewerTest");
        QCoreApplication::setApplicationName("tst_saveintegrity_e2e");
        { QSettings s; s.clear(); }

        m_pg = new PostgresConnection(this);
        QVERIFY2(m_pg->open(pgConfig()), "could not open shared PostgresConnection");

        m_identity = new IdentityManager(this);
        m_identity->setDisplayName("E2EUser");
        m_identity->setColor("#22aa55");

        m_db = new DatabaseManager(this);
        QVERIFY2(m_db->open(pgConfig(), m_identity), "DatabaseManager open failed");
        QVERIFY(m_db->isOpen());
    }

    void cleanupTestCase() {
        if (m_pg && m_pg->isOpen()) wipe();
        if (m_db) { m_db->close(); }
    }

    void init() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) return;
        wipe();
    }

    // ----------------------------------------------------------------------
    // SCENARIO 1 (RC1 + RC2 headline): edit -> save -> reload, every value
    // matches in-memory truth. The user's #1 symptom: "the next day I checked
    // the file to see that it had reverted."
    //
    // Chains BOTH edit paths into ONE session, exactly as production did:
    //   * SOME cells stream through the REAL LiveSync per-cell worker (these
    //     land in the DB blob + bump the row version, but the in-memory struct
    //     keeps the stale value -- it must be PRESERVED from the DB on merge).
    //   * OTHER cells are edited in memory only and marked dirty (the panel
    //     contract: dirtyCells "samples[i].<Metric>"). These NEVER reached
    //     LiveSync (the id<=0 / broken-stream cases of RC2) and must survive
    //     the whole-session save because the dirty-aware merge keeps memory
    //     authoritative for them.
    // Reload by id and assert: LiveSync-streamed cells AND dirty local edits
    // BOTH match. A regression in either arbitration reverts a value here.
    // ----------------------------------------------------------------------
    void scenario1_editSaveReload_valuesMatchMemory() {
        // 1. Persist a fresh session (all scores 5.0). id back-filled.
        SensorySession s = makeSensorySession("Reload Match", "Charlie R1",
                                               "2026-06-01");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        const int sessId = s.id;

        // 2. LiveSync per-cell commits for SOME cells, via the REAL async
        //    worker against the container. Smoothness 5.0 -> 7.0,
        //    Vapor Volume 5.0 -> 6.5. NOT marked dirty: these are the
        //    LiveSync-owned, DB-authoritative cells.
        LiveSync sync(m_pg, m_identity);
        sync.setWorkerConfig(pgConfig());
        QVERIFY(sync.commitCell("sensory_sessions", sessId,
                                "json_path:samples[0].Smoothness", 7.0));
        QVERIFY(sync.commitCell("sensory_sessions", sessId,
                                "json_path:samples[0].Vapor Volume", 6.5));
        // Drain the worker queue and block until Postgres has both writes.
        QVERIFY2(sync.flushNowAndWait(6000),
                 "LiveSync worker did not confirm the per-cell drain");

        // Confirm the streamed cells are actually in the DB blob now.
        QTRY_VERIFY2(qFuzzyCompare(dbScore(sessId, 0, "Smoothness"), 7.0),
                     "LiveSync Smoothness commit never landed");
        QCOMPARE(dbScore(sessId, 0, "Vapor Volume"), 6.5);

        // 3. Mutate OTHER cells IN MEMORY ONLY and mark them dirty. These never
        //    flow through LiveSync (the RC2 data-loss path). The in-memory
        //    struct still carries the STALE version from step 1 (the per-cell
        //    commits bumped the DB version out from under it) -- the v2.4.0
        //    bug treated that as "already synced, skip" and dropped the edit.
        s.samples[0].scores["Overall Liking"] = 8.5;
        s.samples[0].scores["Burnt Taste"]    = 2.0;
        s.dirtyCells << "samples[0].Overall Liking" << "samples[0].Burnt Taste";
        // The struct's other scores are still the seed 5.0 -- the streamed
        // Smoothness/Vapor Volume are STALE 5.0 in memory but NOT dirty, so the
        // merge must keep the DB value for them rather than the memory 5.0
        // (asserted at the DB level in 5b; the Qt read-back limitation is in 5c).

        // 4. Whole-session save. Internally: fresh-version OCC (no spurious
        //    VersionMismatch), read DB blob, dirty-aware merge, UPDATE.
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);

        // 5a. RC2 HEADLINE assertion (the v2.5.0 fix this suite validates):
        //     the dirty-marked LOCAL edits survive the whole-session save and
        //     read back correctly. These are serialized from the in-memory
        //     NUMBER, so the read path parses them. This is the user's "edits
        //     reverted the next day" symptom — and it is FIXED.
        SensorySession reloaded = m_db->loadSensorySession(sessId);
        QCOMPARE(reloaded.id, sessId);
        QCOMPARE(reloaded.samples[0].scores["Overall Liking"], 8.5);
        QCOMPARE(reloaded.samples[0].scores["Burnt Taste"],    2.0);
        // The untouched metric kept the seed value (number) and round-trips.
        QCOMPARE(reloaded.samples[0].scores["Overall Flavor"], 5.0);

        // 5b. LiveSync-streamed cells: assert at the DB level that the merge
        //     PRESERVED them (the whole-session save did NOT clobber them back
        //     to the stale in-memory 5.0). dbScore() reads via ::float8 in SQL,
        //     so it sees the true stored value regardless of its JSON type.
        QCOMPARE(dbScore(sessId, 0, "Smoothness"),   7.0);
        QCOMPARE(dbScore(sessId, 0, "Vapor Volume"), 6.5);

        // 5c. *** DATAVIEWER-4 ROOT-CAUSE FIX (the original "scores reset to 5"
        //     revert) — now GREEN end to end. ***
        //     The streamed scores round-trip through loadSensorySession: they
        //     read back as the streamed 7.0/6.5, NOT the 5.0 default. This was
        //     the v2.0.1-era bug where dve_commit_cell_json stored every JSONB
        //     value via to_jsonb($2::text), so a numeric score was stored as a
        //     JSON STRING ("7.0") and sensorySessionFromJson's
        //     QJsonValue::toDouble(5.0) returned the DEFAULT for that string.
        //     The fix is two-sided and BOTH halves are exercised here:
        //       * writer — ensureSchema (run by m_db->open in initTestCase)
        //         healed the live container's 7-arg dve_commit_cell_json to store
        //         numeric-looking values as JSON NUMBERS, so the LiveSync commits
        //         in step 2 above now land as numbers;
        //       * reader — the tolerant jsonToDouble in sensorySessionFromJson
        //         also coerces any string-typed legacy value (covered directly by
        //         scenario6 below + the focused tst_sensorydataplaceholder unit).
        QCOMPARE(reloaded.samples[0].scores["Smoothness"],   7.0);
        QCOMPARE(reloaded.samples[0].scores["Vapor Volume"], 6.5);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 2 (RC3): broken worker mid-edit, the whole-session save still
    // wins. Production: the worker connection entered a permanently-broken
    // state ("26000 unnamed prepared statement does not exist") and every
    // per-cell commit died silently. The defense is two-layered: the worker
    // reconnects (Task 4) AND, regardless of whether the per-cell commit
    // recovers, the dirty-marked local edit is authoritative on whole-save.
    //
    // We kill the worker's backend (pg_terminate_backend), then commit a cell
    // (may or may not land via reconnect) AND mark it dirty + whole-save.
    // Reload: the edit is present REGARDLESS of the per-cell commit's fate.
    // ----------------------------------------------------------------------
    void scenario2_brokenWorkerMidEdit_saveStillWins() {
        SensorySession s = makeSensorySession("Broken Worker", "Charlie R1",
                                              "2026-06-02");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        const int sessId = s.id;

        LiveSync sync(m_pg, m_identity);
        sync.setWorkerConfig(pgConfig());

        // Prime the worker so its backend exists, then identify+kill it. Snapshot
        // the pre-worker client backends so the worker's is the one new PID.
        // Filter to real client backends (exclude pg_cron's flickering workers).
        static const char* kClientBackendFilter =
            "SELECT pid FROM pg_stat_activity "
            "WHERE datname = current_database() AND pid <> pg_backend_pid() "
            "AND backend_type = 'client backend' "
            "AND coalesce(application_name, '') NOT LIKE 'pg_cron%'";

        QSqlQuery q(m_pg->queryDb());
        QSet<int> prePids;
        QVERIFY(q.exec(kClientBackendFilter));
        while (q.next()) prePids.insert(q.value(0).toInt());

        QVERIFY(sync.commitCell("sensory_sessions", sessId,
                                "json_path:samples[0].Smoothness", 4.0));
        QVERIFY2(sync.flushNowAndWait(6000), "prime drain failed");
        QTRY_VERIFY2(qFuzzyCompare(dbScore(sessId, 0, "Smoothness"), 4.0),
                     "prime commit never landed");

        int workerPid = -1;
        QVERIFY(q.exec(kClientBackendFilter));
        while (q.next()) {
            const int pid = q.value(0).toInt();
            if (!prePids.contains(pid)) { workerPid = pid; break; }
        }
        QVERIFY2(workerPid > 0, "could not identify worker backend PID");
        QVERIFY(q.exec(QString("SELECT pg_terminate_backend(%1)").arg(workerPid)));
        QVERIFY(q.next());
        QTRY_VERIFY2(
            ([&]() {
                return q.exec(QString("SELECT count(*) FROM pg_stat_activity "
                                      "WHERE pid = %1").arg(workerPid))
                    && q.next() && q.value(0).toInt() == 0;
            })(),
            "worker backend did not terminate");

        // Commit a NEW cell on the now-broken connection. The worker MAY
        // reconnect-and-retry (Task 4) or it may fail -- this scenario does NOT
        // depend on which. We do NOT assert the per-cell value landed.
        sync.commitCell("sensory_sessions", sessId,
                        "json_path:samples[0].Overall Liking", 9.0);

        // The actual safety net: mark the cell dirty + whole-session save. The
        // dirty-aware merge makes the in-memory value authoritative no matter
        // what happened to the per-cell stream.
        s.samples[0].scores["Overall Liking"] = 9.0;
        s.dirtyCells << "samples[0].Overall Liking";
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);

        // The dirty edit (serialized from the in-memory NUMBER) is present
        // REGARDLESS of the per-cell commit's fate -- the RC3 safety net holds.
        SensorySession reloaded = m_db->loadSensorySession(sessId);
        QCOMPARE(reloaded.samples[0].scores["Overall Liking"], 9.0);
        // The pre-kill streamed Smoothness is intact at the DB level (the save's
        // merge preserved it -- not clobbered).
        QCOMPARE(dbScore(sessId, 0, "Smoothness"), 4.0);
        // DATAVIEWER-4 fix (was pinned-broken as 5.0): the pre-kill streamed
        // score now round-trips through loadSensorySession as the streamed 4.0,
        // because the commit stored it as a JSON NUMBER (ensureSchema heal) and
        // the reader coerces it. See scenario1 5c.
        QCOMPARE(reloaded.samples[0].scores["Smoothness"], 4.0);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 3 (RC4, June-10 log): rename thrice, no endless loop. The
    // production failure: a session renamed to a name an existing row owns hit
    // 23505, the panel baseline never updated, and EVERY subsequent save
    // (incl. the silent 5s autosave) re-detected the same rename and re-collided
    // forever -- the false "this test name is already in use" the user saw.
    //
    // Here: a session persisted as "New Session"; an existing row owns "T";
    // rename to "T" via the AutoSuffix wrapper (forces INSERT, like the panel's
    // rename branch); save; save again; save again. Assert: zero
    // UniqueViolation results, the resolved name "T_1" is stable across all
    // three saves (the wrapper re-baselines originalSessionName so no re-detect),
    // and exactly ONE suffixed row exists (the loop is dead -- no T_2/T_3/...).
    // ----------------------------------------------------------------------
    void scenario3_renameThriceNoLoop() {
        const QString tester = "Charlie R1";
        const QString date   = "2026-06-03";

        // Pre-existing owner of the name "T" for this (tester, date).
        SensorySession owner = makeSensorySession("T", tester, date);
        QCOMPARE(m_db->tryWriteSensorySession(owner), WriteResult::Success);
        QVERIFY(owner.id > 0);

        // Our session, persisted as "New Session".
        SensorySession s = makeSensorySession("New Session", tester, date);
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        const int newSessionRowId = s.id;

        // Rename to "T" (collides with `owner`). Mirror the panel rename branch:
        // a rename forces INSERT (id=-1) so the old "New Session" row is
        // preserved. buildSession regenerates sessionName from testTitle, so we
        // suffix both in lockstep -- which is exactly what the AutoSuffix
        // wrapper does on collision.
        s.sessionName        = "T";
        s.testTitle          = "T";
        s.id                 = -1;     // rename -> INSERT (preserve old row)
        s.version            = 0;
        s.originalSessionName = "New Session";

        // SAVE #1: collides on "T", auto-suffixes to "T_1", succeeds.
        QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(s), WriteResult::Success);
        QCOMPARE(s.sessionName, QString("T_1"));
        QVERIFY(s.id > 0);
        QVERIFY(s.id != newSessionRowId);   // a new row, old one preserved
        // The wrapper re-baselined so the panel cannot re-detect a fresh rename.
        QCOMPARE(s.originalSessionName, QString("T_1"));
        const int suffixedRowId = s.id;

        // SAVE #2: no edits. With id>0 + matching name this is an in-place
        // UPDATE -- NOT a rename, NOT a collision. No new row, no violation.
        QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(s), WriteResult::Success);
        QCOMPARE(s.sessionName, QString("T_1"));   // stable
        QCOMPARE(s.id, suffixedRowId);             // same row updated in place

        // SAVE #3: same again. The loop, if alive, would re-collide here.
        QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(s), WriteResult::Success);
        QCOMPARE(s.sessionName, QString("T_1"));   // STILL stable -- loop is dead
        QCOMPARE(s.id, suffixedRowId);

        // Row accounting for (tester, date): the original "New Session"
        // (preserved by the rename->INSERT), the pre-existing "T", and exactly
        // ONE "T_1". No T_2/T_3 means the rename loop never fired again.
        QCOMPARE(sensoryRowCountNamed("New Session", tester, date), 1);
        QCOMPARE(sensoryRowCountNamed("T",   tester, date), 1);
        QCOMPARE(sensoryRowCountNamed("T_1", tester, date), 1);
        QCOMPARE(sensoryRowCountNamed("T_2", tester, date), 0);
        // Our session's lineage is exactly 2 rows (the spec's count): the old
        // "New Session" + the one suffixed "T_1". The 3rd row ("T") is the
        // external collision seed, not part of our session.
        QCOMPARE(sensoryRowCount(tester, date), 3);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 4 (RC4 duplicate-friendly naming): three FRESH sessions sharing
    // the SAME (title, tester, date) resolve to "T", "T_1", "T_2" via the
    // AutoSuffix wrapper -- the user's requested "sensory duplicates allowed via
    // _1/_2/_3 iterator", no modal error. All three rows present.
    // ----------------------------------------------------------------------
    void scenario4_duplicateCreateThrice_suffixChain() {
        const QString tester = "Dana R2";
        const QString date   = "2026-06-04";

        SensorySession a = makeSensorySession("T", tester, date);
        QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(a), WriteResult::Success);
        QCOMPARE(a.sessionName, QString("T"));

        SensorySession b = makeSensorySession("T", tester, date);
        QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(b), WriteResult::Success);
        QCOMPARE(b.sessionName, QString("T_1"));

        SensorySession c = makeSensorySession("T", tester, date);
        QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(c), WriteResult::Success);
        QCOMPARE(c.sessionName, QString("T_2"));

        // All three distinct rows present.
        QVERIFY(a.id > 0 && b.id > 0 && c.id > 0);
        QVERIFY(a.id != b.id && b.id != c.id && a.id != c.id);
        QCOMPARE(sensoryRowCountNamed("T",   tester, date), 1);
        QCOMPARE(sensoryRowCountNamed("T_1", tester, date), 1);
        QCOMPARE(sensoryRowCountNamed("T_2", tester, date), 1);
        QCOMPARE(sensoryRowCount(tester, date), 3);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 5 (F6): TPM versioned re-adds. Save a FileResult; re-save the
    // SAME struct (id>0) -> in-place UPDATE, still 1 row. A FRESH struct with
    // the same path (id=-1) -> a NEW version row (the F6 added_at identity, no
    // longer a UniqueViolation on file_path). loadFileByPath returns the NEWEST.
    // ----------------------------------------------------------------------
    void scenario5_tpmReAddAndContinue() {
        const QString path = "/tmp/e2e-tpm.xlsx";

        // Save a fresh file -> 1 row, id back-filled.
        FileResult fr = makeFileResult(path, "VERSION-1");
        QCOMPARE(m_db->tryWriteFile(fr), WriteResult::Success);
        QVERIFY(fr.id > 0);
        const qint64 v1Id = fr.id;
        QCOMPARE(filesRowCount(path), 1);

        // Re-save the SAME struct (id>0) -> in-place UPDATE, still 1 row, same id.
        fr.sheets[0].samples[0].sampleName = "VERSION-1-EDITED";
        QCOMPARE(m_db->tryWriteFile(fr), WriteResult::Success);
        QCOMPARE(fr.id, v1Id);
        QCOMPARE(filesRowCount(path), 1);

        // A FRESH struct (id=-1) with the same path -> a new version row (F6).
        FileResult fr2 = makeFileResult(path, "VERSION-2");
        QCOMPARE(m_db->tryWriteFile(fr2), WriteResult::Success);
        QVERIFY(fr2.id > 0);
        QVERIFY(fr2.id != v1Id);
        QCOMPARE(filesRowCount(path), 2);

        // loadFileByPath returns the NEWEST version (ORDER BY added_at DESC).
        FileResult newest = m_db->loadFileByPath(path);
        QCOMPARE(newest.id, fr2.id);
        QCOMPARE(newest.sheets[0].samples[0].sampleName, QString("VERSION-2"));
    }

    // ----------------------------------------------------------------------
    // SCENARIO 6 (DATAVIEWER-4 reader half): LEGACY string-typed scores already
    // sitting in the production DB read back correctly. The writer fix stops NEW
    // corruption, but the live NAS DB is full of historical rows whose scores
    // were stored as JSON STRINGS by the old to_jsonb(text) commit path. The
    // tolerant jsonToDouble reader must coerce those on load — otherwise every
    // pre-existing session still reverts to 5.0 on the next open.
    //
    // We bypass the app's writer entirely and INSERT a row with raw SQL whose
    // json_data carries STRING-typed scores (exactly what the old function
    // wrote), then loadSensorySession by id and assert the real values, not the
    // 5.0 default.
    // ----------------------------------------------------------------------
    void scenario6_legacyStringScoresReadCorrectly() {
        // Raw JSON with string-typed scores, as the broken commit fn produced.
        const QString json = QStringLiteral(
            "{\"session_name\":\"Legacy String\","
            "\"tester_name\":\"Charlie R1\",\"date\":\"2026-06-06\","
            "\"samples\":[{\"name\":\"Sample 1\","
            "\"Smoothness\":\"7.5\",\"Burnt Taste\":\"2\","
            "\"Vapor Volume\":\"6\",\"Overall Flavor\":\"8\","
            "\"Overall Liking\":\"9\","
            "\"voltage\":\"3.7\",\"resistance\":\"1.2\"}]}");

        int sessId = -1;
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("INSERT INTO sensory_sessions "
                      "(session_name, tester_name, date, json_data, updated_by) "
                      "VALUES (?, ?, ?, ?::jsonb, 'test') RETURNING id");
            q.addBindValue("Legacy String");
            q.addBindValue("Charlie R1");
            q.addBindValue("2026-06-06");
            q.addBindValue(json);
            QVERIFY2(q.exec() && q.next(), qPrintable(q.lastError().text()));
            sessId = q.value(0).toInt();
        }
        QVERIFY(sessId > 0);

        // Sanity: the stored scores really ARE JSON strings (the corruption we
        // are repairing on read), not numbers.
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("SELECT jsonb_typeof(json_data->'samples'->0->'Smoothness') "
                      "FROM sensory_sessions WHERE id = ?");
            q.addBindValue(sessId);
            QVERIFY(q.exec() && q.next());
            QCOMPARE(q.value(0).toString(), QString("string"));
        }

        // The tolerant reader coerces the string-typed values on load.
        SensorySession reloaded = m_db->loadSensorySession(sessId);
        QCOMPARE(reloaded.id, sessId);
        QCOMPARE(reloaded.samples.size(), 1);
        QCOMPARE(reloaded.samples[0].scores["Smoothness"],     7.5);  // was 5.0 (BUG)
        QCOMPARE(reloaded.samples[0].scores["Burnt Taste"],    2.0);
        QCOMPARE(reloaded.samples[0].scores["Vapor Volume"],   6.0);
        QCOMPARE(reloaded.samples[0].scores["Overall Flavor"], 8.0);
        QCOMPARE(reloaded.samples[0].scores["Overall Liking"], 9.0);
        QCOMPARE(reloaded.samples[0].voltage,    3.7);
        QCOMPARE(reloaded.samples[0].resistance, 1.2);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 7 (v2.4.2 A1): app_version is stamped server-side from the
    // connection's application_name on INSERT; an old-client NULL row is FILLED
    // (never blanked) on a v2.4.2 UPDATE; the heal is idempotent.
    // ----------------------------------------------------------------------
    void scenario7_appVersionStamped() {
        // INSERT via the app's own connection (application_name=DataViewer/<ver>).
        SensorySession s = makeSensorySession("Stamp Me", "Charlie R1",
                                              "2026-06-07");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);

        auto appVer = [&](int id) -> QString {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("SELECT app_version FROM sensory_sessions WHERE id = ?");
            q.addBindValue(id);
            return (q.exec() && q.next()) ? q.value(0).toString()
                                          : QStringLiteral("<err>");
        };
        QVERIFY2(appVer(s.id).startsWith("DataViewer/"),
                 qPrintable("app_version not stamped: " + appVer(s.id)));

        // Simulate an OLD client (NULL app_version) by disabling the trigger.
        // ALWAYS re-enable before any assert that could abort the slot — else a
        // failure in this window would leak a disabled trigger to later slots
        // (silent loss of stamping). So we capture results, ENABLE, then assert.
        int oldId = -1;
        QString insertErr;
        bool disabled = false, inserted = false;
        {
            QSqlQuery q(m_pg->queryDb());
            disabled = q.exec("ALTER TABLE sensory_sessions DISABLE TRIGGER "
                              "trg_sensory_sessions_stamp_app_version");
            if (disabled) {
                q.prepare("INSERT INTO sensory_sessions (session_name, "
                          "tester_name, date, json_data, updated_by) "
                          "VALUES (?,?,?,'{}'::jsonb,'old') RETURNING id");
                q.addBindValue("Old Client Row");
                q.addBindValue("Charlie R1");
                q.addBindValue("2026-06-07");
                inserted = q.exec() && q.next();
                if (inserted) oldId = q.value(0).toInt();
                else          insertErr = q.lastError().text();
            }
            // Re-enable UNCONDITIONALLY (idempotent) before any QVERIFY below.
            q.exec("ALTER TABLE sensory_sessions ENABLE TRIGGER "
                   "trg_sensory_sessions_stamp_app_version");
        }
        QVERIFY2(disabled, "could not disable the stamp trigger");
        QVERIFY2(inserted, qPrintable("old-client row INSERT failed: " + insertErr));
        // The old row's stamp is NULL ("pre-v2.4.2").
        QCOMPARE(appVer(oldId), QString());   // NULL -> empty QString

        // A v2.4.2 UPDATE fills the NULL via COALESCE (never blanks a stamp).
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("UPDATE sensory_sessions SET tester_name = tester_name "
                      "WHERE id = ?");
            q.addBindValue(oldId);
            QVERIFY(q.exec());
        }
        QVERIFY2(appVer(oldId).startsWith("DataViewer/"),
                 "NULL app_version should be filled on a v2.4.2 update");

        // Heal idempotence: a second reopen() leaves exactly ONE stamp trigger.
        QVERIFY(m_db->reopen());
        QVERIFY(m_db->reopen());
        QSqlQuery q(m_pg->queryDb());
        QVERIFY(q.exec("SELECT count(*) FROM pg_trigger WHERE tgname = "
                       "'trg_sensory_sessions_stamp_app_version'"));
        QVERIFY(q.next());
        QCOMPARE(q.value(0).toInt(), 1);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 8 (v2.4.2 A2): the 6-arg dve_commit_cell_json overload is healed
    // to numeric JSONB storage, matching the v2.4.1 7-arg heal.
    //
    // NOTE: the 6-arg form is UNCALLABLE while the 7-arg OCC overload
    // (…, p_expected_version INT DEFAULT NULL) coexists — a 6-positional-arg
    // call is ambiguous between the true 6-arg fn and the 7-arg-with-default
    // (SQLSTATE 42725 "function is not unique", surfaced by QPSQL as 26000). So
    // a behavioral call would be testing a fiction that never occurs in
    // production (where both overloads exist). We instead verify the heal at the
    // CATALOG level — the 6-arg body now contains the numeric coercion — which
    // is exactly what the ensureSchema heal guards on. The numeric CASE's
    // runtime behavior is already proven by scenarios 1/2/6 via the 7-arg form
    // (identical body). The heal still matters for single-source-of-truth
    // (init.sql/migration/ensureSchema agree) and the fresh-init.sql-only-deploy
    // edge where ONLY the 6-arg exists and is therefore callable.
    // ----------------------------------------------------------------------
    void scenario8_sixArgCommitJsonHealedToNumeric() {
        auto sixArgBodyHasNumeric = [&]() -> bool {
            QSqlQuery q(m_pg->queryDb());
            if (!q.exec("SELECT prosrc FROM pg_proc WHERE proname = "
                        "'dve_commit_cell_json' AND pronargs = 6") || !q.next())
                return false;
            return q.value(0).toString().contains("to_jsonb($2::numeric");
        };

        // Install the OLD text-only 6-arg body (RED baseline: no numeric marker).
        {
            QSqlQuery q(m_pg->queryDb());
            QVERIFY2(q.exec(
                "CREATE OR REPLACE FUNCTION dve_commit_cell_json("
                "p_table TEXT, p_row_id BIGINT, p_path_text TEXT, "
                "p_path_arr TEXT[], p_value TEXT, p_uuid TEXT) "
                "RETURNS BOOLEAN AS $fn$ DECLARE affected INT; BEGIN "
                "EXECUTE format('UPDATE %I SET json_data = jsonb_set(json_data, $1, "
                "to_jsonb($2::text)::jsonb, true), updated_by = $3 WHERE id = $4', "
                "p_table) USING p_path_arr, p_value, p_uuid, p_row_id; "
                "GET DIAGNOSTICS affected = ROW_COUNT; RETURN affected > 0; "
                "END; $fn$ LANGUAGE plpgsql;"),
                qPrintable(q.lastError().text()));
        }
        QVERIFY2(!sixArgBodyHasNumeric(),
                 "precondition: the old 6-arg body should lack the numeric marker");

        // The heal runs on reopen() and flips the 6-arg body to numeric storage.
        QVERIFY(m_db->reopen());
        QVERIFY2(sixArgBodyHasNumeric(),
                 "6-arg dve_commit_cell_json was not healed to numeric JSONB");
    }

    // ----------------------------------------------------------------------
    // SCENARIO 9 (v2.4.2 R3 — the reset-to-5 KEYSTONE / catch-up guard):
    // After a network blip + reconnect, reloadOpenResourceAfterReconnect()
    // pulls the authoritative DB blob for the open session and dirty-aware-
    // merges it into memory. This scenario validates the MERGE that catch-up
    // relies on — the precise arbitration that stops a stale in-memory save
    // from clobbering a freshly-normalized DB value, while never losing the
    // user's own unsaved (dirty) edit:
    //   * Another client edits Smoothness in the DB during the "offline
    //     window" (Smoothness is NOT dirty locally — still the seed 5.0).
    //   * The local client edits a DIFFERENT cell (Overall Liking) and marks
    //     it dirty; Smoothness is untouched locally.
    //   * Catch-up loads DB truth and merges keeping dirty local: the non-
    //     dirty Smoothness must take the REMOTE value (9.0), and the dirty
    //     Overall Liking must keep the LOCAL value (2.0). Neither is lost.
    //
    // This guards the MERGE arbitration AND the SP2-T4 overlay contract the
    // production catch-up now uses. The original 9be0550 wiring rebuilt the
    // session via sensorySessionFromJson(merged) and refreshed the navigator,
    // which (a) dropped images and reset id/version — breaking OCC — and (b)
    // flushed stale on-screen widgets, reverting the merged remote value. The
    // fix overlays ONLY the scalar scores onto the in-memory struct (the exact
    // loop SensoryPanel::applyMergedScoresToCurrentSession runs), preserving
    // every non-score field. This scenario asserts BOTH the value arbitration
    // and that non-score fields survive — it would FAIL against the broken
    // fromJson round-trip below.
    // ----------------------------------------------------------------------
    void scenario9_reconnectCatchUpMergePreservesBothSides() {
        SensorySession s = makeSensorySession("Catchup", "Eve R3", "2026-06-09");
        // Seed a real attached image + a string-shaped score so the round-trip
        // hazards (Defect 1) are exercised: images + anchors must survive.
        s.imagePaths << "C:/photos/eve-r3.png";
        s.imageLayouts << QRectF(0.1, 0.1, 0.5, 0.5);
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        const int id = s.id;
        // tryWriteSensorySession back-fills id/version (and image ids) on the
        // local struct — this is the live-app state the catch-up overlays onto.
        const int    seededId      = s.id;
        const int    seededVersion = s.version;
        const QStringList seededImages = s.imagePaths;
        QVERIFY(!seededImages.isEmpty());

        // "Offline window": another client edits Smoothness directly in the DB.
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("UPDATE sensory_sessions SET json_data = jsonb_set("
                      "json_data, '{samples,0,Smoothness}', to_jsonb(?::numeric)) WHERE id=?");
            q.addBindValue("9.0");
            q.addBindValue(id);
            QVERIFY2(q.exec(), qPrintable(q.lastError().text()));
        }
        // Local client: a DIFFERENT cell edited + marked dirty, Smoothness
        // untouched locally (still the seed 5.0, NOT dirty).
        s.samples[0].scores["Overall Liking"] = 2.0;
        s.dirtyCells.clear();
        s.dirtyCells << "samples[0].Overall Liking";

        // Catch-up reload + dirty-aware merge (what reloadOpenResourceAfterReconnect
        // does for a sensory resource): load DB truth, merge keeping dirty local.
        SensorySession dbNow = m_db->loadSensorySession(id);
        QJsonObject merged = mergeSensoryPreservingDbScores(
            sensorySessionToJson(s), sensorySessionToJson(dbNow), s.dirtyCells);

        // SP2-T4 overlay-only path: copy ONLY the kSensoryMetrics scalars onto a
        // COPY of the in-memory struct. This calls the SAME shared helper
        // SensoryPanel::applyMergedScoresToCurrentSession runs in production (the
        // GUI panel is not constructible in this headless harness), so the test
        // exercises real production code against the real merged blob — there is
        // no parallel/hand-copied overlay loop to drift out of sync.
        SensorySession result = s;            // keep every in-memory field
        overlayMergedScores(result, merged);

        // --- Value arbitration (the original guard) ---
        // Non-dirty Smoothness took the REMOTE value (9.0), not the stale local 5.0.
        QCOMPARE(result.samples[0].scores["Smoothness"], 9.0);
        // Dirty Overall Liking kept the LOCAL edit (2.0).
        QCOMPARE(result.samples[0].scores["Overall Liking"], 2.0);

        // --- Defect 1: non-score fields MUST survive the overlay ---
        // The broken 9be0550 path (sensorySessionFromJson(merged)) would reset
        // id to -1, version to 0, and drop imagePaths entirely. Assert the
        // persistence anchors + attached image are intact so the next save is an
        // UPDATE-in-place (OCC), not a duplicate INSERT.
        QCOMPARE(result.id, seededId);
        QCOMPARE(result.version, seededVersion);
        QCOMPARE(result.imagePaths, seededImages);

        // Sanity: prove the OLD broken approach actually loses these fields, so
        // this test genuinely discriminates the regression.
        const SensorySession brokenRoundTrip = sensorySessionFromJson(merged);
        QCOMPARE(brokenRoundTrip.id, -1);            // anchor reset
        QCOMPARE(brokenRoundTrip.version, 0);        // anchor reset
        QVERIFY(brokenRoundTrip.imagePaths.isEmpty());  // image dropped
    }

    // ----------------------------------------------------------------------
    // SCENARIO 9b (v2.4.2 SP2-T4 — detailed-sensory catch-up overlay):
    // Twin of scenario9 for the detailed panel. Asserts the same overlay-only
    // contract: a non-dirty remote score is adopted, a dirty local score is
    // kept, and the persistence anchors + attached image survive (the broken
    // detailedSensorySessionFromJson round-trip would drop them).
    // ----------------------------------------------------------------------
    void scenario9b_detailedCatchUpOverlayPreservesNonScoreFields() {
        DetailedSensorySession s;
        s.sessionName  = "DetailCatchup";
        s.testTitle    = "DetailCatchup";
        s.testerName   = "Mallory R3";
        s.assessorName = "Mallory R3";
        s.media        = "Sample B";
        s.date         = "2026-06-09";
        s.timestamp    = "2026-06-09T10:00:00Z";
        DetailedSensorySample sample;
        sample.name = "Sample 1";
        for (const QString& m : kDetailedAllMetrics) sample.scores[m] = 5.0;
        s.samples.append(sample);
        s.imagePaths << "C:/photos/mallory-r3.png";
        s.imageLayouts << QRectF(0.2, 0.2, 0.4, 0.4);

        QCOMPARE(m_db->tryWriteDetailedSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        const int id = s.id;
        const int    seededId      = s.id;
        const int    seededVersion = s.version;
        const QStringList seededImages = s.imagePaths;
        QVERIFY(!seededImages.isEmpty());

        // Pick two real metric keys to arbitrate (first = remote, second = dirty).
        QVERIFY(kDetailedAllMetrics.size() >= 2);
        const QString remoteMetric = kDetailedAllMetrics.at(0);
        const QString dirtyMetric  = kDetailedAllMetrics.at(1);

        // "Offline window": another client edits the remote metric in the DB.
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare(QString("UPDATE detailed_sensory_sessions SET json_data = "
                              "jsonb_set(json_data, '{samples,0,%1}', "
                              "to_jsonb(?::numeric)) WHERE id=?").arg(remoteMetric));
            q.addBindValue("8.0");
            q.addBindValue(id);
            QVERIFY2(q.exec(), qPrintable(q.lastError().text()));
        }
        // Local client: a DIFFERENT metric edited + marked dirty.
        s.samples[0].scores[dirtyMetric] = 3.0;
        s.dirtyCells.clear();
        s.dirtyCells << QString("samples[0].%1").arg(dirtyMetric);

        DetailedSensorySession dbNow = m_db->loadDetailedSensorySession(id);
        QJsonObject merged = mergeDetailedSensoryPreservingDbScores(
            detailedSensorySessionToJson(s), detailedSensorySessionToJson(dbNow),
            s.dirtyCells);

        // SP2-T4 overlay-only path: calls the SAME shared helper
        // DetailedSensoryPanel::applyMergedScoresToCurrentSession runs in
        // production (the GUI panel is not constructible headless), so the test
        // exercises real production code — no parallel overlay loop to drift.
        DetailedSensorySession result = s;     // keep every in-memory field
        overlayMergedScores(result, merged);

        // Value arbitration: remote adopted, dirty kept.
        QCOMPARE(result.samples[0].scores[remoteMetric], 8.0);
        QCOMPARE(result.samples[0].scores[dirtyMetric], 3.0);

        // Defect 1: anchors + image survive the overlay.
        QCOMPARE(result.id, seededId);
        QCOMPARE(result.version, seededVersion);
        QCOMPARE(result.imagePaths, seededImages);

        // Discriminator: the broken round-trip drops them.
        const DetailedSensorySession broken =
            detailedSensorySessionFromJson(merged);
        QCOMPARE(broken.id, -1);
        QCOMPARE(broken.version, 0);
        QVERIFY(broken.imagePaths.isEmpty());
    }

    // ----------------------------------------------------------------------
    // SCENARIO 10 (DATAVIEWER-11 — phone-append survives a desktop save):
    // The phone web form appends a sample to a sensory row the desktop has open.
    // tryWriteSensoryCore's read-merge-write (DatabaseManager.cpp:1781-1799) pulls
    // DB truth and merges it into the wholesale blob before the UPDATE. Pre-fix,
    // mergeSensoryPreservingDbScores truncated to the in-memory sample count, so
    // the desktop's whole-session save silently DELETED the phone's appended
    // sample (the inverse of the owner requirement). This asserts:
    //   (a) the save now PRESERVES the DB-only phone sample (the 1.2 merge fix),
    //   (b) the sample_uid round-trips through the DB write+read (1.1), and
    //   (c) the panel-side apply grows the in-memory struct so the cards/radar
    //       render the new sample (1.3 — overlayMergedScores + the shared
    //       appendDbOnlyTailSamples helper the panel runs; no GUI needed).
    // ----------------------------------------------------------------------
    void scenario10_phoneAppendSurvivesDesktopSave() {
        SensorySession s = makeSensorySession("PhoneAppend", "Zoe R1", "2026-06-25");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        const int id = s.id;
        QCOMPARE(s.samples.size(), 1);             // desktop in-memory: 1 sample

        // Phone web form appends a 2nd sample to the SAME row (the tail-append
        // shape dve_append_sensory_sample produces) while the desktop holds it open.
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("UPDATE sensory_sessions SET json_data = jsonb_set(json_data, "
                      "'{samples}', COALESCE(json_data->'samples','[]'::jsonb) || ?::jsonb) "
                      "WHERE id=?");
            q.addBindValue(QString::fromUtf8(
                "{\"name\":\"S2-phone\",\"sample_uid\":\"uid-phone-1\","
                "\"Burnt Taste\":6,\"Vapor Volume\":6,\"Overall Flavor\":6,"
                "\"Smoothness\":6,\"Overall Liking\":6}"));
            q.addBindValue(id);
            QVERIFY2(q.exec(), qPrintable(q.lastError().text()));
        }

        // Desktop whole-session save of the STILL-1-sample in-memory struct
        // (mirrors onUpdateDatabase flush=true; s.dirtyCells empty — no local edits).
        // The internal read-merge-write must keep the DB-only phone sample.
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);

        // (a)+(b): reload — the phone sample (and its uid) survived the save.
        const SensorySession reloaded = m_db->loadSensorySession(id);
        QCOMPARE(reloaded.samples.size(), 2);
        QCOMPARE(reloaded.samples[1].name, QString("S2-phone"));
        QCOMPARE(reloaded.samples[1].sampleUid, QString("uid-phone-1"));

        // (c): the panel-side apply (live refresh) grows the open in-memory struct
        // via the SAME shared helpers SensoryPanel::applyMergedScoresToCurrentSession
        // runs, so cards/radar render the new sample (the GUI panel is not
        // constructible headless — call the production helpers directly).
        SensorySession open = makeSensorySession("PhoneAppend", "Zoe R1", "2026-06-25");
        open.id = id; open.version = reloaded.version;     // the still-1-sample open struct
        const QJsonObject mergedForUi = mergeSensoryPreservingDbScores(
            sensorySessionToJson(open), sensorySessionToJson(reloaded), open.dirtyCells);
        overlayMergedScores(open, mergedForUi);
        appendDbOnlyTailSamples(open, mergedForUi);
        QCOMPARE(open.samples.size(), 2);
        QCOMPARE(open.samples[1].sampleUid, QString("uid-phone-1"));
    }

    // ----------------------------------------------------------------------
    // SCENARIO 11 (DATAVIEWER-19 audit — removing a sample must NOT resurrect it
    // or smear its scores onto a survivor):
    // The DV-11 size-driven tail-append regression re-appended a removed sample on
    // the next whole-session save, AND the index-based score overlay copied the
    // removed sample's DB scores onto the survivor that slid into its index.
    // Remove the MIDDLE sample, save (read-merge-write), and assert the row shrinks
    // and the surviving samples keep THEIR OWN scores.
    // ----------------------------------------------------------------------
    void scenario11_removeSampleDoesNotResurrectOrCorrupt() {
        SensorySession s = makeSensorySession("RemoveTest", "Ada R1", "2026-06-29");
        s.samples.clear();
        for (int k = 0; k < 3; ++k) {                       // S0=3.0, S1=5.0, S2=7.0
            SensorySample smp;
            smp.name = QString("S%1").arg(k);
            for (const QString& m : kSensoryMetrics) smp.scores[m] = 3.0 + k * 2;
            s.samples.append(smp);
        }
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        const int id = s.id;
        QCOMPARE(m_db->loadSensorySession(id).samples.size(), 3);

        // User removes the MIDDLE sample (S1) -> in-memory now [S0, S2]. `edited`
        // carries id/version so the save is an OCC UPDATE (read-merge-write).
        SensorySession edited = s;
        edited.samples.removeAt(1);
        // Mirror SensoryPanel::onRemoveCard: a removal records into removedSampleUids
        // (a uid-less desktop sample -> sentinel), which puts the read-merge-write
        // into identity-based, no-resurrect mode.
        edited.removedSampleUids.insert(QStringLiteral("__removed__"));
        QCOMPARE(edited.samples.size(), 2);
        QCOMPARE(m_db->tryWriteSensorySession(edited), WriteResult::Success);

        // Reload: the row SHRANK to 2 (S1 not resurrected) and the survivors kept
        // THEIR scores (S2 not smeared with S1's 5.0).
        const SensorySession r = m_db->loadSensorySession(id);
        QCOMPARE(r.samples.size(), 2);
        QCOMPARE(r.samples[0].scores["Smoothness"], 3.0);   // S0 intact
        QCOMPARE(r.samples[1].scores["Smoothness"], 7.0);   // S2 intact, NOT S1's 5.0
    }

    // ----------------------------------------------------------------------
    // SCENARIO 12 (DV-25): a co-open client's STALE whole-session save must not
    // revert another client's committed per-cell edits. "Client A" streams
    // name/comments/voltage per-cell (the real production path for a rename);
    // "client B" holds the pre-edit struct with NO local edits and whole-saves.
    // Today the read-merge-write only preserves DB SCORES, so B's stale name/
    // comments/voltage win -- the reported "name resets after a couple seconds".
    //
    // Renumbered from the plan's "scenario9": scenarios 9/9b/10/11 already exist
    // in this file. The three QEXPECT_FAILs are the DOCUMENTED RED; Task 3's
    // full-field merge coverage turns them green (the failures are removed there).
    // ----------------------------------------------------------------------
    void scenario12_staleWholeSave_doesNotRevertPerCellFields() {
        SensorySession b = makeSensorySession("DV25 Revert", "Charlie R1",
                                              "2026-07-16");
        QCOMPARE(m_db->tryWriteSensorySession(b), WriteResult::Success);
        QVERIFY(b.id > 0);
        const int sessId = b.id;

        // "Client A": per-cell commits via the REAL worker (rename + note + V).
        LiveSync sync(m_pg, m_identity);
        sync.setWorkerConfig(pgConfig());
        QVERIFY(sync.commitCell("sensory_sessions", sessId,
                                "json_path:samples[0].name",
                                QStringLiteral("Renamed by A")));
        QVERIFY(sync.commitCell("sensory_sessions", sessId,
                                "json_path:samples[0].comments",
                                QStringLiteral("A's note")));
        QVERIFY(sync.commitCell("sensory_sessions", sessId,
                                "json_path:samples[0].voltage", 3.7));
        QVERIFY2(sync.flushNowAndWait(6000), "per-cell drain failed");
        QCOMPARE(dbText(sessId, 0, "name"), QStringLiteral("Renamed by A"));

        // "Client B": stale struct (still "Sample 1"), zero local edits,
        // whole-session save -- the autosave tick's exact shape.
        QCOMPARE(b.dirtyCells.size(), 0);
        QCOMPARE(m_db->tryWriteSensorySession(b), WriteResult::Success);

        SensorySession r = m_db->loadSensorySession(sessId);
        // Task 3 fix landed: the stale whole-save no longer reverts A's per-cell
        // name/comments/voltage - these now assert outright (were XFAIL in Task 1).
        QCOMPARE(r.samples[0].name, QStringLiteral("Renamed by A"));
        QCOMPARE(r.samples[0].comments, QStringLiteral("A's note"));
        QCOMPARE(r.samples[0].voltage, 3.7);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 13 (DV-25 detailed twin of scenario12): the detailed-sensory
    // whole-session save has the identical stale-overwrite shape. Client A
    // streams name/comments per-cell; stale client B whole-saves and reverts
    // them. The XFAILs turn green in Task 4 (detailed full-field merge).
    // ----------------------------------------------------------------------
    void scenario13_detailedStaleWholeSave_doesNotRevertPerCellFields() {
        DetailedSensorySession b = makeDetailedSession("DV25 Detailed",
                                                       "Charlie R1", "2026-07-16");
        QCOMPARE(m_db->tryWriteDetailedSensorySession(b), WriteResult::Success);
        QVERIFY(b.id > 0);
        const int sessId = b.id;

        LiveSync sync(m_pg, m_identity);
        sync.setWorkerConfig(pgConfig());
        QVERIFY(sync.commitCell("detailed_sensory_sessions", sessId,
                                "json_path:samples[0].name",
                                QStringLiteral("Renamed by A")));
        QVERIFY(sync.commitCell("detailed_sensory_sessions", sessId,
                                "json_path:samples[0].comments",
                                QStringLiteral("A's note")));
        QVERIFY2(sync.flushNowAndWait(6000), "detailed per-cell drain failed");
        QCOMPARE(dbTextDetailed(sessId, 0, "name"), QStringLiteral("Renamed by A"));

        QCOMPARE(b.dirtyCells.size(), 0);
        QCOMPARE(m_db->tryWriteDetailedSensorySession(b), WriteResult::Success);

        DetailedSensorySession r = m_db->loadDetailedSensorySession(sessId);
        // Task 4 fix landed: the detailed stale whole-save no longer reverts A's
        // per-cell name/comments - assert outright (were XFAIL in Task 1).
        QCOMPARE(r.samples[0].name, QStringLiteral("Renamed by A"));
        QCOMPARE(r.samples[0].comments, QStringLiteral("A's note"));
    }

    // ----------------------------------------------------------------------
    // SCENARIO 14 (DV-28): the save-loop contract the four MainWindow loops must
    // follow - gate, write-if-needed, clear-on-success. A clean persisted session
    // produces NO UPDATE at all (ctid stable); a dirty one physically updates the
    // row; and the post-save clear makes the NEXT pass a no-op again. This is the
    // headless pin for the gate (the GUI loops aren't constructible here); it is
    // green as soon as the Task 2 helpers exist. (Renumbered from the plan's
    // scenario10.)
    // ----------------------------------------------------------------------
    void scenario14_needsSaveGate_cleanSessionIssuesNoUpdate() {
        SensorySession s = makeSensorySession("DV28 NoOp", "Charlie R1",
                                              "2026-07-16");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        s.dirty = false; s.dirtyCells.clear();      // as the loop leaves it on Success
        s.originalSessionName = s.sessionName;       // loaded-baseline shape

        const QString before = rowCtid(s.id);
        QVERIFY(!before.isEmpty());
        // Pass 1: clean -> the gate must skip; no UPDATE reaches the DB.
        if (DVE::sensorySessionNeedsSave(s))
            QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(s), WriteResult::Success);
        QCOMPARE(rowCtid(s.id), before);

        // Pass 2: a real edit flows through and physically updates the row.
        s.samples[0].scores["Smoothness"] = 8.0;
        s.dirtyCells << "samples[0].Smoothness";
        QVERIFY(DVE::sensorySessionNeedsSave(s));
        QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(s), WriteResult::Success);
        s.dirty = false; s.dirtyCells.clear();       // the loop's Success clear
        QVERIFY(rowCtid(s.id) != before);

        // Pass 3: clean again -> skipped again, ctid stable.
        const QString after = rowCtid(s.id);
        if (DVE::sensorySessionNeedsSave(s))
            QCOMPARE(m_db->tryWriteSensorySessionAutoSuffix(s), WriteResult::Success);
        QCOMPARE(rowCtid(s.id), after);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 15 (DV-25 flip guard): with NO per-cell stream, a locally dirty
    // edit to EVERY arbitrated field survives the whole-session save. If a UI edit
    // path ever forgets to record its dirty cell, the merge (now DB-authoritative
    // for these fields) would revert the user's OWN edit - this pins the contract
    // per field, and must be green both before and after the DV-25 flip.
    // (Renumbered from the plan's scenario11.)
    // ----------------------------------------------------------------------
    void scenario15_dirtyLocalEdits_surviveWholeSave_everyField() {
        SensorySession s = makeSensorySession("DV25 Fields", "Charlie R1",
                                              "2026-07-16");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);

        s.samples[0].name              = "Local Rename";
        s.samples[0].comments          = "local note";
        s.samples[0].voltage           = 3.3;
        s.samples[0].resistance        = 1.2;
        s.samples[0].heatingTechnology = "S17B";
        s.samples[0].powerType         = "Constant Power";
        s.samples[0].puffLengthSec     = 4.5;
        s.samples[0].scores["Smoothness"] = 7.5;
        s.dirtyCells << "samples[0].name" << "samples[0].comments"
                     << "samples[0].voltage" << "samples[0].resistance"
                     << "samples[0].heating_technology"
                     << "samples[0].power_type"
                     << "samples[0].puff_length_sec"
                     << "samples[0].Smoothness";
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);

        const SensorySession r = m_db->loadSensorySession(s.id);
        QCOMPARE(r.samples[0].name,              QStringLiteral("Local Rename"));
        QCOMPARE(r.samples[0].comments,          QStringLiteral("local note"));
        QCOMPARE(r.samples[0].voltage,           3.3);
        QCOMPARE(r.samples[0].resistance,        1.2);
        QCOMPARE(r.samples[0].heatingTechnology, QStringLiteral("S17B"));
        QCOMPARE(r.samples[0].powerType,         QStringLiteral("Constant Power"));
        QCOMPARE(r.samples[0].puffLengthSec,     4.5);
        QCOMPARE(r.samples[0].scores["Smoothness"], 7.5);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 16 (DV-25 flip guard, detailed twin): a locally dirty name +
    // comments + one metric survive the detailed whole-session save with no
    // per-cell stream.
    // ----------------------------------------------------------------------
    void scenario16_detailedDirtyLocalEdits_surviveWholeSave() {
        DetailedSensorySession s = makeDetailedSession("DV25 Detailed Fields",
                                                       "Charlie R1", "2026-07-16");
        QCOMPARE(m_db->tryWriteDetailedSensorySession(s), WriteResult::Success);

        const QString metric = kDetailedAllMetrics.at(0);
        s.samples[0].name = "Local Rename";
        s.samples[0].comments = "local note";
        s.samples[0].scores[metric] = 7.0;
        s.dirtyCells << "samples[0].name" << "samples[0].comments"
                     << QStringLiteral("samples[0].%1").arg(metric);
        QCOMPARE(m_db->tryWriteDetailedSensorySession(s), WriteResult::Success);

        const DetailedSensorySession r = m_db->loadDetailedSensorySession(s.id);
        QCOMPARE(r.samples[0].name,     QStringLiteral("Local Rename"));
        QCOMPARE(r.samples[0].comments, QStringLiteral("local note"));
        QCOMPARE(r.samples[0].scores[metric], 7.0);
    }

    // ----------------------------------------------------------------------
    // SCENARIO 17 (Phase 3a / H7): the DB-browser load path keeps its writeback.
    //
    // MainWindow's Load-from-Database (MainWindow.cpp :5079-5102) re-processes
    // the .xlsx, adopts the DB row's id+version onto the freshly-processed
    // FileResult -- whose child ids are ALL still the -1 sentinel, because
    // processFile only reads the workbook -- and persists it once.
    // persistFileCore back-fills sheet/sample/row ids onto the FileResult it is
    // handed (DatabaseOps.cpp:508 / :596 / :663), so that persist MUST take the
    // mutable-reference overload. Routed through the const-ref bool shim, the
    // writeback lands in a throwaway copy (DatabaseManager.cpp:904-911) and the
    // in-memory struct keeps id=-1 on every child; the NEXT whole-file save then
    // re-INSERTs the entire subtree while the Phase-C orphan prune
    // (DatabaseOps.cpp:757-784) deletes the originals -- every child row id
    // churns, on the most common bulk-load path.
    //
    // The scenario models that sequence step for step. It cannot instantiate
    // MainWindow (no test links the GUI TU), so it pins the contract the fixed
    // call site depends on: part A the mutable-ref persist, part B the shim's
    // documented discard, which is exactly why the call site cannot use it.
    // ----------------------------------------------------------------------
    void scenario17_dbBrowserLoadKeepsWriteback() {
        const QString path = "/tmp/e2e-dbbrowser.xlsx";

        // Seed the DB the way the first plain open of this workbook would.
        FileResult seeded = makeFileResult(path, "SEEDED");
        QCOMPARE(m_db->tryWriteFile(seeded), WriteResult::Success);
        QVERIFY(seeded.id > 0);

        // --- Load-from-Database, modelled step for step --------------------
        const FileResult dbRow = m_db->loadFile(seeded.id);          // :5079
        QVERIFY(dbRow.id > 0);

        // processFile() re-reads only the .xlsx, so every id is the sentinel.
        FileResult result = makeFileResult(path, "REPROCESSED");     // :5087
        QCOMPARE(result.sheets[0].id,                   qint64(-1));
        QCOMPARE(result.sheets[0].samples[0].id,        qint64(-1));
        QCOMPARE(result.sheets[0].samples[0].rows[0].id, qint64(-1));

        result.id      = dbRow.id;                                   // :5096
        result.version = dbRow.version;                              // :5097

        // :5102 -- the persist. Must keep its writeback.
        QCOMPARE(m_db->tryWriteFile(result), WriteResult::Success);

        QVERIFY2(result.sheets[0].id > 0,
                 "sheet id discarded: the load path persisted through a const-ref overload");
        QVERIFY2(result.sheets[0].samples[0].id > 0,
                 "sample id discarded: the load path persisted through a const-ref overload");
        QVERIFY2(result.sheets[0].samples[0].rows[0].id > 0,
                 "data-row id discarded: the load path persisted through a const-ref overload");

        const qint64 testId   = result.sheets[0].id;
        const qint64 sampleId = result.sheets[0].samples[0].id;
        const qint64 rowId    = result.sheets[0].samples[0].rows[0].id;
        QCOMPARE(soleId("tests"),     testId);
        QCOMPARE(soleId("samples"),   sampleId);
        QCOMPARE(soleId("data_rows"), rowId);

        // The next whole-file save must be a pure UPDATE -- same ids in memory
        // AND in the DB. With the writeback discarded this is where the subtree
        // is re-INSERTed and the originals pruned.
        result.sheets[0].samples[0].rows[0].tpm = 4.25;
        QCOMPARE(m_db->tryWriteFile(result), WriteResult::Success);
        QCOMPARE(result.sheets[0].id,                    testId);
        QCOMPARE(result.sheets[0].samples[0].id,         sampleId);
        QCOMPARE(result.sheets[0].samples[0].rows[0].id, rowId);
        QCOMPARE(soleId("tests"),     testId);
        QCOMPARE(soleId("samples"),   sampleId);
        QCOMPARE(soleId("data_rows"), rowId);

        // Part B: saveFile(const&) stays fire-and-forget by design
        // (DatabaseManager.cpp:904-927). Other callers depend on that, which is
        // why H7 is fixed at the one call site and not in the shim. If this ever
        // starts back-filling, the shim's contract changed under those callers.
        FileResult shim = makeFileResult("/tmp/e2e-shim.xlsx", "SHIM");
        QVERIFY(m_db->saveFile(shim));
        QCOMPARE(shim.id,                             qint64(-1));
        QCOMPARE(shim.sheets[0].id,                   qint64(-1));
        QCOMPARE(shim.sheets[0].samples[0].rows[0].id, qint64(-1));
    }
};

QTEST_MAIN(TstSaveIntegrityE2E)
#include "tst_saveintegrity_e2e.moc"
