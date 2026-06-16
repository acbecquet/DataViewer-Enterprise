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
    // This is a GUARD (mergeSensoryPreservingDbScores already exists), not a
    // strict RED->GREEN; a failure here means the merge the catch-up depends
    // on has a real bug. The production wiring (calling this on reconnect via
    // reloadOpenResourceAfterReconnect) is verified by the app build + the
    // user's smoke test.
    // ----------------------------------------------------------------------
    void scenario9_reconnectCatchUpMergePreservesBothSides() {
        SensorySession s = makeSensorySession("Catchup", "Eve R3", "2026-06-09");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);
        const int id = s.id;

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
        SensorySession result = sensorySessionFromJson(merged);

        // Non-dirty Smoothness took the REMOTE value (9.0), not the stale local 5.0.
        QCOMPARE(result.samples[0].scores["Smoothness"], 9.0);
        // Dirty Overall Liking kept the LOCAL edit (2.0).
        QCOMPARE(result.samples[0].scores["Overall Liking"], 2.0);
    }
};

QTEST_MAIN(TstSaveIntegrityE2E)
#include "tst_saveintegrity_e2e.moc"
