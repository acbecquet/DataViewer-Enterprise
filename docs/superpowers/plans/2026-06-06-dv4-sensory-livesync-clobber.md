# DATAVIEWER-4 — LiveSync = single source of truth (no clobber, exports & close persist) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Make LiveSync per-cell score edits the single source of truth for every DB write **and** every file export, eliminate the "scores reset to 5/0" clobber, and make closing a single file/session persist exactly like closing the whole program (scoped to that one item) — without hurting typing responsiveness.

**Architecture (one invariant, applied everywhere):** At every **deliberate** persistence point — Ctrl+U, per-item Close, any export (report/Excel/CSV/JSON), and program-close — do three things in order: (1) `saveCurrentTester()` to flush live widgets into the in-memory model; (2) **`LiveSync::flushNowAndWait()`** to drain the 200 ms throttle window and block (bounded) until the worker thread has written this client's pending per-cell edits to the DB; (3) write/read with **DB-authoritative scores** via one shared pure merge helper per mode. The merge keeps every per-metric score key from the DB blob (LiveSync owns scores) while taking all non-score data (metadata, structure, names, comments, device props) from the in-memory model. The 5-second **auto-save** timer deliberately does NOT flush-and-wait (it stays fully async) so active typing never blocks. Result: no path can clobber a LiveSync score, exports always reflect the authoritative DB scores, and closing an item saves it first.

**Why this shape:** the in-memory `scores` map can never express "unset" (spin boxes default to 5.0 sensory / 0.0 detailed; `toSample()` writes all metrics unconditionally), so a wholesale write always carries a default for any metric a stale/remote view doesn't reflect, stamping it over LiveSync's value. The merge (DB-authoritative scores) prevents clobbering **others'** edits; the flush-first prevents the merge from reverting **my own** just-typed edit whose async commit hasn't landed. Both are required for "no conflicts."

**Tech Stack:** C++17 / Qt 6.10 (QtTest, QJsonObject/QJsonArray/QJsonDocument, QThread/QEventLoop/QMetaObject), qmake + MinGW, PostgreSQL 16 (QPSQL). MIP: run `python tools/decrypt_via_copy.py --apply` before any build; create/modify source via plaintext (decrypt first), and prefer the Python delete-and-rewrite pattern for brand-new files. Build is `-Werror -Wall -Wextra -Wpedantic`. Test runner: `tests\run-tests.ps1 [-Filter <suite>]`; DB suites need `tests\start-test-postgres.ps1` (sets `DVE_TEST_PG_CONN`, prepends `vendor\libpq-16` to PATH).

**Behavioral changes (intended — flag to Charlie):**
1. A whole-session save **never changes a score** — scores are owned by LiveSync. Excel import is unaffected (INSERT path, no merge).
2. Exports (report/Excel/CSV/JSON) reflect **DB** scores merged with in-memory metadata, so a missed live update or unopened-elsewhere edit no longer ships a stale score.
3. The per-item **Close** button now **auto-saves** the closed item authoritatively before removing it (TPM file, sensory session(s), detailed session(s)). The old per-file "update the database before closing? Yes/No" prompt is **removed** — close always persists (a hard save error or name collision aborts the close so nothing is lost). This matches "Close = program-close scoped to one item."

---

## File structure

- **`src/pipeline/SensoryData.h` / `.cpp`** — add pure `mergeSensoryPreservingDbScores(inMemory, dbCurrent)` beside `sensorySessionToJson`/`fromJson`. Used by BOTH the DB write and sensory exports.
- **`src/pipeline/DetailedSensoryData.h` / `.cpp`** — add pure `mergeDetailedSensoryPreservingDbScores(inMemory, dbCurrent)` (uses `kDetailedAllMetrics`). Used by the DB write and the detailed export.
- **`src/database/LiveSync.h` / `.cpp`** — add public `void flushNowAndWait(int timeoutMs = 4000)`; add a re-entrancy guard `bool m_flushing`.
- **`src/database/LiveSyncWorker.h` / `.cpp`** — add `public slots: void sync();` that emits new `signals: void synced();` (the drain barrier).
- **`src/database/DatabaseManager.cpp`** — wire the merge into the UPDATE branch of `tryWriteSensoryCore` (~1664-1718, bind ~1687) and `tryWriteDetailedSensoryCore` (~2126-2176, bind ~2147). No header/schema change.
- **`src/ui/SensoryPanel.h` / `.cpp`** — add `QVector<SensorySession> dbAuthoritativeSessions(const QVector<SensorySession>&)`; route report (`generateFullReport` ~2112, `generateCombinedPptx` ~2295), Excel (`saveToExcel` ~1557), CSV (`writeStatsCsv` ~2195), JSON (`saveToJson` ~1546) through it.
- **`src/ui/DetailedSensoryPanel.h` / `.cpp`** — add the same helper; route `generateFullReport` (~1456) / `generateCombinedPptx` (~1506) through it.
- **`src/MainWindow.h` / `.cpp`** — `onUpdateDatabase(bool flushPending=false)` (Ctrl+U + program-close pass `true`; the 5 s timer passes `false`). Rework `onCloseFile` (~2212) to flush + authoritative persist. Add `MainWindow::saveSensorySessionsBeforeClose(const QVector<int>&)` and `saveDetailedSensorySessionsBeforeClose(const QVector<int>&)`; call them from the close-button lambdas (~510-520, ~544-554) before `closeSessions`.
- **`DataViewerEnterprise.pro`** — `VERSION` 2.3.3 → 2.3.4.
- **Tests:** `tests/tst_sensorydataplaceholder/`, `tests/tst_detailedsensoryjson/` (pure merge), `tests/tst_livesync/` (flush primitive), `tests/tst_databasemanager/` (two-writer integration), `tests/tst_sensoryreportsource/` (export authority).

No `OfflineSnapshot` change: the snapshot regenerates from Postgres on clean close, inheriting the corrected blob.

---

## Task 1: Pure sensory score-merge helper + unit tests

**Files:**
- Modify: `src/pipeline/SensoryData.h` (declare), `src/pipeline/SensoryData.cpp` (define)
- Test: `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp`

- [ ] **Step 1: Write failing unit tests.** In `tst_sensorydataplaceholder.cpp` (already includes `SensoryData.h`). Add `#include <QJsonArray>` if absent. Add a file-static helper and four `private slots`:

```cpp
// --- DATAVIEWER-4: whole-session save / export must not clobber LiveSync scores -
static QJsonObject oneSampleBlob(const QString& sampleName, double score,
                                 const QString& comments = QString())
{
    QJsonObject sample;
    sample["name"]     = sampleName;
    sample["comments"] = comments;
    for (const QString& m : DVE::kSensoryMetrics) sample[m] = score;
    QJsonArray samples; samples.append(sample);
    QJsonObject root;
    root["session_name"] = "S";
    root["samples"]      = samples;
    return root;
}

void mergeSensory_dbScoreWinsOverInMemoryDefault()
{
    QJsonObject mem = oneSampleBlob("A", 5.0);   // stale/default in-memory view
    QJsonObject db  = oneSampleBlob("A", 8.0);   // LiveSync wrote 8 per-cell
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    const QJsonObject s0 = merged["samples"].toArray()[0].toObject();
    for (const QString& m : DVE::kSensoryMetrics)
        QCOMPARE(s0[m].toDouble(), 8.0);          // DB scores preserved, not 5
}

void mergeSensory_nonScoreKeysComeFromInMemory()
{
    QJsonObject mem = oneSampleBlob("NewName", 5.0, "new comment");
    mem["media"] = "MediaX";
    QJsonObject db  = oneSampleBlob("OldName", 8.0, "old comment");
    db["media"] = "MediaOld";
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    QCOMPARE(merged["media"].toString(), QString("MediaX"));          // metadata in-memory
    const QJsonObject s0 = merged["samples"].toArray()[0].toObject();
    QCOMPARE(s0["name"].toString(), QString("NewName"));             // name in-memory
    QCOMPARE(s0["comments"].toString(), QString("new comment"));     // comments in-memory
    QCOMPARE(s0[DVE::kSensoryMetrics.first()].toDouble(), 8.0);       // score from DB
}

void mergeSensory_newInMemorySampleKeepsItsScores()
{
    QJsonObject mem = oneSampleBlob("A", 5.0);
    QJsonArray ms = mem["samples"].toArray();
    QJsonObject s1; s1["name"] = "B";
    for (const QString& m : DVE::kSensoryMetrics) s1[m] = 7.0;
    ms.append(s1); mem["samples"] = ms;
    QJsonObject db = oneSampleBlob("A", 8.0);                         // only sample 0
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    const QJsonArray out = merged["samples"].toArray();
    QCOMPARE(out.size(), 2);
    QCOMPARE(out[0].toObject()[DVE::kSensoryMetrics.first()].toDouble(), 8.0); // matched -> DB
    QCOMPARE(out[1].toObject()[DVE::kSensoryMetrics.first()].toDouble(), 7.0); // new -> in-memory
}

void mergeSensory_missingDbScoreKeyLeavesInMemoryValue()
{
    QJsonObject mem = oneSampleBlob("A", 5.0);
    QJsonObject db  = oneSampleBlob("A", 8.0);
    QJsonObject s0 = db["samples"].toArray()[0].toObject();
    s0.remove("Smoothness");                                          // DB lacks one metric
    QJsonArray dbs; dbs.append(s0); db["samples"] = dbs;
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    const QJsonObject m0 = merged["samples"].toArray()[0].toObject();
    QCOMPARE(m0["Smoothness"].toDouble(), 5.0);                       // falls back to in-memory
    QCOMPARE(m0["Burnt Taste"].toDouble(), 8.0);                      // present -> DB
}
```

> Confirm the exact metric strings in `kSensoryMetrics` (e.g. "Smoothness", "Burnt Taste") by reading `src/pipeline/SensoryData.h`; adjust the two literal metric names in the last test if they differ.

- [ ] **Step 2: Run to verify it fails (undeclared function).**
Run: `tests\run-tests.ps1 -Filter tst_sensorydataplaceholder`
Expected: build FAILS — `mergeSensoryPreservingDbScores` not declared.

- [ ] **Step 3: Declare the helper** in `src/pipeline/SensoryData.h`, in `namespace DVE`, beside the serializer declarations:

```cpp
// DATAVIEWER-4: merge an in-memory sensory blob with the current DB blob so a
// whole-session write OR an export never clobbers LiveSync-owned per-cell SCORE
// values. Scores are DB-authoritative: for every sample present in BOTH (matched
// by array index), each kSensoryMetrics score key is taken from `dbCurrent`; all
// other keys (metadata, structure, name, comments, device props) come from
// `inMemory`. Samples in `inMemory` beyond `dbCurrent`'s array keep their
// in-memory scores (newly added, not yet in the DB). Pure / no DB.
QJsonObject mergeSensoryPreservingDbScores(const QJsonObject& inMemory,
                                           const QJsonObject& dbCurrent);
```

- [ ] **Step 4: Implement** in `src/pipeline/SensoryData.cpp` (ensure `#include <QJsonArray>`), at the end of the `namespace DVE` block:

```cpp
QJsonObject mergeSensoryPreservingDbScores(const QJsonObject& inMemory,
                                           const QJsonObject& dbCurrent)
{
    QJsonObject merged = inMemory;                       // metadata + structure: in-memory
    const QJsonArray dbSamples  = dbCurrent.value("samples").toArray();
    QJsonArray       memSamples = merged.value("samples").toArray();

    for (int i = 0; i < memSamples.size() && i < dbSamples.size(); ++i) {
        QJsonObject       memSample = memSamples[i].toObject();
        const QJsonObject dbSample  = dbSamples[i].toObject();
        for (const QString& metric : kSensoryMetrics) {
            if (dbSample.contains(metric))               // DB-authoritative score
                memSample[metric] = dbSample.value(metric);
        }
        memSamples[i] = memSample;
    }
    merged["samples"] = memSamples;
    return merged;
}
```

- [ ] **Step 5: Run to verify it passes.**
Run: `tests\run-tests.ps1 -Filter tst_sensorydataplaceholder`
Expected: PASS (existing slots + 4 new).

- [ ] **Step 6: Commit.**
```bash
git add src/pipeline/SensoryData.h src/pipeline/SensoryData.cpp tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp
git commit -m "feat(sensory): pure DB-authoritative score-merge helper (DATAVIEWER-4)"
```

---

## Task 2: Pure detailed-sensory score-merge helper + unit tests

**Files:**
- Modify: `src/pipeline/DetailedSensoryData.h` (declare), `src/pipeline/DetailedSensoryData.cpp` (define)
- Test: `tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp`

> Confirm the metric-list symbol — it is `kDetailedAllMetrics` (serializer loops it at `DetailedSensoryData.cpp:42-44`, default 0.0). Use the actual symbol.

- [ ] **Step 1: Write failing unit tests.** In `tst_detailedsensoryjson.cpp` (includes `DetailedSensoryData.h`). Add `#include <QJsonArray>` if absent; add a static helper + two slots:

```cpp
static QJsonObject oneDetailedSampleBlob(const QString& name, double score)
{
    QJsonObject sample; sample["name"] = name;
    for (const QString& m : DVE::kDetailedAllMetrics) sample[m] = score;
    QJsonArray samples; samples.append(sample);
    QJsonObject root; root["session_name"] = "D"; root["samples"] = samples;
    return root;
}

void mergeDetailed_dbScoreWinsOverInMemoryDefault()
{
    QJsonObject mem = oneDetailedSampleBlob("A", 0.0);   // unset -> 0.0 default
    QJsonObject db  = oneDetailedSampleBlob("A", 6.0);   // LiveSync wrote 6
    QJsonObject merged = DVE::mergeDetailedSensoryPreservingDbScores(mem, db);
    const QJsonObject s0 = merged["samples"].toArray()[0].toObject();
    for (const QString& m : DVE::kDetailedAllMetrics)
        QCOMPARE(s0[m].toDouble(), 6.0);
}

void mergeDetailed_newInMemorySampleKeepsScores()
{
    QJsonObject mem = oneDetailedSampleBlob("A", 0.0);
    QJsonArray ms = mem["samples"].toArray();
    QJsonObject s1; s1["name"] = "B";
    for (const QString& m : DVE::kDetailedAllMetrics) s1[m] = 4.0;
    ms.append(s1); mem["samples"] = ms;
    QJsonObject db = oneDetailedSampleBlob("A", 6.0);
    QJsonObject merged = DVE::mergeDetailedSensoryPreservingDbScores(mem, db);
    const QJsonArray out = merged["samples"].toArray();
    QCOMPARE(out[0].toObject()[DVE::kDetailedAllMetrics.first()].toDouble(), 6.0);
    QCOMPARE(out[1].toObject()[DVE::kDetailedAllMetrics.first()].toDouble(), 4.0);
}
```

- [ ] **Step 2: Run to verify it fails (undeclared function).**
Run: `tests\run-tests.ps1 -Filter tst_detailedsensoryjson`
Expected: build FAILS.

- [ ] **Step 3: Declare** in `src/pipeline/DetailedSensoryData.h` (in `namespace DVE`, beside the serializer decls):

```cpp
// DATAVIEWER-4: detailed-sensory counterpart of mergeSensoryPreservingDbScores.
// DB-authoritative for every kDetailedAllMetrics score key on matched samples.
QJsonObject mergeDetailedSensoryPreservingDbScores(const QJsonObject& inMemory,
                                                   const QJsonObject& dbCurrent);
```

- [ ] **Step 4: Implement** in `src/pipeline/DetailedSensoryData.cpp` (ensure `#include <QJsonArray>`):

```cpp
QJsonObject mergeDetailedSensoryPreservingDbScores(const QJsonObject& inMemory,
                                                   const QJsonObject& dbCurrent)
{
    QJsonObject merged = inMemory;
    const QJsonArray dbSamples  = dbCurrent.value("samples").toArray();
    QJsonArray       memSamples = merged.value("samples").toArray();
    for (int i = 0; i < memSamples.size() && i < dbSamples.size(); ++i) {
        QJsonObject       memSample = memSamples[i].toObject();
        const QJsonObject dbSample  = dbSamples[i].toObject();
        for (const QString& metric : kDetailedAllMetrics) {
            if (dbSample.contains(metric))
                memSample[metric] = dbSample.value(metric);
        }
        memSamples[i] = memSample;
    }
    merged["samples"] = memSamples;
    return merged;
}
```

- [ ] **Step 5: Run to verify it passes.**
Run: `tests\run-tests.ps1 -Filter tst_detailedsensoryjson`
Expected: PASS.

- [ ] **Step 6: Commit.**
```bash
git add src/pipeline/DetailedSensoryData.h src/pipeline/DetailedSensoryData.cpp tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp
git commit -m "feat(detailed-sensory): pure DB-authoritative score-merge helper (DATAVIEWER-4)"
```

---

## Task 3: `LiveSync::flushNowAndWait()` drain-and-block primitive + tests

**Files:**
- Modify: `src/database/LiveSyncWorker.h` / `.cpp` (add `sync()` slot + `synced()` signal)
- Modify: `src/database/LiveSync.h` / `.cpp` (add `flushNowAndWait`)
- Test: `tests/tst_livesync/tst_livesync.cpp`

**Design:** `flushNowAndWait` (1) stops the throttle timer and calls `onThrottleTick()` to dispatch all `m_pendingCommits` to the worker immediately; (2) posts a `sync()` call to the worker via a queued connection — because the worker thread processes its event queue in order, `sync()` runs only after every previously-queued `commitJson` has executed; (3) waits on a local `QEventLoop`, quit either by the worker's `synced()` signal or a `QTimer` timeout (bounded, so a stalled DB can never hang the UI forever). A re-entrancy guard makes a nested call a no-op. When no worker is wired (sync-fallback used by some tests) `onThrottleTick()` already ran the commits synchronously, so it returns immediately.

- [ ] **Step 1: Write the failing test.** In `tst_livesync.cpp`, read the existing setup to reuse its worker-config helper (look for `setWorkerConfig` usage and the `DVE_TEST_PG_CONN` skip guard). Add a slot:

```cpp
void flushNowAndWait_drainsPendingToDbSynchronously()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
    // Build a LiveSync wired to the worker (mirror this suite's existing
    // "with worker" construction — setWorkerConfig(pgConfig())).
    DVE::LiveSync sync(/* conn */ pgConn(), /* identity */ identity());
    sync.setWorkerConfig(pgConfigForWorker());

    // Pick a table+row the allowlist accepts and that exists in the test DB.
    // Reuse whatever row this suite already seeds; here assume a sensory row id.
    const qint64 rowId = seedSensoryRow();          // suite helper or inline INSERT

    QVERIFY(sync.commitCell("sensory_sessions", rowId,
                            "json_path:samples[0].Smoothness", 8.0));
    QVERIFY(sync.pendingCount() > 0);               // queued, not yet written

    sync.flushNowAndWait(4000);

    QCOMPARE(sync.pendingCount(), 0);               // throttle queue drained
    // And the value is in the DB right now (no sleep/spin needed):
    QSqlQuery q(workerVerifyDb());
    q.prepare("SELECT (json_data->'samples'->0->>'Smoothness') FROM sensory_sessions WHERE id = ?");
    q.addBindValue(rowId);
    QVERIFY(q.exec() && q.next());
    QCOMPARE(q.value(0).toString(), QString("8"));  // stored as JSON string by dve_commit_cell_json
}
```

> Read the suite for the exact construction helpers (connection, identity, how other tests seed a row and verify). If the suite lacks a worker-config helper, construct the worker connection from `pgConfig()` the same way `LiveSync::setWorkerConfig` expects. The assertion's key point: after `flushNowAndWait`, the DB already has the value with NO polling.

- [ ] **Step 2: Run to verify it fails.**
Run: `tests\start-test-postgres.ps1` then `tests\run-tests.ps1 -Filter tst_livesync`
Expected: build FAILS — `flushNowAndWait` not declared.

- [ ] **Step 3a: Add the worker barrier.** In `src/database/LiveSyncWorker.h`, add to the worker class:

```cpp
public slots:
    // DATAVIEWER-4: drain barrier. Because the worker thread runs its event
    // queue in order, this slot executes only AFTER every commitJson() queued
    // before it has run. Emitting synced() lets the UI thread know the queue
    // is empty up to this point.
    void sync();

signals:
    void synced();
```

In `src/database/LiveSyncWorker.cpp`:

```cpp
void LiveSyncWorker::sync()
{
    emit synced();
}
```

- [ ] **Step 3b: Add `flushNowAndWait`.** In `src/database/LiveSync.h`, add a public method after `pendingCount()` and a private guard near `m_pendingCommits`:

```cpp
    // DATAVIEWER-4: drain the throttle queue NOW and block (bounded by
    // timeoutMs) until the worker has written this client's pending per-cell
    // edits to the DB. Call at every DELIBERATE persist point (Ctrl+U, Close,
    // export, program-close) BEFORE a whole-session write/read so the DB holds
    // the freshest scores and the merge can't revert a just-typed value. The
    // 5 s auto-save does NOT call this (it stays fully async) so typing never
    // blocks. No-op (returns immediately) when nothing is pending or no worker
    // is wired (sync fallback already wrote synchronously). Re-entrant calls
    // are no-ops.
    void flushNowAndWait(int timeoutMs = 4000);
```
```cpp
    bool m_flushing = false;   // DATAVIEWER-4 re-entrancy guard
```

In `src/database/LiveSync.cpp` add the includes `#include <QEventLoop>` and `#include <QTimer>` if missing, and implement:

```cpp
void LiveSync::flushNowAndWait(int timeoutMs)
{
    if (m_flushing) return;                  // re-entrancy guard
    m_flushing = true;

    // 1) Dispatch everything queued in the throttle window right now.
    if (m_throttleTimer && m_throttleTimer->isActive())
        m_throttleTimer->stop();
    onThrottleTick();                        // drains m_pendingCommits -> worker (queued)

    // 2) No worker => onThrottleTick ran commits synchronously; done.
    if (!m_worker || !m_workerThread) { m_flushing = false; return; }

    // 3) Wait (bounded) for the worker to drain its queue up to a sync barrier.
    QEventLoop loop;
    QTimer timeout;
    timeout.setSingleShot(true);
    connect(&timeout, &QTimer::timeout, &loop, &QEventLoop::quit);
    connect(m_worker, &LiveSyncWorker::synced, &loop, &QEventLoop::quit);
    QMetaObject::invokeMethod(m_worker, "sync", Qt::QueuedConnection);
    timeout.start(timeoutMs);
    loop.exec();

    m_flushing = false;
}
```

> `onThrottleTick()` is a private slot; calling it directly from a member is fine. Confirm `m_throttleTimer`, `m_worker`, `m_workerThread`, `m_pendingCommits` names against `LiveSync.h` (they are as listed). The `QEventLoop` briefly processes events; the `m_flushing` guard prevents a re-entrant `flushNowAndWait`. Other queued UI events (e.g. NOTIFY applies) may process during the wait — that is harmless and desirable (keeps the UI painting).

- [ ] **Step 4: Run to verify it passes.**
Run: `tests\run-tests.ps1 -Filter tst_livesync`
Expected: PASS — `pendingCount()==0` and the DB has "8" immediately after `flushNowAndWait`.

- [ ] **Step 5: Commit.**
```bash
git add src/database/LiveSync.h src/database/LiveSync.cpp src/database/LiveSyncWorker.h src/database/LiveSyncWorker.cpp tests/tst_livesync/tst_livesync.cpp
git commit -m "feat(livesync): bounded flushNowAndWait drain-and-block primitive (DATAVIEWER-4)"
```

---

## Task 4: Wire merge into `tryWriteSensoryCore` UPDATE + sensory integration regression test

**Files:**
- Modify: `src/database/DatabaseManager.cpp` — UPDATE branch of `tryWriteSensoryCore` (~1664-1718, json_data bind ~1687)
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp`

- [ ] **Step 1: Write the failing integration test.** In `tst_databasemanager.cpp` add a slot that reproduces the clobber: INSERT a session, simulate a LiveSync per-cell edit out-of-band (which bumps `version` via the trigger), re-read the bumped version into a stale (all-5.0) in-memory copy, whole-session-save it, then assert the DB kept the LiveSync score and applied the metadata change.

```cpp
void sensoryWholeSessionSave_preservesLiveSyncScores()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
    DVE::DatabaseManager db;
    QVERIFY(openDb(db));

    DVE::SensorySession s = makeSensorySession("ClobberTest", "Charlie", "2026-06-06");
    DVE::WriteResult wr0 = db.tryWriteSensorySession(s);   // by-ref -> populates s.id, s.version
    QCOMPARE(wr0, DVE::WriteResult::Success);
    const qint64 sid = s.id;
    QVERIFY(sid > 0);

    // Simulate LiveSync per-cell commit out-of-band; bump_version trigger bumps version.
    runOob([&](QSqlQuery& q){
        q.prepare("UPDATE sensory_sessions "
                  "SET json_data = jsonb_set(json_data, '{samples,0,Smoothness}', '8'::jsonb, true), "
                  "    updated_by = 'livesync-sim' WHERE id = ?");
        q.addBindValue(static_cast<qlonglong>(sid));
        QVERIFY(q.exec());
    });

    int bumpedVersion = -1;
    runOob([&](QSqlQuery& q){
        q.prepare("SELECT version FROM sensory_sessions WHERE id = ?");
        q.addBindValue(static_cast<qlonglong>(sid));
        QVERIFY(q.exec() && q.next());
        bumpedVersion = q.value(0).toInt();
    });

    // Stale in-memory view: all scores 5.0, CURRENT id+version, changed metadata.
    DVE::SensorySession stale = makeSensorySession("ClobberTest", "Charlie", "2026-06-06");
    stale.id = sid;
    stale.version = bumpedVersion;
    stale.media = "MergedMedia";
    for (const QString& m : DVE::kSensoryMetrics) stale.samples[0].scores[m] = 5.0;
    QCOMPARE(db.tryWriteSensorySession(stale), DVE::WriteResult::Success);   // UPDATE lands

    DVE::SensorySession loaded = db.loadSensorySession(sid);
    QCOMPARE(loaded.samples[0].scores.value("Smoothness"), 8.0);   // LiveSync value preserved
    QCOMPARE(loaded.samples[0].scores.value("Burnt Taste"), 5.0);  // untouched metric
    QCOMPARE(loaded.media, QString("MergedMedia"));                 // metadata applied
}
```

> Helper names to confirm against the suite: `openDb`, `makeSensorySession`, `loadSensorySession(qint64)`. For the out-of-band edit, mirror the suite's existing pattern (the prior plan referenced a `bumpRowVersionOutOfBand`-style helper near line ~168 that opens a second `QPSQL` connection). If no `runOob` lambda helper exists, inline a second-connection block as that existing helper does, then `QSqlDatabase::removeDatabase(name)`.

- [ ] **Step 2: Run to verify it FAILS (reproduces the clobber).**
Run: `tests\start-test-postgres.ps1` then `tests\run-tests.ps1 -Filter tst_databasemanager`
Expected: FAIL at `QCOMPARE(loaded...Smoothness, 8.0)` — current code wrote 5.0 over it.

- [ ] **Step 3: Implement the merge in the UPDATE branch.** In `tryWriteSensoryCore`, ensure the file has `#include <QJsonDocument>`, `<QJsonObject>`, `<QJsonArray>`, and `"pipeline/SensoryData.h"` (SensoryData.h is already included — it now declares the helper). Just before the UPDATE is prepared (inside `if (s.id != -1 && s.version > 0)`), compute the merged blob and bind it instead of `jsonStr`:

```cpp
    if (s.id != -1 && s.version > 0) {
        // DATAVIEWER-4: read-merge-write. Pull the current blob (this txn) and
        // keep LiveSync-owned per-cell scores so the wholesale write can't reset
        // them to the serializer's 5.0 default. Falls back to the raw blob if the
        // row is gone (the guarded UPDATE then returns RowDeleted/VersionMismatch).
        QString jsonToWrite = jsonStr;
        {
            QSqlQuery sel(db);
            sel.prepare("SELECT json_data FROM sensory_sessions WHERE id = ?");
            sel.addBindValue(s.id);
            if (sel.exec() && sel.next()) {
                const QJsonObject dbRoot =
                    QJsonDocument::fromJson(sel.value(0).toString().toUtf8()).object();
                const QJsonObject memRoot =
                    QJsonDocument::fromJson(jsonStr.toUtf8()).object();
                jsonToWrite = QString::fromUtf8(QJsonDocument(
                    mergeSensoryPreservingDbScores(memRoot, dbRoot))
                        .toJson(QJsonDocument::Compact));
            }
        }

        QSqlQuery q(db);
        q.prepare(R"( UPDATE sensory_sessions SET ... json_data = CAST(? AS JSONB), ...
                      WHERE id = ? AND version = ? RETURNING id, version )");
        ...
        q.addBindValue(jsonToWrite);   // <-- was: q.addBindValue(jsonStr);
        ...
    }
```

Only the `json_data` bind changes (`jsonStr` -> `jsonToWrite`). All other binds/branches and the INSERT branch are untouched.

- [ ] **Step 4: Run to verify it PASSES** + suite regression.
Run: `tests\run-tests.ps1 -Filter tst_databasemanager`
Expected: the new slot PASSES. Known-unrelated failures may exist (`testSensoryLayoutPersistence`, `testSaveSensorySessionPreservesLayoutOnReSave`, `sensoryHeaderPresets_roundTrip`) — these are an `init.sql`-drift issue (ephemeral container lacks `dve_commit_session_layout` + `sensory_header_presets`), NOT this change. Confirm the failure COUNT does not increase beyond those.

- [ ] **Step 5: Commit.**
```bash
git add src/database/DatabaseManager.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "fix(db): merge-preserve LiveSync scores on whole-session sensory save (DATAVIEWER-4)"
```

---

## Task 5: Wire merge into `tryWriteDetailedSensoryCore` UPDATE + detailed integration test

**Files:**
- Modify: `src/database/DatabaseManager.cpp` — UPDATE branch of `tryWriteDetailedSensoryCore` (~2126-2176, json_data bind ~2147)
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp`

- [ ] **Step 1: Write the failing integration test** mirroring Task 4 for detailed sessions: INSERT a `DetailedSensorySession` (suite helper `makeDetailedSensorySession`), out-of-band `jsonb_set` a metric from `kDetailedAllMetrics` to a non-default value (e.g. 6), re-read version, stale-save with that metric at 0.0 + a metadata change, reload, assert the metric stays 6 and the metadata applied. Use `tryWriteDetailedSensorySession` + the suite's detailed load helper (`loadDetailedSensorySession`).

```cpp
void detailedWholeSessionSave_preservesLiveSyncScores()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
    DVE::DatabaseManager db;
    QVERIFY(openDb(db));
    DVE::DetailedSensorySession s = makeDetailedSensorySession("DetClobber", "Charlie", "2026-06-06");
    QCOMPARE(db.tryWriteDetailedSensorySession(s), DVE::WriteResult::Success);
    const qint64 sid = s.id; QVERIFY(sid > 0);

    const QString metric = DVE::kDetailedAllMetrics.first();
    runOob([&](QSqlQuery& q){
        q.prepare("UPDATE detailed_sensory_sessions "
                  "SET json_data = jsonb_set(json_data, ?, '6'::jsonb, true) WHERE id = ?");
        q.addBindValue(QString("{samples,0,%1}").arg(metric));
        q.addBindValue(static_cast<qlonglong>(sid));
        QVERIFY(q.exec());
    });
    int bumped = -1;
    runOob([&](QSqlQuery& q){
        q.prepare("SELECT version FROM detailed_sensory_sessions WHERE id = ?");
        q.addBindValue(static_cast<qlonglong>(sid));
        QVERIFY(q.exec() && q.next()); bumped = q.value(0).toInt();
    });
    DVE::DetailedSensorySession stale = makeDetailedSensorySession("DetClobber", "Charlie", "2026-06-06");
    stale.id = sid; stale.version = bumped; stale.media = "DetMergedMedia";
    for (const QString& m : DVE::kDetailedAllMetrics) stale.samples[0].scores[m] = 0.0;
    QCOMPARE(db.tryWriteDetailedSensorySession(stale), DVE::WriteResult::Success);

    DVE::DetailedSensorySession loaded = db.loadDetailedSensorySession(sid);
    QCOMPARE(loaded.samples[0].scores.value(metric), 6.0);   // preserved
    QCOMPARE(loaded.media, QString("DetMergedMedia"));        // metadata applied
}
```

> Note the detailed `fromJson` reads scores with a `qBound`/default-1.0 clamp; if the metric's valid range excludes 6, pick a value inside the range (read `DetailedSensoryData.cpp` fromJson). The jsonb path is built with a bind to avoid quoting issues; if the suite prefers literal paths, format it inline.

- [ ] **Step 2: Run to verify it FAILS.**
Run: `tests\run-tests.ps1 -Filter tst_databasemanager`
Expected: FAIL on the preserved-metric assertion.

- [ ] **Step 3: Implement** the same read-merge-write in the UPDATE branch of `tryWriteDetailedSensoryCore`. Ensure `#include "pipeline/DetailedSensoryData.h"` is present. Compute `jsonToWrite` from `SELECT json_data FROM detailed_sensory_sessions WHERE id = ?` and `mergeDetailedSensoryPreservingDbScores(memRoot, dbRoot)`; change the `json_data` bind (~2147) from `jsonStr` to `jsonToWrite`. INSERT branch untouched.

```cpp
    if (s.id != -1 && s.version > 0) {
        QString jsonToWrite = jsonStr;
        {
            QSqlQuery sel(db);
            sel.prepare("SELECT json_data FROM detailed_sensory_sessions WHERE id = ?");
            sel.addBindValue(s.id);
            if (sel.exec() && sel.next()) {
                const QJsonObject dbRoot =
                    QJsonDocument::fromJson(sel.value(0).toString().toUtf8()).object();
                const QJsonObject memRoot =
                    QJsonDocument::fromJson(jsonStr.toUtf8()).object();
                jsonToWrite = QString::fromUtf8(QJsonDocument(
                    mergeDetailedSensoryPreservingDbScores(memRoot, dbRoot))
                        .toJson(QJsonDocument::Compact));
            }
        }
        // ... existing UPDATE prepare ...
        q.addBindValue(jsonToWrite);   // <-- was jsonStr
    }
```

- [ ] **Step 4: Run to verify it PASSES** + suite regression check (same 3-known-unrelated caveat as Task 4).
Run: `tests\run-tests.ps1 -Filter tst_databasemanager`

- [ ] **Step 5: Commit.**
```bash
git add src/database/DatabaseManager.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "fix(db): merge-preserve LiveSync scores on whole-session detailed-sensory save (DATAVIEWER-4)"
```

---

## Task 6: DB-authoritative SENSORY exports (report / Excel / CSV / JSON)

**Files:**
- Modify: `src/ui/SensoryPanel.h` (declare helper), `src/ui/SensoryPanel.cpp` (define + route exports)
- Test: `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp` (pure round-trip of the merge-on-export at the JSON layer) and a guard in `tests/tst_sensoryreportsource/` if it constructs sessions directly.

**Approach:** add a panel helper that, for a list of in-memory sessions, returns DB-authoritative copies: flush LiveSync once, then for each session with `id > 0` and a DB present, fetch the DB blob and run `fromJson(merge(toJson(inMem), dbBlob))`; sessions with `id <= 0` (never persisted) or no DB pass through unchanged. Route all four export entry points through it.

> Confirm member names by reading `SensoryPanel.cpp`: the DB pointer used by `generateFullReport` (`m_db`) and the LiveSync pointer used by the `cellCommitted` lambda (`m_liveSync`). Use those exact names. Confirm `db->loadSensorySession(qint64)` returns a `SensorySession` and `sensorySessionToJson` / `sensorySessionFromJson` signatures.

- [ ] **Step 1: Write the failing test.** Because the panel needs a GUI + DB, unit-test the *transform* at the JSON layer (the panel helper is a thin wrapper over it). In `tst_sensorydataplaceholder.cpp` add:

```cpp
void export_usesDbScoresWithInMemoryMetadata()
{
    // Simulates dbAuthoritativeSessions' core: in-memory has fresh metadata +
    // stale 5.0 scores; DB has LiveSync 8.0 scores + old metadata. The export
    // blob must carry in-memory metadata and DB scores.
    QJsonObject mem = oneSampleBlob("A", 5.0, "memo");
    mem["media"] = "FreshMedia";
    QJsonObject db = oneSampleBlob("A", 8.0, "old");
    db["media"] = "OldMedia";
    QJsonObject exportBlob = DVE::mergeSensoryPreservingDbScores(mem, db);
    QCOMPARE(exportBlob["media"].toString(), QString("FreshMedia"));
    const QJsonObject s0 = exportBlob["samples"].toArray()[0].toObject();
    QCOMPARE(s0["comments"].toString(), QString("memo"));
    QCOMPARE(s0[DVE::kSensoryMetrics.first()].toDouble(), 8.0);
}
```

Run: `tests\run-tests.ps1 -Filter tst_sensorydataplaceholder` — Expected PASS already if Task 1 is in (this test pins the export contract; it should pass once the helper exists). If you prefer a red first, write it before Task 1's helper — but here it documents the export requirement and guards against regressions.

- [ ] **Step 2: Declare the helper** in `src/ui/SensoryPanel.h` (private section):

```cpp
    // DATAVIEWER-4: return DB-authoritative copies of `inMem` for export. Flushes
    // LiveSync first so the DB holds this client's latest per-cell scores, then
    // overlays DB scores onto in-memory metadata via mergeSensoryPreservingDbScores.
    // Sessions with id <= 0 (never saved) or when no DB is available pass through.
    QVector<SensorySession> dbAuthoritativeSessions(const QVector<SensorySession>& inMem);
```

- [ ] **Step 3: Implement** in `src/ui/SensoryPanel.cpp` (ensure includes `<QJsonDocument>`, `<QJsonObject>`; `SensoryData.h` already included):

```cpp
QVector<SensorySession> SensoryPanel::dbAuthoritativeSessions(
        const QVector<SensorySession>& inMem)
{
    if (m_liveSync) m_liveSync->flushNowAndWait();     // DB now holds our latest scores
    if (!m_db) return inMem;

    QVector<SensorySession> out;
    out.reserve(inMem.size());
    for (const SensorySession& s : inMem) {
        if (s.id <= 0) { out.append(s); continue; }    // never persisted -> in-memory
        const SensorySession dbSess = m_db->loadSensorySession(s.id);
        if (dbSess.id <= 0) { out.append(s); continue; } // row gone -> in-memory
        const QJsonObject merged = mergeSensoryPreservingDbScores(
            sensorySessionToJson(s), sensorySessionToJson(dbSess));
        SensorySession authoritative = sensorySessionFromJson(merged);
        authoritative.id      = s.id;       // preserve identity/version anchors
        authoritative.version = s.version;
        out.append(authoritative);
    }
    return out;
}
```

> If `sensorySessionFromJson` drops `id`/`version`/`originalSessionName`, re-stamp them from `s` after the call (as shown for id/version). Read `sensorySessionFromJson` to confirm which fields survive the round-trip and re-stamp any export-relevant field (testerName, date, testTitle) that the serializer does not emit.

- [ ] **Step 4: Route the four export paths** through the helper. In each, replace the in-memory selected-session vector with the authoritative one *before* it is consumed:
  - `generateFullReport` (~2112): after `saveCurrentTester()` builds `selected` from `m_sessions`, wrap it: `selected = dbAuthoritativeSessions(selected);` before constructing `SensoryReportSource`.
  - `generateCombinedPptx` (~2295): same — transform the `sessions` argument source at the call site, or apply the helper at the top of the function to its input.
  - `saveToExcel` (~1557): build the session(s) it exports via `dbAuthoritativeSessions({sess})[0]` (or the multi-session list) before reading scores.
  - `saveToJson` (~1546): serialize from `dbAuthoritativeSessions({sess})[0]`.
  - `writeStatsCsv` (~2195): build its `buildSession()` source through the helper.

> Read each method first; apply the transform at the single point where the export's session list is finalized. Keep it one line per site (`x = dbAuthoritativeSessions(x);`). Do NOT change report rendering logic.

- [ ] **Step 5: Build + run.**
Run: `tests\run-tests.ps1 -Filter tst_sensorydataplaceholder` and `tests\run-tests.ps1 -Filter tst_sensoryreportsource`
Expected: PASS. Then a full app release build to confirm the panel compiles under `-Werror` (Task 10 does the clean build; here just `mingw32-make -j8` in `build`).

- [ ] **Step 6: Commit.**
```bash
git add src/ui/SensoryPanel.h src/ui/SensoryPanel.cpp tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp
git commit -m "feat(sensory): DB-authoritative scores for report/Excel/CSV/JSON exports (DATAVIEWER-4)"
```

---

## Task 7: DB-authoritative DETAILED-sensory export (report)

**Files:**
- Modify: `src/ui/DetailedSensoryPanel.h` (declare helper), `src/ui/DetailedSensoryPanel.cpp` (define + route report)
- Test: `tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp` (JSON-layer contract guard)

- [ ] **Step 1: Write the contract test** in `tst_detailedsensoryjson.cpp`:

```cpp
void export_detailed_usesDbScoresWithInMemoryMetadata()
{
    QJsonObject mem = oneDetailedSampleBlob("A", 0.0); mem["media"] = "FreshMedia";
    QJsonObject db  = oneDetailedSampleBlob("A", 6.0); db["media"]  = "OldMedia";
    QJsonObject ex = DVE::mergeDetailedSensoryPreservingDbScores(mem, db);
    QCOMPARE(ex["media"].toString(), QString("FreshMedia"));
    QCOMPARE(ex["samples"].toArray()[0].toObject()[DVE::kDetailedAllMetrics.first()].toDouble(), 6.0);
}
```
Run: `tests\run-tests.ps1 -Filter tst_detailedsensoryjson` — Expected PASS (helper from Task 2 exists).

- [ ] **Step 2: Declare + implement** `DetailedSensoryPanel::dbAuthoritativeSessions(const QVector<DetailedSensorySession>&)` mirroring Task 6 but with `mergeDetailedSensoryPreservingDbScores`, `detailedSensorySessionToJson`, `detailedSensorySessionFromJson`, `m_db->loadDetailedSensorySession(id)`, and the panel's LiveSync pointer (confirm its name; the panel routes edits via `commitSampleField` -> LiveSync). Re-stamp `id`/`version` (and per-image `imageIds`/`imageVersions` if the serializer drops them — read `detailedSensorySessionFromJson`).

```cpp
QVector<DetailedSensorySession> DetailedSensoryPanel::dbAuthoritativeSessions(
        const QVector<DetailedSensorySession>& inMem)
{
    if (m_liveSync) m_liveSync->flushNowAndWait();
    if (!m_db) return inMem;
    QVector<DetailedSensorySession> out; out.reserve(inMem.size());
    for (const DetailedSensorySession& s : inMem) {
        if (s.id <= 0) { out.append(s); continue; }
        const DetailedSensorySession dbSess = m_db->loadDetailedSensorySession(s.id);
        if (dbSess.id <= 0) { out.append(s); continue; }
        const QJsonObject merged = mergeDetailedSensoryPreservingDbScores(
            detailedSensorySessionToJson(s), detailedSensorySessionToJson(dbSess));
        DetailedSensorySession a = detailedSensorySessionFromJson(merged);
        a.id = s.id; a.version = s.version;
        // Re-stamp any non-serialized fields needed by the report (read fromJson).
        out.append(a);
    }
    return out;
}
```

- [ ] **Step 3: Route the report** through it: in `generateFullReport` (~1456) after `saveCurrentTester()` builds `selected`, set `selected = dbAuthoritativeSessions(selected);` before `generateCombinedPptx`. If `generateCombinedPptx` (~1506) is also reachable directly, apply the helper to its input at the top.

- [ ] **Step 4: Build + run.**
Run: `tests\run-tests.ps1 -Filter tst_detailedsensoryjson` + `mingw32-make -j8` in `build`.
Expected: PASS + clean compile.

- [ ] **Step 5: Commit.**
```bash
git add src/ui/DetailedSensoryPanel.h src/ui/DetailedSensoryPanel.cpp tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp
git commit -m "feat(detailed-sensory): DB-authoritative scores for report export (DATAVIEWER-4)"
```

---

## Task 8: Flush-before-persist on the DELIBERATE save paths (Ctrl+U + program-close)

**Files:**
- Modify: `src/MainWindow.h` (signature), `src/MainWindow.cpp` (`onUpdateDatabase`, Ctrl+U connect ~1415, `promptSaveDatabase` ~5071, the 5 s timer connect ~383)

**Why:** `onUpdateDatabase` is invoked from three places: the Ctrl+U action (deliberate), the program-close save (deliberate), and the 5-second auto-save timer (background). The deliberate ones must `flushNowAndWait` first so the merge in Tasks 4/5 sees this client's latest per-cell scores (no self-revert window). The timer must NOT block (responsiveness), so it keeps the async path.

- [ ] **Step 1: Change the signature.** In `src/MainWindow.h`, change `void onUpdateDatabase();` to:

```cpp
    // flushPending=true at DELIBERATE save points (Ctrl+U, program-close): drains
    // LiveSync to the DB first so the whole-session merge sees the latest scores.
    // The 5 s auto-save timer calls with false (stays fully async; never blocks).
    void onUpdateDatabase(bool flushPending = false);
```

- [ ] **Step 2: Flush at the top when asked.** At the start of `onUpdateDatabase` (~4416), before the TPM loop:

```cpp
void MainWindow::onUpdateDatabase(bool flushPending)
{
    if (flushPending && m_liveSync) {
        flushExcelWrites();              // drain debounced TPM cell edits too
        m_liveSync->flushNowAndWait();   // block (bounded) until per-cell edits hit the DB
    }
    int saved = 0, failed = 0;
    ...
```

> Confirm `m_liveSync` is the MainWindow LiveSync member name (it constructs LiveSync and passes it to the panels). `flushExcelWrites()` is already a MainWindow method.

- [ ] **Step 3: Update the three call sites.**
  - Ctrl+U action (~1415): `connect(dbUpdateAct, &QAction::triggered, this, [this]{ onUpdateDatabase(true); });`
  - `promptSaveDatabase` (~5071): `onUpdateDatabase(true);`
  - 5 s timer (~383): keep `onUpdateDatabase()` (defaults to `false`).

> The Ctrl+U connect currently binds the method pointer directly (`&MainWindow::onUpdateDatabase`). With a defaulted bool, a method-pointer connect would call it with `false`; change that one site to a lambda passing `true` as shown.

- [ ] **Step 4: Build + smoke test.**
Run: `mingw32-make -j8` in `build` (clean compile under `-Werror`). Run `tests\run-tests.ps1 -Filter tst_mainwindow_remotecell` to confirm the MainWindow harness still builds/links and passes.
Expected: clean build, harness PASS.

- [ ] **Step 5: Commit.**
```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(persist): flush LiveSync before Ctrl+U and program-close saves (DATAVIEWER-4)"
```

---

## Task 9: Close = authoritative scoped save before removal (TPM file + sensory/detailed sessions)

**Files:**
- Modify: `src/MainWindow.h` (declare two helpers), `src/MainWindow.cpp` (`onCloseFile` ~2212; sensory close lambda ~510-520; detailed close lambda ~544-554; add the two helpers)
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp` (a "close persists then the row reflects the closed item" integration check at the DB layer)

**Why:** today `onCloseFile` uses the weak `saveFile()` behind a Yes/No prompt and never flushes LiveSync; `SensoryPanel::closeSessions` / `DetailedSensoryPanel::closeSessions` just `removeAt` with NO save and NO `saveCurrentTester()` — losing live-widget edits. Make Close persist the closed item authoritatively first (program-close behavior, scoped to that item), then remove.

- [ ] **Step 1: Write a DB-layer integration guard** in `tst_databasemanager.cpp` (the GUI close path can't run headless, so pin the save semantics the close path relies on): save a sensory session, edit a score via the whole-session save (now merge-safe), and assert a reload reflects it — i.e. that `tryWriteSensorySession` is a complete authoritative persist. (This is effectively covered by Task 4; add a short `closeSavesAreAuthoritative_sensory` alias test that documents the close contract, calling the same save+reload and asserting scores + metadata both land.) Keep it minimal; its purpose is documentation + regression.

Run: `tests\run-tests.ps1 -Filter tst_databasemanager` — Expected PASS after Task 4.

- [ ] **Step 2: Rework `onCloseFile`** (~2212). Replace the prompt-and-`saveFile` block (lines ~2222-2236) with an unconditional authoritative persist:

```cpp
    // H2: drain debounced Excel writes + LiveSync per-cell edits BEFORE removal
    // so the closing file picks up its last edits (DATAVIEWER-4: Close == a
    // scoped program-close — always persist, never prompt-to-discard).
    flushExcelWrites();
    if (m_liveSync) m_liveSync->flushNowAndWait();

    const QString fp = m_loadedFiles[m_currentFileIndex].filePath;
    if (m_modifiedFilePaths.contains(fp)) {
        persistLoadedFile(m_currentFileIndex);    // WriteResult-aware + OCC retry
        // If it still couldn't save (hard error / offline), keep it open so the
        // user doesn't lose data; persistLoadedFile already surfaced the reason.
        if (m_modifiedFilePaths.contains(fp)) {
            const auto resp = QMessageBox::warning(
                this, tr("Close Without Saving?"),
                tr("'%1' could not be saved to the database.\n\n"
                   "Close anyway and lose unsaved database changes?")
                    .arg(QFileInfo(fp).fileName()),
                QMessageBox::Yes | QMessageBox::No, QMessageBox::No);
            if (resp != QMessageBox::Yes) return;   // abort close, keep the file
        }
    }
```

Leave the rest of `onCloseFile` (exclusion cleanup, `removeAt`, UI reset) unchanged.

> `persistLoadedFile(int)` already exists (~2307) and does tryWriteFile + one-shot OCC re-inherit + status + clears the dirty flag on success. Reusing it makes per-file close match `onUpdateDatabase`'s TPM rigor exactly. The offline case (`RetryOffline`) keeps the file dirty by design; the warning lets the user choose — same data-safety posture as program-close's Cancel.

- [ ] **Step 3: Add `saveSensorySessionsBeforeClose`** in `MainWindow.cpp` (declare in `.h`):

```cpp
// DATAVIEWER-4: authoritatively persist the given sensory sessions (by panel
// index) before they are closed, so closing never drops edits. Returns the
// indices that FAILED to save (caller should keep those open). Mirrors the
// per-session save in onUpdateDatabase (rename->INSERT, UniqueViolation skip)
// but scoped to the closing set and quiet on success.
QVector<int> MainWindow::saveSensorySessionsBeforeClose(const QVector<int>& indices)
{
    QVector<int> failed;
    if (!m_sensoryPanel || !m_db) return failed;
    if (m_liveSync) m_liveSync->flushNowAndWait();        // our scores -> DB first

    QVector<SensorySession> sessions = m_sensoryPanel->allSessions();  // flushes widgets
    for (int idx : indices) {
        if (idx < 0 || idx >= sessions.size()) continue;
        SensorySession& sess = sessions[idx];
        if (DVE::isPlaceholderSession(sess)) continue;

        const bool isRename = sess.id > 0 && !sess.originalSessionName.isEmpty()
                              && sess.originalSessionName != sess.sessionName;
        if (isRename) { sess.id = -1; sess.version = 0; }   // preserve old row

        DVE::WriteResult r = m_db->tryWriteSensorySession(sess);
        if (r == DVE::WriteResult::UniqueViolation) {
            QMessageBox::information(this, tr("Sensory Session Name Taken"),
                tr("Another sensory session named \"%1\" already exists for "
                   "tester \"%2\" on %3.\n\nThe rename was not saved; pick a "
                   "different Test Title before closing.")
                   .arg(sess.sessionName, sess.testerName, sess.date));
            failed.append(idx);
        } else if (r != DVE::WriteResult::Success
                   && r != DVE::WriteResult::VersionMismatch
                   && r != DVE::WriteResult::RowDeleted) {
            failed.append(idx);    // hard error/offline -> keep open
        }
    }
    m_sensoryPanel->syncSavedSessionState(sessions);   // back-fill id/version
    updateDbSyncIndicator();
    return failed;
}
```

Add `saveDetailedSensorySessionsBeforeClose(const QVector<int>&)` symmetrically using `m_detailedSensoryPanel`, `tryWriteDetailedSensorySession`, `loadDetailedSensorySession`-free (no rename branch — detailed has no `originalSessionName`), and `syncSavedSessionState`.

> Confirm `isPlaceholderSession`, `syncSavedSessionState`, `allSessions` exist (they are used in `onUpdateDatabase`). `allSessions()` calls `saveCurrentTester()` internally, so live widgets are flushed.

- [ ] **Step 4: Call them from the close lambdas before `closeSessions`.** Sensory lambda (~510-520):

```cpp
    connect(m_homeSensCloseBtn, &QToolButton::clicked, this, [this]() {
        if (!m_sensoryPanel) return;
        QVector<int> indices;
        for (auto* item : m_sensoryNav->selectedItems())
            indices.append(m_sensoryNav->row(item));
        if (indices.isEmpty() && m_sensoryPanel->currentSessionIndex() >= 0)
            indices.append(m_sensoryPanel->currentSessionIndex());
        if (indices.isEmpty()) return;

        const QVector<int> failed = saveSensorySessionsBeforeClose(indices);
        // Don't close sessions that failed to save (name clash / hard error).
        QVector<int> toClose;
        for (int i : indices) if (!failed.contains(i)) toClose.append(i);
        if (toClose.isEmpty()) { updateImageButton(); return; }

        m_sensoryPanel->closeSessions(toClose);
        updateImageButton();
    });
```

Detailed lambda (~544-554): identical shape using `saveDetailedSensorySessionsBeforeClose`, `m_detailedSensoryNav`, `m_detailedSensoryPanel`.

> `closeSessions` itself is left unchanged (still just removes) — the save now happens before the call. Index validity holds because no removal happens between save and close within the lambda.

- [ ] **Step 5: Build + run.**
Run: `mingw32-make -j8` in `build`; `tests\run-tests.ps1 -Filter tst_databasemanager`.
Expected: clean compile under `-Werror`; DB suite green (modulo the 3 known init.sql-drift skips).

- [ ] **Step 6: Commit.**
```bash
git add src/MainWindow.h src/MainWindow.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "fix(close): authoritatively persist a file/session before closing it (DATAVIEWER-4)"
```

---

## Task 10: Regression sweep, version bump v2.3.4, build, installer, Plane

**Files:** `DataViewerEnterprise.pro`, any test that encoded old clobber/close behavior.

- [ ] **Step 1: Full DB-suite regression + semantic check.** Run the whole `tst_databasemanager` suite; confirm no NEW failures beyond the 3 known `init.sql`-drift ones. Scan existing sensory/detailed slots for any test that UPDATEs a session with a changed score and asserts the new score via the whole-session blob — that is now (correctly) DB-authoritative; if one asserted the old clobber behavior, drive the score change through a per-cell `jsonb_set` instead, with a comment referencing DATAVIEWER-4. Do NOT weaken an assertion without understanding it.
Run: `tests\start-test-postgres.ps1` then `tests\run-tests.ps1 -Filter tst_databasemanager`

- [ ] **Step 2: Build the full Qt test suite** (catches any TU including the changed headers — `LiveSync.h`, `SensoryData.h`, `DetailedSensoryData.h`).
Run: `tests\run-tests.ps1`
Expected: green except the 3 known-unrelated.

- [ ] **Step 3: Bump VERSION** in `DataViewerEnterprise.pro` from `2.3.3` to `2.3.4` (plaintext file; careful edit or Python rewrite).

- [ ] **Step 4: Clean release build** (VERSION change needs a clean rebuild):
```
python tools\decrypt_via_copy.py --apply
cd build
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ CONFIG+=release ..\DataViewerEnterprise.pro
mingw32-make clean && mingw32-make -j8
```
Expected: `release\DataViewer.exe` FileVersion = `2.3.4.0`, warning-free under `-Werror`.

- [ ] **Step 5: Build the installer** (do NOT touch Synology):
```
tools\prepare_python_embed.bat   :: only if release\python_bundle.zip is absent
build_installer.bat              :: -> dist\DataViewer-setup.exe
```

- [ ] **Step 6: Commit the version bump.**
```bash
git add DataViewerEnterprise.pro
git commit -m "chore(release): v2.3.4 internal -- LiveSync single-source-of-truth + close-persist (DATAVIEWER-4)"
```

- [ ] **Step 7: Set DATAVIEWER-4 to *Ready for Release* in Plane** with a resolution comment: root cause (wholesale `json_data` replace + 5.0/0.0 serializer default clobbered LiveSync scores; exports + per-item close read/dropped stale in-memory state); fix (pure DB-authoritative score merge reused by the whole-session UPDATE and all exports; `LiveSync::flushNowAndWait` drain barrier at every deliberate persist point; Close now authoritatively persists the item before removal; 5 s autosave stays async for responsiveness; OCC + INSERT branches unchanged); verification (pure merge unit tests, flush-drain integration test, two-writer DB regression for both session types, full suite); v2.3.4 installer built, NOT deployed. Versioning: internal v2.3.4; wraps into v2.4.0 with DV-3 (v2.3.3) and DV-2 (v2.3.5).

---

## Risks & notes

- **Responsiveness:** the only blocking call is `flushNowAndWait`, invoked solely on deliberate actions (Ctrl+U, Close, export, program-close). It is a no-op when nothing is pending and bounded by `timeoutMs` (default 4 s) so a stalled DB degrades to "save proceeded without the last async cell" rather than a frozen UI. The 5 s autosave never blocks. Typing path is unchanged (200 ms async coalescing on the worker thread).
- **Behavioral change — Close auto-saves (no discard prompt).** Intended per the user's requirement ("Close == program-close for one file; saves must carry over"). A hard save failure or name collision aborts the close so nothing is lost. Flag in the summary.
- **Index-based sample matching** in the merge pairs `samples[i]` by position (consistent with LiveSync's own `samples[%1]` addressing). Structural add/remove/reorder stays in-memory-authoritative via the length guard. Documented limitation, not a regression.
- **JSON string-vs-number for scores:** `dve_commit_cell_json` stores scores as JSON strings ("8"); the serializer writes numbers (8). The merge copies the DB value verbatim, so a LiveSync score stays a string in the merged blob. `*FromJson` uses `.toDouble()`, which parses both — numerically safe. Full normalization is a Tier-3 / §8 follow-up, deliberately out of scope here.
- **Rename-via-INSERT** (Test-Title rename forces `id=-1` -> new row) is unchanged; the new row is seeded from the in-memory snapshot. This is the existing never-delete-the-old-row policy; the §8 follow-up revisits it.
- **OCC preserved:** merge SELECT runs in the caller's transaction; the guarded `WHERE id=? AND version=?` UPDATE and `classifyMissingUpdate` paths are unchanged. The Ctrl+U "skip — already live-synced" reconciliation still works.
- **No schema / no OfflineSnapshot change.** Snapshot regenerates from the corrected Postgres state on clean close.
