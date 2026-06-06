# DATAVIEWER-4 — Sensory scores reset to 5 (whole-session save clobbers LiveSync) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** Stop the whole-session sensory/detailed-sensory save from overwriting LiveSync's per-cell score values with the serializer's default (5.0 sensory / 0.0 detailed).

**Architecture:** Read-merge-write. The whole-session `UPDATE … json_data = CAST(? AS JSONB)` becomes non-destructive to score sub-paths: before writing, SELECT the current DB blob (same transaction) and overlay it so that, for every sample present in both, the per-metric **score keys are taken from the DB** (LiveSync is the authoritative per-cell writer). All non-score data (session metadata, sample structure, names, comments, device props) stays in-memory-authoritative. The merge is a **pure function** (two `QJsonObject`s in, merged out) so the core logic is unit-tested with no database. The existing OCC `WHERE id=? AND version=?` guard is preserved. The INSERT branch (fresh sessions, Test-Title renames forced to id=-1) is unchanged — there is no DB row to preserve.

**Why DB-authoritative (not "omit unset"):** the in-memory `SensorySample::scores` map can never express "unset" — the spin boxes default to 5.0, `SampleCard::toSample()` writes all 5 metrics unconditionally, and the (de)serializer coerces the full set. So a wholesale write always carries a 5.0 for any metric the in-memory view doesn't currently reflect (stale/remote/multi-user), stamping it over LiveSync's value. Since every UI score edit fires a LiveSync per-cell commit, the DB always holds the freshest score, so preserving DB scores is correct.

**Tech Stack:** C++17 / Qt 6.10 (QtTest, QJsonObject/QJsonArray/QJsonDocument), qmake + MinGW, PostgreSQL 16 (QPSQL). MIP: run `python tools/decrypt_via_copy.py --apply` before any build; create/modify source via plaintext (decrypt first). Build is `-Werror -Wall -Wextra -Wpedantic`.

**Behavioral change (intended, spec §4.2):** after this fix, a whole-session save **never changes a score** — scores are owned by the LiveSync per-cell path. Excel import is unaffected (it INSERTs new sessions → no merge). This is the direction of the deferred §8 "LiveSync authoritative" follow-up.

---

## File structure

- **`src/pipeline/SensoryData.h` / `.cpp`** — add pure `mergeSensoryPreservingDbScores(inMemory, dbCurrent)`. Natural home beside `sensorySessionToJson`/`fromJson`.
- **`src/pipeline/DetailedSensoryData.h` / `.cpp`** — add pure `mergeDetailedSensoryPreservingDbScores(inMemory, dbCurrent)` (uses `kDetailedAllMetrics`).
- **`src/database/DatabaseManager.cpp`** — wire the merge into the UPDATE branch of `tryWriteSensoryCore` (~1664) and `tryWriteDetailedSensoryCore` (~2126). No header change. No schema change.
- **`tests/tst_sensorydataplaceholder/`** — pure unit tests for the sensory merge (no DB).
- **`tests/tst_detailedsensoryjson/`** — pure unit tests for the detailed merge (no DB).
- **`tests/tst_databasemanager/`** — integration regression tests (two-writer clobber) for both session types.
- **`DataViewerEnterprise.pro`** — `VERSION` 2.3.3 → 2.3.4.

No `OfflineSnapshot` change: the sensory save path writes only Postgres; the SQLite snapshot regenerates from Postgres on clean close, so it inherits the corrected blob.

---

## Task 1: Pure sensory score-merge helper + unit tests

**Files:**
- Modify: `src/pipeline/SensoryData.h` (declare), `src/pipeline/SensoryData.cpp` (define)
- Test: `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp`

- [ ] **Step 1: Write failing unit tests.** Add to the test class in `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp` (it already includes `SensoryData.h`, `<QJsonObject>`, `<QJsonArray>`). Add `#include <QJsonArray>` if absent. Add four slots:

```cpp
// --- DATAVIEWER-4: whole-session save must not clobber LiveSync scores ------
// Helper: a minimal sensory blob with one sample whose 5 metrics carry `score`.
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
    // in-memory added a 2nd sample the DB doesn't have yet → keep in-memory scores.
    QJsonObject mem = oneSampleBlob("A", 5.0);
    QJsonArray ms = mem["samples"].toArray();
    QJsonObject s1; s1["name"] = "B";
    for (const QString& m : DVE::kSensoryMetrics) s1[m] = 7.0;
    ms.append(s1); mem["samples"] = ms;
    QJsonObject db = oneSampleBlob("A", 8.0);                         // only sample 0
    QJsonObject merged = DVE::mergeSensoryPreservingDbScores(mem, db);
    const QJsonArray out = merged["samples"].toArray();
    QCOMPARE(out.size(), 2);
    QCOMPARE(out[0].toObject()[DVE::kSensoryMetrics.first()].toDouble(), 8.0); // matched → DB
    QCOMPARE(out[1].toObject()[DVE::kSensoryMetrics.first()].toDouble(), 7.0); // new → in-memory
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
    QCOMPARE(m0["Burnt Taste"].toDouble(), 8.0);                      // present → DB
}
```

Register each slot under `private slots:`.

- [ ] **Step 2: Run to verify it fails (compile error: undeclared function).**

Run: `tests\run-tests.ps1 -Filter tst_sensorydataplaceholder`
Expected: build FAILS — `mergeSensoryPreservingDbScores` not declared.

- [ ] **Step 3: Declare the helper** in `src/pipeline/SensoryData.h`, in the `namespace DVE` block beside the existing serializer declarations (after line ~150):

```cpp
// DATAVIEWER-4: merge an in-memory sensory blob with the current DB blob so a
// whole-session write never clobbers LiveSync-owned per-cell SCORE values.
// Scores are DB-authoritative: for every sample present in BOTH (matched by
// array index), each kSensoryMetrics score key is taken from `dbCurrent`; all
// other keys (metadata, structure, name, comments, device props) come from
// `inMemory`. Samples in `inMemory` beyond `dbCurrent`'s array keep their
// in-memory scores (newly added, not yet in the DB). Pure / no DB.
QJsonObject mergeSensoryPreservingDbScores(const QJsonObject& inMemory,
                                           const QJsonObject& dbCurrent);
```

- [ ] **Step 4: Implement** in `src/pipeline/SensoryData.cpp` (add `#include <QJsonArray>` if not already present — it is). Add at the end of the `namespace DVE` block:

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
Expected: PASS (all existing slots + the 4 new ones).

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

> Before writing, confirm the metric-list symbol name in `DetailedSensoryData.h` — it is `kDetailedAllMetrics` (the serializer loops it at `DetailedSensoryData.cpp:42-44` with default 0.0). Use the actual symbol; if it differs, adjust both the helper and the tests.

- [ ] **Step 1: Write failing unit tests.** In `tests/tst_detailedsensoryjson/tst_detailedsensoryjson.cpp` (includes `DetailedSensoryData.h`), add `#include <QJsonArray>` if absent, and two slots mirroring Task 1 but using `DVE::kDetailedAllMetrics` and default 0.0:

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
    QJsonObject mem = oneDetailedSampleBlob("A", 0.0);   // unset → 0.0 default
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

- [ ] **Step 4: Implement** in `src/pipeline/DetailedSensoryData.cpp` (add `#include <QJsonArray>` if absent):

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

## Task 3: Wire merge into `tryWriteSensoryCore` + sensory integration regression test

**Files:**
- Modify: `src/database/DatabaseManager.cpp` — UPDATE branch of `tryWriteSensoryCore` (~1664-1719)
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp`

- [ ] **Step 1: Write the failing integration test.** In `tst_databasemanager.cpp` add a slot. It reproduces the clobber: save a session, simulate a LiveSync per-cell score edit out-of-band (which bumps `version` via the `bump_version` trigger), then re-read the bumped version into a stale (all-5.0) in-memory copy and whole-session-save it. Assert the DB keeps the LiveSync score and applies the metadata change.

```cpp
void sensoryWholeSessionSave_preservesLiveSyncScores()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
    DVE::DatabaseManager db;
    QVERIFY(openDb(db));

    // 1) INSERT a fresh session, all scores 5.0.
    DVE::SensorySession s = makeSensorySession("ClobberTest", "Charlie", "2026-06-06");
    QVERIFY(db.saveFile, "sanity");                 // (remove; placeholder)
    DVE::WriteResult wr0 = db.tryWriteSensorySession(s);   // by-ref → populates s.id, s.version
    QCOMPARE(wr0, DVE::WriteResult::Success);
    const qint64 sid = s.id;
    QVERIFY(sid > 0);

    // 2) Simulate a LiveSync per-cell commit: set sample 0 "Smoothness" = 8 via
    //    jsonb_set out-of-band. The bump_version trigger increments version.
    {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_livesync_sim";
        QSqlDatabase oob = QSqlDatabase::addDatabase("QPSQL", cname);
        oob.setHostName(cfg.host); oob.setPort(cfg.port);
        oob.setDatabaseName(cfg.database); oob.setUserName(cfg.user);
        oob.setPassword(cfg.password);
        QVERIFY(oob.open());
        QSqlQuery q(oob);
        q.prepare("UPDATE sensory_sessions "
                  "SET json_data = jsonb_set(json_data, '{samples,0,Smoothness}', '8'::jsonb, true), "
                  "    updated_by = 'livesync-sim' "
                  "WHERE id = ?");
        q.addBindValue(static_cast<qlonglong>(sid));
        QVERIFY(q.exec());
        oob.close();
        QSqlDatabase::removeDatabase(cname);
    }

    // 3) Re-read the bumped version so the whole-session UPDATE will LAND
    //    (version coincidence — the exact case that clobbers today).
    int bumpedVersion = -1;
    {
        DVE::DbConfig cfg = pgConfig();
        const QString cname = "tst_dbm_ver";
        QSqlDatabase oob = QSqlDatabase::addDatabase("QPSQL", cname);
        oob.setHostName(cfg.host); oob.setPort(cfg.port);
        oob.setDatabaseName(cfg.database); oob.setUserName(cfg.user);
        oob.setPassword(cfg.password);
        QVERIFY(oob.open());
        QSqlQuery q(oob);
        q.prepare("SELECT version FROM sensory_sessions WHERE id = ?");
        q.addBindValue(static_cast<qlonglong>(sid));
        QVERIFY(q.exec() && q.next());
        bumpedVersion = q.value(0).toInt();
        oob.close();
        QSqlDatabase::removeDatabase(cname);
    }

    // 4) Stale in-memory view: all scores 5.0, but with the CURRENT id+version
    //    and a changed metadata field (media). This is the clobber trigger.
    DVE::SensorySession stale = makeSensorySession("ClobberTest", "Charlie", "2026-06-06");
    stale.id = sid;
    stale.version = bumpedVersion;
    stale.media = "MergedMedia";
    for (const QString& m : DVE::kSensoryMetrics) stale.samples[0].scores[m] = 5.0;
    QCOMPARE(db.tryWriteSensorySession(stale), DVE::WriteResult::Success);  // UPDATE lands

    // 5) Reload and assert: Smoothness preserved at 8 (NOT reset to 5), other
    //    metrics still 5, metadata change applied.
    DVE::SensorySession loaded = db.loadSensorySession(sid);
    QCOMPARE(loaded.samples[0].scores.value("Smoothness"), 8.0);   // LiveSync value preserved
    QCOMPARE(loaded.samples[0].scores.value("Burnt Taste"), 5.0);  // untouched metric
    QCOMPARE(loaded.media, QString("MergedMedia"));                 // metadata applied
}
```

> NOTE to implementer: remove the `QVERIFY(db.saveFile, "sanity")` placeholder line — it was illustrative. Confirm helper names against the suite: `makeSensorySession(...)` exists (line ~114), `loadSensorySession(qint64)` exists (used by `testLoadSensorySessionPopulatesId`). If `loadSensorySession` is by-different-signature, use the suite's existing load helper. Match the out-of-band connection pattern already used by `bumpRowVersionOutOfBand` (line ~168).

- [ ] **Step 2: Run to verify it FAILS (reproduces the clobber).**

Run: `tests\start-test-postgres.ps1` then `tests\run-tests.ps1 -Filter tst_databasemanager`
Expected: FAIL at `QCOMPARE(loaded...Smoothness, 8.0)` — current code wrote 5.0 over it.

- [ ] **Step 3: Implement the merge in the UPDATE branch.** In `tryWriteSensoryCore` (`DatabaseManager.cpp` ~1664), ensure `#include <QJsonDocument>`, `<QJsonObject>`, `<QJsonArray>`, and `"pipeline/SensoryData.h"` are present at the top of the file (SensoryData.h already is — it declares the helper). Replace the start of the UPDATE branch so the bound blob is the merged one. Change the `q.addBindValue(jsonStr);` (the `json_data` bind, ~line 1687) to bind `jsonToWrite`, and compute `jsonToWrite` just before preparing the UPDATE:

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
        q.prepare(R"(
            UPDATE sensory_sessions SET
                ...
                json_data     = CAST(? AS JSONB),
                ...
            WHERE id = ? AND version = ?
            RETURNING id, version
        )");
        ...
        q.addBindValue(jsonToWrite);   // <-- was: q.addBindValue(jsonStr);
        ...
    }
```

(Only the `json_data` bind changes from `jsonStr` to `jsonToWrite`; all other binds/branches unchanged. The INSERT branch is untouched.)

- [ ] **Step 4: Run to verify it PASSES** + no regression in the suite.

Run: `tests\run-tests.ps1 -Filter tst_databasemanager`
Expected: the new slot PASSES; pre-existing sensory slots still pass (note: 3 unrelated failures — `testSensoryLayoutPersistence`, `testSaveSensorySessionPreservesLayoutOnReSave`, `sensoryHeaderPresets_roundTrip` — are a separate `init.sql`-drift issue, see the spawned follow-up task; they fail because the ephemeral container lacks `dve_commit_session_layout` + `sensory_header_presets`, NOT because of this change. Confirm the count of failures does not INCREASE beyond those 3.)

- [ ] **Step 5: Commit.**

```bash
git add src/database/DatabaseManager.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "fix(db): merge-preserve LiveSync scores on whole-session sensory save (DATAVIEWER-4)"
```

---

## Task 4: Wire merge into `tryWriteDetailedSensoryCore` + detailed integration test

**Files:**
- Modify: `src/database/DatabaseManager.cpp` — UPDATE branch of `tryWriteDetailedSensoryCore` (~2126-2177)
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp`

- [ ] **Step 1: Write the failing integration test** mirroring Task 3 for detailed sessions: INSERT a `DetailedSensorySession` (use the suite's `makeDetailedSensorySession` helper, line ~140), out-of-band `jsonb_set` a metric from `kDetailedAllMetrics` to a non-default value (e.g. 6), re-read version, stale-save with that metric at 0.0, reload, assert the metric stays 6 and a metadata change applied. Use `tryWriteDetailedSensorySession` + the suite's detailed load helper.

- [ ] **Step 2: Run to verify it FAILS.**

Run: `tests\run-tests.ps1 -Filter tst_databasemanager`
Expected: FAIL on the preserved-metric assertion.

- [ ] **Step 3: Implement** the same merge in the UPDATE branch of `tryWriteDetailedSensoryCore` (~2126). Add `#include "pipeline/DetailedSensoryData.h"` if not present. Compute `jsonToWrite` from a `SELECT json_data FROM detailed_sensory_sessions WHERE id = ?` and `mergeDetailedSensoryPreservingDbScores(memRoot, dbRoot)`; change the `json_data` bind (~line 2147) from `jsonStr` to `jsonToWrite`. INSERT branch untouched.

- [ ] **Step 4: Run to verify it PASSES** + suite regression check (same 3-known-unrelated-failures caveat).

Run: `tests\run-tests.ps1 -Filter tst_databasemanager`

- [ ] **Step 5: Commit.**

```bash
git add src/database/DatabaseManager.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "fix(db): merge-preserve LiveSync scores on whole-session detailed-sensory save (DATAVIEWER-4)"
```

---

## Task 5: Regression sweep, version bump v2.3.4, build, installer

**Files:** `DataViewerEnterprise.pro`, (any existing test that encoded the old clobber behavior)

- [ ] **Step 1: Full DB-suite regression + semantic check.** Run the whole `tst_databasemanager` suite and confirm no NEW failures beyond the 3 known `init.sql`-drift ones. In particular, scan the existing sensory/detailed slots for any test that **UPDATEs a session with a changed score and asserts the new score** — that would now (correctly) be DB-authoritative. If one exists and was asserting the old (clobber) behavior, update it to drive the score change through a per-cell path (`jsonb_set`) instead of the whole-session blob, with a comment referencing DATAVIEWER-4. Do NOT weaken an assertion without understanding it.

Run: `tests\start-test-postgres.ps1` then `tests\run-tests.ps1 -Filter tst_databasemanager`

- [ ] **Step 2: Build the full Qt test suite** (catch any TU that includes the changed headers) and confirm green except the 3 known-unrelated.

Run: `tests\run-tests.ps1`

- [ ] **Step 3: Bump VERSION** in `DataViewerEnterprise.pro` from `2.3.3` to `2.3.4` (via the Python delete-and-rewrite or a careful edit; the file is plaintext).

- [ ] **Step 4: Clean release build** (VERSION change needs a clean rebuild):

```
python tools\decrypt_via_copy.py --apply
qmake CONFIG+=release && mingw32-make clean && mingw32-make -j8
```
Expected: `release\DataViewer.exe` FileVersion = `2.3.4.0`, warning-free under `-Werror`.

- [ ] **Step 5: Build the installer** (do NOT touch Synology):

```
tools\prepare_python_embed.bat   :: only if release\python_bundle.zip is absent
build_installer.bat              :: → dist\DataViewer-setup.exe
```

- [ ] **Step 6: Commit the version bump.**

```bash
git add DataViewerEnterprise.pro
git commit -m "chore(release): v2.3.4 internal -- sensory LiveSync score-clobber fix (DATAVIEWER-4)"
```

- [ ] **Step 7: Set DATAVIEWER-4 to *Ready for Release* in Plane** and add a resolution comment (root cause: wholesale `json_data` replace + 5.0/0.0 serializer default clobbers LiveSync per-cell scores; fix: DB-authoritative score merge on the whole-session UPDATE, OCC guard preserved, INSERT unchanged; verification: pure merge unit tests + two-writer integration regression; commits + v2.3.4 installer NOT deployed). Versioning: internal v2.3.4; wraps into v2.4.0 with DV-3 (v2.3.3) and DV-2 (v2.3.5).

---

## Risks & notes

- **Behavioral change (intended):** scores are now LiveSync-authoritative on whole-session UPDATE. Every UI score edit already fires a per-cell LiveSync commit, so the DB holds the latest; the whole-session save no longer needs to (and must not) write scores. Excel import is unaffected (INSERT path). Flag this to Charlie in the summary.
- **Index-based sample matching:** the merge pairs `samples[i]` by position, consistent with how LiveSync itself addresses cells (`samples[%1]`, via `m_cards.indexOf`). Structural sample add/remove/reorder is in-memory-authoritative (new/removed samples handled by the length guard). Documented limitation; not a regression.
- **Narrow async window:** a score typed microseconds before a whole-session save, whose per-cell LiveSync commit hasn't landed yet, is briefly written as the DB's prior value, then corrected when the queued LiveSync commit lands. Pre-existing; fully resolved by the §8 LiveSync-authoritative follow-up.
- **OCC preserved:** the merge SELECT runs in the caller's transaction; the guarded `WHERE id=? AND version=?` UPDATE and its `classifyMissingUpdate` VersionMismatch/RowDeleted path are unchanged, so the Ctrl+U "skip — LiveSync already saved it" loop keeps working.
- **No schema / no OfflineSnapshot change.** Snapshot regenerates from the corrected Postgres state on clean close.
