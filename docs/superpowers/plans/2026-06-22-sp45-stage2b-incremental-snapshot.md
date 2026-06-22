# SP4.5 Stage 2b — Incremental Offline Snapshot + Save Progress UX — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop the offline-snapshot regen from re-copying 81.5 MB of unchanged image blobs on every close (the 7–12s stall), and show a non-frozen progress bar for the work that genuinely remains.

**Architecture:** `OfflineSnapshot::regenToPath` gains an *incremental* mode: when a valid prior `snapshot.sqlite` exists, copy it to `.tmp` (the blobs ride along), reload only the small data tables, and diff the image tables by `(id, updated_at)` so only changed blobs are pulled from Postgres — then promote with the existing atomic `MoveFileExW`. The three regen call sites that the user waits on (close, manual refresh) drive a determinate `QProgressDialog`. The chatty OCC DB-save path is deliberately untouched.

**Tech Stack:** C++17, Qt 6.10.1 (Sql/Widgets), qmake + MinGW 13.1, SQLite (QSQLITE) + Postgres (QPSQL), QtTest. Namespace `DVE`.

**Spec:** `docs/superpowers/specs/2026-06-22-sp45-stage2b-incremental-snapshot-design.md`

---

## Prerequisites (read before Task 1)

- **Branch:** work on `feature/v2.4.0-bugfix-batch` (already checked out). Do NOT start on `main`.
- **MIP:** before ANY C++ build run `python tools/decrypt_via_copy.py --apply` from the repo root. New `.cpp/.h` must be created via Python delete-and-rewrite (see CLAUDE.md) — the Edit/Write tools may MIP-label a freshly-closed source file.
- **Test DB:** `tst_offlinesnapshot` needs `dve-test-pg` (Docker, port 5433) with `deploy/postgres/init.sql` + every file in `deploy/postgres/migrations/` applied (see memory `test-container-needs-migrations`). If Docker Desktop is off, ASK the owner to start it — do not start it yourself. `tests\start-test-postgres.ps1` sets `DVE_TEST_PG_CONN` and prepends `vendor\libpq-16` to PATH.
- **Build/run one suite:** from repo root, `tests\run-tests.ps1 -Filter offlinesnapshot` builds all but runs only `tst_offlinesnapshot`. To capture QtTest totals reliably, run the built exe directly with `-o results.txt,txt` (backgrounded cmd redirection drops QtTest stdout).
- **Warnings are errors:** the build is `-Werror -Wall -Wextra -Wpedantic`. No new warnings.

## File map

| File | Responsibility | Change |
|---|---|---|
| `src/database/OfflineSnapshot.h` | snapshot API | add `RegenProgress` type, `RegenStats` struct, fingerprint-segment helpers; extend `regenToPath` + add a `regenerate(pg, progress)` overload |
| `src/database/OfflineSnapshot.cpp` | regen body | incremental mode (copy + small-table reload + image diff); progress calls; `INSERT OR REPLACE` for `_snapshot_meta` |
| `src/database/SnapshotRegenWorker.h/.cpp` | background regen | add `regenProgress(int,int,QString)` signal; pass a progress lambda into `regenToPath` |
| `src/MainWindow.h/.cpp` | UI wiring | determinate `QProgressDialog` at close + `onRefreshSnapshotTriggered`; call the new `regenerate(pg, progress)` overload |
| `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp` | tests | fingerprint-segment unit tests; incremental==full, data-only-skips-blobs, image diff, schema-mismatch fallback, failure-leaves-prod-intact |
| `DataViewerEnterprise.pro` | version | `VERSION = 2.4.6` → `2.4.7` (Task 6 only) |

## Phase-count constant for progress

The regen reports against a fixed phase budget. Define near the top of `OfflineSnapshot.cpp`:

```cpp
// Progress phases: 1 (copy/create) + 10 table refreshes + 1 (meta) + 1 (checkpoint/promote).
static constexpr int kRegenPhases = 13;
```

---

### Task 1: Progress plumbing through regenToPath (additive, no behavior change)

**Files:**
- Modify: `src/database/OfflineSnapshot.h` (declarations)
- Modify: `src/database/OfflineSnapshot.cpp` (progress calls in the existing full path)
- Modify: `src/database/SnapshotRegenWorker.h` / `.cpp` (signal + lambda)
- Test: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`

- [ ] **Step 1: Add the progress type, RegenStats, and extend signatures in the header.**

In `OfflineSnapshot.h`, just above `regenToPath`'s declaration (after the `std::function` include already present):

```cpp
    // SP4.5 Stage 2b: regen progress callback (done, total, phase label). Null-safe.
    using RegenProgress = std::function<void(int done, int total, const QString& phase)>;

    // SP4.5 Stage 2b: per-run regen diagnostics (test/log visibility). wasIncremental
    // is true when the prior snapshot's blobs were reused; imageRowsPulledFromPg counts
    // blob rows actually fetched from PG (0 on a data-only change == the win).
    struct RegenStats {
        bool wasIncremental        = false;
        int  smallTablesReloaded   = 0;
        int  imageRowsPulledFromPg = 0;
    };
```

Change the `regenToPath` declaration to append two defaulted params:

```cpp
    static bool regenToPath(PostgresConnection* live,
                            const QString&      destPath,
                            std::atomic<bool>*  cancel,
                            QString*            outFingerprint,
                            QDateTime*          outServerTimeUtc,
                            QString*            outError,
                            const RegenProgress& progress = {},
                            RegenStats*         outStats  = nullptr);
```

Add a progress-aware instance overload next to `bool regenerate(PostgresConnection* live);`:

```cpp
    // SP4.5 Stage 2b: same as regenerate() but drives a progress callback (used by
    // the close + manual-refresh progress dialog). The no-arg overload calls this
    // with an empty callback.
    bool regenerate(PostgresConnection* live, const RegenProgress& progress);
```

- [ ] **Step 2: Write a failing test that the callback is invoked.**

Add to `tst_offlinesnapshot.cpp` (in the existing PG-gated section — guard with the same `DVE_TEST_PG_CONN` skip the other DB tests use):

```cpp
void TestOfflineSnapshot::regenProgress_invokedToCompletion()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN"))
        QSKIP("DVE_TEST_PG_CONN not set");
    OfflineSnapshot snap;
    snap.setOverrideDirForTesting(m_tmpDir.path());
    int lastDone = -1, lastTotal = -1;
    auto cb = [&](int d, int t, const QString&) { lastDone = d; lastTotal = t; };
    QVERIFY(snap.regenerate(m_pg, cb));      // m_pg = the test PostgresConnection
    QCOMPARE(lastTotal, 13);                 // kRegenPhases
    QCOMPARE(lastDone, lastTotal);           // ends at done==total
}
```

- [ ] **Step 3: Run it — expect a COMPILE failure** (the `regenerate(pg, cb)` overload doesn't exist yet).

Run: `tests\run-tests.ps1 -Filter offlinesnapshot`
Expected: build error `no matching function for call to regenerate`.

- [ ] **Step 4: Implement the progress plumbing.**

In `OfflineSnapshot.cpp`:
- Add `static constexpr int kRegenPhases = 13;` near the top (below the includes).
- In `regenToPath`, add a tiny helper at the top of the function body and call it at each phase boundary:

```cpp
    int _phase = 0;
    auto tick = [&](const QString& label) {
        if (progress) progress(++_phase, kRegenPhases, label);
    };
```

Insert `tick("Copying files");`, `tick("Copying tests");`, … before each existing per-table block (10 calls, matching the existing table order), `tick("Writing metadata");` before the `_snapshot_meta` block, and `tick("Finalizing");` before the `wal_checkpoint`. Leave the first phase for the create/copy step added in Task 3 (`tick("Preparing snapshot");` at the very top, right after the `tick` lambda is defined).
- Make the two new params real: the function already returns via `outError`; add `if (outStats) outStats->wasIncremental = false;` near the top (Task 3 flips it true on the incremental path).
- Implement the overloads at the bottom near the existing `regenerate`:

```cpp
bool OfflineSnapshot::regenerate(PostgresConnection* live) {
    return regenerate(live, RegenProgress{});
}

bool OfflineSnapshot::regenerate(PostgresConnection* live, const RegenProgress& progress) {
    m_lastError.clear();
    if (!live || !live->isOpen()) {
        m_lastError = QStringLiteral("regenerate: PostgresConnection is not open");
        return false;
    }
    if (m_open) close();                     // release our read handle before promote
    QString fp, err; QDateTime srv;
    const bool ok = regenToPath(live, path(), /*cancel*/nullptr, &fp, &srv, &err, progress);
    if (ok && srv.isValid()) m_lastRegenServerTime = srv;
    if (!ok) m_lastError = err;
    return ok;
}
```

(The old single-arg body moves into the new overload; the single-arg one now delegates.)

In `SnapshotRegenWorker.h`, add to `signals:`:

```cpp
    void regenProgress(int done, int total, QString phase);
```

In `SnapshotRegenWorker.cpp::requestRegen`, change the `regenToPath` call to pass a progress lambda (emit is thread-safe — Qt queues it to connected UI slots):

```cpp
        QString fp, err;
        OfflineSnapshot::RegenStats stats;
        const bool ok = DVE::OfflineSnapshot::regenToPath(
            m_pg, m_productionPath, &m_cancel, &fp, /*outServerTimeUtc*/nullptr, &err,
            [this](int d, int t, const QString& ph){ emit regenProgress(d, t, ph); },
            &stats);
```

(Logging `stats.wasIncremental` / `stats.imageRowsPulledFromPg` in the existing `qInfo` line is a nice-to-have — add `<< "incremental:" << stats.wasIncremental << "blobsPulled:" << stats.imageRowsPulledFromPg`.)

- [ ] **Step 5: Run the test — expect PASS, and the existing 23 still green.**

Run: `tests\run-tests.ps1 -Filter offlinesnapshot`
Expected: `Totals: 24 passed, 0 failed` (23 prior + the new one).

- [ ] **Step 6: Commit.**

```bash
git add src/database/OfflineSnapshot.h src/database/OfflineSnapshot.cpp \
        src/database/SnapshotRegenWorker.h src/database/SnapshotRegenWorker.cpp \
        tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp
git commit -m "feat(snapshot): regen progress callback + RegenStats (SP4.5 2b T1)"
```

---

### Task 2: Fingerprint segment helpers (pure C++, no DB)

The fingerprint string is 9 `count/epoch` segments joined by `;` in this fixed order (see `snapshotContentFingerprint`, `OfflineSnapshot.cpp:370-387`): `files, tests, samples, data_rows, images, sensory_sessions, sensory_images, detailed_sensory_sessions, detailed_sensory_images`.

**Files:**
- Modify: `src/database/OfflineSnapshot.h` (declare two statics + the table-order array)
- Modify: `src/database/OfflineSnapshot.cpp` (implement)
- Test: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`

- [ ] **Step 1: Write failing unit tests (no DB needed — do NOT QSKIP these).**

```cpp
void TestOfflineSnapshot::fingerprint_segmentChanged_detectsPerTable()
{
    // 9 segments, fixed order. Identical -> nothing changed.
    const QString a = "1/100;2/200;3/300;4/400;5/500;6/600;7/700;8/800;9/900";
    QString b = a;
    QVERIFY(!OfflineSnapshot::segmentChanged(a, b, "images"));      // images = idx 4
    // Change only the images segment (idx 4).
    b = "1/100;2/200;3/300;4/400;5/999;6/600;7/700;8/800;9/900";
    QVERIFY( OfflineSnapshot::segmentChanged(a, b, "images"));
    QVERIFY(!OfflineSnapshot::segmentChanged(a, b, "data_rows"));   // idx 3 unchanged
    // Malformed / arity drift -> treat as changed (safe: forces refresh).
    QVERIFY(OfflineSnapshot::segmentChanged("garbage", b, "images"));
    QVERIFY(OfflineSnapshot::segmentChanged(QString(), b, "images"));
}
```

- [ ] **Step 2: Run — expect COMPILE failure** (`segmentChanged` undeclared).

Run: `tests\run-tests.ps1 -Filter offlinesnapshot`
Expected: build error.

- [ ] **Step 3: Implement the helpers.**

In `OfflineSnapshot.h` (public statics):

```cpp
    // SP4.5 Stage 2b: the fingerprint table order (matches snapshotContentFingerprint).
    static const char* const kFingerprintTables[9];
    // Split a fingerprint into its `;`-separated segments. Empty list if the
    // count != 9 (malformed / schema drift).
    static QStringList fingerprintSegments(const QString& fp);
    // True if `table`'s segment differs between prior and live, OR either string
    // is unparseable (the safe default: refresh). `table` must be one of
    // kFingerprintTables; an unknown name also returns true.
    static bool segmentChanged(const QString& priorFp, const QString& liveFp,
                               const char* table);
```

In `OfflineSnapshot.cpp`:

```cpp
const char* const OfflineSnapshot::kFingerprintTables[9] = {
    "files", "tests", "samples", "data_rows", "images",
    "sensory_sessions", "sensory_images",
    "detailed_sensory_sessions", "detailed_sensory_images"
};

QStringList OfflineSnapshot::fingerprintSegments(const QString& fp) {
    const QStringList parts = fp.split(';');
    return parts.size() == 9 ? parts : QStringList{};
}

bool OfflineSnapshot::segmentChanged(const QString& priorFp, const QString& liveFp,
                                     const char* table) {
    const QStringList a = fingerprintSegments(priorFp);
    const QStringList b = fingerprintSegments(liveFp);
    if (a.isEmpty() || b.isEmpty()) return true;            // unparseable -> refresh
    int idx = -1;
    for (int i = 0; i < 9; ++i)
        if (qstrcmp(kFingerprintTables[i], table) == 0) { idx = i; break; }
    if (idx < 0) return true;                                // unknown table -> refresh
    return a[idx] != b[idx];
}
```

- [ ] **Step 4: Run — expect PASS** (`Totals: 25 passed, 0 failed`).

- [ ] **Step 5: Commit.**

```bash
git add src/database/OfflineSnapshot.h src/database/OfflineSnapshot.cpp \
        tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp
git commit -m "feat(snapshot): per-table fingerprint segment compare (SP4.5 2b T2)"
```

---

### Task 3: Incremental mode — copy prior + reload small tables (image tables still full-copied)

This task makes incremental mode *correct* (output identical to a full rebuild) without the blob optimization yet — a safe intermediate. The image tables are still reloaded in full here; Task 4 makes them skip/diff.

**Files:**
- Modify: `src/database/OfflineSnapshot.cpp` (`regenToPath`)
- Test: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`

- [ ] **Step 1: Write failing tests — incremental==full + schema-mismatch fallback.**

```cpp
// Helper in the test: SHA/rowcount signature of a snapshot table.
static QString tableSig(const QString& dbPath, const QString& table); // SELECT count(*), sum/hash; impl below

void TestOfflineSnapshot::incremental_matchesFullRebuild()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no PG");
    OfflineSnapshot snap; snap.setOverrideDirForTesting(m_tmpDir.path());
    QVERIFY(snap.regenerate(m_pg));                 // full build #1 -> prod exists
    bumpOneDataRowInPg(m_pg);                       // mutate a data_row (helper)
    OfflineSnapshot::RegenStats st; QString fp, err; QDateTime srv;
    QVERIFY(OfflineSnapshot::regenToPath(m_pg, snap.path(), nullptr, &fp, &srv, &err, {}, &st));
    QVERIFY(st.wasIncremental);                     // reused the prior snapshot
    // Build a fresh full snapshot of the SAME post-mutation state to a 2nd path.
    const QString full2 = m_tmpDir.path() + "/full2.sqlite";
    QVERIFY(OfflineSnapshot::regenToPath(m_pg, full2, nullptr, &fp, &srv, &err, {}, nullptr));
    for (const char* t : OfflineSnapshot::kFingerprintTables)
        QCOMPARE(tableSig(snap.path(), t), tableSig(full2, t));   // identical content
}

void TestOfflineSnapshot::incremental_fallsBackOnSchemaMismatch()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no PG");
    OfflineSnapshot snap; snap.setOverrideDirForTesting(m_tmpDir.path());
    QVERIFY(snap.regenerate(m_pg));
    corruptSchemaVersionInSnapshot(snap.path());    // set source_schema_version=0 (helper)
    OfflineSnapshot::RegenStats st; QString fp, err; QDateTime srv;
    QVERIFY(OfflineSnapshot::regenToPath(m_pg, snap.path(), nullptr, &fp, &srv, &err, {}, &st));
    QVERIFY(!st.wasIncremental);                    // forced full rebuild
}
```

Add the three test helpers (`tableSig`, `bumpOneDataRowInPg`, `corruptSchemaVersionInSnapshot`) — `tableSig` opens the SQLite file read-only and returns `QString::number(count) + ":" + group_concat-based hash` (use `SELECT count(*)||'|'||coalesce(group_concat(id||updated_at),'') FROM <t>` for a deterministic content signature that does not depend on blob equality for the non-image tables; for image tables include `length(image_data)`):

```cpp
static QString tableSig(const QString& dbPath, const QString& table) {
    const QString conn = "sig_" + QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    QString out;
    { QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", conn);
      db.setDatabaseName(dbPath); db.setConnectOptions("QSQLITE_OPEN_READONLY");
      QVERIFY2(db.open(), qPrintable(db.lastError().text()));
      QSqlQuery q(db);
      const bool hasBlob = (table == "images" || table == "sensory_images"
                            || table == "detailed_sensory_images");
      const QString expr = hasBlob
        ? "count(*)||'|'||coalesce(group_concat(id||':'||updated_at||':'||length(coalesce(image_data,''))),'')"
        : "count(*)||'|'||coalesce(group_concat(id||':'||updated_at),'')";
      if (q.exec("SELECT " + expr + " FROM " + table) && q.next()) out = q.value(0).toString();
      db.close(); }
    QSqlDatabase::removeDatabase(conn);
    return out;
}
```

- [ ] **Step 2: Run — expect FAIL** (`st.wasIncremental` is false; `regenToPath` always builds full today, and a second full build over an existing prod path works but `wasIncremental` stays false).

- [ ] **Step 3: Implement incremental mode (copy + small-table reload via a lambda).**

In `regenToPath`, after computing `prodPath`/`tmpPath`/`tmpConn` and BEFORE the `QFile::remove(tmpPath...)` wipe, add the **mode decision** (reads the prior snapshot's meta read-only):

```cpp
    // ---- SP4.5 Stage 2b: decide incremental vs full ------------------------
    QString priorFp;
    bool incremental = false;
    if (QFile::exists(prodPath)) {
        const QString roConn = QStringLiteral("dve_snap_ro_") +
            QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
        {
            QSqlDatabase ro = QSqlDatabase::addDatabase("QSQLITE", roConn);
            ro.setDatabaseName(prodPath);
            ro.setConnectOptions("QSQLITE_OPEN_READONLY");
            if (ro.open()) {
                int ver = -1;
                QSqlQuery vq(ro);
                if (vq.exec("SELECT value FROM _snapshot_meta WHERE key='source_schema_version'")
                    && vq.next())
                    ver = vq.value(0).toInt();
                QSqlQuery fq(ro);
                if (ver == kSnapshotSchemaVersion
                    && fq.exec("SELECT value FROM _snapshot_meta WHERE key='content_fingerprint'")
                    && fq.next())
                    priorFp = fq.value(0).toString();
                ro.close();
            }
        }
        QSqlDatabase::removeDatabase(roConn);
        incremental = !priorFp.isEmpty();
    }
    if (outStats) outStats->wasIncremental = incremental;
```

Replace the tmp-build block. Today it is: wipe `.tmp`, `tmpDb.open()`, pragmas, then the `for (kCreateStatements)` schema loop. Change to:

```cpp
    QFile::remove(tmpPath); QFile::remove(tmpPath + "-wal"); QFile::remove(tmpPath + "-shm");
    if (incremental) {
        // Reuse the prior snapshot's blobs wholesale; patch only what changed.
        if (!QFile::copy(prodPath, tmpPath)) {
            incremental = false;                 // copy failed -> safe full rebuild
            if (outStats) outStats->wasIncremental = false;
        }
    }
    // (open tmpDb on tmpPath as today; set WAL + synchronous=FULL as today)
    ...
    if (!incremental) {
        for (const char* stmt : kCreateStatements) { /* existing schema-create loop */ }
    }
```

Add a **small-table reload lambda** after `tmpDb.transaction()` succeeds and after the cancel lambda is defined. It wraps the existing SELECT→bind→exec pattern and prefixes a `DELETE` in incremental mode:

```cpp
    auto bail = [&](const QString& where, const QSqlError& e) -> bool {
        m_lastError = QStringLiteral("regenerate(%1): ").arg(where) + e.text();
        tmpDb.rollback(); rollbackPg(); tmpDb.close();
        tmpDb = QSqlDatabase(); QSqlDatabase::removeDatabase(tmpConn); cleanup();
        return false;
    };

    auto reloadSmall = [&](const char* table, const char* selectSql,
                           const char* insertSql, int expectCols,
                           const QString& label) -> bool {
        tick(label);
        if (incremental) {
            QSqlQuery del(tmpDb);
            if (!del.exec(QStringLiteral("DELETE FROM %1").arg(table)))
                return bail(QStringLiteral("DELETE %1").arg(table), del.lastError());
        }
        QSqlQuery src(pg);
        if (!src.exec(selectSql)) return bail(QStringLiteral("SELECT %1").arg(table), src.lastError());
        const int kCols = src.record().count();
        assertColumnArity(src, expectCols, kCols, table);
        QSqlQuery dst(tmpDb); dst.prepare(insertSql);
        while (src.next()) {
            for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
            if (!dst.exec()) return bail(QStringLiteral("INSERT %1").arg(table), dst.lastError());
        }
        if (outStats) outStats->smallTablesReloaded++;
        return true;
    };
```

Now convert the **seven small-table inline blocks** (files, tests, samples, data_rows, sensory_sessions, detailed_sensory_sessions, settings) to single calls, reusing each block's exact SELECT/INSERT SQL and arity (copy the literal SQL strings from the existing blocks — files `:558`, tests `:594`, samples `:625`, data_rows `:666`, sensory_sessions `:735`, detailed_sensory_sessions `:807`, settings `:877`). Example for files:

```cpp
    if (!reloadSmall("files",
            "SELECT id, file_path, file_name, loaded_at, template_version, "
            "sheet_count, sample_count, added_at, updated_at, updated_by, version, "
            "app_version FROM files ORDER BY id",
            "INSERT INTO files (id, file_path, file_name, loaded_at, template_version, "
            "sheet_count, sample_count, added_at, updated_at, updated_by, version, app_version) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            12, QStringLiteral("Copying files"))) return false;
```

Leave the **three image-table blocks** (images `:700`, sensory_images `:772`, detailed_sensory_images `:842`) as-is FOR NOW, but make them incremental-safe by prefixing a DELETE so the copied rows don't collide on re-INSERT. Wrap each existing image block's `QSqlQuery dst(tmpDb); dst.prepare("INSERT ...")` with:

```cpp
        if (incremental) {
            QSqlQuery del(tmpDb);
            if (!del.exec("DELETE FROM images")) return bail("DELETE images", del.lastError());
        }
```

(Repeat for `sensory_images` / `detailed_sensory_images`.) This keeps Task 3 correct; Task 4 replaces these three with skip/diff.

Change the three `_snapshot_meta` writes from `INSERT INTO` to `INSERT OR REPLACE INTO _snapshot_meta` (the copied tmp already has those keys; `key` is the table's primary key so a plain INSERT would collide in incremental mode).

- [ ] **Step 4: Run — expect PASS** (`incremental_matchesFullRebuild`, `incremental_fallsBackOnSchemaMismatch`, and all prior green). `Totals: 27 passed, 0 failed`.

- [ ] **Step 5: Commit.**

```bash
git add src/database/OfflineSnapshot.cpp tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp
git commit -m "feat(snapshot): incremental regen — copy prior + reload small tables (SP4.5 2b T3)"
```

---

### Task 4: Image-table skip / per-row diff (the actual blob win)

**Files:**
- Modify: `src/database/OfflineSnapshot.cpp` (`regenToPath` image blocks)
- Test: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`

- [ ] **Step 1: Write failing tests — data-only skips blobs; image change pulls only changed.**

```cpp
void TestOfflineSnapshot::dataOnlyChange_pullsZeroImageBlobs()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no PG");
    OfflineSnapshot snap; snap.setOverrideDirForTesting(m_tmpDir.path());
    seedImages(m_pg, 3);                            // 3 image rows (helper)
    QVERIFY(snap.regenerate(m_pg));                 // full build
    const QString before = tableSig(snap.path(), "images");
    bumpOneDataRowInPg(m_pg);                       // data-only change
    OfflineSnapshot::RegenStats st; QString fp, err; QDateTime srv;
    QVERIFY(OfflineSnapshot::regenToPath(m_pg, snap.path(), nullptr, &fp, &srv, &err, {}, &st));
    QCOMPARE(st.imageRowsPulledFromPg, 0);          // THE WIN: no blobs re-read
    QCOMPARE(tableSig(snap.path(), "images"), before); // images byte-identical
}

void TestOfflineSnapshot::imageChange_pullsOnlyChanged()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no PG");
    OfflineSnapshot snap; snap.setOverrideDirForTesting(m_tmpDir.path());
    seedImages(m_pg, 3);
    QVERIFY(snap.regenerate(m_pg));
    updateOneImageInPg(m_pg);                       // change 1 image (new bytes + updated_at)
    OfflineSnapshot::RegenStats st; QString fp, err; QDateTime srv;
    QVERIFY(OfflineSnapshot::regenToPath(m_pg, snap.path(), nullptr, &fp, &srv, &err, {}, &st));
    QCOMPARE(st.imageRowsPulledFromPg, 1);          // only the changed one
    // And the result matches a full rebuild of the same state.
    const QString full2 = m_tmpDir.path() + "/full2.sqlite";
    QVERIFY(OfflineSnapshot::regenToPath(m_pg, full2, nullptr, &fp, &srv, &err, {}, nullptr));
    QCOMPARE(tableSig(snap.path(), "images"), tableSig(full2, "images"));
}
```

Add helpers `seedImages(pg, n)`, `updateOneImageInPg(pg)` (UPDATE one row's `image_data` + `updated_at = now()`), and a delete variant if covering deletes.

- [ ] **Step 2: Run — expect FAIL** (`imageRowsPulledFromPg` is 0 because nothing increments it yet; the data-only test may pass accidentally only if blobs happen to match — the image-change test fails the `== 1`).

- [ ] **Step 3: Implement the image diff lambda; replace the three image blocks.**

Add after `reloadSmall`:

```cpp
    QString liveFp;            // set just before the table refreshes (Step 3b)
    auto refreshImages = [&](const char* table, const char* selectAllSql,
                             const char* insertSql, int expectCols,
                             const QString& label) -> bool {
        tick(label);
        // Full rebuild: straight copy (empty tmp), same as a small table.
        if (!incremental)
            return reloadSmall(table, selectAllSql, insertSql, expectCols, label);
        // Unchanged segment -> the copied blobs are already correct. Skip entirely.
        if (!segmentChanged(priorFp, liveFp, table)) return true;
        // Changed: diff by (id, updated_at); pull blobs ONLY for new/changed ids.
        QHash<qint64, QString> pgRows, tmpRows;
        { QSqlQuery q(pg);
          if (!q.exec(QStringLiteral("SELECT id, updated_at FROM %1").arg(table)))
              return bail(QStringLiteral("diff-pg %1").arg(table), q.lastError());
          while (q.next()) pgRows.insert(q.value(0).toLongLong(), q.value(1).toString()); }
        { QSqlQuery q(tmpDb);
          if (!q.exec(QStringLiteral("SELECT id, updated_at FROM %1").arg(table)))
              return bail(QStringLiteral("diff-tmp %1").arg(table), q.lastError());
          while (q.next()) tmpRows.insert(q.value(0).toLongLong(), q.value(1).toString()); }
        // Deletes: rows in tmp no longer in pg.
        for (auto it = tmpRows.constBegin(); it != tmpRows.constEnd(); ++it) {
            if (!pgRows.contains(it.key())) {
                QSqlQuery d(tmpDb);
                d.prepare(QStringLiteral("DELETE FROM %1 WHERE id = ?").arg(table));
                d.addBindValue(it.key());
                if (!d.exec()) return bail(QStringLiteral("diff-del %1").arg(table), d.lastError());
            }
        }
        // To pull: ids new in pg, or whose updated_at differs.
        QVariantList toPull;
        for (auto it = pgRows.constBegin(); it != pgRows.constEnd(); ++it)
            if (!tmpRows.contains(it.key()) || tmpRows.value(it.key()) != it.value())
                toPull << it.key();
        if (toPull.isEmpty()) return true;
        QStringList ph; for (int i = 0; i < toPull.size(); ++i) ph << "?";
        QSqlQuery src(pg);
        src.prepare(QString(selectAllSql).replace("ORDER BY id",
                    QStringLiteral("WHERE id IN (%1) ORDER BY id").arg(ph.join(","))));
        for (const QVariant& id : toPull) src.addBindValue(id);
        if (!src.exec()) return bail(QStringLiteral("diff-sel %1").arg(table), src.lastError());
        const int kCols = src.record().count();
        assertColumnArity(src, expectCols, kCols, table);
        QSqlQuery del(tmpDb); del.prepare(QStringLiteral("DELETE FROM %1 WHERE id = ?").arg(table));
        QSqlQuery dst(tmpDb); dst.prepare(insertSql);
        while (src.next()) {
            const QVariant id = src.value(0);
            del.addBindValue(id);
            if (!del.exec()) return bail(QStringLiteral("diff-rep-del %1").arg(table), del.lastError());
            for (int c = 0; c < kCols; ++c) dst.bindValue(c, src.value(c));
            if (!dst.exec()) return bail(QStringLiteral("diff-ins %1").arg(table), dst.lastError());
            if (outStats) outStats->imageRowsPulledFromPg++;
        }
        return true;
    };
```

**Step 3b:** compute `liveFp` once, before the table refreshes (it is the same string `snapshotContentFingerprint(pg)` computes; the existing code already calls that at `:925` for the meta — hoist a single call up so both the diff and the meta use it):

```cpp
    liveFp = snapshotContentFingerprint(pg);     // inside the REPEATABLE READ txn
```

Place this right after `tmpDb.transaction()` / cancel-lambda, before the first table refresh. Then reuse `liveFp` for the `contentFp` used in the meta block (replace the `:925` re-computation with `const QString contentFp = liveFp;`).

Replace the three image blocks (now DELETE+INSERT from Task 3) with:

```cpp
    if (cancelled()) return false;
    if (!refreshImages("images",
            "SELECT id, sample_id, sort_order, file_name, image_data, layout_x, layout_y, "
            "layout_w, layout_h, crop_x, crop_y, crop_w, crop_h, updated_at, updated_by, version "
            "FROM images ORDER BY id",
            "INSERT INTO images (id, sample_id, sort_order, file_name, image_data, layout_x, "
            "layout_y, layout_w, layout_h, crop_x, crop_y, crop_w, crop_h, updated_at, "
            "updated_by, version) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            16, QStringLiteral("Copying images"))) return false;
```

(Repeat for `sensory_images` and `detailed_sensory_images`, using each one's exact 16-column SELECT/INSERT from the existing blocks at `:772` and `:842`. Keep the `cancelled()` checks that currently guard the blob tables.)

- [ ] **Step 4: Run — expect PASS** (`dataOnlyChange_pullsZeroImageBlobs`, `imageChange_pullsOnlyChanged`, all prior). `Totals: 29 passed, 0 failed`.

- [ ] **Step 5: Add a failure-leaves-prod-intact test + run.**

```cpp
void TestOfflineSnapshot::regenFailure_leavesProdIntact()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no PG");
    OfflineSnapshot snap; snap.setOverrideDirForTesting(m_tmpDir.path());
    QVERIFY(snap.regenerate(m_pg));
    const QString sig = tableSig(snap.path(), "files");
    m_pg->close();                                  // force every PG read to fail mid-regen
    QString fp, err; QDateTime srv;
    QVERIFY(!OfflineSnapshot::regenToPath(m_pg, snap.path(), nullptr, &fp, &srv, &err, {}, nullptr));
    QCOMPARE(tableSig(snap.path(), "files"), sig);  // prod unchanged; .tmp cleaned up
    QVERIFY(!QFile::exists(snap.path() + ".tmp"));
    reopenPg(m_pg);                                 // restore for later tests
}
```

Run: `tests\run-tests.ps1 -Filter offlinesnapshot` → `Totals: 30 passed, 0 failed`.

- [ ] **Step 6: Commit.**

```bash
git add src/database/OfflineSnapshot.cpp tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp
git commit -m "feat(snapshot): image-table skip/diff — pull only changed blobs (SP4.5 2b T4)"
```

---

### Task 5: Determinate progress dialog at the two call sites the user waits on

**Files:**
- Modify: `src/MainWindow.cpp` (`onRefreshSnapshotTriggered`, `closeEvent`)
- Modify: `src/MainWindow.h` (none expected; `QProgressDialog` already included)

- [ ] **Step 1: Upgrade `onRefreshSnapshotTriggered` (MainWindow.cpp:6283) to a determinate bar.**

Replace the indeterminate dialog + synchronous `regenerate` with a determinate one driven by the progress callback:

```cpp
    QProgressDialog progress(tr("Refreshing offline snapshot..."), QString(), 0, 100, this);
    progress.setWindowModality(Qt::WindowModal);
    progress.setCancelButton(nullptr);
    progress.setMinimumDuration(0);
    progress.setValue(0);
    QApplication::processEvents();

    const bool ok = m_snapshot->regenerate(m_pgConn,
        [&progress](int done, int total, const QString& phase) {
            progress.setLabelText(phase);
            progress.setValue(total > 0 ? (done * 100 / total) : 0);
            QApplication::processEvents();
        });
    progress.setValue(100);
    progress.close();
```

(The `if (ok) … else …` message boxes stay as-is.)

- [ ] **Step 2: In `closeEvent` (MainWindow.cpp:6457-6468), show the same dialog around the synchronous regen.**

Replace the `else { regenerating … m_snapshot->regenerate(m_pgConn) … }` branch with:

```cpp
        } else {
            qInfo() << "[perf] closeEvent: regenerating offline snapshot (DB changed)";
            QProgressDialog progress(tr("Saving offline copy..."), QString(), 0, 100, this);
            progress.setWindowModality(Qt::WindowModal);
            progress.setCancelButton(nullptr);
            progress.setMinimumDuration(0);   // show immediately; common case is <1s
            progress.setValue(0);
            QApplication::processEvents();
            const bool ok = m_snapshot->regenerate(m_pgConn,
                [&progress](int done, int total, const QString& phase) {
                    progress.setLabelText(phase);
                    progress.setValue(total > 0 ? (done * 100 / total) : 0);
                    QApplication::processEvents();
                });
            if (!ok)
                qWarning() << "Snapshot regenerate failed:" << m_snapshot->lastError();
            progress.setValue(100);
            qInfo() << "[perf] closeEvent: snapshot regen done";
        }
```

This keeps the existing `[perf]` markers (so the owner's log still measures the close) and the surrounding `regenWasInFlight` / `isCurrentVsLive` guards untouched. The regen now runs incrementally, so this branch is sub-second on a data-only close and animates per-blob on the rare image change.

- [ ] **Step 3: Build the full app (not just the test) to confirm the wiring compiles `-Werror`.**

Run (from repo root, the rebuild-dataviewer build line, debug is fine for this check):
```
python tools/decrypt_via_copy.py --apply
cd build && set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro && mingw32-make -j4
```
Expected: links to `DataViewer.exe`, no warnings/errors.

- [ ] **Step 4: Commit.**

```bash
git add src/MainWindow.cpp
git commit -m "feat(ui): determinate Saving... progress dialog at close + manual refresh (SP4.5 2b T5)"
```

---

### Task 6: v2.4.7 internal build + full-suite verification

**Files:**
- Modify: `DataViewerEnterprise.pro` (`VERSION`)
- Create: `release_overview/release_overview_v_2_4_7.txt`

- [ ] **Step 1: Run the full test suite first (gate before bumping).**

Run: `tests\run-tests.ps1` (ensure `dve-test-pg` is up with migrations). Expected: all green except the two known pre-existing fails (`tst_excelexporter` Excel round-trip; `tst_storedfns` if the container schema is stale). `tst_offlinesnapshot` = 30/0. The save e2e suites (`tst_saveintegrity_e2e`, `tst_twoclient_e2e`) must stay green (no save-path regression).

- [ ] **Step 2: Bump VERSION and do the clean release rebuild + installer via the rebuild-dataviewer flow.**

Edit `DataViewerEnterprise.pro`: `VERSION = 2.4.6` → `VERSION = 2.4.7`. Then:
```
python tools/decrypt_via_copy.py --apply
cd "C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise" && cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe -spec win32-g++ DataViewerEnterprise.pro CONFIG+=release && mingw32-make.exe clean && mingw32-make.exe -j8 release && .\build_installer.bat"
```

- [ ] **Step 3: Verify the version surfaces.**

Run: `powershell -Command "(Get-Item 'release\DataViewer.exe').VersionInfo | Format-List FileVersion,ProductVersion"` → `2.4.7.0`.
Run: `powershell -Command "(Get-Item 'dist\DataViewer-setup.exe').VersionInfo | Format-List ProductVersion"` → `2.4.7`.

- [ ] **Step 4: Write `release_overview/release_overview_v_2_4_7.txt`** (MIP-safe: rewrite via Python, then `git add` in one command — git reads the plaintext before MIP re-labels). Content: a **Performance** section — "Closing the app after editing is now near-instant: the offline backup copy reuses the images it already saved instead of re-copying them every time, and a progress bar now shows when a save is genuinely in progress." Derive specifics from `git log --oneline 4772879..HEAD`.

- [ ] **Step 5: Commit the version bump + overview.**

```bash
git add DataViewerEnterprise.pro release_overview/release_overview_v_2_4_7.txt
git commit -m "chore(release): v2.4.7 internal — SP4.5 Stage 2b incremental snapshot"
git tag v2.4.7-internal
```

- [ ] **Step 6: Report to the owner** — installer at `dist\DataViewer-setup.exe` (2.4.7), the smoke-test script (edit → Update DB → wait 30s → close should be near-instant; add an image → close shows a brief animated "Saving…" bar; `dataviewer.log` shows `incremental:true blobsPulled:0` on data-only closes). Do NOT push/merge/Synology — that is the owner's gate.

---

## Self-Review

**1. Spec coverage:**
- Incremental copy+patch+promote → Tasks 3 (copy + small-table reload) + 4 (image skip/diff); promotion path untouched. ✓
- Change detection via fingerprint segments → Task 2. ✓
- Progress UX (non-frozen dialog) → Task 1 (callback) + Task 5 (dialogs at the two waited-on sites). ✓
- DB-save rewrite explicitly out of scope → no task touches `tryWriteFile`. ✓
- Crash-safety (atomic promote, full-rebuild fallback) → Task 3 mode decision + fallback test; promotion code unchanged. ✓
- Data-loss guard (close drains persist) → untouched; `closeEvent` edits are confined to the regen branch (Task 5). ✓
- Thread-safety → Task 1 worker keeps its own connection; the read-only meta probe in Task 3 uses its own throwaway connection. ✓
- Tests: unchanged-skip (existing `isCurrentVsLive`), data-only-zero-blobs (T4), image diff (T4), schema-mismatch fallback (T3), incremental==full (T3), failure-leaves-prod (T4), progress callback (T1). ✓

**2. Placeholder scan:** Every code step shows real code; SQL strings reference exact existing line numbers to copy verbatim. Test helpers (`tableSig`, `bumpOneDataRowInPg`, `seedImages`, `updateOneImageInPg`, `corruptSchemaVersionInSnapshot`, `reopenPg`) — `tableSig` is given in full; the others are one-to-three-line PG/SQLite mutations the implementer writes against the test's existing `m_pg` fixture (specified by behavior + the exact SQL effect). No "TODO/handle edge cases" left.

**3. Type consistency:** `RegenProgress` / `RegenStats` / `segmentChanged` / `kFingerprintTables` / `kRegenPhases` names are used identically across Tasks 1–4. `regenToPath` final signature (8 params, last two defaulted) is fixed in Task 1 and matched by every later call. `regenerate(pg, progress)` overload defined in Task 1, used in Task 5. ✓

**Note for the implementer:** verify `_snapshot_meta`'s `key` column is the PRIMARY KEY before relying on `INSERT OR REPLACE` (read the `kCreateStatements` entry for `_snapshot_meta`, just after the `detailed_sensory_images` block). `INSERT OR REPLACE` is correct whether or not it is the PK, but the comment should state the real constraint.
