# v2.4.2 Sub-plan 2 — Convergence (Catch-up, Normalizer, Listen-socket) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Stop the nightly normalizer + flapping-network combination from resurrecting the DATAVIEWER-4 reset-to-5 bug: suppress NOTIFY for maintenance writes, losslessly normalize legacy string-typed scores (one-time + nightly), reload-and-merge the open resource on reconnect, and detect a dead NOTIFY socket.

**Architecture:** Tier-2 of the v2.4.2 batch (spec `docs/superpowers/specs/2026-06-11-v242-backcompat-resilience-design.md`, items R3, R4, R4b). Builds on Sub-plan 1 (transport + `app_version` + the `statement_timeout=10000` that the normalizer must opt out of). SQL changes follow the existing `ensureSchema` catalog-guarded heal pattern, kept in lockstep with `init.sql` + canonical migrations. Wraps to **v2.4.3** with an installer so transport + convergence can be smoke-tested together on a real connection.

**Tech Stack:** C++17 / Qt 6.10 (QtSql/QPSQL), PostgreSQL 16 (jsonb, pg_cron), qmake + MinGW, QtTest against the `dve-test-pg` container.

**Key learnings carried from Sub-plan 1:**
- Every connection now sets `statement_timeout=10000`. The normalizer's full-table UPDATE can exceed that → its SQL body **must** `SET LOCAL statement_timeout = 0`.
- `ensureSchema()` runs as autocommit statements (no wrapping transaction). `pg_advisory_xact_lock` needs a transaction → the one-time heal must wrap its gate+normalize+marker in an explicit `db.transaction()/commit()`.
- The test container provisioner (`start-test-postgres.ps1`) now applies migrations after init.sql, BUT it **skips the pg_cron tail** — so the nightly cron job will NOT exist in tests; tests must exercise `dve_normalize_legacy_json()` directly + via the heal path.
- Reviews flagged: never fail silently; make test cleanup robust to assert-aborts.

---

## Test environment setup

Container `dve-test-pg` on port 5433 must be running (if Docker is off, ASK the user to start it — do not start it yourself). Env + build/run recipe (PowerShell shown; bash equivalents work):
```
$env:DVE_TEST_PG_CONN = 'host=127.0.0.1 port=5433 dbname=dve_test user=test password=test'
```
**Build & run ONE suite** (RUN needs Qt+MinGW DLLs on PATH or the exe silently exits — trust only a freshly-built exe run with `-o`):
```
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/tests/<suite>"
cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe <suite>.pro && mingw32-make.exe -j8"
DVE_TEST_PG_CONN='host=127.0.0.1 port=5433 dbname=dve_test user=test password=test' PATH="C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/vendor/libpq-16:/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH" ./release/<suite>.exe -o res.txt,txt
grep -E "PASS|FAIL|Totals" res.txt ; rm -f res.txt
```
MIP: before edits/builds, if a source file shows `%TSD-Header-###%`, run `python tools/decrypt_via_copy.py --apply`. Create/land new files via Python round-trip + immediate `git add`. Commit trailer: `Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>`.

---

## File Structure

- `deploy/postgres/init.sql` — `notify_row_change` maintenance early-return; new `dve_normalize_legacy_json()` function; nightly `cron.schedule` entry (tail).
- `src/database/DatabaseManager.cpp` — `ensureSchema()` heals: `notify_row_change` (suppression), `dve_normalize_legacy_json` (create-or-replace), and the advisory-locked one-time normalize run.
- `deploy/postgres/migrations/2026-06-11-notify-maintenance-suppression.sql`, `2026-06-11-legacy-score-normalizer.sql` — canonical records.
- `src/MainWindow.cpp` / `.h` — `reloadOpenResourceAfterReconnect()` (catch-up) + presence re-activation in `onConnectionCameOnline`.
- `src/database/PostgresConnection.cpp` / `.h` — `pingListen()`.
- `src/database/NotificationListener.cpp` / `.h` — shared `channels()` + `resubscribeMissing()`.
- `src/database/ConnectionMonitor.cpp` — ping the listen socket too.
- Tests: `tests/tst_databasemanager/` (normalizer + suppression + heal-gate), `tests/tst_saveintegrity_e2e/` (catch-up merge scenario), `tests/tst_livesync/` (listen-socket + resubscribe).
- `DataViewerEnterprise.pro` — VERSION 2.4.2 → 2.4.3.

---

## Task 1: NOTIFY maintenance-suppression GUC (R4 prerequisite)

**Files:** `deploy/postgres/init.sql` (notify_row_change), `src/database/DatabaseManager.cpp` (ensureSchema heal), `deploy/postgres/migrations/2026-06-11-notify-maintenance-suppression.sql` (new), `tests/tst_livesync/tst_livesync.cpp` (test).

- [ ] **Step 1: Write the failing test** in `tst_livesync.cpp`. Add slot decl `void notify_suppressedWhenMaintenanceGucSet();` and impl (model on the existing `commitCell_notifyPayloadCarriesColumnAndValue` which uses a foreign `NotificationListener` + `QSignalSpy`):

```cpp
void TstLiveSync::notify_suppressedWhenMaintenanceGucSet()
{
    PostgresConnection foreign;
    QVERIFY(foreign.open(pgConfig()));
    NotificationListener listener(&foreign);
    QVERIFY(listener.subscribe());
    QSignalSpy spy(&listener, &NotificationListener::rowChanged);

    QSqlQuery q(m_conn->queryDb());
    // With dve.maintenance='1' set on the WRITING session, the trigger must
    // NOT emit a NOTIFY for this UPDATE.
    QVERIFY(q.exec("SET dve.maintenance = '1'"));
    QVERIFY(q.exec(QString("UPDATE data_rows SET draw_pressure = 3.21 "
                           "WHERE id=%1").arg(m_dataRowId)));
    QVERIFY2(!spy.wait(1500),
             "no rowChanged NOTIFY should arrive while dve.maintenance='1'");
    QCOMPARE(spy.count(), 0);

    // Positive control: clear the GUC, update again, NOTIFY must arrive.
    QVERIFY(q.exec("SET dve.maintenance = '0'"));
    QVERIFY(q.exec(QString("UPDATE data_rows SET draw_pressure = 3.22 "
                           "WHERE id=%1").arg(m_dataRowId)));
    QVERIFY2(spy.wait(2000), "rowChanged NOTIFY should arrive with maintenance off");
}
```
Add the slot decl in the `private slots:` block.

- [ ] **Step 2: Run → FAIL** (the suppression does not exist yet, so the first UPDATE emits a NOTIFY → `spy.wait(1500)` returns true → the `!spy.wait` QVERIFY2 fails). Build+run `tst_livesync`.

- [ ] **Step 3: Add the early-return in `init.sql`'s `notify_row_change`.** In `deploy/postgres/init.sql`, the function body begins (around line 324) with `BEGIN` then `payload := jsonb_build_object(...)`. Insert immediately after `BEGIN`:
```sql
  -- v2.4.2 R4: maintenance/bulk writes (the legacy-score normalizer) set this
  -- GUC so they don't fan out a NOTIFY per row and storm every live client.
  IF current_setting('dve.maintenance', true) = '1' THEN RETURN NULL; END IF;
```

- [ ] **Step 4: Heal it in `ensureSchema()`.** In `src/database/DatabaseManager.cpp`, after the existing function-heal blocks (e.g. after the 6-arg `dve_commit_cell_json` block), add a catalog-guarded heal mirroring the `dve_stamp_app_version` style:
```cpp
    // ── notify_row_change: maintenance-suppression early-return (v2.4.2 R4) ──
    // Heal live DBs so a bulk normalize (dve.maintenance='1') doesn't NOTIFY-
    // storm clients. Guard on prosrc so the already-healed path takes no work.
    {
        bool needsHeal = false, probed = false;
        QSqlQuery chk(db);
        if (chk.exec(QStringLiteral(
                "SELECT prosrc FROM pg_proc WHERE proname = 'notify_row_change'"))) {
            probed = true;
            if (chk.next())
                needsHeal = !chk.value(0).toString()
                                 .contains(QStringLiteral("dve.maintenance"));
        } else {
            logDebug(QStringLiteral("ensureSchema: cannot inspect notify_row_change: %1")
                         .arg(chk.lastError().text()));
        }
        if (probed && needsHeal) {
            QSqlQuery repl(db);
            const QString ddl = QStringLiteral(
                "CREATE OR REPLACE FUNCTION notify_row_change() RETURNS TRIGGER AS $fn$ "
                "DECLARE payload JSONB; BEGIN "
                "  IF current_setting('dve.maintenance', true) = '1' THEN RETURN NULL; END IF; "
                "  payload := jsonb_build_object('table', TG_TABLE_NAME, 'op', TG_OP, "
                "    'id', COALESCE(NEW.id, OLD.id), "
                "    'updated_by', COALESCE(NEW.updated_by, OLD.updated_by)); "
                "  IF current_setting('dve.live_column', true) <> '' THEN "
                "    payload := payload || jsonb_build_object('column', "
                "      current_setting('dve.live_column', true), 'new_value', "
                "      current_setting('dve.live_value', true)); END IF; "
                "  PERFORM pg_notify('dataviewer_changes', payload::text); "
                "  RETURN NULL; END; $fn$ LANGUAGE plpgsql;");
            if (repl.exec(ddl))
                logDebug(QStringLiteral("ensureSchema: notify_row_change healed "
                             "(maintenance suppression)"));
            else
                logDebug(QStringLiteral("ensureSchema: could not heal "
                             "notify_row_change: %1").arg(repl.lastError().text()));
        }
    }
```
(The healed body is byte-identical to init.sql's plus the early-return. The trigger itself is unchanged, so no re-CREATE TRIGGER needed — `CREATE OR REPLACE FUNCTION` swaps the body in place.)

- [ ] **Step 5: Run → PASS.** The test container applies migrations + ensureSchema (via the DatabaseManager that... NOTE: `tst_livesync` does NOT open a DatabaseManager). So tst_livesync's container heal won't run from this suite. **Provision the heal for the test:** add `notify_row_change` suppression to the migration (Step 6) — the provisioner applies migrations, so the container gets it. Re-run after Step 6. Expected: PASS.

- [ ] **Step 6: Create the migration** `deploy/postgres/migrations/2026-06-11-notify-maintenance-suppression.sql` (the same `CREATE OR REPLACE FUNCTION notify_row_change()` as init.sql's healed form, idempotent). Land plaintext via Python round-trip + `git add`. **Then re-provision the running container** so the test sees it: `Get-Content deploy/postgres/migrations/2026-06-11-notify-maintenance-suppression.sql | docker exec -i dve-test-pg psql -U test -d dve_test` (or re-run `start-test-postgres.ps1` if the user confirms Docker is up). Re-run the test → PASS.

- [ ] **Step 7: Commit** (init.sql + DatabaseManager.cpp + migration + test).

---

## Task 2: dve_normalize_legacy_json() (R4)

**Files:** `deploy/postgres/init.sql` (function), `src/database/DatabaseManager.cpp` (create-or-replace heal), `deploy/postgres/migrations/2026-06-11-legacy-score-normalizer.sql` (new), `tests/tst_databasemanager/tst_databasemanager.cpp` (test).

**The 16 score keys** (5 sensory `kSensoryMetrics` + 11 detailed `kDetailedAllMetrics`): `Burnt Taste, Vapor Volume, Overall Flavor, Smoothness, Overall Liking, Burn Taste, Flavor Intensity, Throat Irritation, Nasal Irritation, Vapor Quality Overall, Cough, Volume Consistency, Performance Consistency, Vapor Temperature, Vapor vs Oil`. (Note: "Vapor Volume" appears in both lists — dedupe to 15 distinct.) Scores are flat keys on each `samples[i]` object; walk `json_data -> 'samples' -> i -> '<key>'`.

- [ ] **Step 1: Write the failing test** in `tst_databasemanager.cpp` (it has a live DatabaseManager + PostgresConnection; follow its existing fixture pattern). The test: seed a sensory_sessions row via raw SQL whose `json_data` has STRING-typed scores; call `SELECT dve_normalize_legacy_json()`; assert (a) the string scores became JSON numbers, (b) a non-numeric string (a name) is untouched, (c) an already-numeric/clean row is NOT version-bumped (untouched), (d) the function returns a count ≥ 1, (e) a second call returns 0 (idempotent). Use a self-contained connection like the e2e suite's `pgConfig()`. Provide the full slot code:

```cpp
void TstDatabaseManager::normalizeLegacyJson_coercesStringScores()
{
    QSqlQuery q(m_conn->queryDb());   // adapt to this suite's connection handle
    // Damaged row: string-typed numeric scores + a non-numeric string field.
    QVERIFY(q.exec(
        "INSERT INTO sensory_sessions (session_name, tester_name, date, json_data) "
        "VALUES ('Norm','T','2026-06-09', "
        "'{\"samples\":[{\"name\":\"S1\",\"Smoothness\":\"7.5\","
        "\"Overall Liking\":\"9\",\"comments\":\"tasted 5 times\"}]}'::jsonb) "
        "RETURNING id"));
    QVERIFY(q.next());
    const int dmgId = q.value(0).toInt();
    // Clean row: numeric scores already.
    QVERIFY(q.exec(
        "INSERT INTO sensory_sessions (session_name, tester_name, date, json_data) "
        "VALUES ('Clean','T','2026-06-09', "
        "'{\"samples\":[{\"name\":\"S1\",\"Smoothness\":8}]}'::jsonb) RETURNING id"));
    QVERIFY(q.next());
    const int cleanId = q.value(0).toInt();
    int cleanVerBefore = -1;
    { QSqlQuery v(m_conn->queryDb());
      v.exec(QString("SELECT version FROM sensory_sessions WHERE id=%1").arg(cleanId));
      v.next(); cleanVerBefore = v.value(0).toInt(); }

    // Run the normalizer.
    QVERIFY2(q.exec("SELECT dve_normalize_legacy_json()") && q.next(),
             qPrintable(q.lastError().text()));
    QVERIFY(q.value(0).toInt() >= 1);

    // (a) string scores -> numbers
    QVERIFY(q.exec(QString(
        "SELECT jsonb_typeof(json_data->'samples'->0->'Smoothness'), "
        "       jsonb_typeof(json_data->'samples'->0->'Overall Liking'), "
        "       (json_data->'samples'->0->>'Smoothness')::float8, "
        "       jsonb_typeof(json_data->'samples'->0->'comments') "
        "FROM sensory_sessions WHERE id=%1").arg(dmgId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toString(), QString("number"));
    QCOMPARE(q.value(1).toString(), QString("number"));
    QCOMPARE(q.value(2).toDouble(), 7.5);
    // (b) non-numeric string ("tasted 5 times") stays a string
    QCOMPARE(q.value(3).toString(), QString("string"));

    // (c) clean row untouched (no version bump)
    QVERIFY(q.exec(QString("SELECT version FROM sensory_sessions WHERE id=%1").arg(cleanId)));
    QVERIFY(q.next());
    QCOMPARE(q.value(0).toInt(), cleanVerBefore);

    // (e) idempotent: second run finds nothing to do
    QVERIFY(q.exec("SELECT dve_normalize_legacy_json()") && q.next());
    QCOMPARE(q.value(0).toInt(), 0);
}
```
Add the slot decl. (Adapt `m_conn`/fixture names to the suite's actual members by reading its `initTestCase`.)

- [ ] **Step 2: Run → FAIL** (function does not exist). Build+run `tst_databasemanager`.

- [ ] **Step 3: Define the function in `init.sql`** (inside the `BEGIN..COMMIT` block, after the commit-cell functions). Exact body:
```sql
-- ── dve_normalize_legacy_json (v2.4.2 R4) ───────────────────────────────────
-- Losslessly rewrite numeric-looking STRING score values to JSON numbers in
-- sensory_sessions + detailed_sensory_sessions. Idempotent: only rows that
-- still have a string-typed score matching the writer's numeric regex are
-- touched (so clean rows don't version-bump or NOTIFY). Sets dve.maintenance
-- so the per-row UPDATE doesn't storm clients, and SET LOCAL statement_timeout
-- = 0 so the full-table scan isn't aborted by the 10s connection cap (v2.4.2
-- sub-plan 1). Returns the number of rows rewritten.
CREATE OR REPLACE FUNCTION dve_normalize_legacy_json() RETURNS INTEGER AS $$
DECLARE
    keys TEXT[] := ARRAY[
        'Burnt Taste','Vapor Volume','Overall Flavor','Smoothness','Overall Liking',
        'Burn Taste','Flavor Intensity','Throat Irritation','Nasal Irritation',
        'Vapor Quality Overall','Cough','Volume Consistency',
        'Performance Consistency','Vapor Temperature','Vapor vs Oil'];
    total INTEGER := 0;
    n INTEGER;
    tbl TEXT;
BEGIN
    SET LOCAL statement_timeout = 0;
    PERFORM set_config('dve.maintenance', '1', true);  -- tx-local; suppresses NOTIFY
    FOREACH tbl IN ARRAY ARRAY['sensory_sessions','detailed_sensory_sessions'] LOOP
        EXECUTE format($q$
            UPDATE %I s SET json_data = jsonb_set(
                s.json_data, '{samples}',
                (SELECT jsonb_agg(
                    COALESCE(
                      (SELECT jsonb_object_agg(kv.key,
                          CASE WHEN kv.key = ANY($1)
                                AND jsonb_typeof(kv.value) = 'string'
                                AND (kv.value #>> '{}') ~ '^-?[0-9]+(\.[0-9]+)?$'
                               THEN to_jsonb((kv.value #>> '{}')::numeric)
                               ELSE kv.value END)
                       FROM jsonb_each(elem.value) kv),
                      elem.value))
                 FROM jsonb_array_elements(s.json_data->'samples')
                      WITH ORDINALITY AS elem(value, ord)))
            WHERE jsonb_typeof(s.json_data->'samples') = 'array'
              AND EXISTS (
                SELECT 1 FROM jsonb_array_elements(s.json_data->'samples') e,
                              jsonb_each(e.value) kv
                WHERE kv.key = ANY($1)
                  AND jsonb_typeof(kv.value) = 'string'
                  AND (kv.value #>> '{}') ~ '^-?[0-9]+(\.[0-9]+)?$')
        $q$, tbl) USING keys;
        GET DIAGNOSTICS n = ROW_COUNT;
        total := total + n;
    END LOOP;
    RETURN total;
END;
$$ LANGUAGE plpgsql;
```
NOTE: `jsonb_agg` over `WITH ORDINALITY` preserves sample order. The `EXISTS` WHERE clause makes it touch only damaged rows (idempotent + no spurious bumps). The regex matches the writer's (init.sql `dve_commit_cell_json`) exactly, so "normalized" is a stable shape.

- [ ] **Step 4: Heal it onto live DBs in `ensureSchema()`.** Add a block that unconditionally `CREATE OR REPLACE`s `dve_normalize_legacy_json` with the same body (it's cheap; no data touched by the definition). Place it before the one-time run (Task 3). Use the same best-effort `logDebug` pattern. (No prosrc guard needed — replacing a function definition is a fast metadata-only op; or guard on `proname='dve_normalize_legacy_json'` absence + a body marker if you prefer symmetry.)

- [ ] **Step 5: Create the migration** `deploy/postgres/migrations/2026-06-11-legacy-score-normalizer.sql` (the `CREATE OR REPLACE FUNCTION dve_normalize_legacy_json()` body only — the one-time run + cron live elsewhere). Land plaintext + `git add` + re-apply to the running container (`docker exec -i ... psql < migration`).

- [ ] **Step 6: Run → PASS** (all 5 assertions). Build+run `tst_databasemanager`.

- [ ] **Step 7: Commit** (init.sql + DatabaseManager.cpp + migration + test).

---

## Task 3: One-time advisory-locked normalize + nightly cron (R4)

**Files:** `src/database/DatabaseManager.cpp` (ensureSchema one-time run), `deploy/postgres/init.sql` (cron tail), `tests/tst_databasemanager/tst_databasemanager.cpp` (gate test).

- [ ] **Step 1: Write the failing test** — `normalizeLegacyJson_oneTimeHealRunsOnceGated()`: delete the `schema_meta` marker `v242_legacy_score_normalize`; seed a damaged row; call `m_db->reopen()` (runs ensureSchema → the one-time normalize); assert the damaged row is now numeric AND the marker row exists; seed a SECOND damaged row; `reopen()` again; assert the second row is STILL string (the gate prevented a re-run). Full slot:
```cpp
void TstDatabaseManager::normalizeLegacyJson_oneTimeHealRunsOnceGated()
{
    QSqlQuery q(m_conn->queryDb());
    QVERIFY(q.exec("DELETE FROM schema_meta WHERE key='v242_legacy_score_normalize'"));
    auto seedDamaged = [&](const char* name) -> int {
        QSqlQuery i(m_conn->queryDb());
        i.exec(QString("INSERT INTO sensory_sessions (session_name,tester_name,date,json_data) "
            "VALUES ('%1','T','2026-06-10', "
            "'{\"samples\":[{\"name\":\"S1\",\"Smoothness\":\"6.5\"}]}'::jsonb) RETURNING id").arg(name));
        i.next(); return i.value(0).toInt();
    };
    auto smoothType = [&](int id) -> QString {
        QSqlQuery v(m_conn->queryDb());
        v.exec(QString("SELECT jsonb_typeof(json_data->'samples'->0->'Smoothness') "
                       "FROM sensory_sessions WHERE id=%1").arg(id));
        v.next(); return v.value(0).toString();
    };
    const int firstId = seedDamaged("OneTimeA");
    QVERIFY(m_db->reopen());                 // ensureSchema -> one-time normalize
    QCOMPARE(smoothType(firstId), QString("number"));
    QVERIFY(q.exec("SELECT 1 FROM schema_meta WHERE key='v242_legacy_score_normalize'"));
    QVERIFY(q.next());
    const int secondId = seedDamaged("OneTimeB");
    QVERIFY(m_db->reopen());                 // marker present -> normalize must NOT re-run
    QCOMPARE(smoothType(secondId), QString("string"));   // untouched: gate held
}
```

- [ ] **Step 2: Run → FAIL** (no one-time run yet; first row stays string). Build+run.

- [ ] **Step 3: Add the advisory-locked one-time run in `ensureSchema()`** (after the function heals; `dve_normalize_legacy_json` must already be create-or-replaced above). Because ensureSchema is autocommit, wrap in an explicit transaction so `pg_advisory_xact_lock` has scope:
```cpp
    // ── v2.4.2 R4: one-time legacy-score normalize (advisory-locked, gated) ──
    // Serializes concurrent first-connects (advisory xact lock) so N clients
    // don't each run the full-table normalize. schema_meta-gated so it runs
    // ONCE; the nightly pg_cron job (init.sql) handles ongoing convergence.
    // Wrapped in an explicit transaction because ensureSchema is otherwise
    // autocommit and pg_advisory_xact_lock would release immediately.
    if (db.transaction()) {
        bool ok = true;
        { QSqlQuery lk(db);
          ok = lk.exec(QStringLiteral("SELECT pg_advisory_xact_lock(4242002)")); }
        bool gateKnown = false, alreadyRun = false;
        if (ok) {
            QSqlQuery g(db);
            gateKnown = g.exec(QStringLiteral(
                "SELECT 1 FROM schema_meta WHERE key='v242_legacy_score_normalize'"));
            if (gateKnown) alreadyRun = g.next();
        }
        if (ok && gateKnown && !alreadyRun) {
            QSqlQuery n(db);
            if (n.exec(QStringLiteral("SELECT dve_normalize_legacy_json()"))) {
                QSqlQuery m(db);
                m.exec(QStringLiteral(
                    "INSERT INTO schema_meta(key,value) "
                    "VALUES ('v242_legacy_score_normalize', now()::text) "
                    "ON CONFLICT (key) DO NOTHING"));
                logDebug(QStringLiteral("ensureSchema: legacy-score normalize ran (one-time)"));
            } else {
                logDebug(QStringLiteral("ensureSchema: legacy-score normalize failed: %1")
                             .arg(n.lastError().text()));
            }
        }
        if (!db.commit()) { db.rollback();
            logDebug(QStringLiteral("ensureSchema: normalize tx commit failed")); }
    } else {
        logDebug(QStringLiteral("ensureSchema: could not begin normalize tx (skipped)"));
    }
```

- [ ] **Step 4: Run → PASS** (first row numeric + marker present; second row stays string). Build+run.

- [ ] **Step 5: Add the nightly cron entry to `init.sql`'s pg_cron tail** (after the existing `cron.schedule` calls, before EOF):
```sql
SELECT cron.schedule(
  'dve_legacy_score_normalize',
  '17 3 * * *',  -- nightly 03:17; the function sets statement_timeout=0 itself
  $$ SELECT dve_normalize_legacy_json() $$
);
```
Document (comment) that the test container skips this tail, so the nightly job is verified only by the one-time heal + the direct-call test (Task 2). If the app's DB role can't register cron jobs, this line is the NAS-admin record.

- [ ] **Step 6: Commit** (DatabaseManager.cpp + init.sql + test). No new migration file — the function migration (Task 2) carries the body; the cron line is NAS-admin/init.sql only (note this in the commit message).

---

## Task 4: Post-reconnect inbound catch-up (R3 — the reset-to-5 keystone)

**Files:** `src/MainWindow.cpp` / `.h` (new `reloadOpenResourceAfterReconnect()`), `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp` (merge scenario).

- [ ] **Step 1: Write the failing e2e test** — the convergence guard. Simulate: a session is saved; another client (raw SQL) edits a score during the "offline window"; the local in-memory struct has a DIFFERENT dirty edit on a DIFFERENT cell; the dirty-aware merge (the catch-up's core) must keep the remote value for the non-dirty cell AND the local value for the dirty cell — neither is lost. Add to `tst_saveintegrity_e2e.cpp`:
```cpp
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
    // Local client: a DIFFERENT cell edited + marked dirty, Smoothness untouched
    // locally (still the seed 5.0, NOT dirty).
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
```
(Confirm the exact names `sensorySessionToJson` / `sensorySessionFromJson` / `mergeSensoryPreservingDbScores` and includes by reading `SensoryData.h`; adapt if the serializer fns differ.)

- [ ] **Step 2: Run → it should PASS already if the merge is correct** (this scenario validates the merge the catch-up relies on; it's a guard, not strictly RED→GREEN since `mergeSensoryPreservingDbScores` exists). If it fails, the merge has a bug — fix that first. The NEW production behavior (calling this on reconnect) is wired in Steps 3-4 and verified by build + the user's smoke test.

- [ ] **Step 3: Add `reloadOpenResourceAfterReconnect()`** to `MainWindow` (declare in `.h` private; define in `.cpp`). Mirror the per-mode branch structure of the existing recreate block (MainWindow.cpp:930-1036). It reloads the open resource by `m_currentResourceType`/`m_currentResourceId`:
  - `"file"`: find the in-memory `FileResult` in `m_loadedFiles` with matching id; if that file is NOT in `m_modifiedFilePaths` (no local unsaved edits), replace it with `m_db->loadFile(id)` and call `displayCurrentSample()`. If it IS dirty, leave memory (don't clobber local edits) — TPM has no per-cell dirty merge yet (deferred).
  - `"sensory_session"`: `SensorySession dbNow = m_db->loadSensorySession(id)`; if `dbNow.id > 0`, merge via `mergeSensoryPreservingDbScores(toJson(current), toJson(dbNow), current->dirtyCells)`, assign back into the panel's current session, `refreshSensoryNavigator()` / re-render.
  - `"detailed_sensory_session"`: same with the detailed twins.
  Guard the whole method on `m_currentResourceId > 0 && m_db && m_db->isOnline()`. Best-effort + logged.

- [ ] **Step 4: Call it at the end of `onConnectionCameOnline()`** — after `flushPendingEdits()` (line ~5739), add `reloadOpenResourceAfterReconnect();`. (Order matters: outbound drain first so local pending edits land, THEN reload+merge pulls remote changes.)

- [ ] **Step 5: Run the e2e suite → PASS** (scenario9 + all prior). Build+run `tst_saveintegrity_e2e`.

- [ ] **Step 6: Commit** (MainWindow.cpp/.h + test).

---

## Task 5: Re-activate local presence on reconnect (R3)

**Files:** `src/MainWindow.cpp` (`onConnectionCameOnline`).

- [ ] **Step 1: Add presence re-activation** in `onConnectionCameOnline()`, after the NOTIFY re-subscribe (replacing the "reactivation happens lazily" no-op at lines 5719-5721):
```cpp
    // R3: re-activate local presence so peers see us again and the heartbeat
    // restarts (offline deactivated it). Lazy reactivation left the user a
    // ghost to other clients after every blip.
    if (m_presence && !m_currentResourceType.isEmpty() && m_currentResourceId > 0) {
        m_presence->activate(m_currentResourceType, m_currentResourceId,
                             m_presence->activeIntent());
    }
    refreshAllPresence();
```
Remove/replace the old comment block.

- [ ] **Step 2: Verify it compiles** (no unit test — presence is GUI/heartbeat integration; verified by the app build in Task 7 and the user's smoke test). Confirm `activeIntent()` exists (`PresenceManager.h:48`) and returns a sensible default when nothing was active.

- [ ] **Step 3: Commit** (MainWindow.cpp).

---

## Task 6: Listen-socket liveness + per-channel re-subscribe (R4b)

**Files:** `src/database/PostgresConnection.cpp` / `.h` (`pingListen()`), `src/database/NotificationListener.cpp` / `.h` (shared channels + `resubscribeMissing()`), `src/database/ConnectionMonitor.cpp` (ping listen too), `src/MainWindow.cpp` (cameOnline guard), `tests/tst_livesync/tst_livesync.cpp` (tests).

- [ ] **Step 1: Write the failing tests** in `tst_livesync.cpp`:
```cpp
void TstLiveSync::pingListen_detectsListenSocket() {
    // Healthy: both sockets answer SELECT 1.
    QVERIFY(m_conn->ping());        // query socket
    QVERIFY(m_conn->pingListen());  // listen socket (NEW)
}
void TstLiveSync::resubscribeMissing_fillsDroppedChannel() {
    PostgresConnection c; QVERIFY(c.open(pgConfig()));
    NotificationListener l(&c);
    QVERIFY(l.subscribe());
    QVERIFY(l.isSubscribedTo("dataviewer_changes"));
    // Simulate one channel dropped.
    l.unsubscribeFromChannelForTest("dataviewer_changes");   // test-only helper
    QVERIFY(!l.isSubscribedTo("dataviewer_changes"));
    QVERIFY(l.resubscribeMissing());                          // NEW
    QVERIFY(l.isSubscribedTo("dataviewer_changes"));
}
```
Add slot decls. (If a test-only unsubscribe-one helper is undesirable, instead assert `resubscribeMissing()` is a no-op that returns true when all channels are present, and that after a fresh `subscribe()` all three channels are subscribed — a lighter test that still exercises the new method.)

- [ ] **Step 2: Run → FAIL** (`pingListen`/`resubscribeMissing` don't exist; won't compile). Build `tst_livesync`.

- [ ] **Step 3: Add `PostgresConnection::pingListen()`** (`.h` decl + `.cpp` def), mirroring `ping()` but on `m_listenDb`:
```cpp
bool PostgresConnection::pingListen() {
    if (!m_open) return false;
    QSqlQuery q(m_listenDb);
    return q.exec("SELECT 1");
}
```

- [ ] **Step 4: Add shared channels + `resubscribeMissing()` to `NotificationListener`.** Promote the `kChannels` list to a private static accessor `static const QStringList& channels();` (so `subscribe()` and `resubscribeMissing()` share it). Add:
```cpp
bool NotificationListener::resubscribeMissing() {
    if (!m_conn || !m_conn->isOpen()) return false;
    QSqlDriver* drv = m_conn->listenDb().driver();
    if (!drv) return false;
    const bool wasEmpty = m_subscribedChannels.isEmpty();
    for (const QString& ch : channels()) {
        if (m_subscribedChannels.contains(ch)) continue;
        if (drv->subscribeToNotification(ch)) m_subscribedChannels.insert(ch);
    }
    if (wasEmpty && !m_subscribedChannels.isEmpty())
        connect(drv, &QSqlDriver::notification, this,
                &NotificationListener::onNotification, Qt::UniqueConnection);
    return m_subscribedChannels.size() == channels().size();
}
```
Refactor `subscribe()` to use `channels()`. (If Step 1 used the test-only helper, add `void unsubscribeFromChannelForTest(const QString&)` guarded for tests — or drop it per the Step 1 lighter alternative.)

- [ ] **Step 5: Ping the listen socket in `ConnectionMonitor::onPing()`.** Change the offline trigger so a dead LISTEN socket (query still alive) also flips offline — the reconnect path then reopens both sockets, re-subscribes, and runs catch-up (Task 4):
```cpp
void ConnectionMonitor::onPing() {
    if (!m_conn) return;
    if (m_conn->ping() && m_conn->pingListen()) return;   // both healthy
    switchToOfflineMode();
    emit wentOffline();
}
```

- [ ] **Step 6: Use per-channel re-subscribe in `onConnectionCameOnline()`.** Replace the all-or-nothing `if (m_notify && !m_notify->isSubscribed())` guard (line 5713) with:
```cpp
    if (m_notify && !m_notify->resubscribeMissing()) {
        qWarning() << "MainWindow: NOTIFY resubscribe incomplete — "
                   << "some live-update channels did not re-attach";
    }
```

- [ ] **Step 7: Run → PASS** (`tst_livesync` green). Build+run.

- [ ] **Step 8: Commit** (PostgresConnection + NotificationListener + ConnectionMonitor + MainWindow + tests).

---

## Task 7: Version bump 2.4.3, clean rebuild, verify, installer

**Files:** `DataViewerEnterprise.pro` (VERSION → 2.4.3), `release_overview/release_overview_v_2_4_3.txt` (new).

- [ ] **Step 1: Bump** `VERSION = 2.4.2` → `2.4.3`.
- [ ] **Step 2: MIP decrypt + clean rebuild** the app (`python tools/decrypt_via_copy.py --apply`; then `cmd.exe //c "set PATH=...;%PATH% && qmake.exe -spec win32-g++ DataViewerEnterprise.pro CONFIG+=release && mingw32-make.exe clean && mingw32-make.exe -j8 release"`). Must be clean under `-Werror`.
- [ ] **Step 3: Run all affected suites fresh** + the atomic-delete python test: `tst_livesync`, `tst_saveintegrity_e2e`, `tst_databasemanager`. All green.
- [ ] **Step 4: Build the installer** via the `rebuild-dataviewer` skill (clean rebuild + `tools\prepare_python_embed.bat` + `build_installer.bat`). Verify `release\DataViewer.exe` + `dist\DataViewer-setup.exe` both report 2.4.3. **Do NOT touch Synology** — surface the path for the user to transfer/test.
- [ ] **Step 5: Write + commit `release_overview/release_overview_v_2_4_3.txt`** (customer-readable: live-update reliability after reconnects, automatic background cleanup of older data formats, no more disappearing/stale collaborator changes after a blip). MIP round-trip + `git add`.
- [ ] **Step 6: Commit** the bump + overview. Hand off: v2.4.3 installer ready for a real-connection smoke test of transport + convergence; capture learnings to feed Sub-plan 3.

---

## Self-Review

**Spec coverage:** R3 (catch-up reload + dirty merge + presence re-activate) → Tasks 4+5. R4 (NOTIFY suppression + normalizer one-time + nightly) → Tasks 1+2+3. R4b (listen-socket liveness + per-channel re-subscribe) → Task 6. ✓

**Sub-plan-1 learnings applied:** normalizer `SET LOCAL statement_timeout=0` (Task 2); advisory lock in an explicit transaction (Task 3); migrations carry the heals so the provisioned container gets them (Tasks 1,2); pg_cron tail skipped in tests → direct-call + heal-path tests (Tasks 2,3); surface failures loudly (Task 6 qWarning). ✓

**Placeholders:** none — every SQL body and C++ insertion is complete. Two adapt-to-suite notes (tst_databasemanager fixture handle names; exact serializer fn names) are explicit verification steps, not gaps.

**Type/name consistency:** `dve.maintenance` GUC, `dve_normalize_legacy_json()`, `v242_legacy_score_normalize` marker, advisory key `4242002`, `reloadOpenResourceAfterReconnect()`, `pingListen()`, `resubscribeMissing()`, `channels()` — used consistently across tasks. The 15 distinct score keys match `kSensoryMetrics` + `kDetailedAllMetrics` (de-duped on "Vapor Volume").

**Risk note:** the normalizer's `jsonb_object_agg`/`jsonb_agg` rebuild is the highest-risk SQL — Task 2's test asserts coercion, non-numeric preservation, clean-row-untouched, and idempotence to pin it. The catch-up wiring (Task 4 Steps 3-4) is GUI-integration verified by build + the user's smoke test (the merge logic itself is unit-tested).
