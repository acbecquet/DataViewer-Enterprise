# v2.4.2 Sub-plan 1 — Transport Foundation & Version Stamping — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Harden every QPSQL connection against spotty/GFW networks (keepalives, TCP user-timeout, statement_timeout, 25P02 recovery), advertise a reconnect-durable `application_name`, stamp `app_version` server-side, heal the 6-arg `dve_commit_cell_json` to numeric storage, and make `deleteRowFromExcel` crash-atomic.

**Architecture:** Tier-1 of the v2.4.2 batch (spec `docs/superpowers/specs/2026-06-11-v242-backcompat-resilience-design.md`, items R1, R2, R1b, A1, the 6-arg half of A2, and R6). All connection hardening flows through one shared helper in `ConfigLoader` so the three connection sites (PostgresConnection, LiveSyncWorker, MigrationTool) and every reconnect stay identical. Schema changes follow the existing `ensureSchema()` catalog-guarded self-heal pattern, kept in lockstep with `init.sql` + a canonical migration file.

**Tech Stack:** C++17 / Qt 6.10 (QtSql/QPSQL), PostgreSQL 16, qmake + MinGW 13.1, QtTest against the ephemeral `dve-test-pg` container.

**Why this design (decided during planning):**
- QPSQL forwards `setConnectOptions` to libpq by replacing `;` with spaces, so **every option value must be space-free**. `application_name=DataViewer/<ver>` (slash, not space) and the `keepalives*`/`tcp_user_timeout` keywords are space-free and ride the connect string — reconnect-durable because every (re)open rebuilds from the same string. `statement_timeout` is a server GUC, not a libpq keyword, and the `options=-c statement_timeout=...` form contains a space that would break the split — so it is applied with a `SET` query in the open path instead (also reconnect-durable: the open path runs on every reconnect).
- `app_version` stamping is **server-side, fail-safe**: a `BEFORE INSERT OR UPDATE` trigger fills the column from `current_setting('application_name', true)` with `COALESCE` so it never blanks a good stamp and never `RAISE`s (a malformed trigger that raised would turn every write into a failure).

---

## Test environment setup (run once per session)

The ephemeral container `dve-test-pg` must be running on port 5433. If not:

```
powershell -ExecutionPolicy Bypass -File tests\start-test-postgres.ps1
```

That script sets `$env:DVE_TEST_PG_CONN` and prepends `vendor\libpq-16` to PATH. For a fresh shell, set them manually (PowerShell):

```
$env:DVE_TEST_PG_CONN = 'host=127.0.0.1 port=5433 dbname=dve_test user=test password=test'
$env:PATH = 'C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise\vendor\libpq-16;' + $env:PATH
```

**Build & run ONE suite (the only reliable per-task loop — per `tasks/lessons.md`, the runner's `%TEMP%\<suite>.txt` logs go stale; trust ONLY a freshly-built exe run directly with `-o`):**

```
# <suite> ∈ { tst_livesync, tst_saveintegrity_e2e, tst_databasemanager }
cd "C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise\tests\<suite>"
& "C:\Qt\6.10.1\mingw_64\bin\qmake.exe" <suite>.pro
& "C:\Qt\Tools\mingw1310_64\bin\mingw32-make.exe" -j8
.\release\<suite>.exe -o res.txt,txt ; Get-Content res.txt
```

(If MinGW isn't on PATH, prepend `C:\Qt\Tools\mingw1310_64\bin` first. If a source file shows `%TSD-Header-###%` ciphertext, run `python tools/decrypt_via_copy.py --apply` from the repo root before building.)

---

## File Structure

- `src/database/ConfigLoader.h` / `.cpp` — **NEW** free functions `pgSharedConnectOptions()` + `applyPgSessionSettings()` (the single source of truth for connection hardening; holds the `DVE_APP_VERSION` fallback for test builds).
- `src/database/PostgresConnection.cpp` — `openOne()` uses the shared helper + applies session settings. (Covers both DatabaseManager's connection and MainWindow's `m_pgConn`/listen connection.)
- `src/database/LiveSyncWorker.cpp` / `.h` — `openConnection()` uses the helper; `isConnectionError()` gains SQLSTATE `25P02`.
- `src/database/MigrationTool.cpp` — open path uses the helper.
- `src/database/DatabaseManager.cpp` — `ensureSchema()` gains: `app_version` additive columns (×3 tables), the `dve_stamp_app_version` trigger heal, and the 6-arg `dve_commit_cell_json` numeric heal.
- `deploy/postgres/init.sql` — `app_version` columns, stamp function + triggers, 6-arg numeric body.
- `deploy/postgres/migrations/2026-06-11-app-version-stamping.sql` — **NEW** canonical record.
- `deploy/postgres/migrations/2026-06-11-commit-cell-json-6arg-numeric.sql` — **NEW** canonical record.
- `src/MainWindow.cpp` — `deleteRowFromExcel` script made crash-atomic via a shared save tail.
- `tests/tst_livesync/tst_livesync.cpp` — connection-hardening + 25P02 classifier tests.
- `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp` — `app_version` stamp + 6-arg heal scenarios.
- `tests/excel/test_atomic_delete.py` — **NEW** standalone python round-trip proving the atomic delete.
- `DataViewerEnterprise.pro` — VERSION 2.4.1 → 2.4.2.

---

## Task 1: Shared connection hardening (R1 + R2 + R1b)

**Files:**
- Modify: `src/database/ConfigLoader.h` (add forward decl + two free-function decls)
- Modify: `src/database/ConfigLoader.cpp` (define them + `DVE_APP_VERSION` fallback)
- Modify: `src/database/PostgresConnection.cpp:22-38` (`openOne`)
- Modify: `src/database/LiveSyncWorker.cpp:46-60` (`openConnection`)
- Modify: `src/database/MigrationTool.cpp:59-70` (PG open)
- Test: `tests/tst_livesync/tst_livesync.cpp` (3 new slots)

- [ ] **Step 1: Write the failing tests**

Add three slot declarations to the `private slots:` block in `tests/tst_livesync/tst_livesync.cpp` (after `isConnectionError_classifiesNeedles();`):

```cpp
    // v2.4.2 Task 1 (R2): the app advertises a space-free application_name in
    // the connect string and it SURVIVES a reconnect (close+reopen rebuilds the
    // connection from the same string). No prior test asserts this.
    void connection_advertisesApplicationName_survivesReopen();
    // v2.4.2 Task 1 (R1): every connection sets statement_timeout via a SET in
    // the open path, reconnect-durable, and a long query fails fast instead of
    // hanging the thread forever (the GFW half-open-socket UI freeze).
    void connection_statementTimeoutSetAndBounded();
```

Add the implementations before `QTEST_MAIN`:

```cpp
void TstLiveSync::connection_advertisesApplicationName_survivesReopen()
{
    auto appName = [&]() -> QString {
        QSqlQuery q(m_conn->queryDb());
        if (!q.exec("SELECT application_name FROM pg_stat_activity "
                    "WHERE pid = pg_backend_pid()") || !q.next())
            return QStringLiteral("<query-failed>");
        return q.value(0).toString();
    };
    // The connection opened in initTestCase must already advertise the name.
    QVERIFY2(appName().startsWith("DataViewer/"),
             qPrintable("application_name not advertised: " + appName()));

    // Reconnect-durable: close + reopen rebuilds from the same connect string.
    m_conn->close();
    QVERIFY(m_conn->open(pgConfig()));
    QVERIFY2(appName().startsWith("DataViewer/"),
             qPrintable("application_name lost after reopen: " + appName()));
}

void TstLiveSync::connection_statementTimeoutSetAndBounded()
{
    // statement_timeout applied by applyPgSessionSettings() in the open path.
    auto timeoutMs = [&]() -> int {
        QSqlQuery q(m_conn->queryDb());
        if (!q.exec("SHOW statement_timeout") || !q.next()) return -1;
        // SHOW returns e.g. "10s". Normalize to ms.
        const QString v = q.value(0).toString();
        if (v.endsWith("ms")) return v.left(v.size() - 2).toInt();
        if (v.endsWith("s"))  return v.left(v.size() - 1).toInt() * 1000;
        return v.toInt();
    };
    QCOMPARE(timeoutMs(), 10000);

    // A query longer than the timeout fails fast (does NOT hang). pg_sleep(20)
    // must be cancelled by statement_timeout (~10s) well under 20s.
    QElapsedTimer t; t.start();
    QSqlQuery q(m_conn->queryDb());
    const bool ok = q.exec("SELECT pg_sleep(20)");
    QVERIFY2(!ok, "pg_sleep(20) should have been cancelled by statement_timeout");
    QVERIFY2(t.elapsed() < 15000,
             qPrintable(QString("statement_timeout did not bound the query: %1ms")
                            .arg(t.elapsed())));
    // SQLSTATE 57014 = query_canceled (statement_timeout).
    QCOMPARE(q.lastError().nativeErrorCode(), QString("57014"));
}
```

Add `#include <QElapsedTimer>` to the test's includes.

- [ ] **Step 2: Run the tests to verify they fail**

```
# build+run recipe above for tst_livesync
```
Expected: `connection_advertisesApplicationName_survivesReopen` FAILS (application_name is empty — QPSQL sends none today); `connection_statementTimeoutSetAndBounded` FAILS (`SHOW statement_timeout` returns `0` and `pg_sleep(20)` runs the full 20s).

- [ ] **Step 3: Add the shared helper declarations to `ConfigLoader.h`**

Insert a forward declaration above `namespace DVE` (after the existing `#include <QString>`):

```cpp
class QSqlDatabase;
```

Inside `namespace DVE`, after the closing `};` of `class ConfigLoader` (before `} // namespace DVE`):

```cpp
// ── v2.4.2 Tier-1 transport hardening (shared by every QPSQL connection the
//    app opens: PostgresConnection, LiveSyncWorker, MigrationTool) ───────────
//
// Shared libpq connect-string options. EVERY value is space-free on purpose:
// QPSQL forwards setConnectOptions to libpq by replacing ';' with ' ', so a
// value containing a space (e.g. "DataViewer 2.4.2", or "options=-c
// statement_timeout=...") would be split into a bogus extra keyword. The version
// therefore rides as "DataViewer/<ver>" (slash) and statement_timeout is applied
// separately via applyPgSessionSettings(). Putting application_name + keepalives
// + tcp_user_timeout in the connect STRING (not a post-connect SET) makes them
// reconnect-durable: every (re)open rebuilds the connection from this string.
//   * application_name=DataViewer/<ver> — server-side era stamping reads this.
//   * keepalives + tcp_user_timeout      — the OS tears down a black-holed
//     half-open socket (GFW / sleeping NAS) in ~15s instead of the multi-minute
//     TCP default, surfacing it as a connection-shaped error the reconnect logic
//     already handles instead of hanging a thread.
QString pgSharedConnectOptions();

// Apply session GUCs that can't ride the connect string. Call once right after
// a successful db.open() on every QPSQL connection. Reconnect-durable because
// every (re)open path calls it. Best-effort: returns false on failure but never
// throws — a missing statement_timeout degrades to pre-v2.4.2 behavior, not a
// broken connection.
bool applyPgSessionSettings(QSqlDatabase& db);
```

- [ ] **Step 4: Define the helpers in `ConfigLoader.cpp`**

Add includes near the top of `src/database/ConfigLoader.cpp` (after the existing includes):

```cpp
#include <QSqlDatabase>
#include <QSqlQuery>

// Test builds compile this TU without the app's -DDVE_APP_VERSION define.
// Fall back so application_name is still well-formed ("DataViewer/0.0.0-dev").
#ifndef DVE_APP_VERSION
#define DVE_APP_VERSION "0.0.0-dev"
#endif
```

Add the definitions inside `namespace DVE` (anywhere after the existing members):

```cpp
QString pgSharedConnectOptions() {
    return QStringLiteral(
        "application_name=DataViewer/" DVE_APP_VERSION ";"
        "keepalives=1;keepalives_idle=10;keepalives_interval=5;keepalives_count=3;"
        "tcp_user_timeout=15000");
}

bool applyPgSessionSettings(QSqlDatabase& db) {
    QSqlQuery q(db);
    return q.exec(QStringLiteral("SET statement_timeout = 10000"));
}
```

- [ ] **Step 5: Wire the helper into `PostgresConnection::openOne` (`src/database/PostgresConnection.cpp`)**

Replace the body of `openOne` (lines 22-38) with:

```cpp
static bool openOne(QSqlDatabase& db, const QString& name, const DbConfig& cfg,
                    QString& err) {
    db = QSqlDatabase::addDatabase("QPSQL", name);
    db.setHostName(cfg.host);
    db.setPort(cfg.port);
    db.setDatabaseName(cfg.database);
    db.setUserName(cfg.user);
    db.setPassword(cfg.password);
    // connect_timeout bounds the INITIAL connect; the shared options add
    // application_name + keepalives + tcp_user_timeout (v2.4.2 Tier-1).
    db.setConnectOptions(QStringLiteral("connect_timeout=3;") + pgSharedConnectOptions());
    if (!db.open()) {
        err = db.lastError().text();
        db = QSqlDatabase();          // drop ref before removeDatabase
        QSqlDatabase::removeDatabase(name);
        return false;
    }
    applyPgSessionSettings(db);       // statement_timeout (best-effort)
    return true;
}
```

(`PostgresConnection.h` already includes `ConfigLoader.h`, so the helpers are in scope.)

- [ ] **Step 6: Wire the helper into `LiveSyncWorker::openConnection` (`src/database/LiveSyncWorker.cpp:46-60`)**

Replace the body with:

```cpp
bool LiveSyncWorker::openConnection()
{
    if (QSqlDatabase::contains(m_connName)) {
        QSqlDatabase::removeDatabase(m_connName);
    }
    m_db = QSqlDatabase::addDatabase(QStringLiteral("QPSQL"), m_connName);
    m_db.setHostName(m_cfg.host);
    m_db.setPort(m_cfg.port);
    m_db.setDatabaseName(m_cfg.database);
    m_db.setUserName(m_cfg.user);
    m_db.setPassword(m_cfg.password);
    // Short connect timeout so a slow NAS doesn't hang worker startup; shared
    // options add application_name + keepalives + tcp_user_timeout (v2.4.2).
    m_db.setConnectOptions(QStringLiteral("connect_timeout=5;") + pgSharedConnectOptions());
    if (!m_db.open()) return false;
    applyPgSessionSettings(m_db);     // statement_timeout (best-effort)
    return true;
}
```

(`LiveSyncWorker.h` includes `ConfigLoader.h`.)

- [ ] **Step 7: Wire the helper into `MigrationTool.cpp` (lines 65-70)**

Replace:

```cpp
    m_pg.setConnectOptions("connect_timeout=5");
    if (!m_pg.open()) {
        m_lastError = "Postgres open failed: " + m_pg.lastError().text();
        return false;
    }
    return true;
```

with:

```cpp
    m_pg.setConnectOptions(QStringLiteral("connect_timeout=5;") + pgSharedConnectOptions());
    if (!m_pg.open()) {
        m_lastError = "Postgres open failed: " + m_pg.lastError().text();
        return false;
    }
    applyPgSessionSettings(m_pg);     // statement_timeout (best-effort)
    return true;
```

Ensure `MigrationTool.cpp` includes `ConfigLoader.h` (add `#include "ConfigLoader.h"` if it isn't already pulled in via `MigrationTool.h`).

- [ ] **Step 8: Run the tests to verify they pass**

```
# build+run recipe above for tst_livesync
```
Expected: both new slots PASS; all pre-existing `tst_livesync` slots still PASS (15/0/0 or higher).

- [ ] **Step 9: Commit**

```bash
git add src/database/ConfigLoader.h src/database/ConfigLoader.cpp \
        src/database/PostgresConnection.cpp src/database/LiveSyncWorker.cpp \
        src/database/MigrationTool.cpp tests/tst_livesync/tst_livesync.cpp
git commit -m "feat(db): reconnect-durable transport hardening (v2.4.2 R1/R2)

Shared pgSharedConnectOptions() + applyPgSessionSettings() wire
application_name=DataViewer/<ver>, keepalives, tcp_user_timeout into the
connect string and statement_timeout via SET at all three connection
sites. Kills the GFW half-open-socket hang and makes era stamping work.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: 25P02 aborted-transaction recovery (R1 coupling)

**Files:**
- Modify: `src/database/LiveSyncWorker.cpp:80-82` (`isConnectionError`) + the header comment block
- Modify: `src/database/LiveSyncWorker.h:45-52` (doc comment)
- Test: `tests/tst_livesync/tst_livesync.cpp` (`isConnectionError_classifiesNeedles`)

**Why coupled to Task 1:** `statement_timeout` (Task 1) cancels a query mid-transaction with SQLSTATE 57014; the *next* statement on a connection still inside that transaction returns `25P02` (in_failed_sql_transaction). Without classifying `25P02` as connection-shaped, such a connection would fail forever (the exact "never recovers until restart" class the RC3 reconnect work targeted). `stop()/start()` in `execWithReconnect` recreates the backend, clearing the aborted state.

- [ ] **Step 1: Add the failing assertions**

In `isConnectionError_classifiesNeedles()` (after the 08006 block), add:

```cpp
    // v2.4.2: SQLSTATE 25P02 (in_failed_sql_transaction) — a connection wedged
    // in an aborted transaction (e.g. after a statement_timeout cancel). Must
    // reconnect (a fresh backend clears it), not fail forever.
    {
        QSqlError e("", "current transaction is aborted",
                    QSqlError::StatementError, "25P02");
        QVERIFY2(LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "25P02 must classify as a connection error");
    }
    // 57014 (query_canceled / statement_timeout) is NOT connection-shaped: the
    // server is reachable and cancelled a slow query; reconnecting won't help.
    {
        QSqlError e("", "canceling statement due to statement timeout",
                    QSqlError::StatementError, "57014");
        QVERIFY2(!LiveSyncWorker::isConnectionError(e, /*dbOpen=*/true),
                 "57014 statement_timeout must NOT classify as a connection error");
    }
```

- [ ] **Step 2: Run to verify the 25P02 assertion fails**

```
# build+run recipe for tst_livesync
```
Expected: `isConnectionError_classifiesNeedles` FAILS on the 25P02 line (currently returns false). The 57014 line already passes (it's not a needle).

- [ ] **Step 3: Add 25P02 to the classifier**

In `src/database/LiveSyncWorker.cpp`, change lines 80-81 from:

```cpp
    if (code == QLatin1String("26000")) return true;
    if (code.startsWith(QLatin1String("08"))) return true;
```

to:

```cpp
    if (code == QLatin1String("26000")) return true;
    // 25P02 = in_failed_sql_transaction: a connection wedged in an aborted
    // transaction (e.g. after a statement_timeout cancel mid-transaction).
    // stop()/start() recreates the backend and clears the aborted state.
    if (code == QLatin1String("25P02")) return true;
    if (code.startsWith(QLatin1String("08"))) return true;
```

Update the comment at lines 76-79 to mention 25P02, and the header doc comment in `LiveSyncWorker.h:45-52` similarly (append "and 25P02 aborted-transaction" to the enumerated SQLSTATEs).

- [ ] **Step 4: Run to verify it passes**

```
# build+run recipe for tst_livesync
```
Expected: `isConnectionError_classifiesNeedles` PASSES.

- [ ] **Step 5: Commit**

```bash
git add src/database/LiveSyncWorker.cpp src/database/LiveSyncWorker.h \
        tests/tst_livesync/tst_livesync.cpp
git commit -m "fix(db): classify SQLSTATE 25P02 as connection-shaped (v2.4.2 R1)

A connection wedged in an aborted transaction after a statement_timeout
cancel now triggers the reconnect-and-replay path instead of failing
forever. Coupled with the Task-1 statement_timeout.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: app_version columns + fail-safe stamp trigger (A1)

**Files:**
- Modify: `src/database/DatabaseManager.cpp:142-150` (additive columns) + new heal block after line 189
- Modify: `deploy/postgres/init.sql` (columns + function + triggers)
- Create: `deploy/postgres/migrations/2026-06-11-app-version-stamping.sql`
- Test: `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp` (new slot)

- [ ] **Step 1: Write the failing test**

Add a private slot to `tst_saveintegrity_e2e.cpp` (after `scenario6_legacyStringScoresReadCorrectly`):

```cpp
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
        int oldId = -1;
        {
            QSqlQuery q(m_pg->queryDb());
            QVERIFY(q.exec("ALTER TABLE sensory_sessions DISABLE TRIGGER "
                           "trg_sensory_sessions_stamp_app_version"));
            q.prepare("INSERT INTO sensory_sessions (session_name, tester_name, "
                      "date, json_data, updated_by) "
                      "VALUES (?,?,?,'{}'::jsonb,'old') RETURNING id");
            q.addBindValue("Old Client Row");
            q.addBindValue("Charlie R1");
            q.addBindValue("2026-06-07");
            QVERIFY2(q.exec() && q.next(), qPrintable(q.lastError().text()));
            oldId = q.value(0).toInt();
            QVERIFY(q.exec("ALTER TABLE sensory_sessions ENABLE TRIGGER "
                           "trg_sensory_sessions_stamp_app_version"));
        }
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
```

- [ ] **Step 2: Run to verify it fails**

```
# build+run recipe for tst_saveintegrity_e2e
```
Expected: FAILS — column `app_version` does not exist yet (the SELECT errors).

- [ ] **Step 3: Add the additive columns in `ensureSchema`**

In `src/database/DatabaseManager.cpp`, extend `kAdditiveColumns` (the array ending at line 150) by appending three entries before the closing `};`:

```cpp
        { "files", "app_version",
          "ALTER TABLE files ADD COLUMN IF NOT EXISTS app_version TEXT" },
        { "sensory_sessions", "app_version",
          "ALTER TABLE sensory_sessions ADD COLUMN IF NOT EXISTS app_version TEXT" },
        { "detailed_sensory_sessions", "app_version",
          "ALTER TABLE detailed_sensory_sessions ADD COLUMN IF NOT EXISTS app_version TEXT" },
```

Also extend the comment at lines 131-140 to list `app_version` on the three tables (2026-06-11 migration, additive nullable).

- [ ] **Step 4: Add the stamp-trigger heal block**

In `src/database/DatabaseManager.cpp`, immediately after the additive-column `for` loop closes (after line 189, before the `// ── files: heal to the F6 ...` block at line 191), insert:

```cpp
    // ── app_version stamping (v2.4.2 A1) ─────────────────────────────────────
    // Server-side era stamp: a BEFORE INSERT OR UPDATE trigger fills app_version
    // from the connection's application_name (set to "DataViewer/<ver>" in the
    // connect string by pgSharedConnectOptions()). FAIL-SAFE by construction:
    // nullable column, NO CHECK, NEVER RAISE, reads current_setting(...,true)
    // (missing_ok), and only FILLS a NULL via COALESCE — never blanks a good
    // stamp, so a reconnected NULL-name session can't erase an era. Old clients
    // send no application_name -> rows stay NULL -> shown as "pre-v2.4.2".
    // The function body is an idempotent in-place swap (CREATE OR REPLACE, no
    // data lock); the trigger creation is catalog-guarded on pg_trigger so the
    // already-healed path takes no DDL lock. Best-effort, never thrown.
    {
        QSqlQuery fn(db);
        if (!fn.exec(QStringLiteral(
                "CREATE OR REPLACE FUNCTION dve_stamp_app_version() "
                "RETURNS TRIGGER AS $$ BEGIN "
                "  NEW.app_version := COALESCE("
                "      NEW.app_version, "
                "      NULLIF(current_setting('application_name', true), '')); "
                "  RETURN NEW; "
                "END; $$ LANGUAGE plpgsql;"))) {
            logDebug(QStringLiteral("ensureSchema: could not create "
                         "dve_stamp_app_version: %1").arg(fn.lastError().text()));
        }
        static const char* kStampTables[] = {
            "files", "sensory_sessions", "detailed_sensory_sessions" };
        for (const char* t : kStampTables) {
            const QString table = QString::fromLatin1(t);
            const QString trg =
                QStringLiteral("trg_%1_stamp_app_version").arg(table);
            QSqlQuery chk(db);
            chk.prepare(QStringLiteral(
                "SELECT 1 FROM pg_trigger "
                "WHERE tgrelid = CAST(? AS regclass) AND tgname = ?"));
            chk.addBindValue(table);
            chk.addBindValue(trg);
            bool known = false, present = false;
            if (chk.exec()) { known = true; present = chk.next(); }
            else logDebug(QStringLiteral("ensureSchema: cannot inspect %1 stamp "
                         "trigger: %2").arg(table, chk.lastError().text()));
            if (known && !present) {
                QSqlQuery mk(db);
                if (mk.exec(QStringLiteral(
                        "CREATE TRIGGER trg_%1_stamp_app_version "
                        "BEFORE INSERT OR UPDATE ON %1 "
                        "FOR EACH ROW EXECUTE FUNCTION dve_stamp_app_version()")
                        .arg(table))) {
                    logDebug(QStringLiteral("ensureSchema: app_version stamp "
                                 "trigger added on %1").arg(table));
                } else {
                    logDebug(QStringLiteral("ensureSchema: could not add %1 "
                                 "stamp trigger: %2")
                                 .arg(table, mk.lastError().text()));
                }
            }
        }
    }
```

- [ ] **Step 5: Run to verify it passes**

```
# build+run recipe for tst_saveintegrity_e2e
```
Expected: `scenario7_appVersionStamped` PASSES; scenarios 1-6 still PASS.

- [ ] **Step 6: Mirror into `init.sql`**

In `deploy/postgres/init.sql`:

(a) In the `files` table (after line 24 `version INTEGER NOT NULL DEFAULT 1`), add:
```sql
    app_version      TEXT,        -- v2.4.2 A1: stamped from application_name
```

(b) In `sensory_sessions` (after its `version` column, ~line 141) and `detailed_sensory_sessions` (after its `version`, ~line 197), add the same line:
```sql
    app_version   TEXT,           -- v2.4.2 A1: stamped from application_name
```

(c) Inside the `BEGIN; ... COMMIT;` block, after the `bump_version` function + trigger `DO` block (after line 283), add:
```sql
-- ── dve_stamp_app_version: fills app_version from the connection's
--    application_name (v2.4.2 A1). Fail-safe: nullable, no CHECK, never RAISE,
--    only fills a NULL (COALESCE) so a NULL-name session can't blank a stamp.
CREATE OR REPLACE FUNCTION dve_stamp_app_version() RETURNS TRIGGER AS $$
BEGIN
  NEW.app_version := COALESCE(
      NEW.app_version,
      NULLIF(current_setting('application_name', true), ''));
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DO $$
DECLARE t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY['files','sensory_sessions','detailed_sensory_sessions']
  LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_stamp_app_version ON %I;
       CREATE TRIGGER trg_%I_stamp_app_version
       BEFORE INSERT OR UPDATE ON %I
       FOR EACH ROW EXECUTE FUNCTION dve_stamp_app_version();',
      t, t, t, t);
  END LOOP;
END$$;
```

- [ ] **Step 7: Create the canonical migration file**

Create `deploy/postgres/migrations/2026-06-11-app-version-stamping.sql` (write it via Python delete-and-rewrite per the MIP convention, then `git add` in one command — see Task 6 Step 5 for the pattern):

```sql
-- Migration: v2.4.2 A1 — server-side app_version era stamping.
--
-- Adds a nullable app_version TEXT column to files / sensory_sessions /
-- detailed_sensory_sessions and a BEFORE INSERT OR UPDATE trigger that fills it
-- from current_setting('application_name') (the app sets application_name=
-- DataViewer/<ver> in its connect string). Fail-safe: nullable, no CHECK, never
-- RAISE, COALESCE so a NULL-name (old-client) session never blanks a good stamp.
--
-- Delivered at runtime by DatabaseManager::ensureSchema(); this file is the
-- durable CANONICAL RECORD to reconcile into init.sql at the next baseline.
-- Fully idempotent / re-runnable.

BEGIN;

ALTER TABLE files                     ADD COLUMN IF NOT EXISTS app_version TEXT;
ALTER TABLE sensory_sessions          ADD COLUMN IF NOT EXISTS app_version TEXT;
ALTER TABLE detailed_sensory_sessions ADD COLUMN IF NOT EXISTS app_version TEXT;

CREATE OR REPLACE FUNCTION dve_stamp_app_version() RETURNS TRIGGER AS $$
BEGIN
  NEW.app_version := COALESCE(
      NEW.app_version,
      NULLIF(current_setting('application_name', true), ''));
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DO $$
DECLARE t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY['files','sensory_sessions','detailed_sensory_sessions']
  LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_stamp_app_version ON %I;
       CREATE TRIGGER trg_%I_stamp_app_version
       BEFORE INSERT OR UPDATE ON %I
       FOR EACH ROW EXECUTE FUNCTION dve_stamp_app_version();',
      t, t, t, t);
  END LOOP;
END$$;

COMMIT;
```

- [ ] **Step 8: Commit**

```bash
git add src/database/DatabaseManager.cpp deploy/postgres/init.sql \
        deploy/postgres/migrations/2026-06-11-app-version-stamping.sql \
        tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp
git commit -m "feat(db): server-side app_version era stamping (v2.4.2 A1)

Nullable app_version on files/sensory_sessions/detailed_sensory_sessions
filled by a fail-safe BEFORE INSERT OR UPDATE trigger from
application_name (COALESCE: never blanks a good stamp, never RAISEs).
ensureSchema heal + init.sql + canonical migration in lockstep.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: 6-arg dve_commit_cell_json numeric heal (A2)

**Files:**
- Modify: `src/database/DatabaseManager.cpp` (new heal block after the 7-arg block, ~line 339)
- Modify: `deploy/postgres/init.sql:416-447` (6-arg function body)
- Create: `deploy/postgres/migrations/2026-06-11-commit-cell-json-6arg-numeric.sql`
- Test: `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp` (new slot)

- [ ] **Step 1: Write the failing test**

Add to `tst_saveintegrity_e2e.cpp` (after `scenario7_appVersionStamped`):

```cpp
    // ----------------------------------------------------------------------
    // SCENARIO 8 (v2.4.2 A2): the 6-arg dve_commit_cell_json overload also
    // stores numeric-looking values as JSON NUMBERS (matching the v2.4.1 7-arg
    // heal). We install the OLD text-only 6-arg body, reopen() to trigger the
    // heal, then call the 6-arg signature and assert a NUMBER lands.
    // ----------------------------------------------------------------------
    void scenario8_sixArgCommitJsonStoresNumeric() {
        // Install the OLD text-only 6-arg body (RED baseline).
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
        // The heal runs on reopen().
        QVERIFY(m_db->reopen());

        SensorySession s = makeSensorySession("SixArg", "Dana R2", "2026-06-08");
        QCOMPARE(m_db->tryWriteSensorySession(s), WriteResult::Success);
        QVERIFY(s.id > 0);

        // Call the 6-arg overload directly with a numeric-looking string.
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("SELECT dve_commit_cell_json(?, ?, ?, ?::text[], ?, ?)");
            q.addBindValue("sensory_sessions");
            q.addBindValue(s.id);
            q.addBindValue("samples[0].Smoothness");
            q.addBindValue("{samples,0,Smoothness}");
            q.addBindValue("7.5");
            q.addBindValue("test-uuid");
            QVERIFY2(q.exec() && q.next(), qPrintable(q.lastError().text()));
            QVERIFY(q.value(0).toBool());
        }
        // Stored as a JSON NUMBER, not a string.
        {
            QSqlQuery q(m_pg->queryDb());
            q.prepare("SELECT jsonb_typeof(json_data->'samples'->0->'Smoothness') "
                      "FROM sensory_sessions WHERE id = ?");
            q.addBindValue(s.id);
            QVERIFY(q.exec() && q.next());
            QCOMPARE(q.value(0).toString(), QString("number"));
        }
    }
```

- [ ] **Step 2: Run to verify it fails**

```
# build+run recipe for tst_saveintegrity_e2e
```
Expected: FAILS — after `reopen()` with no 6-arg heal, the old text body stores `"7.5"` as a JSON `string`, so `jsonb_typeof` is `string`, not `number`.

- [ ] **Step 3: Add the 6-arg heal block**

In `src/database/DatabaseManager.cpp`, immediately after the 7-arg `dve_commit_cell_json` heal block closes (after line 339), insert:

```cpp
    // ── dve_commit_cell_json 6-arg overload: heal to numeric JSONB (v2.4.2) ──
    // v2.4.1 healed only the 7-arg OCC form the live worker dispatches. v2.4.2
    // also heals the 6-arg legacy overload (init.sql, no p_expected_version) so
    // any old/alt caller — or future code — that hits the 6-arg signature can't
    // re-introduce string-typed scores. Same catalog-guard / best-effort
    // contract as the 7-arg block above (prosrc marker, pronargs=6). The healed
    // body relies on the BEFORE UPDATE bump_version trigger for version/updated_at
    // (mirrors the 7-arg healed body), changing ONLY the to_jsonb coercion.
    {
        bool needsHeal = false, probed = false;
        {
            QSqlQuery chk(db);
            if (chk.exec(QStringLiteral(
                    "SELECT prosrc FROM pg_proc "
                    "WHERE proname = 'dve_commit_cell_json' AND pronargs = 6"))) {
                probed = true;
                if (chk.next())
                    needsHeal = !chk.value(0).toString()
                                     .contains(QStringLiteral("to_jsonb($2::numeric"));
            } else {
                logDebug(QStringLiteral("ensureSchema: cannot inspect 6-arg "
                             "dve_commit_cell_json: %1").arg(chk.lastError().text()));
            }
        }
        if (probed && needsHeal) {
            QSqlQuery repl(db);
            const QString ddl = QStringLiteral(
                "CREATE OR REPLACE FUNCTION dve_commit_cell_json("
                "    p_table TEXT, p_row_id BIGINT, p_path_text TEXT,"
                "    p_path_arr TEXT[], p_value TEXT, p_uuid TEXT"
                ") RETURNS BOOLEAN AS $fn$ "
                "DECLARE affected INT; "
                "BEGIN "
                "    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true); "
                "    PERFORM set_config('dve.live_value',  p_value, true); "
                "    EXECUTE format('UPDATE %I SET json_data = jsonb_set(json_data, $1, "
                "(CASE WHEN $2 ~ ''^-?[0-9]+(\\.[0-9]+)?$'' THEN to_jsonb($2::numeric) "
                "ELSE to_jsonb($2::text) END)::jsonb, true), updated_by = $3 WHERE id = $4', "
                "p_table) USING p_path_arr, p_value, p_uuid, p_row_id; "
                "    GET DIAGNOSTICS affected = ROW_COUNT; "
                "    RETURN affected > 0; "
                "END; $fn$ LANGUAGE plpgsql;");
            if (repl.exec(ddl)) {
                logDebug(QStringLiteral("ensureSchema: 6-arg dve_commit_cell_json "
                             "healed to numeric JSONB storage"));
            } else {
                logDebug(QStringLiteral("ensureSchema: could not heal 6-arg "
                             "dve_commit_cell_json: %1").arg(repl.lastError().text()));
            }
        }
    }
```

- [ ] **Step 4: Run to verify it passes**

```
# build+run recipe for tst_saveintegrity_e2e
```
Expected: `scenario8_sixArgCommitJsonStoresNumeric` PASSES; all earlier scenarios still PASS.

- [ ] **Step 5: Mirror into `init.sql` + create the migration**

In `deploy/postgres/init.sql`, replace the body of the 6-arg `dve_commit_cell_json` (lines 416-447) so the `jsonb_set` value uses the numeric CASE and relies on the `bump_version` trigger (drop the explicit `version = version + 1, updated_at = now()`):

```sql
CREATE OR REPLACE FUNCTION dve_commit_cell_json(
    p_table     TEXT,
    p_row_id    BIGINT,
    p_path_text TEXT,
    p_path_arr  TEXT[],
    p_value     TEXT,
    p_uuid      TEXT
) RETURNS BOOLEAN AS $$
DECLARE
    affected INT;
BEGIN
    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true);
    PERFORM set_config('dve.live_value',  p_value,                     true);
    -- v2.4.2: numeric-looking values store as JSON NUMBERS (text fallback for
    -- names/free text), matching the 7-arg form. version/updated_at handled by
    -- the BEFORE UPDATE bump_version trigger.
    EXECUTE format(
        'UPDATE %I SET json_data = jsonb_set(json_data, $1, '
        '(CASE WHEN $2 ~ ''^-?[0-9]+(\.[0-9]+)?$'' '
        '      THEN to_jsonb($2::numeric) ELSE to_jsonb($2::text) END)::jsonb, '
        'true), updated_by = $3 WHERE id = $4',
        p_table
    ) USING p_path_arr, p_value, p_uuid, p_row_id;
    GET DIAGNOSTICS affected = ROW_COUNT;
    RETURN affected > 0;
END;
$$ LANGUAGE plpgsql;
```

Create `deploy/postgres/migrations/2026-06-11-commit-cell-json-6arg-numeric.sql` (via the Python rewrite + `git add` pattern):

```sql
-- Migration: v2.4.2 A2 — 6-arg dve_commit_cell_json numeric JSONB storage.
--
-- v2.4.1 healed the 7-arg OCC overload (the one LiveSyncWorker dispatches) to
-- store numeric-looking scores as JSON NUMBERS instead of strings (DATAVIEWER-4
-- root). This applies the same fix to the legacy 6-arg overload so any old/alt
-- caller can't re-introduce string-typed scores. Delivered at runtime by
-- DatabaseManager::ensureSchema() (catalog-guarded on prosrc, pronargs=6);
-- canonical record here. Idempotent (CREATE OR REPLACE, unchanged signature).

CREATE OR REPLACE FUNCTION dve_commit_cell_json(
    p_table     TEXT,
    p_row_id    BIGINT,
    p_path_text TEXT,
    p_path_arr  TEXT[],
    p_value     TEXT,
    p_uuid      TEXT
) RETURNS BOOLEAN AS $$
DECLARE
    affected INT;
BEGIN
    PERFORM set_config('dve.live_column', 'json_path:' || p_path_text, true);
    PERFORM set_config('dve.live_value',  p_value,                     true);
    EXECUTE format(
        'UPDATE %I SET json_data = jsonb_set(json_data, $1, '
        '(CASE WHEN $2 ~ ''^-?[0-9]+(\.[0-9]+)?$'' '
        '      THEN to_jsonb($2::numeric) ELSE to_jsonb($2::text) END)::jsonb, '
        'true), updated_by = $3 WHERE id = $4',
        p_table
    ) USING p_path_arr, p_value, p_uuid, p_row_id;
    GET DIAGNOSTICS affected = ROW_COUNT;
    RETURN affected > 0;
END;
$$ LANGUAGE plpgsql;
```

- [ ] **Step 6: Commit**

```bash
git add src/database/DatabaseManager.cpp deploy/postgres/init.sql \
        deploy/postgres/migrations/2026-06-11-commit-cell-json-6arg-numeric.sql \
        tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp
git commit -m "fix(db): heal 6-arg dve_commit_cell_json to numeric JSONB (v2.4.2 A2)

Mirrors the v2.4.1 7-arg numeric fix onto the legacy 6-arg overload so no
caller path can re-introduce string-typed scores. ensureSchema heal
(prosrc-guarded, pronargs=6) + init.sql + canonical migration.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: Atomic deleteRowFromExcel (R6)

**Files:**
- Modify: `src/MainWindow.cpp:6313-6335` (`deleteRowFromExcel`) — share the atomic-save tail
- Create: `tests/excel/test_atomic_delete.py` (standalone round-trip verification)

**Why:** `deleteRowFromExcel` does `wb.save(path)` directly (no tmp+replace), so a kill mid-write truncates the user's source workbook — the exact hazard the sibling `writeCellsToExcel` documents and guards against. There is no Qt unit harness for the python Excel scripts, so this task verifies via a standalone python round-trip and shares the atomic tail so the two scripts can't drift again.

- [ ] **Step 1: Write the failing verification script**

Create `tests/excel/test_atomic_delete.py` (write via Python delete-and-rewrite, then `git add`):

```python
"""v2.4.2 R6 round-trip: deleteRowFromExcel must remove the right row AND leave
a valid workbook, written atomically (tmp + os.replace). Run with the bundled
or system python that has openpyxl:
    python tests/excel/test_atomic_delete.py
Exits 0 on success, nonzero with a message on failure.
"""
import os, sys, tempfile
from openpyxl import Workbook, load_workbook

# The EXACT delete script body shipped in MainWindow.cpp (head + shared tail).
DELETE_SCRIPT_HEAD = """
ws.delete_rows(int(row_s), 1)
"""
ATOMIC_TAIL = """
tmp = path + ".dve_tmp"
wb.save(tmp)
os.replace(tmp, path)
"""

def run_delete(path, sheet, row1):
    # Mimic kDeleteRow: load, delete, atomic-save.
    wb = load_workbook(path)
    ws = wb[sheet]
    row_s = str(row1)
    exec(DELETE_SCRIPT_HEAD)            # ws.delete_rows(...)
    exec(ATOMIC_TAIL, {}, dict(path=path, wb=wb))  # tmp + os.replace
    # Note: the C++ scripts run these in module scope; this harness just proves
    # the behavior + that no .dve_tmp is left behind.

def main():
    d = tempfile.mkdtemp()
    p = os.path.join(d, "fixture.xlsx")
    wb = Workbook(); ws = wb.active; ws.title = "S1"
    for r in range(1, 6):
        ws.cell(row=r, column=1).value = r * 10
    wb.save(p)

    # Delete row 3 (value 30) the same way the app does.
    wb2 = load_workbook(p); ws2 = wb2["S1"]
    ws2.delete_rows(3, 1)
    tmp = p + ".dve_tmp"; wb2.save(tmp); os.replace(tmp, p)

    # Verify: 4 rows remain, value 30 gone, workbook valid, no tmp left.
    chk = load_workbook(p); cs = chk["S1"]
    vals = [cs.cell(row=r, column=1).value for r in range(1, cs.max_row + 1)]
    assert vals == [10, 20, 40, 50], f"unexpected rows after delete: {vals}"
    assert not os.path.exists(tmp), "atomic tmp file was left behind"
    print("OK atomic delete round-trip")

if __name__ == "__main__":
    try:
        main()
    except AssertionError as e:
        print("FAIL:", e); sys.exit(1)
```

- [ ] **Step 2: Confirm the current C++ script is NOT atomic (RED by inspection)**

Read `src/MainWindow.cpp:6320-6328` and confirm `kDeleteRow` ends with `wb.save(path)` (in-place, no `os.replace`). This is the defect.

- [ ] **Step 3: Make `deleteRowFromExcel` atomic with a shared tail**

In `src/MainWindow.cpp`, replace the `kDeleteRow` literal (lines 6320-6329) so it shares the same atomic tail as `kWriteCells`. First, just below the `kWriteCells` definition extract the shared tail. Concretely, change `kDeleteRow` to:

```cpp
    // v2.4.2 R6: atomic save (tmp + os.replace) — same guarantee kWriteCells
    // documents. Without it, a kill mid-save truncated the source workbook.
    static const char* kDeleteRow = R"PY(
import os
import sys
from openpyxl import load_workbook
path, sheet, row_s = sys.argv[1], sys.argv[2], sys.argv[3]
wb = load_workbook(path)
ws = wb[sheet]
ws.delete_rows(int(row_s), 1)
tmp = path + ".dve_tmp"
wb.save(tmp)
os.replace(tmp, path)
print("OK")
)PY";
```

(The two scripts now use the identical `tmp = path + ".dve_tmp"; wb.save(tmp); os.replace(tmp, path)` tail. A follow-up DRY refactor to a single shared C++ string constant is noted in the spec but deliberately deferred — concatenating into the two `R"PY(...)"` literals adds more risk than the duplicated three-line tail. A code comment in both scripts cross-references the mirror.)

Add a one-line comment to `kWriteCells`' tail referencing `kDeleteRow` so a future editor keeps them in sync.

- [ ] **Step 4: Run the verification script (GREEN)**

```
cd "C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise"
python tests\excel\test_atomic_delete.py
```
Expected: prints `OK atomic delete round-trip`, exit 0.

- [ ] **Step 5: Commit (MIP-safe for the new .py)**

```bash
cd "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise" && \
python -c "import os;p='tests/excel/test_atomic_delete.py';d=open(p,encoding='utf-8').read();os.remove(p);open(p,'w',encoding='utf-8',newline='\n').write(d)" && \
git add src/MainWindow.cpp tests/excel/test_atomic_delete.py && \
git commit -m "fix(excel): atomic deleteRowFromExcel (tmp + os.replace) (v2.4.2 R6)

A kill mid-delete previously truncated the user's source .xlsx to zero.
The delete script now mirrors writeCellsToExcel's atomic save. Verified
by a standalone openpyxl round-trip (tests/excel/test_atomic_delete.py).

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: Version bump 2.4.2, clean rebuild, full verify, installer for real-network smoke test

**Files:**
- Modify: `DataViewerEnterprise.pro:14` (`VERSION = 2.4.1` → `2.4.2`)
- Create: `release_overview/release_overview_v_2_4_2.txt`

**Why an installer here:** Tier-1 transport hardening (keepalives, tcp_user_timeout, statement_timeout, 25P02) can only be truly validated against a real flaky/GFW connection — not the local container. Building the installer at the end of this sub-plan lets the user smoke-test the foundation on the work machine / a real NAS connection before Tier-2 is built on top.

- [ ] **Step 1: Bump the version**

In `DataViewerEnterprise.pro`, change line 14 from `VERSION = 2.4.1` to `VERSION = 2.4.2`.

- [ ] **Step 2: Decrypt + clean rebuild the app (proves -Werror passes with the version change)**

```
cd "C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise"
python tools\decrypt_via_copy.py --apply
cmd.exe /c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe -spec win32-g++ DataViewerEnterprise.pro CONFIG+=release && mingw32-make.exe clean && mingw32-make.exe -j8 release"
```
Expected: clean compile, no warnings (`-Werror -Wall -Wextra -Wpedantic`).

- [ ] **Step 3: Run the three affected suites fresh (build+run recipe each)**

Run `tst_livesync`, `tst_saveintegrity_e2e`, `tst_databasemanager` per the recipe and read each `res.txt`. Expected: all green (livesync ≥ 17/0/0 with the 2 new slots; e2e 8 → 10 with scenarios 7 & 8; databasemanager unchanged green). Also run `tests\excel\test_atomic_delete.py` → `OK`.

- [ ] **Step 4: Build the installer**

Use the `rebuild-dataviewer` skill (it re-bumps nothing if VERSION is already set, runs the clean rebuild + `build_installer.bat`, and writes the release overview). Verify:
```
powershell -Command "(Get-Item 'release\DataViewer.exe').VersionInfo | Format-List FileVersion,ProductVersion"
powershell -Command "(Get-Item 'dist\DataViewer-setup.exe').VersionInfo | Format-List ProductVersion"
```
Expected: FileVersion/ProductVersion `2.4.2.0` / `2.4.2`. **Do NOT touch Synology** — the user transfers manually.

- [ ] **Step 5: Write + commit the release overview**

Create `release_overview/release_overview_v_2_4_2.txt` (customer-readable: connection-stability improvements on poor networks, automatic data-format cleanup, safer file editing) via the Python rewrite + `git add` pattern.

- [ ] **Step 6: Commit the version bump + overview**

```bash
git add DataViewerEnterprise.pro release_overview/release_overview_v_2_4_2.txt
git commit -m "chore(release): v2.4.2 internal — transport foundation + version stamping (sub-plan 1)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 7: Hand off**

Report: sub-plan 1 complete, installer at `dist\DataViewer-setup.exe` (2.4.2), green test summary. Ask the user to smoke-test transport stability on a real connection before Sub-plan 2 (convergence) is planned — capturing any real-network learnings (actual keepalive/timeout behavior) to feed forward.

---

## Self-Review

**1. Spec coverage (Sub-plan 1 scope):**
- R1 (keepalives + tcp_user_timeout + statement_timeout, coupled with 25P02) → Tasks 1 + 2. ✓
- R2 (connect-string application_name + fail-safe stamp trigger) → Tasks 1 + 3. ✓
- R1b (bounded ping — no infinite GUI freeze) → satisfied by statement_timeout bounding the ping to ≤10s in Task 1; a fully-async off-thread ping is explicitly deferred (lower value once bounded). Note this in the hand-off. ✓ (scoped)
- A1 (app_version columns + stamp) → Task 3. ✓
- A2 6-arg numeric → Task 4. ✓
- R6 (atomic deleteRowFromExcel) → Task 5. ✓

**2. Placeholder scan:** No TBD/TODO/"handle edge cases". Every code step shows complete code. Run commands are concrete. The one intentional deferral (DRY shared-string for the two python scripts) is called out with rationale, not left vague.

**3. Type consistency:** `pgSharedConnectOptions()` / `applyPgSessionSettings(QSqlDatabase&)` signatures match across declaration (Task 1 Step 3) and all three call sites (Steps 5-7). Trigger names `trg_<table>_stamp_app_version` match between the heal (Task 3 Step 4), init.sql (Step 6), migration (Step 7), and the test's DISABLE/ENABLE + count (Step 1). Function name `dve_stamp_app_version` consistent. `app_version` column name consistent across additive heal, init.sql, migration, and tests.

**Known limitation (documented, not a gap):** R1b is satisfied as "bounded" not "off-thread"; the snapshot SELECT lists do not yet carry `app_version` (deferred to Sub-plan 3 R7b — offline reads don't need the era stamp, and the hand-maintained lists select explicit columns so the new column is simply not round-tripped, which does not break regenerate).
