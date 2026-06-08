# DATAVIEWER-2 — Test-scoped sample-name dropdown + bounded popup — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.

**Goal:** The SampleCard "sample name" dropdown should show only the names used for the **current test**, not the entire global pool (which currently covers the screen). Scope sample-name presets by test (`sensory_header_presets.test_name`) and cap the popup height. Preserve all existing data; backfill historical (test → sample-name) pairs so the scoped dropdown is useful immediately.

**Architecture:** The schema change rides the existing **connect-time self-heal `DatabaseManager::ensureSchema()`** (the DV-3 mechanism, run on every `open()`/`reopen()`), NOT `MigrationTool` (which is a one-shot SQLite importer). `ensureSchema()` is extended to ensure `sensory_header_presets` is current: create the table if missing (new shape with `test_name`), add the `test_name` column if missing, and swap the uniqueness from `(kind, value)` to a unique index on `(kind, value, COALESCE(test_name,''))` — all idempotent + catalog-guarded so the common path is lock-free. A one-time gated backfill (marker in `schema_meta`) reconstructs (test → sample-name) pairs from `sensory_sessions` / `detailed_sensory_sessions` history. `saveSensoryHeaderPresets` associates `test_name` with `sample_name` rows; a new `loadSampleNamesForTest(testName)` returns the scoped list; the panel dropdown providers pass the current Test Title. A canonical `.sql` migration file is written for the record + eventual init.sql reconciliation.

**Bonus:** because the test container lacks `sensory_header_presets` entirely (init.sql drift), making `ensureSchema()` create it also **fixes the pre-existing `sensoryHeaderPresets_roundTrip` failure** and the `testSensoryLayoutPresence`/init() wipe errors — verify this.

**Tech Stack:** C++17 / Qt 6.10 (QtTest), qmake + MinGW, PostgreSQL 16 (QPSQL; `jsonb_array_elements` for backfill). MIP: `python tools/decrypt_via_copy.py --apply` before builds. Build is `-Werror -Wall -Wextra -Wpedantic`. Test PG via `tests\start-test-postgres.ps1` (sets `DVE_TEST_PG_CONN`); the running `dve-test-pg` is reused (do NOT `docker rm`).

**Data-integrity invariants (spec §6):** idempotent migration (`IF [NOT] EXISTS`), never DROP data; unique-index + `ON CONFLICT` change together (test asserts repeated saves don't error); legacy NULL-`test_name` rows are KEPT (acceptance: count of distinct historical sample names never decreases); offline snapshot behavior unchanged (presets are not mirrored today — documented out of scope).

**Versioning:** This is the LAST fix of the batch. DV-2 lands, then the batch wraps directly into deployable **v2.4.0** (no separate v2.3.6 installer — v2.4.0 IS the wrap). DV-3 (v2.3.3 hotfix), DV-4 (v2.3.4), Zero-Save-As (v2.3.5) are already in.

---

## File structure

- **`src/database/DatabaseManager.cpp`** — extend `ensureSchema()` (structure heal + gated backfill); test-scope `saveSensoryHeaderPresets`; add `loadSampleNamesForTest`.
- **`src/database/DatabaseManager.h`** — declare `loadSampleNamesForTest`.
- **`src/ui/SensoryPanel.cpp`** — scope the sample-name provider lambda to `m_testTitleEdit`; cap the `attachPresetDropdown` QMenu height.
- **`src/ui/DetailedSensoryPanel.cpp`** — same (its own `attachPresetDropdown` copy + sample-name wiring).
- **`deploy/postgres/migrations/2026-06-07-dv2-sample-name-by-test.sql`** — canonical migration (new-shape create, add column, index swap, triggers, backfill).
- **Tests:** `tests/tst_databasemanager/tst_databasemanager.cpp` (structure heal, scoping, repeated-save, backfill).
- **`DataViewerEnterprise.pro`** — `VERSION` 2.3.5 → 2.4.0 (the wrap, Task 6).

---

## The exact schema shapes (single source of truth for all tasks)

**New `sensory_header_presets` shape** (additive — only `test_name` is new):
```sql
CREATE TABLE IF NOT EXISTS sensory_header_presets (
    id          BIGSERIAL PRIMARY KEY,
    kind        TEXT NOT NULL CHECK (kind IN ('test_name','media','sample_name')),
    value       TEXT NOT NULL CHECK (length(trim(value)) > 0),
    test_name   TEXT,                                   -- NEW; NULL for kind in ('test_name','media')
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    created_by  TEXT,
    version     INTEGER NOT NULL DEFAULT 1,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT
);
```
**Uniqueness swap** (old → new), idempotent:
```sql
ALTER TABLE sensory_header_presets ADD COLUMN IF NOT EXISTS test_name TEXT;
ALTER TABLE sensory_header_presets DROP CONSTRAINT IF EXISTS sensory_header_presets_kind_value_key;
DROP INDEX IF EXISTS idx_sensory_header_presets_kind;     -- old non-unique (kind,value) index
CREATE UNIQUE INDEX IF NOT EXISTS uq_shp_kind_value_test
    ON sensory_header_presets (kind, value, COALESCE(test_name, ''));
CREATE INDEX IF NOT EXISTS idx_shp_kind_test
    ON sensory_header_presets (kind, test_name);          -- supports loadSampleNamesForTest
```
**Matching upsert** (the conflict target MUST equal the unique-index expression):
```sql
INSERT INTO sensory_header_presets (kind, value, test_name, created_by, updated_by)
VALUES (?, ?, ?, ?, ?)
ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING;
```
- `sample_name` rows: bind the real `test_name`. `test_name`/`media` rows: bind NULL → `COALESCE`→`''` so they dedup by `(kind, value)` exactly as before.

> The triggers on the table (`bump_version`, `notify_row_change`) are registered by a `DO $$ ... $$` block in `deploy/postgres/migrations/2026-05-20-sensory-header-presets.sql:35-45`. When `ensureSchema()` creates the table fresh, it must also register those triggers (copy that block, idempotent via `DROP TRIGGER IF EXISTS` then `CREATE TRIGGER`). If the table already existed, the triggers are already there.

---

## Task 1: `ensureSchema()` heals `sensory_header_presets` structure + test

**Files:** `src/database/DatabaseManager.cpp` (`ensureSchema`, ~127-183); test in `tests/tst_databasemanager/tst_databasemanager.cpp`.

> Read `ensureSchema()` first: it currently loops `kAdditiveColumns[]` with a `pg_attribute` guard. Read the trigger `DO $$` block in `deploy/postgres/migrations/2026-05-20-sensory-header-presets.sql`. Read the out-of-band test helpers (`columnExistsOob`, `dropColumnOob`, the 2nd-connection pattern) and `ensureSchema_reAddsDroppedAdditiveColumns` (~1463) — your test mirrors it.

- [ ] **Step 1: Write the failing test.** Add `sensoryHeaderPresets_ensureSchemaHealsStructure`:
  - Out-of-band: `DROP TABLE IF EXISTS sensory_header_presets CASCADE;`
  - Call `db.ensureSchema();`
  - Assert (out-of-band queries): the table exists, has a `test_name` column, has a unique index whose definition contains `COALESCE(test_name`, and does NOT have the old `sensory_header_presets_kind_value_key` constraint.
  - Idempotency: call `db.ensureSchema()` again → still exactly one unique index, no error.
  - Round-trip: `QVERIFY(db.saveSensoryHeaderPresets("TestA", "Media1", {"S1","S2"}));` succeeds (proves the table is usable after heal — this is also the fix for the pre-existing `sensoryHeaderPresets_roundTrip` failure).

```cpp
void sensoryHeaderPresets_ensureSchemaHealsStructure()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
    DVE::DatabaseManager db; QVERIFY(openDb(db));
    runOob([](QSqlQuery& q){ QVERIFY(q.exec("DROP TABLE IF EXISTS sensory_header_presets CASCADE")); });

    db.ensureSchema();

    bool tableExists=false, hasTestName=false, hasCoalesceIdx=false, hasOldConstraint=false;
    runOob([&](QSqlQuery& q){
        q.exec("SELECT to_regclass('public.sensory_header_presets') IS NOT NULL"); if(q.next()) tableExists=q.value(0).toBool();
    });
    QVERIFY(tableExists);
    runOob([&](QSqlQuery& q){
        q.exec("SELECT 1 FROM information_schema.columns WHERE table_name='sensory_header_presets' AND column_name='test_name'"); hasTestName=q.next();
    });
    QVERIFY(hasTestName);
    runOob([&](QSqlQuery& q){
        q.exec("SELECT indexdef FROM pg_indexes WHERE tablename='sensory_header_presets'");
        while(q.next()){ if(q.value(0).toString().contains("COALESCE(test_name")) hasCoalesceIdx=true; }
    });
    QVERIFY(hasCoalesceIdx);
    runOob([&](QSqlQuery& q){
        q.exec("SELECT 1 FROM pg_constraint WHERE conname='sensory_header_presets_kind_value_key'"); hasOldConstraint=q.next();
    });
    QVERIFY(!hasOldConstraint);

    db.ensureSchema();                                    // idempotent: no throw, still healthy
    QVERIFY(db.saveSensoryHeaderPresets("TestA", "Media1", {"S1","S2"}));
}
```
> Use the suite's real out-of-band runner (the DV-4 tasks added `setSensoryScoreOutOfBand`/used a 2nd-connection lambda; reuse that pattern — name it `runOob` or inline the 2nd `QPSQL` connection as the existing helpers do).

- [ ] **Step 2: Run — verify FAIL.** `tests\run-tests.ps1 -Filter tst_databasemanager` (PG up). Expected: FAIL — after dropping the table, `ensureSchema()` (today) does not recreate it, so `to_regclass` is NULL / the save returns false.

- [ ] **Step 3: Extend `ensureSchema()`.** After the existing `kAdditiveColumns` loop, add a `sensory_header_presets` reconciliation. Use `m_pg->queryDb()` (the same `QSqlDatabase&` the function already uses). Catalog-guard each expensive step so the common path is lock-free:
```cpp
    // DATAVIEWER-2: ensure sensory_header_presets exists with the test-scoped shape.
    // Also heals the init.sql drift (the table lives only in a migration file).
    {
        QSqlDatabase& db = m_pg->queryDb();
        QSqlQuery chk(db);
        chk.exec("SELECT to_regclass('public.sensory_header_presets')");
        const bool tableMissing = !(chk.next() && !chk.value(0).isNull());
        if (tableMissing) {
            QSqlQuery c(db);
            if (!c.exec(R"(CREATE TABLE IF NOT EXISTS sensory_header_presets (
                id BIGSERIAL PRIMARY KEY,
                kind TEXT NOT NULL CHECK (kind IN ('test_name','media','sample_name')),
                value TEXT NOT NULL CHECK (length(trim(value)) > 0),
                test_name TEXT,
                created_at TIMESTAMPTZ NOT NULL DEFAULT now(), created_by TEXT,
                version INTEGER NOT NULL DEFAULT 1,
                updated_at TIMESTAMPTZ NOT NULL DEFAULT now(), updated_by TEXT))"))
                logDebug(QStringLiteral("ensureSchema: create sensory_header_presets failed: %1").arg(c.lastError().text()));
            // triggers (copy the DO$$ block from the 2026-05-20 migration; idempotent)
            QSqlQuery t(db);
            t.exec(R"(DO $$ BEGIN
                DROP TRIGGER IF EXISTS trg_shp_bump_version ON sensory_header_presets;
                CREATE TRIGGER trg_shp_bump_version BEFORE UPDATE ON sensory_header_presets
                    FOR EACH ROW EXECUTE FUNCTION bump_version();
                DROP TRIGGER IF EXISTS trg_shp_notify ON sensory_header_presets;
                CREATE TRIGGER trg_shp_notify AFTER INSERT OR UPDATE OR DELETE ON sensory_header_presets
                    FOR EACH ROW EXECUTE FUNCTION notify_row_change();
            END $$;)");
            // ^ match the EXACT trigger + function names used in the 2026-05-20 migration; read it and copy.
        }
        // additive column (cheap pg_attribute guard like kAdditiveColumns)
        QSqlQuery colq(db);
        colq.prepare("SELECT 1 FROM pg_attribute WHERE attrelid = CAST(? AS regclass) AND attname='test_name' AND NOT attisdropped");
        colq.addBindValue(QStringLiteral("sensory_header_presets"));
        bool hasCol = colq.exec() && colq.next();
        if (!hasCol) { QSqlQuery a(db); if(!a.exec("ALTER TABLE sensory_header_presets ADD COLUMN IF NOT EXISTS test_name TEXT")) logDebug(...); }
        // index swap, guarded so it runs at most once
        QSqlQuery idxq(db);
        idxq.exec("SELECT 1 FROM pg_indexes WHERE tablename='sensory_header_presets' AND indexdef LIKE '%COALESCE(test_name%'");
        const bool newIdxExists = idxq.next();
        if (!newIdxExists) {
            QSqlQuery s(db);
            s.exec("ALTER TABLE sensory_header_presets DROP CONSTRAINT IF EXISTS sensory_header_presets_kind_value_key");
            s.exec("DROP INDEX IF EXISTS idx_sensory_header_presets_kind");
            if(!s.exec("CREATE UNIQUE INDEX IF NOT EXISTS uq_shp_kind_value_test ON sensory_header_presets (kind, value, COALESCE(test_name, ''))"))
                logDebug(...);
            s.exec("CREATE INDEX IF NOT EXISTS idx_shp_kind_test ON sensory_header_presets (kind, test_name)");
        }
    }
```
> Confirm `m_pg`, `queryDb()`, `logDebug`, and the trigger function names (`bump_version`/`notify_row_change`) against the actual code + the 2026-05-20 migration. Keep every op best-effort (log, never throw) consistent with the rest of `ensureSchema()`. Fill the `logDebug(...)` placeholders with real messages.

- [ ] **Step 4: Run — verify PASS** + confirm `sensoryHeaderPresets_roundTrip` (the pre-existing failure) now PASSES too. `tests\run-tests.ps1 -Filter tst_databasemanager`.

- [ ] **Step 5: Commit.**
```bash
git add src/database/DatabaseManager.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "feat(db): ensureSchema heals sensory_header_presets to test-scoped shape (DATAVIEWER-2)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 2: Test-scope `saveSensoryHeaderPresets` + add `loadSampleNamesForTest` + test

**Files:** `src/database/DatabaseManager.cpp` (`saveSensoryHeaderPresets` ~2746-2789, `loadSensoryHeaderPresets` ~2791-2812), `src/database/DatabaseManager.h`; test in `tst_databasemanager.cpp`. Depends on Task 1 (the table/column/index exist).

> Read `saveSensoryHeaderPresets` — it receives `(const QString& testName, const QString& media, const QStringList& sampleNames)` and currently inserts sample_name rows WITHOUT `test_name`, `ON CONFLICT (kind, value) DO NOTHING`. Read `loadSensoryHeaderPresets(kind) const`.

- [ ] **Step 1: Write the failing test.** `sampleNames_scopedByTest`:
```cpp
void sampleNames_scopedByTest()
{
    if (qEnvironmentVariableIsEmpty("DVE_TEST_PG_CONN")) QSKIP("no test PG");
    DVE::DatabaseManager db; QVERIFY(openDb(db));
    db.ensureSchema();
    runOob([](QSqlQuery& q){ q.exec("DELETE FROM sensory_header_presets"); });

    QVERIFY(db.saveSensoryHeaderPresets("TestA", "M", {"Alpha","Bravo"}));
    QVERIFY(db.saveSensoryHeaderPresets("TestB", "M", {"Charlie"}));
    QVERIFY(db.saveSensoryHeaderPresets("TestA", "M", {"Alpha","Bravo"}));   // repeat: must NOT error (ON CONFLICT)

    QStringList a = db.loadSampleNamesForTest("TestA");
    QStringList b = db.loadSampleNamesForTest("TestB");
    QCOMPARE(a, (QStringList{"Alpha","Bravo"}));     // scoped to A (ORDER BY lower(value))
    QCOMPARE(b, (QStringList{"Charlie"}));           // scoped to B; A's names absent
    // test_name / media presets stay global + still de-dup by (kind,value)
    QVERIFY(db.loadSensoryHeaderPresets("test_name").contains("TestA"));
    QVERIFY(db.loadSensoryHeaderPresets("media").contains("M"));
}
```

- [ ] **Step 2: Run — verify FAIL** (`loadSampleNamesForTest` undeclared; or A/B not scoped).

- [ ] **Step 3: Implement.** In `saveSensoryHeaderPresets`, change the sample-name insert to bind `test_name` and use the new conflict target. The exact insert (match the index expression):
```cpp
    q.prepare("INSERT INTO sensory_header_presets (kind, value, test_name, created_by, updated_by) "
              "VALUES (?, ?, ?, ?, ?) "
              "ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING");
```
  - For `sample_name` rows bind `testName`; for `test_name`/`media` rows bind a NULL `QVariant(QMetaType(QMetaType::QString))` (or `QVariant()`), so `COALESCE`→`''`. Keep the empty-value skip guard. (If the function uses a shared `insertOne` lambda, thread `test_name` through it.)
  - Declare + implement `QStringList DatabaseManager::loadSampleNamesForTest(const QString& testName) const`:
```cpp
QStringList DatabaseManager::loadSampleNamesForTest(const QString& testName) const
{
    if (!m_online || !isOpen()) return {};
    QStringList out; QSqlQuery q(m_pg->queryDb());
    q.prepare("SELECT value FROM sensory_header_presets "
              "WHERE kind='sample_name' AND test_name = ? ORDER BY lower(value)");
    q.addBindValue(testName);
    if (q.exec()) while (q.next()) out << q.value(0).toString();
    return out;
}
```
  Declare in `DatabaseManager.h` beside `loadSensoryHeaderPresets`.

- [ ] **Step 4: Run — verify PASS** + suite regression.

- [ ] **Step 5: Commit.**
```bash
git add src/database/DatabaseManager.cpp src/database/DatabaseManager.h tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "feat(db): store + query sample-name presets per test (DATAVIEWER-2)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 3: One-time gated backfill in `ensureSchema()` + test

**Files:** `src/database/DatabaseManager.cpp` (`ensureSchema`); test in `tst_databasemanager.cpp`. Depends on Tasks 1-2.

> The backfill reconstructs (test → sample-name) from session history so the scoped dropdown is useful immediately. It runs ONCE, gated by a `schema_meta` marker, so it's not re-run on every connect. Confirm `schema_meta(key TEXT PK, value TEXT)` exists (init.sql ~238-242); if `ensureSchema` can't assume it, `CREATE TABLE IF NOT EXISTS schema_meta (key TEXT PRIMARY KEY, value TEXT)` first.

- [ ] **Step 1: Write the failing test.** `sampleNames_backfilledFromSessions`:
  - Out-of-band: ensure a sensory session row exists with `session_name='HistTest'` and `json_data` containing `samples:[{name:'Hist1'},{name:'Hist2'}]` (insert via the suite's session helper `makeSensorySession`+save, or a raw OOB insert). Also delete the backfill marker: `DELETE FROM schema_meta WHERE key='dv2_sample_name_backfill';` and `DELETE FROM sensory_header_presets WHERE test_name='HistTest';`
  - Call `db.ensureSchema();`
  - Assert `db.loadSampleNamesForTest("HistTest")` contains `Hist1` + `Hist2`.
  - Idempotency: delete the marker again, call `ensureSchema()` again → no duplicate rows / no error (ON CONFLICT). (Optionally assert the marker is set so a normal second connect skips the scan.)

- [ ] **Step 2: Run — verify FAIL.**

- [ ] **Step 3: Implement the gated backfill** at the END of `ensureSchema()` (after the structure block from Task 1):
```cpp
    // DATAVIEWER-2: one-time backfill of (test -> sample-name) pairs from session
    // history, so the scoped dropdown is useful from day one. Gated by a schema_meta
    // marker so it runs once, not on every connect. Idempotent (ON CONFLICT).
    {
        QSqlDatabase& db = m_pg->queryDb();
        QSqlQuery g(db);
        g.exec("SELECT 1 FROM schema_meta WHERE key='dv2_sample_name_backfill'");
        if (!g.next()) {
            QSqlQuery b(db);
            const char* kSensoryBackfill =
                "INSERT INTO sensory_header_presets (kind, value, test_name, created_by, updated_by) "
                "SELECT DISTINCT 'sample_name', smp->>'name', ss.session_name, 'backfill', 'backfill' "
                "FROM sensory_sessions ss "
                "CROSS JOIN LATERAL jsonb_array_elements(ss.json_data->'samples') AS smp "
                "WHERE ss.session_name IS NOT NULL AND length(trim(ss.session_name))>0 "
                "  AND smp->>'name' IS NOT NULL AND length(trim(smp->>'name'))>0 "
                "ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING";
            if (!b.exec(kSensoryBackfill)) logDebug(QStringLiteral("ensureSchema: sensory backfill failed: %1").arg(b.lastError().text()));
            QSqlQuery d(db);
            const char* kDetailedBackfill =
                "INSERT INTO sensory_header_presets (kind, value, test_name, created_by, updated_by) "
                "SELECT DISTINCT 'sample_name', smp->>'name', ds.session_name, 'backfill', 'backfill' "
                "FROM detailed_sensory_sessions ds "
                "CROSS JOIN LATERAL jsonb_array_elements(ds.json_data->'samples') AS smp "
                "WHERE ds.session_name IS NOT NULL AND length(trim(ds.session_name))>0 "
                "  AND smp->>'name' IS NOT NULL AND length(trim(smp->>'name'))>0 "
                "ON CONFLICT (kind, value, COALESCE(test_name, '')) DO NOTHING";
            if (!d.exec(kDetailedBackfill)) logDebug(QStringLiteral("ensureSchema: detailed backfill failed: %1").arg(d.lastError().text()));
            QSqlQuery m(db);
            m.exec("INSERT INTO schema_meta (key, value) VALUES ('dv2_sample_name_backfill', now()::text) "
                   "ON CONFLICT (key) DO NOTHING");
        }
    }
```
> Confirm the detailed table is named `detailed_sensory_sessions` and that its `json_data` uses the same `samples[].name` shape (the explorer confirmed both do). Confirm `schema_meta`'s PK column is `key` (init.sql). If a backfill statement errors on a DB that lacks a session table, the `logDebug` + continue keeps it best-effort.

- [ ] **Step 4: Run — verify PASS** + suite regression (no new failures; the backfill marker means a re-open doesn't re-scan).

- [ ] **Step 5: Commit.**
```bash
git add src/database/DatabaseManager.cpp tests/tst_databasemanager/tst_databasemanager.cpp
git commit -m "feat(db): one-time backfill of sample-name presets from session history (DATAVIEWER-2)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 4: UI — scope the dropdown to current test + cap popup height (both panels)

**Files:** `src/ui/SensoryPanel.cpp`, `src/ui/DetailedSensoryPanel.cpp`. GUI — verify by clean build + review (no headless test).

> Read `attachPresetDropdown(QLineEdit*, std::function<QStringList()>)` (SensoryPanel.cpp ~51-76, and the copy in DetailedSensoryPanel.cpp): it builds a fresh `QMenu` per click and `menu.exec(pos)` — unbounded height. Read the sample-name provider wiring: SensoryPanel `addSampleCard` (~890-893) sets `card->attachNamePresetDropdown([this]{ return m_db->loadSensoryHeaderPresets("sample_name"); })`; DetailedSensoryPanel (~360-363) similar for `m_sampleNameEdit`. Confirm both panels have `m_testTitleEdit` + `m_db`.

- [ ] **Step 1: Scope the sample-name providers to the current test.** SensoryPanel `addSampleCard`:
```cpp
    card->attachNamePresetDropdown([this]() -> QStringList {
        if (!m_db) return {};
        const QString test = m_testTitleEdit->text().trimmed();
        if (test.isEmpty()) return {};                 // no test yet -> empty (avoids the old global flood)
        return m_db->loadSampleNamesForTest(test);
    });
```
DetailedSensoryPanel sample-name wiring: same change with `loadSampleNamesForTest(m_testTitleEdit->text().trimmed())`.
> Only the `sample_name` provider changes. The `test_name` and `media` dropdowns keep using the global `loadSensoryHeaderPresets(...)` (those stay global). Confirm which `attach*` calls are which and leave test_name/media alone.

- [ ] **Step 2: Cap the popup height** in BOTH `attachPresetDropdown` copies, before `menu.exec(pos)`:
```cpp
        // DATAVIEWER-2: bound the popup so a long list can never cover the screen.
        const int rowH = edit->sizeHint().height() > 0 ? edit->sizeHint().height() : 24;
        menu.setMaximumHeight(rowH * 12 + 8);          // ~12 rows; QMenu paginates with scroll arrows past this
```
> `QMenu` honors `maximumHeight` and shows scroll arrows beyond it. `setMaxVisibleItems` does NOT apply to QMenu. If a cleaner cap is available (e.g. the menu's `view()`), use it; the goal is ~10-12 visible rows + scroll. Apply to BOTH copies (sensory + detailed); if you prefer, hoist `attachPresetDropdown` to a shared helper, but a 2-line duplicate edit is acceptable and lower-risk.

- [ ] **Step 3: Build** the app (incremental release in `build\` or repo root): warning-free under -Werror.

- [ ] **Step 4: Commit.**
```bash
git add src/ui/SensoryPanel.cpp src/ui/DetailedSensoryPanel.cpp
git commit -m "feat(ui): scope sample-name dropdown to current test + cap popup height (DATAVIEWER-2)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 5: Canonical migration `.sql` + OfflineSnapshot note

**Files:** create `deploy/postgres/migrations/2026-06-07-dv2-sample-name-by-test.sql`; (no OfflineSnapshot code change — document only).

- [ ] **Step 1: Write the canonical migration file** (via Python delete-and-rewrite per MIP convention, or a normal write — it's SQL, low MIP risk). Mirror what `ensureSchema()` does, as the durable record + for the eventual init.sql reconciliation (task `task_9999578f`) and any manual NAS application:
```sql
-- 2026-06-07 DATAVIEWER-2: sample-name presets scoped per test.
-- Idempotent. Mirrors DatabaseManager::ensureSchema() (the runtime self-heal).
BEGIN;
CREATE TABLE IF NOT EXISTS sensory_header_presets ( ... new shape incl. test_name ... );  -- (copy from the schema block above)
ALTER TABLE sensory_header_presets ADD COLUMN IF NOT EXISTS test_name TEXT;
ALTER TABLE sensory_header_presets DROP CONSTRAINT IF EXISTS sensory_header_presets_kind_value_key;
DROP INDEX IF EXISTS idx_sensory_header_presets_kind;
CREATE UNIQUE INDEX IF NOT EXISTS uq_shp_kind_value_test ON sensory_header_presets (kind, value, COALESCE(test_name, ''));
CREATE INDEX IF NOT EXISTS idx_shp_kind_test ON sensory_header_presets (kind, test_name);
-- triggers: copy the DO$$ block from 2026-05-20-sensory-header-presets.sql
-- backfill (sensory + detailed) — the two INSERT...SELECT...jsonb_array_elements statements from Task 3
INSERT INTO schema_meta (key, value) VALUES ('dv2_sample_name_backfill', now()::text) ON CONFLICT (key) DO NOTHING;
COMMIT;
```
Fill in the elided parts exactly matching Tasks 1+3. This file is NOT auto-applied by the test harness; it's the canonical record (the runtime delivery is `ensureSchema()`).

- [ ] **Step 2: Document OfflineSnapshot scope.** Add a brief comment near the OfflineSnapshot table list (or in the migration file header) noting: `sensory_header_presets` is intentionally NOT mirrored to the offline SQLite snapshot (it never was); the sample-name dropdown is empty offline today and remains so — offline preset suggestions are a possible future enhancement, out of scope for DV-2. (No code change — this preserves current behavior and avoids scope creep. Acceptance per spec §6: no data loss, snapshot logic unchanged.)

- [ ] **Step 3: Commit.**
```bash
git add deploy/postgres/migrations/2026-06-07-dv2-sample-name-by-test.sql
git commit -m "docs(db): canonical SQL migration for test-scoped sample-name presets (DATAVIEWER-2)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

## Task 6: Wrap the batch into deployable v2.4.0 — regression, build, installer, evidence review

**Files:** `DataViewerEnterprise.pro`.

- [ ] **Step 1: Full regression sweep.** `tests\run-tests.ps1` (PG up). REQUIRED evidence:
  - New DV-2 tests pass (`sensoryHeaderPresets_ensureSchemaHealsStructure`, `sampleNames_scopedByTest`, `sampleNames_backfilledFromSessions`).
  - **`sensoryHeaderPresets_roundTrip` now PASSES** (it failed all sprint due to the missing table; Task 1 fixes it). This is a key evidence point — confirm the known-failure count DROPS.
  - All DV-3/DV-4/Zero-Save-As tests still pass.
  - Remaining failures are only the truly-unrelated pre-existing ones (`tst_storedfns` probe, `tst_responsivelayout` segfault, stale `tst_excelexporter` artifact). REGRESSIONS: none.

- [ ] **Step 2: Bump VERSION** `2.3.5` → `2.4.0` in `DataViewerEnterprise.pro`.

- [ ] **Step 3: Clean release build** (VERSION change needs clean rebuild). Use a `.bat`/PowerShell wrapper invoked by absolute path (the Bash tool mangles chained `qmake && make`):
```
python tools\decrypt_via_copy.py --apply
qmake CONFIG+=release (repo root, DESTDIR=release\) ; mingw32-make clean ; mingw32-make -j8
```
Verify `release\DataViewer.exe` FileVersion = `2.4.0.0`, warning-free under -Werror.

- [ ] **Step 4: Build the installer** (do NOT touch Synology): `build_installer.bat` → `dist\DataViewer-setup.exe` (version guard must pass at 2.4.0).

- [ ] **Step 5: Commit.**
```bash
git add DataViewerEnterprise.pro
git commit -m "chore(release): v2.4.0 deployable -- DV-3 + DV-4 + zero-Save-As + DV-2 (sprint wrap)

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 6: Final evidence-based review (whole sprint).** Dispatch a final integration reviewer over the full v2.4.0 diff (since v2.3.1 baseline or the batch branch point) focused on cross-fix interactions: the save/close/guard/auto-name/scoped-dropdown paths must compose; ensureSchema's growing set of heals (puffing_regime, raw_grid, sensory_header_presets + backfill) must all be idempotent + lock-free on the common path; no fix regresses another. Record the verdict.

- [ ] **Step 7: Prepare manual Plane updates** (Plane MCP is disconnected). Produce, in the final summary to Charlie, paste-ready text to: set DATAVIEWER-2 → *Ready for Release* (with a resolution comment), and file the deferred §8 follow-up issue ("Remove Ctrl+U; make LiveSync authoritative; eliminate the batch overwrite").

---

## Risks & notes

- **Index/ON CONFLICT must match exactly** (`(kind, value, COALESCE(test_name,''))`). Mismatch → every insert throws. The `sampleNames_scopedByTest` repeated-save assertion is the guard.
- **NULL vs '' for test_name on global rows:** `COALESCE(test_name,'')` makes NULL-test rows dedup by `(kind,value)`, preserving the old behavior for `test_name`/`media`. Bind NULL for those kinds.
- **ensureSchema on the live multi-user DB:** the index swap takes a brief lock but runs at most once (guarded by `indexdef LIKE '%COALESCE(test_name%'`); `sensory_header_presets` is tiny. Common path is catalog-checked and lock-free. Backfill is gated once via `schema_meta`.
- **Legacy rows preserved:** old NULL-`test_name` sample_name rows are not deleted; they simply stop matching `loadSampleNamesForTest`. Acceptance: distinct historical sample-name count never decreases (the backfill only INSERTs).
- **Empty test title → empty dropdown** (not the old global flood). Combined with Zero-Save-As (DATAVIEWER-8) requiring a test name before save, users will normally have a title set; the scoped list fills in.
- **Offline unchanged:** presets were never in the snapshot; offline dropdown stays empty. Documented, not a regression.
- **Drift reconciliation:** making `ensureSchema()` create `sensory_header_presets` also fixes the long-standing `sensoryHeaderPresets_roundTrip` failure. The `.sql` file + `task_9999578f` still fold this into `init.sql` so fresh containers are correct without the self-heal.
