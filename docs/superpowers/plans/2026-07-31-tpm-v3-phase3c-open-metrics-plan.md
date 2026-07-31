# TPM v3 Phase 3c (open metrics persist) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `DataRow::extra` and `SampleResult::extra` - the open metric maps Phase 2 already fills for inference and manifest sheets - survive a database round-trip, in Postgres and in the offline snapshot.
The success criterion is concrete and smoke-testable: **the manifest demo workbook's custom `Coil Temp (C)` column survives save, close, and reload-from-database.** That has never worked.

**Architecture:** Open metrics move read and write **together**, because they have no wide-column equivalent and therefore cannot go stale (Phase 3 index, SEQUENCING CORRECTION).
The standard 13 columns are deliberately **not touched** - they keep flowing through the existing wide path byte-for-byte, so the regression surface is close to zero and the byte-identity referee cannot move.
Everything added here is additive: one supplementary read query per file load, one supplementary write block per save, three mirrored tables in the snapshot.
This also closes H1's live instance - the documented gap at `src/MainWindow.cpp:1181-1186` where a load-from-database followed by a save silently drops extras.

**Tech Stack:** C++17 / Qt 6.10 (qmake + MinGW), Qt Test, PostgreSQL 16 (`dve-test-pg`, port 5433), SQLite (offline snapshot).

---

## Machine + repo rules (read first)

- Create all new source files with the Write tool ONLY - never python file writes (MIP labeling), never heredoc/echo.
- **NEVER use `git stash`** - it rewrites the working tree, re-triggers the MIP labeler and re-encrypts source mid-session, and the stash stack is shared across worktrees and other sessions. Use a temporary commit instead.
- Ciphertext (`%TSD-Header-###%`): read via `git show HEAD:<path>`; run `python tools/decrypt_via_copy.py --apply` from repo root before any build.
- Public repo: `tests/corpus/` gitignored; never commit real workbooks or `results.txt` artifacts.
- Branch `worktree-tpm-template-v3-research`. Commit per task; plain dashes; NO Co-Authored-By.
- Qt Test stdout is INVISIBLE - always `-o results.txt,txt`, and `/c/Qt/6.10.1/mingw_64/bin` MUST be on PATH when running test exes (silent death otherwise).
- Suite inner loop (from suite dir): `export PATH="/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH"`, `export DVE_TEST_PG_CONN="host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"`, prepend `vendor/libpq-16` to PATH, qmake, mingw32-make, `release/<suite>.exe -o results.txt,txt`.
- Run PG-dependent suites ONE AT A TIME - they wipe shared tables.
- Full-suite gate: `powershell -ExecutionPolicy Bypass -File tests/run-tests.ps1` from repo root. NEVER concurrently with an app build.
- Never touch the production/NAS database. NEVER kill Excel processes.
- -Werror -Wall -Wextra -Wpedantic.

## Reference: verified seams

- **`DataRow::extra`** (`src/pipeline/ReportData.h:32-36`) and **`SampleResult::extra`** (`:82`) are `QMap<QString, QVariant>` keyed by canonical metric key. The header comment says outright they are "NOT yet persisted to Postgres (Phase 3)". They already round-trip through the recovery JSON via the typed envelope (`{"n"|"i"|"s"|"b"|"a"}`, `src/pipeline/ReportDataJson.cpp:72-157`) - no change needed there.
- **Write path:** `persistFileCore` (`src/database/DatabaseOps.cpp:166-803`) is one transaction (`:182` -> `:786`). Phase A captures pre-image child ids (`:327-371`), Phase B upserts depth-first (`:449-754`), Phase C prunes orphans child-first (`:759-784`). Child ids ARE cascaded back onto the passed-in struct (`:508` sheet, `:596` sample, `:663` row) - the contract comment at `:795-799` claiming otherwise is wrong and was flagged in 3a.
- **Read path:** `DatabaseManager::loadFile` (`:957-1240`); the `data_rows` bulk SELECT is `:1089-1097` with positional reads `:1102-1119`; samples `:1037-1045` / `:1050-1074`.
- **Offline snapshot:** `kCreateStatements[]` (`src/database/OfflineSnapshot.cpp:80-278`) mirrors 10 tables with PG-identical column names. `kSnapshotSchemaVersion = 3` at `:55`, enforced on open (`:1369-1375`) and as the incremental gate (`:510`). The freshness fingerprint hardcodes **9 tables across 5 coupled sites**: the SQL at `:377-385`, `kFingerprintTables[9]` at `:422-426`, the `parts.size() != 9` reject at `:430`, the `i < 9` loop at `:439`, and `OfflineSnapshot.h:127`. `checkColumnArity` (3a) is now release-active and aborts the regen on drift - every new copy block must call it.
- **New in 3b:** `metric_defs(id, kind, key, display_name, value_type, unit, role, tags, audit)` with `UNIQUE(kind, key)`; `measurements(id, sample_id, metric_id, sort_order, value_num, value_text, audit)` with `UNIQUE(sample_id, metric_id, sort_order)`; `sample_headers(id, sample_id, field_id, value_num, value_text, audit)` with `UNIQUE(sample_id, field_id)`. None of them carry a `notify_row_change` trigger (index D3). `ensureSchema` creates them and upserts the compiled registry.
- **The demo fixture** `tests/data/manifest_demo.xlsx` has a "Custom Coil Test" sheet whose manifest declares a custom `coil_temp` column; it currently parses into `DataRow::extra` and is invisible everywhere downstream.

## Non-goals

- **NO standard-metric migration.** The 13 `data_rows` value columns and the 22 `samples` value columns keep living in the wide tables and keep being read and written exactly as today. `dve_migrate_to_long_format()` is still never called. That is 3d.
- NO display of custom metrics in the UI, plots, or reports - that is Phase 4. 3c is verified by tests and by a database round-trip, not by eyeballs.
- NO LiveSync / per-cell commit changes, no `dve_commit_measurement`, no notify trigger - 3d.
- NO changes to the recovery JSON shape (extras already round-trip there).
- NO changes to the standard read/write SQL statement text - the additions are separate statements.
- NO production database access.

---

### Task 1: Metric-id resolution with auto-registration

**Files:**
- Create: `src/database/MetricDefCache.h`, `src/database/MetricDefCache.cpp`
- Modify: `DataViewerEnterprise.pro`
- Test: `tests/tst_v3longformat/tst_v3longformat.cpp`

**Why:** every extras write needs `metric_key -> metric_defs.id`. Standard keys are pre-seeded, but a custom column like `coil_temp` is not in `MetricRegistry` at all and must be registered on first save, or its data has nowhere to go.

**Steps:**

- [ ] A small class that loads `(kind, key) -> id` for a connection and resolves on demand.
- [ ] `ensureMetric(kind, key, valueTypeHint)` inserts a missing key with `INSERT ... ON CONFLICT (kind, key) DO NOTHING RETURNING id`, then re-selects on conflict so concurrent writers converge. Never update an existing row from here - `ensureSchema` owns registry-sourced fields.
- [ ] For an auto-registered custom key, set `display_name` to the key itself and `role` to `measured`. **Document that this is deliberate:** the human-facing display name lives in the workbook manifest, which is not plumbed down to the database layer, and custom columns are not displayed until Phase 4. Enriching it is a Phase 4 task, not a reason to plumb the schema through now.
- [ ] Infer `value_type` from the `QVariant`: floating/integral -> `number`, bool -> `bool`, `QByteArray` -> `image`, list-of-doubles -> `numberlist`, else `text`. Mirror the recovery JSON envelope's discrimination order (`ReportDataJson.cpp:124-140`) so the two agree.
- [ ] Red first: resolving a known seeded key returns its id without inserting; an unknown key inserts exactly one row and returns it; resolving the same unknown key twice inserts only once.

---

### Task 2: Write extras inside the save transaction

**Files:**
- Modify: `src/database/DatabaseOps.cpp`
- Test: `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp`, `tests/tst_v3longformat/tst_v3longformat.cpp`

**Steps:**

- [ ] Inside `persistFileCore`'s existing transaction, after a `data_rows` row has been upserted and `dr.id` is known, upsert that row's `extra` entries into `measurements` keyed `(sample_id = sr.id, metric_id, sort_order = ri)`. Use `INSERT ... ON CONFLICT (sample_id, metric_id, sort_order) DO UPDATE`.
- [ ] Same for `SampleResult::extra` into `sample_headers` keyed `(sample_id = sr.id, field_id)`, kind `header`.
- [ ] Route to `value_num` / `value_text` by the resolved `metric_defs.value_type`. Carry `updated_by`; let the trigger own `updated_at` / `version`.
- [ ] **Batch the writes.** Statement count is already `O(children)` (H8) and this multiplies it. Use multi-row `VALUES` per sample rather than a round-trip per extra.
- [ ] **Extend the orphan prune (Phase C) to `measurements` and `sample_headers`** so that removing a custom column actually removes its stored data. Capture the pre-image alongside the existing four tables and delete what was not re-written.
  **Guard this carefully and comment it:** the prune is only safe because in 3c `measurements` holds *exclusively* extras. When 3d puts standard metrics there, an in-memory model that fails to reproduce them would delete real data (H1). State that constraint in the code so 3d cannot miss it, and scope the delete to the file's own samples.
- [ ] Red first: a save with two extras writes exactly two `measurements` rows; a second save with one extra removed prunes exactly that row and leaves the other; an empty `extra` map writes nothing at all (sparse rule holds here too).
- [ ] Confirm the existing 16 save-integrity scenarios plus 3a's scenario 17 stay green - the standard path must be untouched.

---

### Task 3: Read extras back

**Files:**
- Modify: `src/database/DatabaseManager.cpp`
- Test: `tests/tst_databasemanager/tst_databasemanager.cpp`

**Steps:**

- [ ] After the existing `data_rows` bulk SELECT and positional reads, issue **one** supplementary query per file joining `measurements` to `metric_defs` and to the file's samples, ordered so results can be walked alongside the already-built rows. Populate `dr.extra[key]`. Do NOT modify the existing SELECT or its positional indices.
- [ ] Same for `sample_headers` -> `sr.extra`.
- [ ] Convert `value_num` / `value_text` back to `QVariant` using `metric_defs.value_type` so a round-trip preserves the type (a number comes back as a double, not a string). The `numberlist` type needs the same list representation `DataRow::extra` uses today - check how `draw_pressure_per_puff` is stored in memory before choosing.
- [ ] Red first: save a file with typed extras (double, string, bool, number list), reload it, and assert every value AND its `QVariant` type matches.
- [ ] Assert the standard 13 fields are byte-identical across the round-trip - this is the regression guard on Task 2.

---

### Task 4: Offline snapshot schema version 4

**Files:**
- Modify: `src/database/OfflineSnapshot.cpp`, `src/database/OfflineSnapshot.h`
- Test: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`

**Steps:**

- [ ] Mirror `metric_defs`, `measurements`, and `sample_headers` in `kCreateStatements[]`, matching PG column names exactly per the file's stated convention.
- [ ] Add a regen copy block for each, following the existing SELECT / INSERT / bind-loop triple pattern - and **call `checkColumnArity` for each** (3a made it release-active and fatal; a new block without it reintroduces exactly the drift H14 was about).
- [ ] Bump `kSnapshotSchemaVersion` 3 -> 4.
- [ ] Update the fingerprint from 9 tables to 12, in **all five coupled sites at once**: the SQL, `kFingerprintTables`, the `parts.size()` reject, the loop bound, and the declaration in the header. Prefer deriving the count from the array so the number cannot drift again.
- [ ] Populate `extra` in the offline `loadFile` read path, same conversion rules as Task 3.
- [ ] Note in the commit that every existing user snapshot is invalidated by the version bump and will do one full rebuild - safe by design, the fallback path is a regen.
- [ ] Confirm the incremental blob-diff scenarios still pass (`testDataOnlyChangePullsZeroImageBlobs`, `testImageChangePullsOnlyChanged`, `testIncrementalMatchesFullRebuild`).

---

### Task 5: The end-to-end gate

**Files:**
- Test: `tests/tst_v3longformat/tst_v3longformat.cpp`

- [ ] The headline scenario, using `tests/data/manifest_demo.xlsx`: process the workbook, save it to the database, discard the in-memory result, reload from the database, and assert the custom `coil_temp` values are present, correct, and correctly typed on every data row. Fail loudly if any is missing - this is the phase's reason to exist.
- [ ] Round-trip it a second time (reload, save, reload) and assert the values are still there - that is the H1 prune regression guard, and it is the exact sequence that silently loses data today.
- [ ] Repeat the round-trip against the offline snapshot rather than Postgres.

---

### Task 6: Gates and wrap

- [ ] `python tools/decrypt_via_copy.py --apply`, then a clean app build - `-Werror` clean.
- [ ] Full suite, not concurrent with a build. Record counts.
- [ ] Corpus shadow and corpus round-trip (`DVE_TEST_CORPUS_DIR=tests/corpus`) must still be exactly 26/0/6 and 38/0/0. 3c must not be able to move them; if either moves, stop and investigate before continuing.
- [ ] Independent review of the whole 3c diff against this plan.
- [ ] Cut an internal smoke installer (version bump, clean root release rebuild, `build_installer.bat`) and write a release overview whose headline check is: open the manifest demo, save to database, close, reload from the database browser, and confirm the custom column's data survived. Copy the installer, the demo workbook, and the overview to the main repo's `dist\` / `release_overview\`.
- [ ] **NEVER place any 2.10.x installer on Synology** - the public release is frozen until v3.0.0.
- [ ] Update the tracker and the v3 memory topic file.
