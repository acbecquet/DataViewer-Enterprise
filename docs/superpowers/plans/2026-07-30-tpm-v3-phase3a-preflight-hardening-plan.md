# TPM v3 Phase 3a (pre-flight hardening) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Close the six defects the Phase 3 pre-flight audit surfaced, so that the schema migration in 3b/3c/3d cannot be silently masked by broken tooling or compounded by an existing save-path bug.
No schema changes, no new tables, no behavior a user would notice - except that three currently-silent failures become loud.

**Architecture:** Every item here is surgical and independently verifiable today.
Two of them (tasks 2 and 3) are places where a correctness failure is currently invisible in release builds; one (task 1) is the verification tool that every later Phase 3 gate reads through, so it is fixed first.
Nothing in this phase touches `DatabaseOps::persistFileCore`, the wide column lists, or any SQL statement text - that is 3b onward.

**Tech Stack:** C++17 / Qt 6.10 (qmake + MinGW), Qt Test, PowerShell 5.1, Docker (`dve-test-pg`, postgres:16 on port 5433).

---

## Machine + repo rules (read first)

- Create all new source files with the Write tool ONLY - never python file writes (MIP labeling), never heredoc/echo.
- Ciphertext (`%TSD-Header-###%`): read via `git show HEAD:<path>`; run `python tools/decrypt_via_copy.py --apply` from repo root before any build.
- Public repo: `tests/corpus/` gitignored; never commit real workbooks or `results.txt` artifacts.
- Branch `worktree-tpm-template-v3-research`. Commit per task; plain dashes; NO Co-Authored-By.
- Qt Test stdout is INVISIBLE - always `-o results.txt,txt`, and `/c/Qt/6.10.1/mingw_64/bin` MUST be on PATH when running test exes (silent death otherwise).
- Suite inner loop (from suite dir): `export PATH="/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH"`, qmake, mingw32-make, `release/<suite>.exe -o results.txt,txt`.
- Full-suite gate: `powershell -ExecutionPolicy Bypass -File tests/run-tests.ps1` from repo root. NEVER run it concurrently with an app build (-j8 CPU starvation makes python-subprocess suites flake).
- `dve-test-pg` is already running. Never touch the production/NAS database.
- NEVER kill Excel processes.
- -Werror -Wall -Wextra -Wpedantic.

## Reference: verified seams (2026-07-30)

- `tests/start-test-postgres.ps1`: `:57-67` slices `init.sql` from the first `BEGIN;` to the first `COMMIT;`. Verified: `init.sql` has exactly one `BEGIN;` (line 5) and one `COMMIT;` (line 571), and lines 573-613 are the pg_cron tail (`CREATE EXTENSION pg_cron` + three `SELECT cron.schedule(...)`), which stock `postgres:16` cannot run. The slicing is therefore correct **today** and silently drops anything added after line 571 **tomorrow**.
- `tests/start-test-postgres.ps1:79-87`: applies every `deploy/postgres/migrations/*.sql` with `-v ON_ERROR_STOP=0` and `| Out-Null`, with no per-file exit check. The in-file comment (`:69-78`) justifies `ON_ERROR_STOP=0` by "the migrations are idempotent, so it tolerates the expected already-exists notices". `IF NOT EXISTS` / `CREATE OR REPLACE` do not raise errors, so this justification does not hold for a fresh container; verify empirically before relaxing anything.
- `deploy/postgres/migrations/2026-06-25-dv15-rekey-forked-sensory.sql` carries a header saying it must NOT be auto-applied (run manually after a backup). The loop applies it anyway. Harmless on an empty container, but it must become an explicit, documented skip rather than an accident.
- `DatabaseManager::tryWriteFile(const FileResult&)` (`src/database/DatabaseManager.cpp:904-911`) intentionally copies and discards the writeback; `saveFile(const FileResult&)` (`:921-927`) is the bool shim over it. The overloads are fine - the bug is one caller.
- `MainWindow.cpp:5102` (DB-browser load path) calls `m_db->saveFile(result)` immediately after adopting the loaded row's `id`/`version` at `:5096-5097`, so every Load-from-Database mints child ids and throws them away.
- `assertColumnArity` (`src/database/OfflineSnapshot.cpp:304-313`) is `Q_ASSERT_X` + `Q_UNUSED`, i.e. compiled out in release. It is called for `files` (`:805`), `sensory_sessions` (`:967`), `detailed_sensory_sessions` (`:1019`), and images (`:714`, `:768`) - and **not** for `tests`, `samples`, or `data_rows`, whose bind loops use the literals `13` (`:847`), `28` (`:888`), `19` (`:924`).
- `isSameLoadedPath(const QString&, const QString&)` is a file-static helper at `src/MainWindow.cpp:2445`, used at `:2459`, `:2572`, `:3265`, `:3274`, `:5110`, `:7097`. The recovery-restore dedup at `:5654` uses a raw `==` instead.
- `MainWindow::onStoryCellEdited` (`:3496-3500`) calls `m_liveSync->commitCell(...)` and discards the `bool`. `LiveSync::commitCell` returns `false` from the table/column allowlist gate (`src/database/LiveSync.cpp:141-151`) after only a `qWarning`, with no offline enqueue and no unsynced-counter bump.
- `MainWindow::markFileModified` (`:5766`) carries a comment claiming "all 7 TPM edit sites route here"; an exhaustive grep finds 5 call sites (`:2292`, `:3492`, `:7460`, `:7497`, `:7632`).

## Non-goals

- NO new tables, columns, views, or stored functions - that is 3b.
- NO changes to any SQL statement text in `DatabaseOps.cpp` or `DatabaseManager.cpp`'s read path.
- NO changes to `LiveSync`'s allowlist contents or to `liveColumnForDataCol` - those get replaced wholesale in 3d; 3a only makes their rejection visible.
- NO wiring of `VersionLookup` / re-enabling per-cell OCC (H6) - that decision belongs with the per-measurement versioning it exists to serve.
- NO touching `UniqueViolationDialog`, the inert `cell_focus` loop, or the unreachable `samples`/`tests`/`files` allowlist entries. They are logged in the index as dead code to resolve in 3d, not now.
- NO production database access of any kind.

---

### Task 1: Test-container provisioning fails loudly

**Files:**
- Modify: `tests/start-test-postgres.ps1`

**Problem:** a failing migration currently produces a clean-looking provisioning run and a silently wrong schema, and any DDL added to `init.sql` after its `COMMIT;` is invisible to the whole suite. Every Phase 3 gate reads through this script.

**Steps:**

- [ ] Empirically establish the baseline first: recreate the container from scratch with the current script, but with `ON_ERROR_STOP=1` and without `Out-Null`, and record which migrations (if any) actually fail on a fresh database. Do not guess - the in-file justification for `ON_ERROR_STOP=0` is unverified.
- [ ] Switch the migration loop to `-v ON_ERROR_STOP=1`, keep psql's output visible (capture it and echo on failure at minimum), and check the exit status per file. A non-zero exit throws with the migration's filename and psql's output.
- [ ] Any migration that legitimately cannot be applied to a fresh container goes into an explicit, named skip list with a one-line reason, checked by filename. `2026-06-25-dv15-rekey-forked-sensory.sql` is the known member (its own header forbids auto-application).
- [ ] Replace the "first `BEGIN;` to first `COMMIT;`" slice with a slice that ends at the documented pg_cron tail, and throw if the tail contains any statement that is not `CREATE EXTENSION ... pg_cron` or `SELECT cron.`. New DDL appended to `init.sql` must never be silently skipped again.
- [ ] Apply the same exit-status check to the `init.sql` block application (it is also piped to `Out-Null` today).

**Verification (this is the task's real gate):**

- [ ] Drill: copy a migration to a temp file inside `deploy/postgres/migrations/`, introduce a deliberate syntax error, run the script against a freshly recreated container, and confirm it throws and names the file. Remove the temp file afterwards and confirm a clean run succeeds.
- [ ] Recreate `dve-test-pg` from scratch with the fixed script and confirm the schema is complete: both `dve_commit_cell` overloads (`pronargs` 5 and 6), both `dve_commit_cell_json` overloads (6 and 7), and `data_rows.puffing_regime` all present.
- [ ] Re-run the PG-dependent suites (`tst_livesync`, `tst_storedfns`, `tst_databasemanager`) against the rebuilt container - all green.

---

### Task 2: DB-browser load keeps its writeback (H7)

**Files:**
- Modify: `src/MainWindow.cpp` (the `:5102` call site)
- Test: `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp`

**Problem:** `MainWindow.cpp:5102` calls `m_db->saveFile(result)`, which routes through the const-ref `tryWriteFile` overload and discards the post-save `id`/`version` writeback by design. Every Load-from-Database therefore mints child ids and throws them away, so the *next* save re-INSERTs the whole subtree and the orphan prune deletes the originals. Harmless-ish today (row churn); under long format it becomes a full measurement-table rewrite on the most common bulk-load path, and it widens the H1 prune window every time.

**Steps:**

- [ ] Red first: add an E2E scenario that loads a file from the database, saves it, and asserts that the in-memory `FileResult`'s sheet/sample/row ids are non-negative afterwards - and that a second save performs UPDATEs rather than INSERT-plus-prune (assert stable child ids across the two saves).
- [ ] Change the call site to use the mutable-reference overload so the writeback lands. Do NOT change the semantics of the `const&` overloads or of `saveFile` - other callers rely on the fire-and-forget shim, and the discard is documented as intentional there.
- [ ] Confirm the new scenario passes and the existing 16 save-integrity scenarios stay green.

---

### Task 3: Snapshot column-arity guard works in release builds and covers the migrating tables (H14)

**Files:**
- Modify: `src/database/OfflineSnapshot.cpp`
- Test: `tests/tst_offlinesnapshot/tst_offlinesnapshot.cpp`

**Problem:** `assertColumnArity` is a `Q_ASSERT_X` that compiles out in release, and it is not called at all for `tests`, `samples`, or `data_rows` - precisely the three tables Phase 3 rewrites. Their bind loops use the bare literals `13`, `28`, and `19`. A column-count drift silently mis-copies snapshot data, and the offline read path is the one place a user sees data with no server to cross-check against.

**Steps:**

- [ ] Convert the helper into a runtime check that is active in release builds: on mismatch, log at critical severity with the table name and all three counts, and report failure to the caller.
- [ ] Make the regeneration abort on a mismatch rather than continuing. A stale-but-correct previous snapshot is strictly better than a silently mis-copied fresh one; the atomic tmp-then-rename write already guarantees the old snapshot survives an aborted regen.
- [ ] Wire the check for `tests`, `samples`, and `data_rows`, replacing the literal loop bounds with the checked count so the three numbers cannot drift apart again.
- [ ] Red first: a test that drives a regeneration whose SELECT and INSERT arities disagree and asserts the regen fails loudly instead of producing a snapshot. If the copy blocks cannot be reached directly from a test, exercise the helper itself and assert the failure report; state in the commit which of the two you did and why.
- [ ] Keep `-Werror -Wextra -Wpedantic` clean - the existing `Q_UNUSED` lines exist for that reason and will no longer be needed.

---

### Task 4: Recovery-restore dedup uses the shared path comparison (H11)

**Files:**
- Modify: `src/MainWindow.cpp` (`:5654` area)
- Test: `tests/tst_recoverymanager/tst_recoverymanager.cpp` if the seam is reachable, otherwise a focused unit test on the helper

**Problem:** the dirty set `m_modifiedFilePaths` is keyed by file path, and every other load path dedups with `isSameLoadedPath` (`:2459`, `:3274`, `:5110`, `:7097`) while the recovery restore at `:5654` uses a raw `==`. A normalization mismatch yields a dirty path with no matching loaded file; the save loop at `:5186-5188` iterates `m_loadedFiles` and can never reach it, so the file shows as permanently modified and its edits are never saved. `unsavedInventory` has a fallback for this case (`:5880-5881`), which is evidence the mismatch is considered reachable.

**Steps:**

- [ ] Red first if a seam exists: restore a recovery entry whose path differs from the loaded path only by the normalization `isSameLoadedPath` absorbs, and assert one loaded file rather than two.
- [ ] Replace the raw comparison with `isSameLoadedPath`.
- [ ] Confirm `tst_recoverymanager` stays green.

---

### Task 5: A rejected per-cell commit is visible (H4, partial)

**Files:**
- Modify: `src/MainWindow.cpp` (`:3496-3500`)

**Problem:** `LiveSync::commitCell` returns `false` when the table/column allowlist rejects the write, after only a `qWarning`, with no offline enqueue and no unsynced-edit counter bump. The sole TPM caller discards the return. In 3d the allowlist is replaced by metric-key routing, and during that cutover a rejected commit is exactly the failure mode that would look like "the edit just didn't save."

**Steps:**

- [ ] At the call site, check the return and log at warning severity with the file, sheet, sample, row, and resolved column when it is `false`. The whole-file save still covers the edit (`markFileModified` has already run), so this is a diagnosability fix, not a data-loss fix - say so in the comment and do not add a user-facing dialog.
- [ ] Keep it to the call site. Do NOT change `LiveSync`'s allowlist behavior or its return contract; 3d replaces both.

---

### Task 6: `markFileModified`'s comment matches reality

**Files:**
- Modify: `src/MainWindow.cpp` (`:5776-5778`)

**Problem:** the comment claims "all 7 TPM edit sites route here"; an exhaustive grep finds 5 (`:2292`, `:3492`, `:7460`, `:7497`, `:7632`). Either the comment is stale or two edit sites bypass the chokepoint - and the chokepoint is what guarantees recovery `noteDirty()` coverage for every TPM change.

**Steps:**

- [ ] Re-verify the call-site count independently. If it is 5, correct the number. If two sites genuinely bypass `markFileModified`, that is a recovery-coverage bug - fix it and say so in the commit message rather than editing the comment.

---

### Task 7: Gates and wrap

- [ ] `python tools/decrypt_via_copy.py --apply` from repo root, then a clean debug build (`-Werror` must stay clean).
- [ ] Full suite: `powershell -ExecutionPolicy Bypass -File tests/run-tests.ps1`, not concurrent with any build. Record the pass/fail/skip counts.
- [ ] Confirm the corpus shadow and round-trip harnesses are still green (Phase 2's byte-identity referee must not move).
- [ ] Independent review of the whole 3a diff against this plan.
- [ ] Update `docs/sprint-tracker.html` and the v3 memory topic file with 3a's outcome and any new ledger items.
