# v2.5.0 Save/Sync Integrity Fixes — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax.
> **Root-cause dossier (read first):** `docs/superpowers/specs/2026-06-10-v24-save-sync-regressions-evidence.md`

**Goal:** eliminate the v2.4.0 data-loss/false-duplicate regressions: a save is NEVER silently dropped; the local client's own edits are authoritative for cells it touched; rename/duplicate collisions self-resolve with `_1/_2/_3` suffixes; close never hard-blocks.

**Architecture:** (1) fresh-version OCC inside the tryWrite* cores (read version+json in one SELECT, bind the fresh version) so routine VersionMismatch disappears; (2) caller-side mismatch/rowdeleted → bounded retry / re-INSERT, never skip-and-clear-dirty; (3) panels keep a per-session dirty-cell set; merge preserves DB values only for non-dirty cells; (4) LiveSync worker reconnects on broken connections, tryWrite* error paths ROLLBACK; (5) collision auto-suffix loop replaces the modal error; (6) close offers Name/Discard/Cancel.

**Tech:** C++17/Qt 6.10.1, qmake+MinGW, namespace DVE, `-Werror -Wall -Wextra`. Tests: Qt Test vs ephemeral Postgres (`tests\start-test-postgres.ps1`, container `dve-test-pg` already running, env `DVE_TEST_PG_CONN`). MIP: write source files via Python delete-and-rewrite; run `python tools/decrypt_via_copy.py --apply` before builds.

**Versioning:** Batch 1 (Tasks 1–5) → v2.4.1 internal. Batch 2 (Tasks 6–8) → v2.4.2 internal. Wrap → v2.5.0 deployable.

---

## Batch 1 — data integrity (v2.4.1)

### Task 1: Fresh-version OCC + ROLLBACK in the write cores (RC1 server side)

**Files:** Modify `src/database/DatabaseManager.cpp` (tryWriteSensoryCore ~1913, tryWriteDetailedSensoryCore ~2400s, tryWriteFile UPDATE branch); Test `tests/tst_databasemanager/`.

- [ ] **RED:** add tests:
  - `sensoryUpdate_staleInMemoryVersion_succeeds` — save a session (version=1), bump the row version twice out-of-band (simulating LiveSync per-cell commits: `UPDATE sensory_sessions SET version = version + 2 WHERE id=?`), then `tryWriteSensorySession` with the stale struct (version=1) and **changed non-score fields + changed score**. Expect `Success` (not VersionMismatch) and the changed fields present in the DB row.
  - `detailedUpdate_staleInMemoryVersion_succeeds` — twin.
  - `tpmUpdate_staleInMemoryVersion_succeeds` — same shape for tryWriteFile.
  - `sensoryInsert_uniqueViolation_leavesConnectionUsable` — force a 23505 INSERT, then immediately run an unrelated SELECT on the same DatabaseManager and expect it to succeed (regression for the aborted-transaction poisoning; June 10 log evidence).
- [ ] Run: `tests\run-tests.ps1 -Filter tst_databasemanager` → new tests FAIL.
- [ ] **GREEN:**
  - In each UPDATE branch: extend the merge-read SELECT to `SELECT json_data, version FROM ... WHERE id = ?` and bind the **freshly-read version** in `WHERE id = ? AND version = ?` (keep s.version only as fallback when the SELECT failed). True concurrent races still return VersionMismatch — that is now rare and retryable.
  - Trace the transaction wrapper (tryWriteSensorySession / tryWriteDetailedSensorySession / tryWriteFile): every error return path must `ROLLBACK` (or `db.rollback()`) so a failed INSERT can't leave the shared connection in "current transaction is aborted".
- [ ] Run tests → PASS. Full `tst_databasemanager` green. Commit `fix(db): fresh-version OCC in write cores + rollback on error paths (RC1/RC3)`.

### Task 2: Callers never skip-and-drop — retry on mismatch, re-INSERT on deleted (RC1 client side)

**Files:** Modify `src/MainWindow.cpp` onUpdateDatabase (~4567-4590 TPM, ~4653-4665 sensory, ~4729-4741 detailed) and the close-path twins `saveSensorySessionsBeforeClose` / `saveDetailedSensorySessionsBeforeClose` (~2403-2500); Test `tests/tst_databasemanager` (core-level) + manual log assertion.

- [ ] **RED:** core-level test `sensoryUpdate_rowDeleted_reinsertsFreshRow`: save, delete the row out-of-band, `tryWriteSensorySession` again → after the fix the WRAPPER (DatabaseManager level, see GREEN) re-INSERTs: expect Success and a new row containing the in-memory data.
- [ ] **GREEN:**
  - In `DatabaseManager::tryWriteSensorySession`/`tryWriteDetailedSensorySession`/`tryWriteFile` (wrapper level): on `RowDeleted` → reset id/version on a local copy and run the INSERT branch (data must never vanish because someone deleted the row). On `VersionMismatch` → retry the core once (fresh read happens inside); only after a second mismatch return it.
  - In MainWindow: delete the `VersionMismatch/RowDeleted ⇒ "skipped — already up to date via LiveSync"` blocks (all three). Treat any non-Success as `++failed` and KEEP the dirty flag so the next tick retries. Remove the misleading log lines.
- [ ] Run tests → PASS. Commit `fix(persist): retry-on-mismatch, reinsert-on-deleted; never skip-and-clear-dirty (RC1)`.

### Task 3: Dirty-cell tracking + dirty-aware merge (RC2)

**Files:** Modify `src/pipeline/SensoryData.h/.cpp` (merge signature), `src/pipeline/DetailedSensoryData.h/.cpp`, `src/ui/SensoryPanel.h/.cpp`, `src/ui/DetailedSensoryPanel.h/.cpp`, `src/database/DatabaseManager.cpp/.h` (plumb dirty set into tryWrite*), `src/database/LiveSync.h/.cpp` (flushNowAndWait returns bool); Tests `tests/tst_sensorydataplaceholder` or new `tests/tst_dirtymerge` + `tst_databasemanager`.

- [ ] **RED:** pure-function tests:
  - `merge_keepsDbForUntouchedCells` (existing behavior, regression-lock)
  - `merge_keepsMemoryForDirtyCells` — dirty set contains `samples[1].Smoothness`; DB has 5.0, memory has 7.5 → merged = 7.5; untouched metrics still take DB.
  - DB-level: `sensorySave_localEditWinsWhenLiveSyncNeverRan` — reproduce the user's revert: session in DB with score 5.0; in-memory edit to 7.5 marked dirty; NO per-cell commit; whole save; reload → 7.5. (This is the headline regression test.)
- [ ] **GREEN:**
  - `SensorySession`/`DetailedSensorySession` gain `QSet<QString> dirtyCells` (not serialized to JSON).
  - Merge signature: `mergeSensoryPreservingDbScores(mem, db, dirtyCells)` — skip DB-preserve for dirty paths (path format `samples[<idx>].<Metric>`; match the existing kSensoryMetrics naming).
  - Panels: in the same handler that calls `LiveSync::commitCell` (SensoryPanel.cpp ~875, detailed twin) — and crucially also when `activeSessionId() <= 0` — insert the cell path into the current session's dirtyCells. buildSession()/saveCurrentTester() carry the set. Clear the set only after a `Success` whole-session write (syncSavedSessionState) or a confirmed per-cell commit for that path (cellChanged echo with own UUID — if simple; otherwise clear only on whole-save Success, which is safe because keeping a cell dirty merely means memory wins for a cell the user actually edited this run).
  - `LiveSync::flushNowAndWait` returns `bool` (false on timeout); callers log a WARN with count of pending commits on timeout and proceed (now safe).
- [ ] Run tests → PASS. Commit `fix(sensory): dirty-aware merge — local edits authoritative for touched cells (RC2)`.

### Task 4: LiveSync worker reconnect + surfaced failures (RC3)

**Files:** Modify `src/database/LiveSyncWorker.cpp/.h`, `src/database/LiveSync.cpp/.h`, `src/MainWindow.cpp` (indicator hook); Test `tests/tst_livesync`.

- [ ] **RED:** `worker_reconnectsAfterConnectionLoss` — open LiveSync against the test container, kill the worker's backend (`SELECT pg_terminate_backend(pid)` from a second connection targeting the worker's application_name/pid), commit a cell, expect: commit eventually lands (reconnect+retry) OR commitFailed is emitted (assert signal). No silent nothing.
- [ ] **GREEN:**
  - In `commitScalar`/`commitJson`/`focusCell`/`blurCell`: on `q.exec()` failure where the error looks connection-shaped (nativeErrorCode 26000, 08xxx, empty + db not open, "server closed", "Unable to send query") → `stop(); start();` and retry the statement once. Log the reconnect at WARN.
  - `LiveSync` tracks a failed-commit counter (increment on commitFailed/commitConflict, reset on successful sync); expose `int unsyncedEditCount()` + signal; MainWindow's DB sync indicator shows "N edits not synced" state (red) when non-zero.
- [ ] Run tst_livesync → PASS. Commit `fix(livesync): reconnect-and-retry on broken worker connection; surface unsynced edits (RC3)`.

### Task 5: Kill rename loop; auto-suffix duplicates `_1/_2/_3` (RC4)

**Files:** Modify `src/MainWindow.cpp` (sensory rename/UniqueViolation block ~4612-4651, detailed ~4711-4727), `src/ui/SensoryPanel.cpp` (syncSavedSessionState baseline + title widget refresh), `src/ui/DetailedSensoryPanel.cpp` twin, `src/utils/OutputPaths.cpp/.h` (suffix helper); Tests new `tests/tst_namesuffix` (pure) + `tst_databasemanager` (collision path).

- [ ] **RED:**
  - Pure: `nextSuffixedName("T") == "T_1"`, `nextSuffixedName("T_1") == "T_2"`, `nextSuffixedName("T_9") == "T_10"`.
  - DB: `sensoryInsert_duplicateKey_autoSuffixes` — insert ("T","Charlie R1",date); save a second fresh session with the same key → expect Success, sessionName "T_1", both rows present. And `sensoryRename_collision_resolvesAndBaselineUpdated` — reproduce the June 10 loop: existing rows "New Session" and "T"; rename "New Session"→"T"; save → Success with "T_1"; save AGAIN with no edits → expect Success/no-op UPDATE, **no new row, no UniqueViolation** (the loop is dead).
- [ ] **GREEN:**
  - Helper `DVE::nextSuffixedName(const QString&)` in OutputPaths (strip trailing `_<digits>`, increment; first collision → `_1`).
  - In both UniqueViolation blocks: loop (≤25): suffix `sess.sessionName` AND `sess.testTitle`, retry `tryWrite*`; on Success log INFO "auto-renamed to ..." + show a non-modal status-bar message (NO QMessageBox); panel sync then refreshes title widget for the current session (so buildSession can't regenerate the unsuffixed name).
  - `syncSavedSessionState` already refreshes `originalSessionName` on success — verify it now runs for the suffixed result; detailed panel needs the same baseline handling it lacks (no originalSessionName there — confirm rename loop can't occur, or add equivalent).
  - 5s autosave path: rename detection allowed, but collisions resolve the same silent way.
- [ ] Run tests → PASS. Commit `fix(sensory): auto-suffix duplicate/renamed sessions; rename baseline always updated (RC4)`.

### Task 5b: v2.4.1 internal wrap
- [ ] `python tools/decrypt_via_copy.py --apply`; bump VERSION 2.4.0 → 2.4.1; clean release rebuild (repo-root `release\`, NOT build\release\); full test suite run; commit `chore(release): v2.4.1 internal — save/sync integrity batch 1`.

## Batch 2 — UX + identity (v2.4.2)

### Task 6: Close offers Name / Discard / Cancel (RC5)

**Files:** Modify `src/MainWindow.cpp` (`saveSensorySessionsBeforeClose` ~2403, `saveDetailedSensorySessionsBeforeClose` ~2450, promptSaveDatabase ~5184, onCloseFile TPM-failure prompt ~2251); uses existing `DatabaseManager::removeSensorySession/removeDetailedSensorySession`.

- [ ] For each non-savable session at close: QMessageBox::question with THREE buttons — `Save with name…` (focus title field, abort close), `Discard session` (destructive: if id>0 remove DB row; if m_savePath exists QFile::remove; drop from panel list; continue close), `Cancel` (abort close). Multi-session close aggregates one dialog per offending session.
- [ ] TPM persist-failure prompt gains `Retry` alongside Close-anyway/Cancel.
- [ ] Program-close (promptSaveDatabase) path must route through the same options (Discard there also deletes).
- [ ] Manual verification matrix in tests/tst_mainwindow_remotecell style is impractical — cover the discard helper (`discardSensorySession(idx)`) with a DB-level test: create session with DB row + temp file, discard, assert row gone + file gone + list shrunk.
- [ ] Commit `feat(close): Name/Discard/Cancel options — never hard-block (RC5)`.

### Task 7: TPM versioned re-adds (F6)

**Files:** Modify `src/database/DatabaseManager.cpp` (ensureSchema + tryWriteFile INSERT/match path), `deploy/postgres/migrations/2026-06-10-files-added-at-identity.sql` (canonical record), `src/MainWindow.cpp`/file-tree display.

- [ ] ensureSchema (DV-2 pattern, catalog-guarded): add `files.added_at TIMESTAMPTZ DEFAULT now()`; replace `UNIQUE(file_path)` with `UNIQUE(file_path, added_at)`-equivalent expression index; keep legacy rows valid.
- [ ] Load-from-disk flow: when file_path already has DB row(s), do NOT adopt the latest row's id by default — INSERT a new row stamped now(); display name shown as `file_name (yyyy-MM-dd HH:mm)` for stamped duplicates. In-session saves keep updating their own row by id (unchanged).
- [ ] DB browser/load lists every (file, added_at) version.
- [ ] Tests: `tpmReAdd_sameFilePath_createsNewVersionRow`; `tpmInSession_resave_updatesSameRow`.
- [ ] Commit `feat(db): versioned TPM file re-adds via added_at identity (F6)`.

### Task 8: E2E regression harness + v2.4.2 + v2.5.0 wrap (F7)

- [ ] New suite `tests/tst_saveintegrity_e2e` (SUBDIRS entry): scripted end-to-end scenarios chaining the real DatabaseManager + LiveSync against the container: (1) edit→save→reload→values match memory (the headline revert case); (2) broken-worker (pg_terminate_backend) mid-edit → save → reload → values still match; (3) rename→save→save→save → exactly 2 rows ever (original + renamed), zero UniqueViolations; (4) duplicate create ×3 → `T`, `T_1`, `T_2`; (5) discard removes all traces. Assert no occurrence of "skipped — already up to date" in captured log output.
- [ ] Full suite run; v2.4.2 bump + clean rebuild + installer; then wrap VERSION → 2.5.0, clean rebuild, `build_installer.bat`, update `tasks/v2.4.0-acceptance-checklist.md` → v2.5.0 checklist with the new scenarios.
- [ ] Commit `chore(release): v2.5.0 deployable — save/sync integrity + close options + versioned re-adds`.

---

## Self-review notes
- Task 1's fresh-version read makes Task 2's retry path rare but still required (true races, RowDeleted).
- Task 3 dirty-set clearing: conservative rule (clear only on whole-save Success) is correct because a dirty cell only forces memory-wins for cells the user edited in THIS run — by definition the freshest value.
- Task 5 MUST suffix testTitle as well as sessionName (buildSession regenerates sessionName from testTitle — dossier RC4).
- Detailed panel has no originalSessionName/rename branch — Task 5 verifies collision handling there uses the same suffix loop on plain INSERT collisions.
- All ensureSchema changes need the canonical migration file mirror (single-source-of-truth rule).
