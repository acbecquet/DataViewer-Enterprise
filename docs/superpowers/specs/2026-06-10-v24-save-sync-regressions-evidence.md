# v2.4.0 Save/Sync Regressions — Root-Cause Evidence Dossier

**Date:** 2026-06-10 · **Status:** Phase 1 (root-cause investigation) COMPLETE — all causes evidence-backed.
**Sources:** live log `%LOCALAPPDATA%\SDR\DataViewer Enterprise\dataviewer.log` (user's production sessions, Jun 8–10 = pure v2.4.0) + code reads + 4 parallel code-trace investigations.

**User-reported symptoms:**
1. Saves don't stick — "the next day I checked the file to see that it had reverted."
2. Ctrl+U in sensory mode frequently errors "this test name is already in use" — almost always wrongly.
3. Closing an improperly-named session hard-blocks ("add a name") with no discard option.
4. Wants: TPM files get a date+time suffix in the DB; sensory duplicates allowed via `_1/_2/_3` iterator; never block the user, always offer options.

---

## ROOT CAUSE 1 — Saves silently discarded by the "VersionMismatch ⇒ already up to date via LiveSync" heuristic (P0, data loss)

`MainWindow::onUpdateDatabase` treats `WriteResult::VersionMismatch`/`RowDeleted` as *"LiveSync already wrote this row per-cell, skipping is safe"* and **drops the dirty flag**:
- TPM: [MainWindow.cpp:4577-4586](../../src/MainWindow.cpp) — `m_modifiedFilePaths.remove(oldPath); ++filesSkipped;`
- Sensory: [MainWindow.cpp:4655-4662], Detailed: [MainWindow.cpp:4731-4738].

The assumption (v2.0.5 era) is FALSE whenever any edit did *not* flow through LiveSync. Then the skip = **silent discard of the user's in-memory edits**.

**Log evidence (v2.4.0 window, June 8–10): 23 skip events while the user was actively editing:**
```
2026-06-08T20:16:08.971 tryWriteFile(UPDATE files): version mismatch (id=174, expected version=2)
2026-06-08T20:16:08.983 [onUpdateDatabase] TPM file S25B vs S25B1 Curaleaf Curve Test ... skipped — already up to date via LiveSync (result= 1 )
  (11× for this TPM file; 12× sensory: 5× "Tallboy MPX Rosin All Voltages", 4× "D1520 CloudEx 3 Heaters Sensory 2.8V", 3× "Sunday x AirOne")
```
The version-mismatch arises from the app's own bookkeeping (whole-save bumps DB version; in-memory copy keeps the stale version; per-cell commits also bump version), so the skip fires routinely — it is NOT evidence the data landed.

## ROOT CAUSE 2 — DB-authoritative merge has zero freshness arbitration; discards any edit that didn't reach LiveSync (P0, data loss)

`mergeSensoryPreservingDbScores` / detailed variant ([SensoryData.cpp:97-114](../../src/pipeline/SensoryData.cpp)) **unconditionally** copy DB score values over in-memory values for index-matched samples — no updated_at/version/dirty comparison. `tryWriteSensoryCore` UPDATE branch applies it on every whole-session save; exports route through `dbAuthoritativeSessions` (same overlay) so **the .xlsx file also gets the stale values** → "file reverted".

**Edits that never reach LiveSync (traced):**
- Sessions with `id <= 0`: [SensoryPanel.cpp:878](../../src/ui/SensoryPanel.cpp) gates `commitCell` on `activeSessionId() > 0`. Fresh imports / not-yet-persisted sessions stream nothing.
- Broken worker connection (see RC3) — commits fail or are never dispatched.
- `flushNowAndWait()` ([LiveSync.cpp:264-300](../../src/database/LiveSync.cpp)) returns void; on its 4 s timeout **all 6 call sites proceed to the merge anyway** (MainWindow.cpp:2247, 2407, 2454, 4545; SensoryPanel.cpp:2122; DetailedSensoryPanel.cpp:1521).
- `commitFailed`/`commitConflict` signals have **no UI handler** — failures vanish (MainWindow.cpp:235 comment).
- Offline with no snapshot: `commitCell` returns false silently (LiveSync.cpp:153-158).

## ROOT CAUSE 3 — LiveSync worker connection breaks and never recovers (P0 enabler)

```
2026-06-08T20:16:00.407 [WARN] LiveSyncWorker::blurCell failed: "ERROR: unnamed prepared statement does not exist (26000) QPSQL: Unable to send query"
  (10× June 8 20:16–20:18; 126× total since May — known-bad state, persists for minutes until restart)
```
`LiveSyncWorker` ([LiveSyncWorker.cpp](../../src/database/LiveSyncWorker.cpp)) has no reconnect/retry on statement failure; a poisoned connection keeps failing. Zero `commitScalar/commitJson failed` lines in the log + RC1/RC2 depending on LiveSync health = the system has no defense when this happens.

Related connection poisoning: after a 23505 INSERT failure the shared connection is left with an aborted transaction:
```
2026-06-10T09:46:49.759 [WARN] Unable to free statement: ERROR: current transaction is aborted, commands ignored until end of transaction block
```
→ `tryWriteSensoryCore` error path must ROLLBACK.

## ROOT CAUSE 4 — Rename→force-INSERT loop produces the false "name already in use" (P0, blocks saves)

Rename detection ([MainWindow.cpp:4621-4631](../../src/MainWindow.cpp)): `id > 0 && originalSessionName non-empty && != sessionName` → forces `id=-1` → INSERT (old row intentionally preserved). `sessionName` is recomposed from the Test Title verbatim on every widget flush ([SensoryPanel.cpp:935-941](../../src/ui/SensoryPanel.cpp)).

**Failure loop:** when the rename-INSERT hits 23505 (`UniqueViolation` → `++cancelled; continue`), `syncSavedSessionState` ([SensoryPanel.cpp:1088-1094]) only refreshes the baseline `originalSessionName` **on success** — so the panel keeps the stale baseline and **every subsequent save (incl. the silent 5 s autosave, which runs the same rename branch) re-detects the same rename and re-collides, forever**. The 5 s autosave can also INSERT the rename first, making the user's own Ctrl+U collide moments later.

**Log evidence (June 10):**
```
09:45:58.565 sensory rename detected: New Session → D1520 CloudEx 3 Heaters Sensory 2.8V — routing to INSERT
09:45:58.964 sensory rename detected: New Session → D1520 CloudEx 3 Heaters Sensory 2.8V — routing to INSERT   ← same rename, 400 ms later
09:46:49 / 09:46:53 / 09:46:55 / 09:47:08 / 09:47:11 / 10:08:58  INSERT sensory_sessions: duplicate key (idx_sensory_sessions_key)  ← the loop
```
Natural key = `UNIQUE(session_name, tester_name, date)` (init.sql idx_sensory_sessions_key; same for detailed). Historical "New Session" rows in the DB (May logs show many) seed these collisions; never-delete policy means duplicates breed.

## ROOT CAUSE 5 — Close flow hard-blocks unnamed sessions (UX, by design in v2.4.0 — design wrong)

`saveSensorySessionsBeforeClose` ([MainWindow.cpp:2403-2444]) + detailed twin: non-savable session → warning "Name Required to Close" (OK only) → session forced to stay open. No discard path. Deletion APIs already exist: `DatabaseManager::removeSensorySession(id)` / `removeDetailedSensorySession(id)` (DatabaseManager.h:167/181); autosaved file path is tracked (`m_savePath`).

## NOT broken (verified)

- `ensureSchema` self-heal worked on the live DB: June 6 added `data_rows.puffing_regime`; June 8 19:46 healed `sensory_header_presets` (test_name + index swap + backfill). The 42703 storm ended June 6.
- No 42883/42P10 in the v2.4.0 window; DV-2 preset scoping shows no errors.

## Identity map (for the duplicate-friendly features)

- TPM files: `UNIQUE(files.file_path)` (init.sql idx_files_path); same path re-saved = UPDATE in place. Feature: re-adding a file later must create a NEW row → identity needs an added-at stamp (additive column + index swap, ensureSchema-healable, mirrors DV-2 pattern); display name gets " (yyyy-MM-dd HH:mm)".
- Sensory/detailed: `UNIQUE(session_name, tester_name, date)`. Feature: on real duplicate, auto-suffix session_name `_1, _2, _3` (bounded retry), no dialog; suffix must propagate to in-memory struct, UI label, auto-save filename.

## Fix architecture (v2.4.1+ internal patches → v2.5.0 deployable)

Principle (user's words): *in-memory edits of the active client are authoritative for the cells THAT CLIENT touched; never silently drop a save; never hard-block — always offer options.*

1. **F1 — Never discard a save (RC1):** remove the VersionMismatch⇒skip heuristic. On VersionMismatch: re-read row (version + JSON), re-merge dirty-aware, retry UPDATE (bounded). On RowDeleted: INSERT fresh. Dirty flags cleared only after a confirmed write.
2. **F2 — Dirty-aware merge (RC2):** panels track locally-edited cell paths per session (set on edit, cleared on confirmed commit/whole-save). Merge keeps DB values ONLY for non-dirty cells. flushNowAndWait returns bool; timeout → log + proceed (now safe because dirty cells survive the merge).
3. **F3 — LiveSync resilience (RC3):** worker reconnect-and-retry-once on statement failure; ROLLBACK on failed transactions in tryWrite* error paths; commitFailed/commitConflict surfaced in the sync indicator ("N edits not synced").
4. **F4 — Kill the rename loop + duplicate-friendly naming (RC4):** on INSERT 23505 auto-suffix `_1.._N` and retry (sensory + detailed); update rename baseline after ANY resolution; no modal error. Within-pass dedup of identical keys.
5. **F5 — Close offers options (RC5):** non-savable session on close → [Name it / Discard session & delete data / Cancel]; discard = remove DB rows (if id>0) + autosaved .xlsx + list entry. Same options at program close. TPM close keeps persist-first but offers [Retry / Close anyway / Cancel] on failure.
6. **F6 — TPM versioned re-adds:** files identity = (file_path, added_at stamp); fresh add of an existing path = new row with timestamp-suffixed display name; in-session saves keep updating their own row by id.
7. **F7 — Regression harness:** new E2E test suite against the ephemeral Postgres container reproducing each RC (failing first, TDD): mismatch-retry, dirty-merge, broken-worker commit loss, rename-collision loop, duplicate auto-suffix, close-discard. Plus log-pattern assertions (no silent skip lines).

Versioning: fixes land as v2.4.1, v2.4.2… internal; wrap into **v2.5.0** deployable.
