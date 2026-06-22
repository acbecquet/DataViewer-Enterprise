# SP4.5 Stage 2b — Incremental Offline Snapshot + Save Progress UX (Design)

**Date:** 2026-06-22
**Branch:** `feature/v2.4.0-bugfix-batch`
**Ships in:** deployable **v2.5.0** (internal build **v2.4.7** for owner smoke-test first)
**Status:** design approved by owner 2026-06-22; this doc is the spec the implementation plan is written from.

---

## Goal

Eliminate the ~9s close stall that remains after Stage 2a, and make every unavoidable
save wait *visible* rather than a frozen window. Two prongs:

1. **Incremental snapshot regen** — stop re-copying the offline snapshot's unchanged
   image blobs (95% of the file) on every close. Reuse them; refresh only what changed.
2. **Non-frozen "Saving…" progress dialog** — whenever close (or an explicit Ctrl+U)
   must do real work, show a progress bar driven from the worker, so the app reads as
   "working" instead of "hung." This also covers the chatty DB save we are deliberately
   *not* rewriting in this pass.

## Root cause (proven, not hypothesized)

Evidence gathered 2026-06-22 from the owner's own machine (no production-DB queries):

- The regen rebuilds `%LOCALAPPDATA%\SDR\DataViewer Enterprise\snapshot.sqlite`, which is
  **85.9 MB**. Breakdown by table (read directly from the live file):

  | table | rows | blob bytes |
  |---|---:|---:|
  | **images** | **14** | **85.4 MB** (~6 MB each) |
  | data_rows | 30,921 | — |
  | samples | 1,135 | — |
  | tests | 271 | — |
  | sensory_sessions | 205 | — |
  | files | 90 | — |
  | sensory_images | 1 | 0.06 MB |
  | everything else | — | ~0 |

  **81.5 MB / 95% of the snapshot is 14 image blobs.** All editable TPM data is ~4.5 MB.

- The user's `dataviewer.log` `[perf]` markers show the **entire** close cost is the regen
  step: e.g. `regenerating offline snapshot (DB changed)` 09:24:23.742 →
  `snapshot regen done` 09:24:32.623 = **8.9s**; every other close step (excel flush,
  persist drain, save prompt, image-cache wipe, settings, recovery) is <30 ms.
- The regen takes 7–12s **whether it runs in the foreground or the Stage-2a background
  worker** — it is the bulk copy itself, not the UI thread. It is *not* network-bound:
  81 MB over the >1 Gb/s link to the local NAS is ~0.6s ideal. The cost is four passes
  over 81 MB: PG `BYTEA` read-marshalling → row-by-row SQLite `INSERT` → `wal_checkpoint`
  rewrite → `MoveFileExW` write-through.
- **Secondary bug:** the background regen completes (09:24:16) yet the close 7s later
  re-runs the full regen (`DB changed`, 09:24:23). Not a fingerprint *computation* bug
  (both sides use the identical `snapshotContentFingerprint` SQL against PG) — it is
  correct-but-expensive staleness detection losing a race against a write that lands
  during/after the slow 8s regen. Making the regen cheap dissolves it.
- **MIP is a non-issue here:** the at-rest `snapshot.sqlite` first bytes are
  `SQLite format 3\0` (verified with a non-allowlisted reader), **not** `%TSD-Header-###%`.
  So reusing the old file's blobs is a plain local operation — no python-decrypt path.

## Design

### A. Incremental regen — copy, patch, promote

`OfflineSnapshot::regenToPath()` keeps its exact crash-safe **atomic promotion**
(`MoveFileExW(... MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)`); only how the
`.tmp` is built changes:

- **Today:** create `.tmp` from an empty schema, then re-pull all ~10 tables from PG
  (including 81.5 MB of blobs) and row-insert them.
- **New (incremental path):** when a *valid prior* `snapshot.sqlite` exists (file present,
  openable, `source_schema_version` == `kSnapshotSchemaVersion`):
  1. **Copy** `snapshot.sqlite` → `.tmp` via `QFile::copy` (a ~0.3s local copy now that
     the file is plaintext). The 81.5 MB of blobs ride along — never re-read from PG.
  2. Open `.tmp`, `PRAGMA journal_mode=WAL` / `synchronous=FULL`, **one transaction**:
     - Open the same `BEGIN ISOLATION LEVEL REPEATABLE READ READ ONLY` PG txn used today
       (consistency for the fingerprint + copies).
     - For each **small** table whose fingerprint segment changed
       (files, tests, samples, data_rows, sensory_sessions, detailed_sensory_sessions,
       settings): `DELETE FROM <t>` then re-`INSERT` from PG (the existing per-table
       SELECT→prepare→bind→exec loops, reused verbatim). The schema has **no FK
       constraints** and `foreign_keys` is off, so delete/reinsert ordering is irrelevant.
     - For each **image** table (images, sensory_images, detailed_sensory_images) whose
       fingerprint segment changed: **diff by `(id, updated_at)`** against the rows already
       in `.tmp` — `INSERT` new ids, `UPDATE` (or delete+insert) changed ids, `DELETE`
       removed ids. When the segment is unchanged, skip the table entirely (the copied
       blobs are already correct and byte-identical).
     - Refresh `_snapshot_meta` (snapshot_taken_at from PG server clock,
       source_schema_version, content_fingerprint) exactly as today.
     - `COMMIT`, `wal_checkpoint(TRUNCATE)`, close.
  3. Promote `.tmp` → prod with the unchanged `MoveFileExW` path.
- **Full-rebuild fallback (unchanged code path):** when there is no prior snapshot, it
  cannot be opened, or the schema version differs, build from empty as today. Guarantees a
  first run / schema migration / corrupted-snapshot recovery still works.

**Why this approach** (alternatives considered):
- *In-place patch of the live prod file* — technically atomic via the WAL txn, but it
  mutates the single crash-recovery file directly. Copy+promote keeps the R7 guarantee
  that prod is only ever **replaced atomically**, never written in place.
- *Two-file data/blob split* — architecturally clean for "don't rewrite blobs," but it
  rewrites the offline **read** path (ATTACH + cross-file joins) and adds a two-file
  consistency problem. Unnecessary now that MIP is a non-issue. YAGNI.

### B. Change detection (mostly already built)

`snapshotContentFingerprint()` already emits a `count/max(updated_at)::bigint` segment per
table, `;`-joined in a fixed order. Stage 2b parses the **stored** fingerprint (from the
prior snapshot's `_snapshot_meta`) and the **live** fingerprint into per-table segments and
compares them:
- all segments equal → `isCurrentVsLive()` true → skip the whole regen (today's instant
  close, unchanged);
- a segment differs → refresh only that table (small tables: full reload; image tables:
  row diff).

A `regenToPath` caller may pass the prior fingerprint so the function knows which segments
changed without recomputing. (If parsing the prior fingerprint fails for any reason, fall
back to refreshing all tables on the copied file — still correct, just less optimal.)

### C. Progress UX — the "it's saving" signal

A single reusable, **non-frozen** progress dialog (`QProgressDialog`-based; the codebase
already uses `QProgressDialog`/`QProgressBar` in MainWindow — match that style) shown
whenever close or an explicit Ctrl+U actually has work to do:

- The heavy work runs on the worker thread (`SnapshotRegenWorker` for the regen; the
  existing save path for the DB save). The worker emits **progress signals**; the UI thread
  updates the bar and stays responsive (animating, not frozen).
- **Determinate where we can quantify**, indeterminate otherwise:
  - DB save: per-file / per-section progress ("Saving file 3 of 5…").
  - Snapshot regen: coarse phases — copy base file, refresh data tables, copy changed
    image blobs (per-blob increment, the only heavy part), checkpoint+promote.
- New worker signal: `SnapshotRegenWorker::regenProgress(int done, int total, QString phase)`
  (and a matching `OfflineSnapshot::regenToPath` progress callback,
  `std::function<void(int,int,const QString&)>`, null-safe — tests pass `nullptr`).
- The dialog is **not cancellable** during the snapshot promotion (a half-written promotion
  must never be user-abortable); a cancel button, if shown, only sets the existing atomic
  `cancel` flag *before* promotion (same cooperative cancel Stage 2a already polls).

### D. Close / background flow

- **During the session:** the debounced background incremental regen
  (`m_snapshotRegenTimer` → `SnapshotRegenWorker::requestRegen`) keeps the snapshot fresh.
  It is now cheap, so it converges quickly and the close usually finds the snapshot current.
- **At close** (`MainWindow::closeEvent`, order preserved): `finishExcelWritesBlocking` →
  `drainPersistWorkerBlocking` (existing data-loss guard) → save prompt. Then:
  - If `isCurrentVsLive()` → skip regen, instant close (today's behavior).
  - Else → show the "Saving…" dialog and run the **incremental** regen on the worker while
    the bar animates; close when done. Common case (data-only) is sub-second; the rare
    image-change case copies only the changed blobs and is shown, never frozen
    (owner's "block briefly to stay current, with a progress bar" choice).
  - The Stage-2a `regenWasInFlight` guard stays: if a background regen is mid-flight at
    close, cancel it and run the close-time regen instead of racing two writers on `.tmp`.

### E. Scope boundary — DB save path NOT rewritten

The second root cause (the per-row N+1 NAS round-trips in `DatabaseManager::tryWriteFile`)
is the OCC save path with the v2.4.0 data-loss history. We **do not** rewrite it in this
pass. Instead the progress dialog (C) makes its residual cost acceptable — exactly the
owner's "a progress bar mitigates what we can't 100% resolve." Batch-write optimization is
explicitly deferred (future, only if the shown save is still judged too slow).

## Safety

- **Crash-safety:** prod `snapshot.sqlite` is only ever replaced by the atomic
  `MoveFileExW`. The copy+patch happens entirely on `.tmp`; a crash mid-patch leaves prod
  untouched. Full-rebuild fallback covers missing/unopenable/old-schema snapshots.
- **No data-loss regression:** close still drains the background persist worker and flushes
  writeback before the save prompt (Stage 2a guard, unchanged). The save path itself is
  untouched.
- **Thread-safety:** `SnapshotRegenWorker` keeps its **own** `PostgresConnection` and the
  regen opens its **own** uniquely-named SQLite connection; QSqlDatabase is never shared
  across threads. Teardown via `cancel()` + `quit()`/`wait()` (never a
  `BlockingQueuedConnection` on the blocking slot).
- **Correctness:** the incremental result must be byte-for-byte equivalent in *content* to a
  full rebuild for the same DB state (verified by a test that diffs both).

## Testing

Unit tests (extend `tst_offlinesnapshot`, needs `dve-test-pg` with migrations applied):
1. **Unchanged DB** → `isCurrentVsLive` true → regen skipped (no `.tmp` written).
2. **Data-only change** (edit a data_row) → incremental regen: data_rows refreshed; the
   `images` rows are **byte-identical** to before and the image tables were **not re-read
   from PG** (assert the image-segment-unchanged branch was taken — e.g. a test-visible
   counter of image rows pulled from PG is 0, or the `images` table's blobs match a hash
   captured before the regen).
3. **Image add / change / delete** → image diff produces the correct row set (count + blob
   bytes match a full rebuild).
4. **Schema-version mismatch / missing prior snapshot** → full-rebuild fallback runs and
   produces a complete snapshot.
5. **Incremental == full**: for a given DB state, an incremental regen over a stale snapshot
   yields the same table contents (all rows, all blobs) as a from-empty full rebuild.
6. **Rollback leaves prod intact**: a forced failure mid-patch (e.g. bad SQL) leaves the
   previous `snapshot.sqlite` unchanged and removes the `.tmp`.
7. **Progress callback**: `regenToPath` invokes the progress callback with non-decreasing
   `done` ≤ `total` and a final `done == total`; passing `nullptr` is safe.

Manual / real-DB verification (owner, v2.4.7 internal build, `[perf]` markers):
- Edit a cell → Update DB → wait 30s → close: close is sub-second (or shows a brief,
  smoothly-animating "Saving…" bar), not a 9s freeze.
- Add an image → close: "Saving…" bar animates while the new blob copies; close completes.
- `dataviewer.log` shows the incremental path (`regen: copied base + refreshed N tables, 0
  image bytes from PG` style markers) on data-only closes.

## Out of scope / deferred
- Rewriting `tryWriteFile` into batched multi-row statements (the N+1 save). Mitigated by C.
- Any change to the offline **read** path / snapshot schema.
- The two pre-existing test failures (tst_excelexporter Excel round-trip; tst_storedfns
  container schema — needs `deploy/postgres/migrations/` re-applied).

## Risks
- **R1 — verbatim reuse of the per-table copy loops.** The incremental path reuses the
  existing SELECT→bind→exec loops; preserve their column arity asserts. Don't retype bind
  indices.
- **R2 — fingerprint segment parsing.** Must split on the exact `;` order the SQL emits; a
  drift silently refreshes everything (safe but slow) — add an assert that the parsed
  segment count matches the table list.
- **R3 — copy of an open file.** The UI holds a read handle on prod `snapshot.sqlite`; the
  Stage-2a pattern (`m_snapshot->close()` before regen, `openReadOnly()` after) already
  releases it. The worker copies prod→tmp; ensure the copy source isn't the UI's open
  handle path on Windows (close-first already handles this).
- **R4 — progress dialog re-entrancy at close.** Use the existing
  `ExcludeUserInputEvents` discipline so a second close-click can't re-enter `closeEvent`
  while the bar is up.
