# Bug-fix batch -- DATAVIEWER-3 / -4 / -2 -- Design

**Date:** 2026-06-05
**Status:** Approved (design); implementation pending. Per-issue plans written as each is started.
**Plane issues:** DATAVIEWER-3 (TPM upload), DATAVIEWER-4 (sensory reset-to-5), DATAVIEWER-2 (sample-name dropdown)
**Order of work:** DV-3 (High / Blocking) -> DV-4 (High / data-loss) -> DV-2 (Medium)

## 1. Overview

Three triaged Plane bugs, fixed in priority order. This spec lays out all three at the approach
level; each gets its own detailed implementation plan (in `docs/superpowers/plans/`, written
immediately before its implementation) and ships through the project fix-bug cadence:
plan -> implement -> test -> build -> set the Plane issue to *Ready for Release*. No deployment
and no Synology drop -- that remains the manual human checkpoint.

## 2. Shared principles

- **Data integrity first.** No fix may drop, truncate, or silently default existing data. Every
  schema migration is idempotent and preserves existing rows; every change to how a JSONB blob
  is written must preserve all keys it does not explicitly set. See section 6.
- **Surface failures, never swallow them.** Silent `false` / no-op returns are the through-line
  cause of DV-3 and a contributor to DV-4.
- **Verification constraints.** This (work) machine runs the Qt suite against an ephemeral
  `postgres:16` (`tests\start-test-postgres.ps1`) but cannot reproduce true multi-user
  concurrency, offline/Synology behavior, or the deployed binary. Those are covered by automated
  simulations here plus the work-machine deployment self-test. C++ builds require
  `python tools/decrypt_via_copy.py --apply` first (MIP); new source files use the Python
  delete-and-rewrite convention.

## 3. DATAVIEWER-3 -- TPM data not uploading   (approach: Harden + surface)

**Root cause.** `MainWindow::onFileLoadFinished()` (src/MainWindow.cpp ~2300-2392) calls
`m_db->saveFile()` at lines 2361 and 2385 but discards the `WriteResult`, and clears the dirty
flag unconditionally right after (2362 / 2386). `saveFile()` (DatabaseManager.cpp:796) is a bool
shim over `tryWriteFile()` (190-801) that collapses `OfflineReadOnly` / `VersionMismatch` /
`RowDeleted` / SQL error to `false`. The Ctrl+U batch (`onUpdateDatabase` ~4354-4392) only
re-writes files in `m_modifiedFilePaths`, so a load-save that failed is never retried. Net: a
failed first save is both invisible and permanent.

**Fix.**
1. In `onFileLoadFinished()`, replace the two `saveFile()` calls with `tryWriteFile()`,
   capturing the `WriteResult` (mirrors the batch path at 4375).
2. Clear `m_modifiedFilePaths` **only** on `WriteResult::Success`. On any other result, keep /
   insert the file dirty so the close-flush and Ctrl+U batch retry it.
3. Map results to visible state via the DB status indicator (`updateDbSyncIndicator` ~4865-4925)
   plus a `qWarning` line carrying the `WriteResult` + `m_lastError`:
   - `Success` -> normal saved state.
   - `OfflineReadOnly` -> "Offline -- N file(s) not saved"; retried on reconnect.
   - `VersionMismatch` / `RowDeleted` -> re-inherit id/version via `loadFileByPath`, retry once;
     if still failing, mark dirty + log (no retry loop).
   - other (SQL/exception) -> show `m_lastError` snippet in status, mark dirty.
4. One-line log of the result on every load-time save so the "nothing since 6-1" cause is
   visible going forward.

**Files:** src/MainWindow.cpp only. No schema change. `tryWriteFile` is transactional
(BEGIN..COMMIT) so a failure leaves no half-written row.

**Verification.** Qt test vs ephemeral Postgres: (a) healthy save persists + clears dirty;
(b) simulated failure (closed DB / offline) keeps the file dirty and the status reflects
"not saved"; (c) no exception / half-write. The actual production root cause is then confirmed
on the work machine using the new visible status + log.

**Risk:** low; localized. Subtlety: the OCC retry must not loop.
**Out of scope:** redesigning the batch-save model (that is the DV-4 follow-up).

## 4. DATAVIEWER-4 -- Sensory data resets to 5   (approach: Phased)

**Root cause (pinned).** Scores persist in `sensory_sessions.json_data` (JSONB; init.sql:127).
LiveSync per-cell commits update sub-paths via `jsonb_set('{samples,i,scores,metric}', ...)` on
a background worker (LiveSync.h:25-38). The whole-session save `tryWriteSensorySession`
(DatabaseManager.cpp ~1571-1820) issues `UPDATE sensory_sessions SET ... json_data = CAST(? AS
JSONB)` (~1601-1609) -- a **wholesale replace** of the entire blob built from the in-memory
session. When that copy lacks a metric, serialization falls back to `5.0` (SensoryData.cpp:90),
stamping a 5 over LiveSync's per-cell value. Triggered by Ctrl+U (`onUpdateDatabase`:4424),
on-close flush, and Excel-import. OCC skips on version mismatch, but the clobber still occurs
when the saving client holds the current version yet has a stale/partial in-memory blob.

**Phased fix** (keep Ctrl+U as a safe manual flush; full removal deferred to section 8):
1. Make the whole-session write **non-destructive to LiveSync-owned score paths** -- it must not
   wholesale-replace `json_data` and lose keys it does not hold. The plan chooses between:
   - (a) **Read-merge-write:** SELECT current `json_data`, deep-merge the in-memory session over
     it (preserving any key the in-memory copy lacks), write the merged blob under the existing
     OCC version guard. Simpler, fully general. Preferred unless it regresses the rename/INSERT
     branches.
   - (b) **Path-scoped write:** the whole-session UPDATE writes metadata columns + `layout_json`
     + only non-LiveSync json paths, leaving `samples[].scores.*` (everything `isLiveSyncColumn`
     owns) to LiveSync. More surgical; must track the allowlist precisely.
2. Un-entered metrics must not be serialized as a destructive real 5: only user-present score
   keys participate in the write; unset keys never overwrite a DB value.
3. Apply symmetrically to detailed-sensory (`tryWriteDetailedSensorySession` ~2080-2264) -- same
   wholesale-replace shape.
4. Regression test: per-cell `jsonb_set` edit, then a whole-session save from a session whose
   in-memory scores map is missing that metric; assert the metric retains the edited value (not
   5). Cover sensory + detailed; assert all other keys preserved (integrity).

**Files:** src/database/DatabaseManager.cpp (`tryWriteSensorySession` +
`tryWriteDetailedSensorySession`; const/by-ref overloads share a body); possibly the `json_data`
serializer in pipeline/SensoryData.cpp + DetailedSensoryData.cpp. No schema change. Ctrl+U stays
wired (MainWindow.cpp:1406-1415).

**Verification.** New Qt test vs ephemeral Postgres reproducing the two-writer clobber, with
integrity assertions.
**Risk:** medium -- touches the central session-write path; the OCC / rename / INSERT branches
(onUpdateDatabase:4403-4471) must keep working. Mitigated by the regression test + the existing
DB suite.
**Out of scope (-> section 8):** removing Ctrl+U, making LiveSync authoritative for every field,
eliminating the batch save.

## 5. DATAVIEWER-2 -- Sample-name dropdown   (approach: scope + bound)

**Root cause.** `sensory_header_presets` is keyed `(kind, value)`; sample_name rows carry no
test association (`saveSensoryHeaderPresets` DatabaseManager.cpp:2638-2681 inserts `(kind, value)`
`ON CONFLICT (kind, value) DO NOTHING`). `loadSensoryHeaderPresets("sample_name")` (2683-2704)
returns the entire global pool; the SampleCard dropdown (SensoryPanel.cpp:890,
DetailedSensoryPanel.cpp:361) renders all of it, unbounded, covering the screen.

**Fix.**
1. **Scope to current test:**
   - Schema (MigrationTool, idempotent, schema_version bump): `ADD COLUMN IF NOT EXISTS
     test_name TEXT`; replace the `(kind, value)` unique index with a unique index on
     `(kind, value, COALESCE(test_name,''))` so test_name/media rows (NULL) stay unique by
     `(kind, value)` and sample_name rows are unique per `(value, test)`.
   - `saveSensoryHeaderPresets`: store `test_name` on sample_name rows (the function already
     receives `testName`); update the `ON CONFLICT` target to match the new index.
   - New `loadSampleNamesForTest(testName)`: `SELECT value WHERE kind='sample_name' AND
     test_name = ?`.
   - UI: the dropdown callbacks pass the current Test Title.
2. **Bound the popup:** cap the SampleCard name dropdown to ~10-12 visible rows + scroll (max
   popup height) so it can never cover the screen -- a safety net independent of scoping.

**Files:** DatabaseManager.cpp/.h, MigrationTool.cpp, OfflineSnapshot (schema mirror -- section
6), SensoryPanel.cpp + DetailedSensoryPanel.cpp + the SampleCard popup widget.
**Verification.** DB test: names saved under test A are returned for A, not for B; popup height
capped.
**Risk:** low-medium (schema migration) -- governed by section 6.
**Out of scope:** a cross-test "all names" view (legacy NULL rows stay queryable but unshown).

## 6. Data integrity, migrations & backfill   (cross-cutting -- per Charlie's request)

- **Migration idempotency.** `ADD COLUMN IF NOT EXISTS`, `CREATE UNIQUE INDEX IF NOT EXISTS`,
  gated by a schema_version bump so re-running is a no-op. Never DROP data.
- **Unique-index + ON CONFLICT change together.** Changing the unique constraint without
  updating `saveSensoryHeaderPresets`'s `ON CONFLICT` target would make inserts throw. Both
  change in the same step; a test asserts repeated saves do not error.
- **Backfill test associations from sessions.** Existing sample_name rows have NULL test_name
  and cannot be perfectly attributed. Reconstruct real (test -> sample-name) pairs from
  `sensory_sessions`: for each session, `session_name` is the test and `json_data` lists sample
  names; `INSERT (kind='sample_name', value=<name>, test_name=<session_name>) ON CONFLICT DO
  NOTHING`. Makes the scoped dropdown immediately useful from historical data.
- **Preserve legacy rows (no data loss).** Pre-existing NULL-test sample_name rows are kept, not
  deleted. They no longer match a specific-test query (so they stop flooding the dropdown) but
  remain in the table. Acceptance: the count of distinct historical sample names never decreases.
- **Offline snapshot parity.** `OfflineSnapshot` mirrors these tables and regenerates on clean
  close. The plan verifies the snapshot CREATE TABLE / copy logic includes `test_name`, and that
  DV-4's json_data merge does not corrupt the SQLite mirror.
- **DV-4 blob preservation.** The whole-session merge must preserve every `json_data` key it does
  not explicitly set (samples, comments, structure, layout). Regression test asserts an unrelated
  sample / metric / comment survives a whole-session save unchanged.
- **Transactional safety.** `tryWriteFile` / `tryWriteSensorySession` run inside BEGIN..COMMIT;
  a failed write rolls back leaving no half-row. DV-3's change preserves this property.

## 7. Deliverables, versioning & sequencing

**Versioning policy** (Charlie's standing scheme; semantic `x.y.z`):
- **Patch (`z`)** -- internal build, NOT deployed.
- **Minor (`y`)** -- deployable release.
- **Major (`x`)** -- fundamental changes.

Current shipped version is v2.3.1. Each fix below is its own internal (patch) build; when all
three are approved they wrap into a single deployable minor release **v2.4.0**.

1. This spec (committed on branch `feature/v2.4.0-bugfix-batch`).
2. DV-3 -> **v2.3.2** (internal): plan -> implement -> test -> build; set *Ready for Release*.
3. DV-4 -> **v2.3.3** (internal): plan -> implement -> test -> build; set *Ready for Release*.
4. DV-2 -> **v2.3.4** (internal): plan -> implement -> test -> build; set *Ready for Release*.
5. Wrap all three into **v2.4.0** (deployable) once approved; file the deferred follow-up (section 8).

Each internal build bumps `VERSION` in the .pro and does a clean rebuild (`mingw32-make clean`
then rebuild -- qmake does not detect VERSION changes). Each issue's plan is written immediately
before its implementation. No deploy / no Synology drop at any step; the v2.4.0 wrap is handed to
Charlie. The semver policy above is also recorded in CLAUDE.md (added during the DV-3 / v2.3.2
build).

## 8. Deferred follow-up

New Plane issue, filed at the end: **"Remove Ctrl+U; make LiveSync the sole authoritative
sensory/detailed-sensory persistence path; eliminate the whole-session batch overwrite; keep the
UI thread off Postgres."** Captures DV-4's full vision beyond the Phased data-loss fix.
