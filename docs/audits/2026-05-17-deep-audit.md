# v2.0.2 deep-audit — impact-and-fix table

This document closes out the deep audit of the v2.0.1 save/database/sync
chain, performed 2026-05-17. The audit surfaced 25 findings spanning
correctness, hardening, and polish; this release closes all 25 across
four phased commits + Phase 3/4 fixes.

The findings are tagged by severity: **C** = critical (data-loss / OCC
break), **H** = hardening (UX correctness, race), **M** = polish (perf,
quality-of-life, docs).

## Findings by closing commit

| Phase | Finding | Severity | Title | Closing SHA |
|---|---|---|---|---|
| 1 | C1 | critical | OCC absent on per-cell writes (stored fn signature) | 6c00523 |
| 1 | C2 | critical | LiveSyncWorker discards stored fn BOOLEAN return | 6c00523 |
| 1 | M4 | polish | commitConflict signal missing from LiveSync | 6c00523 |
| 1 | M5 | polish | No version anchor on LiveSync API (setVersionLookup) | 6c00523 |
| 1 | M6 | polish | Migration is backwards-compatible (DEFAULT NULL) | 6c00523 |
| 1 | M7 | polish | qHash chain across PendingKey fields | 6c00523 |
| 1 | H4 | hardening | OCC opt-out path for legacy callers | 6c00523 |
| 2 | C3 (TPM) | critical | tryWriteFile DELETE-cascade-rebuild wipes concurrent edits | d9092c8 |
| 2 | C3 (sensory) | critical | tryWriteSensorySession + tryWriteDetailedSensorySession same bug | 9e4a9a7 |
| 2 | C4 | critical | reconnect drain order (file-level before per-cell) | bd6b322 |
| 2 | C5 | critical | drainPendingEdits drops queue on DELETE failure | bd6b322 |
| 2 | H7 | hardening | flushPendingEdits failure-count surfaces in OfflineBanner | bd6b322 |
| 3 | C1/C2/M4 follow-through | critical | **OCC activation** — setVersionLookup wired from MainWindow | 5f6fed6 |
| 3 | C6 | critical | onRemoteCellChanged clobbers dirty cell | 0dd62b9 |
| 3 | C7 | critical | Two-lambda rowChanged race — decoration vs apply ordering | 0dd62b9 |
| 3 | H6 | hardening | Echo on programmatic remote writes via itemChanged | 0dd62b9 |
| 3 | M8 | polish | Stale [this, selfUuid] capture in rowChanged lambda | 0dd62b9 |
| 3 | H1 | hardening | ~LiveSync drops 200 ms throttle queue on close | 222d589 |
| 3 | H9 | hardening | ~LiveSync BlockingQueuedConnection deadlock on dead worker | 222d589 |
| 3 | H2 | hardening | closeEvent/onCloseFile lose debounced Excel edits | 618a2ea |
| 3 | H3 | hardening | writeCellsToExcel silently swallows openpyxl failures | 618a2ea |
| 3 | H5 | hardening | NotificationListener subscribe is all-or-nothing | ade6f52 |
| 3 | H10 | hardening | RowDeletedBanner dismissed unconditionally | 52bead1 |
| 4 | M1 | polish | loadFile N+1 round trips → 5 bulk SELECTs | 6b9d484 |
| 4 | M9 | polish | classifyMissingUpdate transactional race docs | 6b9d484 |
| 4 | H8 | hardening | ImageCache content-hash dedup + close wipe | f28e9af |
| 4 | M2 | polish | runPython fixed-name script file races | 0c76325 |
| 4 | M3 | polish | openpyxl wb.save() leaves truncated file on crash | 0c76325 |

**Total: 25 findings, 9 commits on `worktree-v2.0.2-fixes`** (4 prior
commits landed Phase 1+2; this audit closes the remaining 14).

## Impact summary

### Correctness (data loss / OCC break)

- **C1/C2/M4 (activation)** — Phase 1 plumbed the OCC chain end-to-end but
  no caller registered a `VersionLookup` callback. Phase 3 (`5f6fed6`)
  wires the MainWindow lookup; OCC is now live on every per-cell write
  for the six protected tables. Confirmed via `tst_versionlookup` (12
  slots).

- **C3 (TPM + sensory)** — `tryWriteFile`, `tryWriteSensorySession`, and
  `tryWriteDetailedSensorySession` used DELETE-cascade-rebuild for child
  rows, which wiped concurrent users' edits whenever one user saved.
  Phase 2 replaced the destructive path with a three-phase id-aware
  upsert (pre-image SELECT → UPDATE-or-INSERT → post-prune). Verified by
  `tst_databasemanager::tryWriteFile_idAwareUpsert*` and
  `tst_databasemanager::tryWriteSensorySession_imagesUpsertedNotWiped`.

- **C4/C5/H7 (offline drain)** — the reconnect handler called file-level
  saveFile before draining the per-cell queue, so cells with pending
  offline edits were overwritten by a stale FileResult. The drain itself
  cleared the SQLite queue unconditionally on bulk DELETE failure,
  dropping unreplayed edits. Phase 2 (`bd6b322`) reversed the order,
  added a `replayed_at` sentinel with UPDATE fallback, and surfaced the
  failure count in OfflineBanner.

- **C6/C7 (rowChanged race)** — Phase 3 (`0dd62b9`) consolidated two
  rowChanged subscribers (decoration painter + LiveSync dispatch) whose
  order Qt's signal/slot dispatch did not guarantee. The same commit
  added the dirty-cell guard in `applyRemoteValueToCell` so a remote
  write to a cell the user is actively editing paints yellow instead of
  clobbering.

### Hardening

- **H1/H9** (`222d589`) — `~LiveSync` now drains the 200 ms throttle
  queue before tearing down the worker thread, and guards the worker
  stop invoke against a dead thread (`Qt::BlockingQueuedConnection`
  would deadlock waiting for a slot that never runs).

- **H2/H3** (`618a2ea`) — `closeEvent` and `onCloseFile` flush debounced
  Excel writes before any teardown. `writeCellsToExcel` now returns
  bool; `flushExcelWrites` retains failed entries for retry and
  rate-limits a `QMessageBox` warning.

- **H5** (`ade6f52`) — `NotificationListener::subscribe` attempts each
  of the three channels independently and records successes in
  `QSet<QString> m_subscribedChannels`. A single LISTEN failure no
  longer leaves the listener with no subscriptions at all.

- **H6** (`0dd62b9`) — `m_applyingRemote` flag guards
  `onDataTableItemChanged` against synthesizing a LiveSync commit from
  programmatic remote-applied text changes.

- **H8** (`f28e9af`) — ImageCache uses content-hash filenames
  (`<md5>_<filename>`) so identical blobs across samples/sessions share
  one on-disk file; `closeEvent` wipes the cache directory.

- **H10** (`52bead1`) — `RowDeletedBanner::dismiss()` is gated on save
  success; the sensory + detailed-sensory recreate branches refetch the
  new server row so child image_ids round-trip before the next OCC
  commit.

### Polish (perf + quality)

- **M1** (`6b9d484`) — `loadFile` issues 5 bulk SELECTs filtered on
  `file_id` instead of one-per-parent. A typical 5×20×10 file went from
  ~210 round trips to 5.

- **M2/M3** (`0c76325`) — `QTemporaryFile` for the Python script (race
  fix), `os.replace(tmp, path)` for atomic openpyxl save (corruption
  fix on crash mid-write).

- **M7** (`6c00523`) — `qHash(LiveSync::PendingKey)` chains seeds across
  the three fields so collisions in one component don't collapse to
  O(n) lookup.

- **M9** (`6b9d484`) — comment block on `classifyMissingUpdate`
  explaining the benign transactional race between VersionMismatch and
  RowDeleted classification.

## Test coverage delta

The audit landed:

- **+2 new test binaries** — `tst_versionlookup` (12 slots, pure-function
  unit tests for the OCC version resolver) and
  `tst_mainwindow_remotecell` (3 slots, QApplication-bound tests for
  the dirty-cell guard helper).
- **+4 new test slots** in existing binaries — `tst_livesync` gained
  `destructor_flushesPendingCommits` and
  `destructor_safeWhenWorkerThreadStopped`; `tst_notificationlistener`
  gained `subscribe_tracksEachChannelIndividually` and
  `unsubscribe_clearsAllChannels`.
- Pre-existing Phase 1/2 tests (`tst_livesync::commitCell_occGuard*`,
  `tst_databasemanager::tryWriteFile_idAwareUpsert*`,
  `tst_offlinesnapshot::drainPendingEdits_respectsReplayedAtSentinel`)
  continue to cover the corresponding paths.

**Final state: 34 binaries, 100% pass, 0 skipped.** (Up from 32 binaries
at v2.0.1.)

## What was intentionally not changed

- The `loadSensorySession` / `loadDetailedSensorySession` paths still
  use per-session `loadImagesFor` rather than a bulk SELECT. Sensory
  sessions are typically 1-2 images; the M1-style bulk-ification gains
  little and the path through `loadImagesFor` is shared with the C3
  upsert callers (which need the per-session granularity). Deferred to
  v2.0.3 if telemetry shows it.
- The unused `cachePrefix` / `sessionId+index` parameters in
  `loadImagesFor` are tagged `Q_UNUSED` rather than removed — the
  function is called from three sites and tightening the signature is
  a wider refactor than this audit's scope.
