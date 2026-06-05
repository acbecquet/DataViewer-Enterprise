# Plan C — Auto-Recovery + Crash Snapshot — Design

**Date:** 2026-06-04
**Status:** Approved (design); ready for implementation planning
**Branch:** `feature/plan-c-auto-recovery`
**Part of:** the v2.2.x critical-bug initiative. Siblings: **A** = save paths + Settings tab (SHIPPED v2.2.4); **B** = DB-load render fix (SHIPPED, v2.2.5). **C** = this — **Bug 1, the biggest**. Also riding this branch ahead of Plan C: the Settings-paths broadening (commit `dee5fb4`).

---

## Goal

On any **non-clean termination** — a crash *or* the auto-updater's forced restart — restore the user's complete in-memory work across all three modes (TPM files, Sensory sessions, Detailed Sensory sessions), **including edits never written to `.xlsx` or the database**. Plus a gentler, consolidated clean-close save flow so normal exits don't silently drop untitled work.

## Background — why data is lost today (Bug 1)

- **Three in-memory stores** hold all live work: TPM `MainWindow::m_loadedFiles` (`QVector<FileResult>`), `SensoryPanel::m_sessions`, `DetailedSensoryPanel::m_sessions`. The panels already expose `allSessions()` / `loadSessions()` get/set hooks.
- **No autosave of unsaved in-memory edits exists.** The only persistence is debounced `.xlsx` write-back (TPM cell edits only) + DB save (5 s timer + on-close prompt). The offline `snapshot.sqlite` is a pure DB mirror — it holds **zero** unsaved edits.
- **The updater hard-kills the process:** `UpdateChecker`'s "Update Now" calls `std::_Exit(0)`, which bypasses `closeEvent`, destructors, and every flush — identical to a crash. **Therefore any recovery snapshot must be flushed proactively (on-edit / timer), never only at close**, or it fails the exact update case it targets.

## Locked decisions (from brainstorming)

1. **Full in-memory state snapshot** (not open-set + reload). Recovery is byte-faithful, including untitled/never-saved work. **Images stored as path refs**, not bytes, to keep blobs small.
2. **Consolidated close prompt** (Save All / Discard / Cancel), not a Save-As per file. Save All persists named files in place and pops a Save-As **only** for untitled items.
3. **Per-file blobs + index**, with a **move-to-`_prev`** model on startup; the updater flushes synchronously before `std::_Exit`.

---

## Components

### 1. `RecoveryManager` — `src/utils/RecoveryManager.{h,cpp}`
Owns the rolling snapshot and the crash/clean detection. Sketch:
- `noteDirty(kind, id)` / `scheduleFlush()` — marks an item dirty and arms a **~2 s debounce**.
- `flushNow(bool synchronous=false)` — writes dirty blobs + index **atomically (temp + `rename`)**, **off the UI thread** (`QtConcurrent`) in the normal case; a synchronous variant for the updater path. (We explicitly avoid the synchronous-DB-on-UI-thread pattern that caused the v2.0.6 "Not Responding" freezes.)
- `bool hasRecoverable() const` — `recovery_prev/index.json` present and non-empty.
- `QVector<RecoveryEntry> recoverableItems() const` — feeds the reopen prompt and `RecoverDialog`.
- `adoptPreviousSession()` — startup: rename `recovery/` → `recovery_prev/`, create fresh `recovery/`.
- `clear()` — teardown on clean close (removes both dirs).

### 2. Snapshot store — `%LOCALAPPDATA%/DataViewer/recovery/`
- `index.json`: one entry per open item `{ kind: tpm|sensory|detailed, id, displayName, sourcePath?, dirty, blob }`.
- Per-item blobs: `tpm_<id>.json`, `sensory_<id>.json`, `detailed_<id>.json` — full serialized state. **Per-file so a single edit only rewrites its own blob**, keeping writes cheap even for large multi-file sessions.
- **Cadence:** 2 s debounce after the last mutating action (cell edit, field change, row add/remove, file open/close/rename) + a 30 s safety flush + a synchronous flush before `std::_Exit` (updater). Clean `closeEvent` tears the store down instead.

### 3. Detection state machine (no false positives after updates)
- **Startup:** if `recovery/` exists, the previous instance died uncleanly (crash *or* updater `std::_Exit`). `SingleInstance` guarantees only one running instance, so a leftover `recovery/` is unambiguous. → `adoptPreviousSession()` renames it to `recovery_prev/` and starts a fresh `recovery/`. The reopen prompt + Recover button read `recovery_prev/`; the live snapshot writes `recovery/`.
- **Clean close** (user chose Save All or Discard): `clear()` removes both dirs → no prompt next launch.
- **Updater path:** `UpdateChecker` calls `RecoveryManager::flushNow(synchronous=true)` immediately before `std::_Exit(0)` and does **not** delete — so a self-update recovers fully (the original top pain point).

### 4. Serialization (the net-new work)
- **`ReportDataJson.{h,cpp}` (src/pipeline/):** `toJson`/`fromJson` for `FileResult → SheetResult → SampleResult → DataRow`. The field set **mirrors the DB column mapping** (`DatabaseManager`) to stay a single source of truth; a round-trip unit test plus a field-parity check guard against drift.
- **Reuse** `SensorySession` JSON (`sensorySessionToJson`/`FromJson` in `SensoryData.cpp`).
- **Promote** `DetailedSensorySession` JSON out of the anonymous namespace in `DatabaseManager.cpp` into public functions in `DetailedSensoryData.cpp`, with a round-trip test.

### 5. Clean-close flow (consolidated)
- `closeEvent` builds an **unsaved inventory** across all three stores: TPM (`m_modifiedFilePaths` + any never-saved file), Sensory (`m_sensorySessionsDirty` / placeholder sessions via `isPlaceholderSession()`), Detailed (`m_detailedSensorySessionsDirty` — **newly honored; this fixes the pre-existing gap** where `promptSaveDatabase` ignored it).
- One dialog: **Save All / Discard / Cancel**. Save All persists named files in place and pops a Save-As only for untitled items. Then `RecoveryManager::clear()` and accept. Cancel vetoes the close.

### 6. Recovery flow
- **Auto-prompt** at startup when `hasRecoverable()`: *"N files were open in the previous session. Reload them?"* → **Reload all** restores every blob into the right store, marks each modified (so the user can re-save), and switches to the relevant mode. **Not now** keeps `recovery_prev/` so the Recover button still works this session.
- **Tools → Recover button** (`MainWindow::buildToolsTab` + `RibbonGroup::addLargeButton`): opens `RecoverDialog` (`src/ui/`) listing previous items (name, mode, saved/unsaved) with checkboxes for **selective** reload.

### 7. Restore reconciliation
Restored items keep their DB `id`/`version` but are marked modified; the next save goes through the normal optimistic-concurrency path (the panels' `inheritExistingIdsAndVersions()` covers natural-key cases), so recovery can't trigger spurious unique-violation / stale-row dialogs.

---

## Files

**Create:**
- `src/utils/RecoveryManager.{h,cpp}`
- `src/pipeline/ReportDataJson.{h,cpp}`
- `src/ui/RecoverDialog.{h,cpp}`
- `tests/tst_reportdatajson/` and `tests/tst_recoverymanager/` (+ wire into `tests/tests.pro`)

**Modify:**
- `src/MainWindow.{h,cpp}` — edit signals → `noteDirty`; startup recovery prompt; consolidated `closeEvent` flow + unsaved inventory; Tools Recover button.
- `src/utils/UpdateChecker.cpp` — synchronous `flushNow` before `std::_Exit`.
- `src/pipeline/DetailedSensoryData.{h,cpp}` — promote `DetailedSensorySession` JSON to public.
- `src/ui/SensoryPanel.*`, `src/ui/DetailedSensoryPanel.*` — expose/confirm dirty state for the inventory (get/set session APIs already exist).
- `DataViewerEnterprise.pro` — new sources/headers; `VERSION` bump at the end.

## Tests
- `ReportData` JSON round-trip + DB-field-parity guard.
- `DetailedSensorySession` JSON round-trip (post-promotion).
- `RecoveryManager`: snapshot write→read fidelity; detection (orphan `recovery/` → adopted as recoverable); `clear()` teardown; synchronous updater flush writes a complete blob set.
- Close inventory: a dirty Detailed session is now detected (regression guard for the fixed gap).

## Mixed / edge cases
- **Image path-refs that moved/deleted** between sessions: restore the entry, show a "image missing" placeholder rather than failing the whole reload.
- **Very large sessions:** writes are per-blob and off-thread; if a single blob is huge, the atomic temp+rename still applies (no partial files).
- **User declines recovery then works and closes cleanly:** clean close clears both `recovery/` and `recovery_prev/`, so the declined data does not resurface later.

## Non-goals
- Auto-update **download/replace** internals (Plan C only guarantees the snapshot is flushed before the existing `std::_Exit`).
- Cross-machine recovery.
- Undo/redo-stack recovery.
- Per-keystroke snapshots (debounced only).
- **(I-1) Sensory/Detailed image associations are not recovered.** Sensory and
  Detailed recovery restores session **data** (scores, comments, header/session
  fields) but **NOT** image associations: the `SensorySession` /
  `DetailedSensorySession` JSON serializers (`sensorySessionToJson` /
  `detailedSensorySessionToJson`) deliberately omit images — image bytes live in
  the Postgres `images` table, not in the session blob — so the recovery
  snapshot, which reuses those serializers, inherits the same omission. (TPM
  recovery, by contrast, goes through `fileResultToJson` and **does** include
  image refs.) Changing this would require altering the shared session JSON
  contract — and therefore the on-disk DB blob format — which is out of scope
  for Plan C; recovered sensory/detailed sessions simply re-associate images
  from the DB on next save.

## Open items (resolve during planning/implementation)
- Exact debounce / safety intervals (2 s / 30 s proposed).
- Whether to cap snapshot size or warn on pathologically large sessions.
- Final placement of `RecoveryManager` (utils vs a new `src/recovery/`).

## Build / version
A **minor** `VERSION` bump (→ `2.3.0`, this is a feature) + installer happens at the **end** of implementation, bundling the already-committed Settings-paths fix. Then: user installer eyeball-test → merge `feature/plan-c-auto-recovery` → `main` → user's Synology drop.
