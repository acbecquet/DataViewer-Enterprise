# DataViewer Enterprise Changelog

## [2.1.0] — 2026-05-27

### Fixes
- Sensory comments-box outline: bumped to a real 3 px border. The v2.0.10 stylesheet set 3 px but the global `QTextEdit { border: 1 px }` rule in `AppTheme.cpp` tied on selector specificity and won; switching to a `QTextEdit#sensoryCommentsEdit` ID selector raises specificity so the inline rule applies.
- Added 4 px of vertical breathing room between the "Comments:" label and the box so they no longer visually touch.

### UpdateChecker hardening
- Added diagnostic logging across `latestAvailable()` / `check()` — the scan now reports the update root, every subfolder it considers, why each one is skipped (not a version, no installer), the best candidate, and the final dialog/no-dialog decision. When the prompt doesn't appear, the log file tells you which branch returned early instead of failing silently.
- Trim whitespace on folder names before parsing — Synology sync occasionally appends trailing spaces.
- If `DataViewer-setup.exe` isn't at the expected path, fall back to a case-insensitive wildcard match (`DataViewer*setup*.exe`) inside the version subdir. Catches installer filenames the sync layer normalised differently.

## [2.0.10] — 2026-05-26

### Fixes
- Load Excel from sensory / detailed sensory mode now routes through file-type detection. A TPM file loaded while in sensory mode auto-switches to TPM mode (and vice versa) instead of being parsed as the wrong type.
- Sensory sample cards widened (220 → 245 px at 3-up, 240 → 265 px at 2-up) so the V/R/HT row fits without the heating-tech combo getting clipped by the card edge.
- Removed up/down arrow buttons from the 5 sensory metric spinboxes and the puff-length spinbox — values are always typed, the buttons just added visual noise.
- Sensory metric form is now centered inside each sample card instead of pinned to the left edge.
- Sample-card comments box gets a clearer outline + focus highlight so it's obvious where to type.
- Voltage field auto-disables under "Variable Voltage", "Variable Power", and "Constant Power" power types where V is not a direct input; resistance and heating-tech remain editable.
- Fixed sensory-session save failing with duplicate-key constraint violation on `idx_sensory_sessions_key`. Three converging bugs:
  - `SensoryPanel::loadFile()` (Excel import path used by the ribbon's Load Excel button) never called `inheritExistingIdsAndVersions()`, so re-imports always entered as `id=-1` and tripped INSERT.
  - The default placeholder name "New Session" now gets replaced with the user's Test Title on save, so two same-day sessions for the same tester no longer collide.
  - `onUpdateDatabase` saved through a copy of `m_sessions`, so server-assigned ids from successful INSERTs never propagated back to panel state. After the first save, the second Ctrl+U / auto-save tick attempted INSERT again on the same natural key. Fixed by calling `inheritExistingIdsAndVersions()` after the sensory save loop to back-fill ids.
- Sensory + detailed-sensory edits now kick the 5-second debounced auto-save timer. Previously only TPM edits started it, so sensory changes lit up the "modified (Ctrl+U)" indicator but never auto-saved.
- Test Title edits in sensory mode now rename the underlying database session (the `session_name` natural-key column follows the displayed Test Title). Renaming into a name another session already holds prompts an override dialog — accepting deletes the existing row and writes the current edits in its place; declining leaves the rename unsaved.
- Sample-card comments box outline tripled from 1px to 3px so it's clearly visible against the white card background, at rest and on focus.

### Removed
- LiveSync optimistic-concurrency check disabled. The v2.0.2 version-lookup callback sent the local row version with every per-cell commit, but the server bumped the row's version on every successful write and there was no back-propagation to update the local cache — so the **second** edit on any row hit OCC miss, the worker emitted `commitConflict`, and no slot was listening, dropping the edit silently. Switched to last-writer-wins, matching the project's "no merging concerns" stance. Cross-user concurrent edits to the same cell will now resolve to whoever wrote last; remote NOTIFY-driven updates still surface in real time for other rows.

### Walked back from earlier v2.0.10 work
- Ctrl+U restored as an explicit save trigger and the "modified (Ctrl+U)" status indicator reinstated. The "kick the timer on sensory edit" experiment from earlier in this cycle pulled the full save loop onto the UI thread every 5 s — on slower LANs that produced "Not Responding" freezes. LiveSync still persists per-cell sensory edits live; Ctrl+U is the session-level flush + TPM file commit.

## [2.0.9] — 2026-05-21

### UI Polish
- AppTheme refactored into named tokens (colors, spacing, radius, type, elevation).
- New palette: lighter app surface (#F5F6F8), slate table headers (#2C3E50), status bar redesigned with light background (#ECEEF1) and semantic dots.
- Replaced Qt stock pixmaps with 16 Lucide SVG icons across the ribbon.
- Added responsive layout: app reflows to compact mode below 1100 px window width (sidebar collapses to icon strip, ribbon goes icons-only, breadcrumb truncates, sensory cards drop columns, detailed sensory dual charts stack).
- Per-screen polish for TPM workspace, Sensory, Detailed Sensory, Database Browser.
- 7 modal dialogs got consistent padding, drop shadows, primary/destructive button hierarchy.

### Behind the scenes
- New `DVE::ResponsiveLayout` singleton with breakpoint detection and 50 ms debounced resize events.
- 1 new test class (`tst_responsivelayout`); now 35 total Qt tests.
- Lucide icon attribution (ISC license) bundled in resources/icons/.

### Known follow-ups (for v2.0.10)
- Sidebar overlay: clicking an icon-strip button restores the full sidebar but doesn't auto-collapse on outside click (needs Qt::Popup overlay).
- Sample card label column at 72 px elides "Overall Flavor"/"Overall Liking" in narrow fonts — bump to ~90 px if it reads badly in production.
- PropertiesPanel spinbox fixed-width 75 px should be verified at 125%+ DPI for clipping.
- tst_responsivelayout intermittently crashes with heap corruption (0xC0000374) under test-runner load — singleton event-filter lifetime fragility; mitigated by retry, root cause investigation deferred.

## 2026-05-21 - v2.0.8 .xlsm Template Support + VBA Upload Sidecar

Adds `.xlsm` (macro-enabled workbook) support to DataViewer's file-type
gates so it can ingest the upload-button-equipped Automated Testing
Template directly. Pipeline path was already format-agnostic (openpyxl
and the COM reader both treat `.xlsm` and `.xlsx` identically); only
the GUI gates needed to open.

Ships alongside a VBA sidecar module
(`excel-sidecar/DataViewerUpload.bas`) for the workbook's two upload
buttons: data-population checklist + copy-to-paths + DataViewer
shell-out. No DB schema changes.

### `.xlsm` gates (src/MainWindow.cpp)
- `detectFileType` now accepts `.xlsm` in addition to `.xlsx`/`.xls`/`.json`.
- The Open File dialog's filter advertises `*.xlsm`.
- Drag-and-drop routes `.xlsm` URLs through `routeFile()`.

### Excel sidecar (excel-sidecar/DataViewerUpload.bas)
- `Btn_UploadAll`: runs the checklist; if it passes, saves the
  workbook, copies the `.xlsm` to `DV_SynologyPath` and `DV_LocalPath`,
  materialises a temporary `.xlsx` (via `SaveAs FileFormat:=51`), and
  shells `DataViewer.exe` on the temp file. DataViewer's
  `onFileLoadFinished` auto-saves the parsed file to Postgres.
- `Btn_DryRunChecklist`: runs only the checklist and reports results
  in `DV_Status` / `DV_Log`.
- Checklist rules: required named ranges non-empty; all 10 canonical
  data sheets present; per populated sample block — unique sample ID
  within sheet, heating technology filled, puff sequence strictly
  increasing, mass-before > mass-after, TPM within `[0, 50]` mg/puff.
- Optional `DV_DataViewerExe` named range overrides the default
  `C:\Program Files\DataViewer Enterprise\DataViewer.exe` path.

## 2026-05-20 - v2.0.7 Codebase Cleanup

No user-visible changes. Internal cleanup release; existing functionality
preserved exactly.

### Git topology
- Removed 2 abandoned worktrees (`cranky-hofstadter-afa8a2`, `v2.0.2-fixes`).
- Deleted 11 merged local branches and pruned 8 merged remote branches.

### Repo-tree
- Removed `test_rcc_output.cpp` (27k-line Qt rcc fallback output tracked
  by accident) and added it to `.gitignore` alongside other generated
  build artifacts.
- Removed `Makefile*`, `.qmake.stash`, `build/`, `debug/`, `release/` from
  working tree (all already gitignored).
- Archived `DATAVIEWER_UPDATES.txt` (pre-CHANGELOG.md 48KB plain-text log)
  to `DataViewer-Archive/`.

### Tests
- `tst_sopLoader::loadsKnownTemplate` and
  `tst_reportgenerator::loadSopRows_filtersToRequestedTests` now skip
  gracefully when their `.xlsx` fixture is MIP-encrypted on the
  developer's machine. Runs cleanly on CI and deployment machines
  without Microsoft Information Protection labels; restores 34/34 green
  on the developer environment.
- `--self-test` extended with `application_version` and
  `zipwriter_roundtrip` smoke methods to catch stale-Makefile builds
  and zlib/ZipWriter linkage regressions.

### Source cleanup
- Removed orphan `src/ui/SensoryDialog.{cpp,h}` (1830 LOC). Never built
  into the binary (not in `.pro`), never `#include`d anywhere — was
  replaced by `SensoryPanel` during the v2.0.0 work but the source
  files lingered. Contained its own duplicates of `FlowLayout` and
  `SampleCard`.
- Removed unused `DatabaseManager::findSensorySessionByKey` and
  `DatabaseManager::findDetailedSensorySessionByKey` (the v2.0.5
  single-key form). Replaced by the v2.0.6 bulk variants which are the
  only callers; the single-key form had zero call sites.
- `PptxWriter::makeTextBox` and `PptxWriter::makeTableCell` now share
  `XmlBuilder::escapeXml` instead of hand-rolling 4 `.replace()` calls
  each. Apostrophe escapes (`&apos;`) now work in shape text.

### Docs
- Archived 25 shipped plan/spec/handoff docs (v1.3.0 through v2.0.5
  work) and one unused diagnostic script (`tools/db_integrity_check.py`)
  to `DataViewer-Archive/2026-05-20-cleanup-pass-3-doc-rot/`.
- Updated `CLAUDE.md` test-class count from 27 → 34 and clarified the
  v2.0.7 status of `test_rcc_output.cpp`.

### Deferred to v2.0.8+
Large file-split refactors (MainWindow.cpp, DatabaseManager.cpp,
SensoryPanel.cpp, DetailedSensoryPanel.cpp, OfflineSnapshot.cpp,
PptxWriter.cpp) were identified in the audit but deferred. See
`docs/superpowers/specs/2026-05-21-post-v2.0.7-refactors.md`.

### Net repo impact
~27,200 fewer lines in tree (mostly from `test_rcc_output.cpp` + the
25 shipped doc archives + SensoryDialog).

## 2026-05-17 - v2.0.2 Save & DB Correctness (Phase 1 of 4)

Phase 1 of a coordinated bug-fix release addressing 25 findings from a deep
audit of the save/DB paths. Phase 1 closes the optimistic-concurrency chain
opened by the v2.0.1 perf fix.

### Postgres migration (apply to NAS before installing v2.0.2 clients)

`deploy/postgres/migrations/2026-05-17-v2.0.2-fixes.sql`:

- **C1 + M6**: `dve_commit_cell` / `dve_commit_cell_json` accept an optional
  `p_expected_version INT` parameter that, when non-NULL, adds
  `AND version = $N` to the WHERE clause for optimistic-concurrency. The
  redundant explicit `version = version + 1` was removed from the SET clause
  — the `bump_version` BEFORE-trigger owns version increments. The new
  parameter has `DEFAULT NULL` so v2.0.1 callers keep working during rollout.
- **M4**: `dve_focus_cell` now scopes its DELETE to the user's *other*
  focus cells; a user can hold focus in multiple cells across panels.
  Uses `ON CONFLICT DO UPDATE` to refresh `started_at` on re-focus.
- **M5**: `dve_cell_focus_cleanup` pg_cron schedule replaced — the v2.0.1
  `'*/30 * * * * *'` was an invalid 6-field expression and the job never
  ran, leaving ghost focus rows forever. Now `'* * * * *'`.
- **H4**: new `dve_commit_session_layout` stored function for OCC-checked
  sensory layout saves with the same trigger-payload enrichment as the
  cell-commit functions.

### C++ wiring

- `LiveSyncWorker::commitScalar` / `commitJson` take a new `qint64
  expectedVersion` argument, bind it as NULL when `< 0`, and read the
  function's BOOLEAN return. On FALSE the worker emits a new
  `commitConflict` signal (distinct from `commitFailed` so the offline
  snapshot does NOT re-enqueue a stale write).
- `LiveSync` accepts a registered version-lookup callback via
  `setVersionLookup` and forwards the expected version to the worker on
  every commit. Without a lookup callback, all commits run with OCC
  disabled (matches v2.0.1 behavior for tests / legacy callers).
- `LiveSync` forwards the worker's `commitConflict` upstream via its own
  `commitConflict` signal (consumed by MainWindow in a later phase).
- `DatabaseManager::saveSensoryLayout` routed through the new
  `dve_commit_session_layout` stored function with inline read-then-write
  retry-once on OCC miss. Removes the version-staleness side effect of
  the old open-coded UPDATE.
- `M7`: `qHash(PendingKey)` now chains seeds across the three fields
  instead of XOR-with-shared-seed.

## 2026-04-01 - Detailed Sensory Mode UI Polish & Fixes

### UI Layout
- **Unified question form**: Replaced two separate column containers with a single grid layout, eliminating the visual separator between left and right columns. The form now appears as one cohesive area.
- **Question numbering**: Added numbers 1-14 to all questions so the left-to-right, top-to-bottom reading order is immediately clear.
- **4-quadrant alignment**: The question form column split now aligns with the dual radar chart split below, creating a clean 4-quadrant visual layout.
- **Input width capping**: Combo boxes capped at 280px, line edits at 220px, spin boxes at 70px. Inputs no longer stretch to fill all available horizontal space.
- **Combined header row**: Merged the header fields (Test Title, Assessor, Tester, Media, Date) and sample navigation (prev/next, Add Sample, Remove) into a single row. Header input fields narrowed to ~90px to fit everything.
- **Margin fixes**: Added 8px left/right padding inside the question grid for label readability. Equalized top/bottom margins on the header row.
- **Comments border**: Added a solid outline (QFrame::Box) to the Comments text edit for visual clarity.

### Data & Charts
- **Inverted radar charts**: Radar plot normalization is now inverted so that a score of 1 (best) maps to the outermost ring (9) and the worst score maps to the innermost ring (1). Good results now fill the plot area instead of appearing mostly empty.
- **Mouthpiece/Draw Resistance**: Changed from a free-text QLineEdit to a 5-option dropdown combo box with descriptive answers:
  1. Very easy pull. Good design overall
  2. Draw resistance fine, mouthpiece needs improvement
  3. Mouthpiece fine, draw resistance too small
  4. Mouthpiece fine, draw resistance too high
  5. Mouthpiece and draw resistance made it very hard to puff

### Files Modified
- `src/ui/DetailedSensoryPanel.cpp` - Form layout, header/nav merge, mouthpiece combo
- `src/ui/DetailedSensoryPanel.h` - Changed `m_mouthpieceEdit` (QLineEdit) to `m_mouthpieceCombo` (QComboBox)
- `src/pipeline/DetailedSensoryData.h` - Added `kMouthpieceOptions`, inverted `normalizeToRadar()`
