# DataViewer Enterprise Changelog

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
