# Live Collaborative Editing — Design (Track B / v2.0.1)

- **Date:** 2026-05-16
- **Status:** Draft (pending user review)
- **Version target:** v2.0.1 (patch release on top of v2.0.0)
- **Predecessor specs:**
  - `2026-05-11-postgres-multiuser-design.md` — the v2.0 save-driven concurrency layer this design retires
  - `2026-05-15-postwave-bugfix-batch-design.md` — the batch of v2.0 polish landed immediately before this work

## Goal

Replace the save-driven optimistic-locking layer with continuous per-cell
synchronization over Postgres NOTIFY. Editing feels like Google Sheets:
every committed cell is durable instantly, every other open client sees
the change in under one second, and the user-visible "save" flow becomes
"export to Excel" only. The v2.0 conflict-dialog layer
(SaveCoordinator / ConflictResolver / VersionMismatchDialog /
RowDeletedDialog) is deleted in this release.

## Scope

Applies to all three editable workflows: **TPM mode** (data_rows table),
**Sensory mode** (json_data JSONB samples), and **Detailed Sensory mode**
(json_data JSONB samples). Every editable widget in those workflows
broadcasts cell focus + cell commits the same way.

## Decisions captured during brainstorming

1. Cell decoration is a 2px border in the editor's user color plus a
   small name flag positioned above the cell ("Variant C" in the
   2026-05-16 brainstorm session).
2. Same-cell collision resolves with last-writer-wins. The user whose
   value was overwritten sees their cell flash for ~200 ms when the
   remote value arrives so the overwrite is visible.
3. The Save button is renamed "Export to Excel". The debounced Excel
   write-back timer (`m_excelWriteTimer`) keeps running and now writes
   remote edits to the local .xlsx in addition to local edits. The
   button forces an immediate flush.
4. Offline behavior: per-cell commits queue in the existing pending-edit
   infrastructure (extended to cell granularity) and replay on
   reconnect. LWW resolves any drift.
5. Single-shot migration. v2.0.1 deletes the v2.0 conflict layer; both
   code paths do not coexist. The small office team (≈10 users on the
   LAN) updates together.

## Architecture

One new module plus targeted replumbing of three existing layers.

### `src/database/LiveSync.{h,cpp}` (new)

The single chokepoint for cell-level writes and remote applies. Owns:

- `commitCell(table, rowId, column, value)` → executes the targeted
  `UPDATE`, bumps `version`, sets `updated_by`. Returns the new version.
  Errors propagate to callers so the offline queue can capture them.
- `applyRemoteCellChange(RowChange)` slot wired to
  `NotificationListener::rowChanged` — looks up the row in the
  in-memory model, applies the new column value, sets the
  `kFlashRole` for ~200 ms.
- `focusCell(table, rowId, column)` / `blurCell()` — upserts/deletes a
  row in the new `cell_focus` table and pg_notifies a new
  `cell_focus` channel so other clients can draw the border + name flag.

### `src/widgets/LiveTableModel.{h,cpp}` (new)

A `QAbstractTableModel` for TPM `data_rows`. Replaces the current
`QTableWidget`-with-manual-`setItem` pattern. Exposes:

- Standard `data()` / `setData()`. `setData()` calls
  `LiveSync::commitCell` and updates the local cache only after the
  commit returns successfully (or queues to the offline pending-edit
  store when offline).
- Custom roles `kRemoteFocusUserRole`, `kRemoteFocusColorRole`,
  `kFlashRole` consumed by the delegate.
- `applyRemoteRow(rowId, column, value)` invoked by `LiveSync`.

### `src/widgets/CellFocusDelegate.{h,cpp}` (new)

Paints the Variant-C decoration:

- 2px solid border in the remote user's hex color, painted inside the
  cell rect so the next cell's border doesn't double up.
- A name flag positioned above the cell — small rounded-top rectangle
  in the same color carrying the user's display name in white.
- 200 ms flash animation when `kFlashRole` is set, driven by a single
  `QVariantAnimation` running at view scope so we don't allocate one
  per cell.

### `src/ui/SensoryPanel.cpp` and `DetailedSensoryPanel.cpp`

Each `QLineEdit`, `QDoubleSpinBox`, `QComboBox`, and `QTextEdit` inside
a `SampleCard` (regular and detailed) gets:

- `editingFinished` (or `currentTextChanged` for combos) →
  `LiveSync::commitCell("sensory_sessions", sessionId,
  "json_path:samples[i].<field>", newValue)`.
- Remote applies arrive through `LiveSync::applyRemoteCellChange` and
  patch the underlying `SensorySample` struct, then re-invoke
  `SampleCard::fromSample()` with signals blocked so the widget updates
  without re-emitting `changed()`. The widget briefly flashes via its
  stylesheet for ~200 ms.

**Sensory cards intentionally skip the Variant-C cell border.** Each
card is already a QGroupBox with its own border; per-widget borders
inside would be visually noisy and harder to paint without subclassing
QLineEdit / QDoubleSpinBox / etc. The existing file-level presence
(avatar bar at the top of the panel + the navigator dot) carries the
"who is here" signal for sensory work. Per-cell focus broadcasts are
therefore omitted for sensory in v2.0.1 — only commits go through
`LiveSync`. If real-world use shows users want per-card decoration we
can add it as a follow-up by introducing a small `WidgetFocusAdorner`
that overlays a colored border via a child `QFrame`.

### Save-button and Excel write-back

- Rename the Ribbon "Save" button to "Export to Excel".
- Keep `m_excelWriteTimer` running. When a remote cell change arrives
  via `applyRemoteCellChange`, push the same write into the local
  `m_pendingWrites` set so the next `flushExcelWrites()` includes it.
- The Export-to-Excel button calls `flushExcelWrites()` synchronously.

### Offline behavior

The existing `PendingEditQueue` extends to per-cell commits. While the
`ConnectionMonitor` reports offline:

- `LiveSync::commitCell` queues the (table, rowId, column, value) tuple
  to the in-memory queue and to the SQLite snapshot.
- The `OfflineBanner` pending-count reflects total queued cell ops.
- On reconnect, the queue replays in order. LWW resolves any drift.

## Schema additions

```sql
-- New: per-cell focus state for live presence borders.
CREATE TABLE IF NOT EXISTS cell_focus (
    user_uuid    UUID        NOT NULL,
    table_name   TEXT        NOT NULL,
    row_id       BIGINT      NOT NULL,
    column_name  TEXT        NOT NULL,
    started_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (user_uuid, table_name, row_id, column_name)
);
CREATE INDEX IF NOT EXISTS idx_cell_focus_target
    ON cell_focus(table_name, row_id);

-- For Sensory cells the column_name is an opaque JSON path string,
-- e.g. "json_path:samples[2].scores.smoothness". The server does not
-- interpret it; clients use the table_name to know which workflow it
-- belongs to.
```

### Trigger updates

The existing row-changed trigger pg_notifies a payload like
`{table, op, id, updated_by}`. Extended to include `column` and
`new_value` for UPDATE ops on the live-sync tables so the receiver
doesn't have to re-SELECT just to know what changed:

```json
{"table":"data_rows","op":"UPDATE","id":12345,
 "column":"draw_pressure","new_value":1.42,
 "updated_by":"a1b2c3..."}
```

For Sensory JSONB rows, `column` carries the JSON path and `new_value`
carries the scalar at that path. Server-side helpers use `jsonb_set` so
two clients editing different paths inside the same `json_data` blob
don't clobber each other.

## What gets retired

Deleted in v2.0.1:

- `src/database/SaveCoordinator.{h,cpp}`
- `src/database/ConflictResolver.{h,cpp}`
- `src/database/VersionMismatchDialog.{h,cpp}`
- `src/database/RowDeletedDialog.{h,cpp}`
- All call sites routing through `SaveCoordinator`
  (`MainWindow::onUpdateDatabase` sensory loop,
  `SensoryPanel::save`, `DetailedSensoryPanel::save`).
- Their tests if any exist.

The placeholder-session predicate (`isPlaceholderSession`) stays — it
now gates `LiveSync::commitCell` for sensory sessions instead of
`SaveCoordinator::saveSensorySession`.

Retained:

- `UniqueViolationDialog` — unique-key violations still happen on
  initial INSERT (e.g. creating a new file or sensory session with a
  duplicate natural key).
- `OfflineSnapshot`, `ConnectionMonitor`, `OfflineBanner` — extended
  to per-cell granularity but otherwise unchanged.
- `PresenceManager` — file/session-level presence stays. The new
  `cell_focus` table is a separate channel for the finer-grained
  cell decoration.

## Testing strategy

New test classes:

- `tst_livesync` — exercises `LiveSync::commitCell` against the
  ephemeral `postgres:16` test container started by
  `tests/start-test-postgres.ps1`. Verifies the UPDATE happens,
  version bumps, NOTIFY payload arrives within 1 s.
- `tst_livetablemodel` — drives a `LiveTableModel` with mock data,
  verifies `setData()` calls `LiveSync`, `applyRemoteRow` updates the
  model, role transitions correctly, flash state clears after the
  animation interval.
- `tst_cellfocusdelegate` — pixel test in the same spirit as
  `tst_presencedotsdelegate`. Renders a cell with a remote focus role
  and asserts the border color, border thickness, and name flag are
  in the expected pixel regions.

Existing tests updated:

- `tst_twoclient_e2e` — extended to drive a per-cell edit on client A
  and assert client B sees it within 1 s.
- `tst_offlinesnapshot` — verifies queued cell commits replay on
  reconnect and apply LWW correctly.

Tests removed alongside the deleted classes (if any exist):
`tst_savecoordinator`, `tst_conflictresolver`,
`tst_versionmismatchdialog`, `tst_rowdeleteddialog`.

## Out of scope

- Real-time cursor presence beyond the cell border (no
  character-by-character caret broadcast).
- Operational transform for text fields. Last-writer-wins applies to
  whole-cell values including multi-character `Notes` / `Comments`.
- Undo/redo across remote edits. Local undo stack still works for the
  user's own edits; remote writes are not undoable from the
  receiving side.
- Audit-log / history UI. `updated_at` / `updated_by` are still
  recorded but not surfaced.

## Risks

- **NOTIFY storm.** A user arrow-keying through 50 cells could fire 50
  focus changes in five seconds. Mitigation: debounce `focusCell` by
  150 ms; cancel a pending focus broadcast if `blurCell` fires first.
- **Per-cell roundtrip latency.** Every commit is a network round-trip.
  Sub-100 ms expected on LAN; profile under load before declaring done.
- **JSONB merge correctness.** Two users editing different fields of the
  same sample must not clobber each other. Server-side update uses
  `jsonb_set` keyed on the supplied JSON path; deep paths require careful
  path-parsing.
- **Cell-focus cleanup on crash.** A user whose process dies leaves
  stale `cell_focus` rows. The 30 s heartbeat-staleness window cleans
  them up, but a crashed user's border lingers for up to 30 s on
  other clients.

## Execution plan

After this spec is reviewed and approved, transition to writing-plans
to produce `2026-05-16-live-collab-plan.md` as a sequenced
implementation plan. The plan is expected to be substantially larger
than the v2.0.1 bug-fix batch — schema migration, three new modules
(LiveSync, LiveTableModel, CellFocusDelegate), retirement of four
existing modules, and tests for all of it. Likely structured as
five or six ordered task groups: schema → LiveSync core → TPM
integration (model + delegate) → Sensory integration → conflict-layer
retirement → end-to-end and offline tests.
