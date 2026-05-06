# Sensory Report Preview — Design

- **Date:** 2026-05-06
- **Status:** Draft (pending user review)
- **Scope:** Sensory mode only. TPM/detailed-sensory unchanged.

## Goal

Add a WYSIWYG preview/editor that opens when the user clicks "Sensory Report".
Lets the user:

1. See every body slide in the staged report exactly as it'll render.
2. Switch between slides via a thumbnail strip.
3. Drag/resize the table, radar chart, legend, and title on each slide.
4. Sort the table by clicking column headers.
5. Toggle samples in/out of the report via a checkbox panel (does not modify
   underlying session data).
6. Edit the cumulative summary slide(s) the same way.

The preview state IS the report — clicking "Create Report" generates a PPTX
that matches whatever's on screen.

Layout edits **persist** so reopening the same Excel file or the same DB
session shows the previous layout.

## Non-goals

- Editing cover or group-divider slides (templated; will be inserted by the
  PPTX writer using existing template assets at generation time).
- TPM-mode report preview (separate future port).
- Layout editing for cumulative reports across runs/users (cumulative layout
  is a global user preference, not a per-(session-set) setting).
- Undo/redo on the canvas (v2).
- Layout templates/presets (v2).

## Architecture

### UI: Modal QDialog

`SensoryReportPreviewDialog` (new, in `src/ui/`) replaces the current
"save dialog → generate" flow. The existing `SensoryPanel::generateFullReport()`
and `generateCombinedPptx()` invoke this dialog instead of going straight
to PPTX writing.

```
┌─ SensoryReportPreviewDialog ─────────────────────────────────────────────┐
│ ┌────────────┐ ┌──────────────────────────────────────────────────────┐ │
│ │ Slide      │ │  QGraphicsView canvas                                │ │
│ │ thumbnails │ │  800 × 450 px  (= 13.33" × 7.5", 60 px/inch)         │ │
│ │ ──────     │ │                                                      │ │
│ │ □ S1 Body  │ │   ┌───────────────┐  ┌───────────────┐               │ │
│ │ □ S1 Imgs  │ │   │  Table        │  │  Radar chart  │               │ │
│ │ □ S2 Body  │ │   │  (sortable)   │  │               │               │ │
│ │ □ S2 Imgs  │ │   └───────────────┘  └───────────────┘               │ │
│ │ □ Cumul.   │ │                                                      │ │
│ │            │ │   Legend ─────────────────────────                   │ │
│ ├────────────┤ │                                                      │ │
│ │ Samples    │ │  (Selected items get drag/resize handles, same as    │ │
│ │ ─────      │ │   ImageViewDialog's ResizableImageItem.)             │ │
│ │ ▼ Session 1│ │                                                      │ │
│ │  ✓ Briq 2-1│ │                                                      │ │
│ │  ✓ Briq 2-2│ │                                                      │ │
│ │  ✗ Briq 2-3│ │                                                      │ │
│ │ ▼ Session 2│ │                                                      │ │
│ │  ✓ Briq 2-1│ │                                                      │ │
│ │  ✓ Briq 2-2│ │                                                      │ │
│ └────────────┘ └──────────────────────────────────────────────────────┘ │
│                                              [ Cancel ] [ Create Report ]│
└──────────────────────────────────────────────────────────────────────────┘
```

### Slide types and editability

| Slide | Shown in preview? | Editable elements |
|---|---|---|
| Cover (template) | No | — |
| Group dividers (template) | No | — |
| Content (per session, "S1 Body") | Yes | Title, Table, Radar chart, Legend |
| Image (per session, "S1 Imgs") | Yes | Image positions, sizes, crops (reuse ImageViewDialog logic) |
| Cumulative summary ("Cumul.") | Yes | Title, Table, Radar chart, Legend |

### Editable canvas elements

Each editable slide element is a subclass of `ResizableSlideItem` (new),
which generalises `ImageViewDialog`'s `ResizableImageItem`:

- `PlotItem` — wraps a re-rendered `QPixmap` of the radar chart. On resize,
  the chart re-renders at the new aspect ratio.
- `TableItem` — renders a styled `QGraphicsTextItem`-based grid. Column
  headers are clickable for sort.
- `LegendItem` — colour swatches + labels.
- `TextItem` — title text, double-click to edit.

Drag/resize behaviour mirrors `ResizableImageItem`: corner handle for resize,
body for move. Aspect-ratio policy per item type:

| Item | Aspect ratio |
|---|---|
| `PlotItem` (radar chart) | Locked (1:1) — distortion looks bad on radar |
| `TableItem` | Free — rows/cols flex |
| `LegendItem` | Free |
| `TextItem` | Free |
| Image items (existing) | Locked (existing behaviour) |

Snap-to-grid and alignment guides omitted in v1.

### Sample checkbox panel

Below the thumbnail strip, scrollable. Grouped by session, each sample as a
`QCheckBox` (defaults checked).

Unchecking re-renders the affected slide(s) immediately:

- The session's content slide table loses the row, radar chart loses the polygon.
- The session's image slide is unaffected (images aren't tied to samples).
- The cumulative summary recomputes: a sample is in cumulative IFF at least
  one of its host sessions has it checked.

Checkbox state lives in the layout JSON as `excludedSamples` per session.

### Sortable column headers

Click sequence on a column header:

1. First click → sort descending by that column (best-first for sensory).
2. Second click → sort ascending.
3. Third click → revert to insertion order.

Sort state lives in the layout JSON as `tableSort: { column, order }`.

### Persistence

#### Per-session layouts

- **DB:** new column `sensory_sessions.layout_json TEXT` (nullable; NULL = use
  defaults). Migration adds the column with `DEFAULT NULL`.
- **Excel:** workbook custom property `dve_layout` (JSON string). Round-tripped
  via openpyxl in the existing `ExcelReader` Python helper and the existing
  cell-write-back Python helper.

JSON shape:

```json
{
  "version": 1,
  "tableSort": { "column": "Overall Liking", "order": "desc" },
  "excludedSamples": ["Briq 2-3"],
  "contentSlide": {
    "title":      [x, y, w, h],
    "table":      [x, y, w, h],
    "radarChart": [x, y, w, h],
    "legend":     [x, y, w, h]
  },
  "imageSlide": {
    "imageLayouts": [[x, y, w, h], ...],
    "imageCrops":   [[x, y, w, h], ...]
  }
}
```

`imageLayouts` and `imageCrops` move from per-sample storage into the unified
layout JSON. Backwards compat: when `layout_json` is NULL, fall back to legacy
per-sample fields.

Coordinates are in inches (canvas: 13.33" × 7.5"), matching the existing
PPTX EMU conversion in `PptxWriter`.

#### Cumulative layout

- **DB:** `settings` table, key `sensory.cumulative_layout`, value = JSON.
- **Excel:** **not stored.** A cumulative report can span N Excel files; no
  natural home. DB is the only source of truth.

JSON shape:

```json
{
  "version": 1,
  "tableSort": { "column": "Overall Liking", "order": "desc" },
  "title":      [x, y, w, h],
  "table":      [x, y, w, h],
  "radarChart": [x, y, w, h],
  "legend":     [x, y, w, h]
}
```

#### Write timing

- **Auto-save** on every edit, debounced 500ms (same pattern as
  `MainWindow::m_excelWriteTimer`).
- On dialog close, flush pending writes immediately before destruction.
- Auto-save writes to DB synchronously, queues an Excel write that runs
  via the existing Python subprocess pipeline.

#### Conflict resolution

- **On load:** if both DB and Excel hold a layout for the session, take the
  one with the newer `loaded_at` (DB) / `lastModified` (Excel). Same
  last-write-wins convention used elsewhere for sensory data.
- **On save:** DB first (synchronous), then queue Excel write. If the Excel
  write fails (file locked, network share unavailable), log a warning and
  keep going — DB is authoritative.

### Code structure

#### New files

- `src/ui/SensoryReportPreviewDialog.{h,cpp}` — the dialog itself
- `src/ui/SlideCanvasItems.{h,cpp}` — `ResizableSlideItem` base class +
  `PlotItem`, `TableItem`, `LegendItem`, `TextItem` subclasses
- `src/reporting/SensoryReportLayout.{h,cpp}` — JSON model: serialize,
  deserialize, defaults, version migration

#### Modified files

- `src/database/DatabaseManager.{h,cpp}`
  - Schema migration: add `sensory_sessions.layout_json` column
  - `loadLayoutForSession(int sessionId)` / `saveLayoutForSession(...)`
  - `loadCumulativeLayout()` / `saveCumulativeLayout(...)` via settings table
- `src/ExcelReader.{h,cpp}` — read `dve_layout` custom property in the
  Python helper, expose it on the parsed result
- Excel write-back Python helper in MainWindow / SensoryPanel — accept and
  write `dve_layout` custom property
- `src/ui/SensoryPanel.{h,cpp}`
  - `generateFullReport()` opens `SensoryReportPreviewDialog` instead of
    going straight to PPTX
  - `generateCombinedPptx()` accepts `QHash<int, SensoryReportLayout>` (session-id
    → layout) and a `SensoryReportLayout cumulativeLayout`; falls back to
    defaults when missing
- `src/reporting/PptxWriter.{h,cpp}` — `addContentSlide(...)` accepts
  `SlideElementLayout` overrides for table/plot/legend/title positions
  (currently hardcoded inches)

#### Touched lightly

- `src/MainWindow.cpp` — wire the new dialog flow
- `tests/` — see Testing below

## Data flow

```
┌──────────────────────────────────────────────────────────────────────┐
│ User clicks "Sensory Report" in ribbon                              │
└──┬───────────────────────────────────────────────────────────────────┘
   │
   ▼
┌──────────────────────────────────────────────────────────────────────┐
│ SensoryPanel::generateFullReport (or generateCombinedPptx)          │
│   1. Collect selected sessions                                       │
│   2. For each session: load layout JSON (DB > Excel > defaults)      │
│   3. Load cumulative layout from DB settings (or defaults)           │
│   4. Open SensoryReportPreviewDialog with sessions + layouts         │
└──┬───────────────────────────────────────────────────────────────────┘
   │
   ▼
┌──────────────────────────────────────────────────────────────────────┐
│ SensoryReportPreviewDialog                                          │
│   - User edits canvas → updates SensoryReportLayout in memory       │
│   - Debounced auto-save → DatabaseManager + Excel writeback         │
│   - "Create Report" → emit accept(), close, hand layouts back       │
│   - "Cancel" → emit reject(), edits already auto-saved              │
└──┬───────────────────────────────────────────────────────────────────┘
   │ (on accept)
   ▼
┌──────────────────────────────────────────────────────────────────────┐
│ SensoryPanel calls PptxWriter with sessions + per-session layouts   │
│ + cumulative layout                                                  │
│   - PptxWriter.addContentSlide(layout-aware)                         │
│   - PptxWriter.addImageSlide(uses imageLayouts)                      │
│   - PptxWriter.addContentSlide(cumulative, with cumulativeLayout)    │
└──────────────────────────────────────────────────────────────────────┘
```

## Testing

New tests:

- `tests/tst_sensoryreportlayout/` — round-trip JSON serialization, version
  migration, defaults computation
- Extend `tests/tst_databasemanager/` — `layout_json` column persistence,
  `sensory.cumulative_layout` settings round-trip, schema migration from
  pre-layout DB

Manual deployment self-test addition (`tests/deployment/Test-Deployment.ps1`):

- Open sensory report preview, edit one element, close preview, reopen,
  verify layout restored.

## Risks & mitigations

| Risk | Mitigation |
|------|-----------|
| Excel custom-property write fails on Synology share | DB authoritative; log warning, don't fail report |
| Layout JSON shape drift across releases | `version` field; migration helpers; never delete unknown fields on load |
| Plot rerender on resize is slow | Render at low DPI in canvas (~96), high DPI only for final PPTX |
| Modal dialog blocks main window during edit | Acceptable — staging a report is a focused activity |
| User unchecks all samples in a session → empty content slide | Show inline warning on that slide; allow but warn before "Create Report" |
| Stale layout when session adds/removes samples in DB after layout saved | Layout JSON references samples by name; missing names ignored, new samples default to checked |

## v2 candidates (out of scope here)

- Port to TPM and detailed-sensory reports
- Cover/divider slide editing
- Layout templates / presets (save current layout as a named preset)
- Undo/redo
- Snap-to-grid, alignment guides
- Export layout as standalone JSON for sharing
