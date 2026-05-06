# Report Preview — Design

- **Date:** 2026-05-06
- **Status:** Draft (pending user review of expanded scope)
- **Scope:** All three report modes — Sensory, TPM, Detailed Sensory.

## Goal

Add a WYSIWYG preview/editor that opens when the user clicks **any** "Report"
button. Lets the user:

1. See every body slide in the staged report exactly as it'll render.
2. Switch between slides via a thumbnail strip.
3. Drag/resize the table, plots, legend, and title on each slide.
4. Sort the table by clicking column headers.
5. Toggle samples in/out of the report via a checkbox panel (does not
   modify underlying data).
6. Edit cover and group-divider slides — title text only, template assets
   (logos, colour theme, fixed layout) stay locked.
7. Edit cumulative summary slides the same as content slides.
8. Use snap-to-grid and alignment guides while dragging.
9. Undo/redo any edit.
10. Save the current layout as a named preset, apply a saved preset,
    export the layout as a JSON file, import one from a JSON file.

The preview state IS the report — clicking "Create Report" generates a
PPTX that matches whatever's on screen.

Layout edits **persist** across sessions per data anchor (the file or the
DB row that produced the slides). Reopening the same source shows the
previous layout.

## Non-goals

- Real-time collaborative editing (one user at a time).
- Version history per layout — last-write-wins.
- Diff view between layouts.
- A PowerPoint plugin.

## Architecture

### Modal QDialog: `ReportPreviewDialog`

Replaces the current direct-to-PPTX flow in all three modes. Mirrors
`ImageViewDialog`'s `QGraphicsView`-based pattern, scaled to handle a
multi-slide report.

```
┌─ ReportPreviewDialog ─────────────────────────────────────────────────────────────────┐
│ ┌────────────┐ ┌──────────────────────────────────────────┐ ┌─────────────────────┐  │
│ │ Slide      │ │  QGraphicsView canvas                    │ │ Properties panel    │  │
│ │ thumbs     │ │  800 × 450 px (= 13.33"×7.5")            │ │ ──────────────      │  │
│ │ ──────     │ │                                          │ │ Selected: Table     │  │
│ │ □ Cover    │ │   ┌───────────┐  ┌───────────┐           │ │ x: 0.46"  y: 0.92"  │  │
│ │ □ Section1 │ │   │  Table    │  │ Radar     │           │ │ w: 6.50"  h: 4.20"  │  │
│ │ □ S1 Body  │ │   │ (sortable)│  │ chart     │           │ │ [Bring Forward]     │  │
│ │ □ S1 Imgs  │ │   └───────────┘  └───────────┘           │ │ [Send Backward]     │  │
│ │ □ Section2 │ │                                          │ │                     │  │
│ │ □ S2 Body  │ │   Legend ────────                        │ │ Sort: Overall ▾ desc│  │
│ │ □ Cumul.   │ │                                          │ │                     │  │
│ │            │ │   Alignment guides + snap-to-grid live   │ │                     │  │
│ ├────────────┤ │                                          │ ├─────────────────────┤  │
│ │ Samples    │ │                                          │ │ Toolbar             │  │
│ │ ──────     │ │                                          │ │ Snap[✓]  Grid 0.1"  │  │
│ │ ▼ Session1 │ │                                          │ │ Preset: ▾  [Save…]  │  │
│ │  ✓ Briq2-1 │ │                                          │ │ [Import…] [Export…] │  │
│ │  ✓ Briq2-2 │ │                                          │ │ [Undo] [Redo]       │  │
│ │  ✗ Briq2-3 │ │                                          │ │                     │  │
│ │ ▼ Session2 │ │                                          │ │                     │  │
│ │  ✓ Briq2-1 │ │                                          │ │                     │  │
│ └────────────┘ └──────────────────────────────────────────┘ └─────────────────────┘  │
│                                                          [ Cancel ] [ Create Report ] │
└───────────────────────────────────────────────────────────────────────────────────────┘
```

### Mode polymorphism: `IReportSource`

Mode-specific differences (data shape, persistence target, PPTX writer
invocation) are encapsulated behind an interface so the dialog stays
mode-agnostic:

```cpp
class IReportSource {
public:
    virtual ~IReportSource() = default;

    // Identity
    virtual QString modeId() const = 0;          // "sensory" | "tpm" | "detailed_sensory"
    virtual QString sourceLabel() const = 0;     // shown in dialog title bar

    // Slides
    virtual int slideCount() const = 0;
    virtual SlideKind slideKind(int idx) const = 0;     // Cover|Divider|Content|Image|Cumulative
    virtual ReportSlideSpec buildSlide(int idx, const ReportLayout&,
                                        const QSet<QString>& excludedSamples) const = 0;

    // Sample enumeration (for the checkbox panel)
    virtual QVector<SampleRef> allSamples() const = 0;  // {sessionId/sheetId, sampleId, displayName}

    // Persistence
    virtual ReportLayout loadLayout() const = 0;
    virtual void saveLayout(const ReportLayout&) = 0;

    // Final PPTX write
    virtual bool writePptx(const QString& outPath, const ReportLayout&,
                            const QSet<QString>& excludedSamples,
                            QString* errorOut) = 0;
};
```

Concrete implementations:

- `SensoryReportSource` — wraps one or more `SensorySession`
- `TpmReportSource` — wraps one `FileResult` (Test or Full report)
- `CombinedTpmReportSource` — wraps multiple `FileResult` (Combined Full)
- `DetailedSensoryReportSource` — wraps one or more `DetailedSensorySession`

Each source class is the only place that knows mode-specific schema and
file paths. The dialog talks to the interface only.

### Slide types and editability

| Slide kind | Editable elements |
|---|---|
| Cover | Title text, subtitle/date text |
| Section divider | Title text |
| Content (per session/sheet) | Title text, table, plot(s), legend |
| Image (per session/sheet) | Image positions, sizes, crops |
| Cumulative summary | Title text, table, plot(s), legend |

Logos, colour theme, background graphics on cover/divider slides are
**not** editable in v1 — they remain template-driven so the brand stays
consistent. v2 candidate: full template-edit mode.

### Editable canvas items

Each editable element is a subclass of `ResizableSlideItem` (new),
generalising `ImageViewDialog::ResizableImageItem`:

| Item | Aspect ratio | Notes |
|---|---|---|
| `PlotItem` | Locked (1:1 for radar; 4:3 for line/bar) | Re-renders pixmap on resize |
| `TableItem` | Free | Column headers clickable for sort |
| `LegendItem` | Free | Auto-wraps swatches to width |
| `TextItem` | Free | Double-click to edit; supports basic formatting (bold/italic) |
| `ImageItem` (existing) | Locked | Crop mode preserved |

Drag/resize/move from `ResizableImageItem`. Snap-to-grid + alignment
guides layer on top (see below).

### Sample checkbox panel

Below the thumbnail strip; scrollable. Grouped by session/sheet, one
`QCheckBox` per sample, defaults **all checked**.

Toggling re-renders affected slides immediately:

- Content slides drop the row from the table and the polygon/line/bar
  from the plot.
- Image slides are unaffected (images aren't tied to specific samples).
- Cumulative slides recompute: a sample appears in cumulative IFF at
  least one of its host sessions/sheets has it checked.

Excluded set lives **outside** the layout JSON (it's per-report-staging,
not persisted). Holds in `ReportPreviewDialog::m_excludedSamples` for the
duration of the dialog.

### Sortable column headers

Click sequence on a column header in any `TableItem`:

1. First click → sort descending by that column
2. Second click → sort ascending
3. Third click → revert to insertion order

Sort persists into the layout JSON as `tableSort: { column, order }`.

### Snap-to-grid + alignment guides

- Grid spacing: 0.1" default; toggleable in the toolbar.
- Snap targets: grid lines + every other item's edges + slide centerlines.
- While dragging, magenta dashed lines show active guides; snap-to within
  6 px (image space) of a guide.
- Toggle state remembered in `QSettings` per user.

### Undo / redo

Command-pattern stack in the dialog (max 100 entries). Every mutation
goes through a `LayoutCommand`:

```cpp
class LayoutCommand {
public:
    virtual ~LayoutCommand() = default;
    virtual void apply(ReportLayout&) = 0;
    virtual void undo(ReportLayout&) = 0;
    virtual QString description() const = 0;
};

class MoveItemCommand    : public LayoutCommand;
class ResizeItemCommand  : public LayoutCommand;
class SortColumnCommand  : public LayoutCommand;
class EditTextCommand    : public LayoutCommand;
class ToggleSampleCommand: public LayoutCommand;  // (excluded set, not layout JSON)
class ZOrderCommand      : public LayoutCommand;
```

Bindings: `Ctrl+Z` undo, `Ctrl+Y` / `Ctrl+Shift+Z` redo. Stack is
in-memory only — closing the dialog discards it.

### Presets / templates

New DB table:

```sql
CREATE TABLE layout_presets (
    id          INTEGER PRIMARY KEY,
    mode_id     TEXT NOT NULL,           -- 'sensory' | 'tpm' | 'detailed_sensory'
    name        TEXT NOT NULL,
    layout_json TEXT NOT NULL,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(mode_id, name)
);
```

Toolbar UI:

- **Preset dropdown** — lists presets for the current mode
- **Save as Preset…** — prompts for name, writes current layout
- **Manage Presets…** — small dialog to delete/rename

Presets store slide layouts, default sort, and Z-order. They do **not**
store excluded samples or image-specific layouts (those are
per-source-specific; presets are reusable across sources).

### Layout import / export

- **Export Layout…** — file dialog → writes JSON file with the current
  layout + a `mode` field for cross-validation on import
- **Import Layout…** — file dialog → validates schema → applies to
  current source

Import refuses cross-mode files (e.g., applying a TPM layout to a sensory
source). Same JSON shape as the DB column.

### Persistence

#### Per-source layouts

Layout JSON is anchored to the data source row in the DB and the file in
disk:

| Mode | DB anchor | New column | Excel custom property |
|---|---|---|---|
| Sensory | `sensory_sessions(id)` | `layout_json TEXT` | `dve_layout` |
| TPM (Test Report) | `sheets(id)` | `layout_json TEXT` | `dve_layout_<sheetName>` |
| TPM (Full Report) | `files(id)` | `layout_json TEXT` (whole-file layout) | `dve_layout` |
| TPM (Combined Full) | `settings` table | key `tpm.combined_layout` | — |
| Detailed Sensory | `detailed_sensory_sessions(id)` | `layout_json TEXT` | `dve_layout` |

Excel custom-property keys are namespaced (`dve_layout` for whole-workbook
layout; `dve_layout_<sheetName>` if the same workbook hosts multiple
TPM sheets each with their own report).

JSON shape (per source):

```json
{
  "version": 1,
  "mode": "sensory",
  "tableSort": { "column": "Overall Liking", "order": "desc" },
  "slides": {
    "cover":     { "title": [x,y,w,h], "subtitle": [x,y,w,h] },
    "divider_<id>": { "title": [x,y,w,h] },
    "content_<sessionId>": {
      "title":   [x,y,w,h],
      "table":   [x,y,w,h],
      "plots":   [[x,y,w,h], ...],
      "legend":  [x,y,w,h]
    },
    "image_<sessionId>": {
      "imageLayouts": [[x,y,w,h], ...],
      "imageCrops":   [[x,y,w,h], ...]
    },
    "cumulative": {
      "title":   [x,y,w,h],
      "table":   [x,y,w,h],
      "plots":   [[x,y,w,h], ...],
      "legend":  [x,y,w,h]
    }
  },
  "zOrder": ["table", "legend", "plots[0]", "title"]
}
```

The existing per-sample `imageLayouts`/`imageCrops` storage is folded
into this JSON. Backwards compat: if `layout_json` is NULL, fall back to
legacy per-sample fields.

Coordinates in inches.

#### Cumulative layout (cross-source) for combined reports

Combined reports (Sensory cumulative across multiple sessions, TPM
Combined Full Report across files, Detailed Sensory cumulative) span
multiple sources, so per-source storage doesn't fit. Stored globally:

| Mode | `settings` key |
|---|---|
| Sensory cumulative | `sensory.cumulative_layout` |
| TPM combined | `tpm.combined_layout` |
| Detailed Sensory cumulative | `detailed_sensory.cumulative_layout` |

#### Write timing

- **Auto-save** on every edit, debounced 500ms (matches existing
  `m_excelWriteTimer` pattern in `MainWindow`).
- On dialog close, flush pending writes immediately before destruction.
- Auto-save: DB synchronously, Excel via the existing Python subprocess
  pipeline.

#### Conflict resolution

- **On load:** if both DB and Excel hold a layout for the source, take
  the one with the newer timestamp. Same last-write-wins convention used
  for sensory data.
- **On save:** DB first, then Excel write queued. Excel failure logged
  but doesn't abort the report — DB is authoritative.

## Code structure

### New files

- `src/ui/ReportPreviewDialog.{h,cpp}`
- `src/ui/SlideCanvasItems.{h,cpp}` — base `ResizableSlideItem` + subclasses
- `src/ui/SamplesCheckboxPanel.{h,cpp}`
- `src/ui/PropertiesPanel.{h,cpp}` — selected-item position/size/Z-order editor
- `src/ui/PresetManagerDialog.{h,cpp}`
- `src/reporting/IReportSource.h`
- `src/reporting/SensoryReportSource.{h,cpp}`
- `src/reporting/TpmReportSource.{h,cpp}`
- `src/reporting/CombinedTpmReportSource.{h,cpp}`
- `src/reporting/DetailedSensoryReportSource.{h,cpp}`
- `src/reporting/ReportLayout.{h,cpp}` — JSON model: serialize/deserialize, defaults, version migration
- `src/reporting/LayoutCommand.{h,cpp}` — undo/redo command base + concrete commands
- `src/reporting/PresetStore.{h,cpp}` — DB access for `layout_presets`

### Modified files

- `src/database/DatabaseManager.{h,cpp}` — schema migration adds:
  - `sensory_sessions.layout_json`
  - `sheets.layout_json`
  - `files.layout_json`
  - `detailed_sensory_sessions.layout_json`
  - `layout_presets` table
- `src/ExcelReader.{h,cpp}` and the writeback Python helper — read/write `dve_layout` custom property
- `src/ui/SensoryPanel.{h,cpp}` — `generateFullReport()` and `generateCombinedPptx()` route through `ReportPreviewDialog` via `SensoryReportSource`
- `src/ui/DetailedSensoryPanel.{h,cpp}` — same via `DetailedSensoryReportSource`
- `src/MainWindow.cpp` — TPM report buttons (`onGenerateTestReport`, `onGenerateFullReport`) route through `ReportPreviewDialog` via `TpmReportSource` / `CombinedTpmReportSource`
- `src/reporting/PptxWriter.{h,cpp}` — `addContentSlide(...)`, `addCoverSlide(...)`, `addSectionDividerSlide(...)` all accept layout overrides for element positions and editable text
- `src/reporting/ReportGenerator.{h,cpp}` — accepts a `ReportLayout` parameter; uses overrides where present, kPlotLayout defaults otherwise

## Data flow

```
User clicks any "Report" button
              │
              ▼
Mode-specific entry creates an IReportSource, opens ReportPreviewDialog
              │
              ▼
Dialog loads layout (DB > Excel > defaults)
Dialog populates thumbnails, canvas, sample checkboxes
              │
              ▼
User edits → LayoutCommand → ReportLayout updated → debounced auto-save
                                                  → IReportSource.saveLayout()
                                                  → DB + Excel writeback
              │
              ▼
User clicks "Create Report" → IReportSource.writePptx(outPath, layout, excluded)
                            → existing PptxWriter with overrides applied
              │
              ▼
Dialog closes; existing onReportFinished plumbing handles success/failure
```

## Testing

New tests:

- `tests/tst_reportlayout/` — round-trip JSON serialization, version migration, defaults computation, cross-mode rejection on import
- `tests/tst_layoutcommand/` — apply/undo/redo invariants for every concrete command
- `tests/tst_presetstore/` — DB CRUD with the new table

Extended tests:

- `tests/tst_databasemanager/` — `layout_json` column persistence on each mode's table, `layout_presets` table, settings table for cumulative layouts, schema migration from pre-layout DB

Manual deployment self-test additions (`tests/deployment/Test-Deployment.ps1`):

- Open report preview for each mode, edit one element, close, reopen, verify layout restored.
- Save a preset, change layout, re-apply preset, verify restored.
- Export a layout JSON, import it on a fresh source, verify applied.

## Risks & mitigations

| Risk | Mitigation |
|------|-----------|
| Three-mode scope is 3× sensory-only | Polymorphic `IReportSource` keeps per-mode code isolated; phase the build (see below) |
| Excel custom-property write fails on Synology share | DB authoritative; log warning; report still generates |
| Layout JSON shape drift across releases | `version` field; migration helpers; never delete unknown fields on load |
| Plot rerender on resize is slow | Render at low DPI in canvas (~96), high DPI only for final PPTX |
| Cover/divider title editing risks brand inconsistency | Title text only — logos and template assets stay locked |
| Undo stack memory grows unboundedly | Cap at 100 commands |
| Preset name collisions in shared DB | Allow overwrite-on-save with confirmation prompt |
| Cross-mode layout import causes confusion | Reject import if `mode` field doesn't match current source |
| User unchecks all samples → empty content slide | Inline warning on the slide; prompt before "Create Report" |
| Stale layout when source data changes (sample renamed/deleted) | Layout references samples by name; missing names ignored, new samples default to checked |

## Recommended phasing

The implementation plan (`writing-plans` skill output, next step) will
break the work into phases that each deliver a shippable improvement, so
we can stop early or cut a phase if it proves unnecessary in practice:

1. **Sensory minimum** — `IReportSource` interface, `SensoryReportSource`,
   `ReportPreviewDialog` shell with canvas + drag/resize for table+plot+
   legend+title, sample checkboxes, sort, DB-only persistence on
   `sensory_sessions.layout_json`. Sensory cumulative layout in settings.
2. **Excel custom-property round-trip** for sensory.
3. **Snap-to-grid + alignment guides.**
4. **Undo/redo.**
5. **Presets** (DB table + UI).
6. **Layout JSON import/export.**
7. **TPM port** — `TpmReportSource`, `CombinedTpmReportSource`, schema
   migration on `sheets.layout_json` + `files.layout_json` + Excel
   custom property; route TPM ribbon buttons through the dialog.
8. **Detailed-sensory port** — `DetailedSensoryReportSource`, schema
   migration on `detailed_sensory_sessions.layout_json` + Excel.
9. **Cover/divider title-text editing** across all modes.

Phases 1–6 give sensory full feature parity. Phases 7–9 generalise.
Each phase is independently revertible — if (5) feels unnecessary in
practice, we drop it without disturbing 6+.
