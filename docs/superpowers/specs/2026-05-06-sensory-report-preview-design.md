# Sensory Report Preview — Design

- **Date:** 2026-05-06
- **Status:** Draft (pending user review of revised scope)
- **Scope (v1):** Sensory mode only. TPM is out of scope entirely. Detailed
  Sensory follows in a small follow-on project (architecturally pre-wired
  here so it's mostly an adapter add).

## Goal

Add a WYSIWYG preview/editor that opens when the user clicks any sensory
"Report" button. Lets the user:

1. See every body slide in the staged report exactly as it'll render
   (defaults match the current report layout exactly — title on top,
   table full-width below, radar chart centered below the table,
   properties textbox bottom-right).
2. Switch between slides via a thumbnail strip.
3. Drag/resize the table, radar chart, title, and properties textbox
   on each slide.
4. Sort the table by clicking column headers.
5. Toggle samples in/out of the report via a checkbox panel (does not
   modify underlying data).
6. Edit cover and group-divider slides — title text only; logos, colour
   theme, and template assets stay locked.
7. Edit cumulative summary slides the same as content slides.
8. Use snap-to-grid and alignment guides while dragging.
9. Undo/redo any edit.
10. Save the current layout as a named preset, apply a saved preset,
    export the layout as a JSON file, import one from a JSON file.

The preview state IS the report — clicking "Create Report" generates a
PPTX matching whatever's on screen.

Layout edits **persist** across sessions per data anchor (the file or
the DB row that produced the slides). Reopening the same source shows
the previous layout.

## Non-goals (v1)

- TPM mode. Out of scope entirely. TPM data hierarchy (file → sheet →
  sample → row) and the existing TPM report flow stay untouched.
- Detailed Sensory mode. Post-v1, plugged in via a second adapter; the
  dialog and canvas are written generic enough to host it without
  changes.
- Real-time collaborative editing.
- Version history per layout (last-write-wins).
- Diff view between layouts.
- Editing logos / colour theme / template assets (only title text on
  cover/divider slides).

## Architecture

### Modal QDialog: `ReportPreviewDialog`

Replaces the current direct-to-PPTX flow for sensory reports. Mirrors
`ImageViewDialog`'s `QGraphicsView`-based pattern, scaled to handle a
multi-slide report. The dialog is **mode-agnostic** — it talks to the
data through `IReportSource`. v1 ships only `SensoryReportSource`.

```
┌─ ReportPreviewDialog ─────────────────────────────────────────────────────────────────┐
│ ┌────────────┐ ┌──────────────────────────────────────────┐ ┌─────────────────────┐  │
│ │ Slide      │ │  QGraphicsView canvas                    │ │ Properties panel    │  │
│ │ thumbs     │ │  800 × 450 px (= 13.33"×7.5")            │ │ ──────────────      │  │
│ │ ──────     │ │                                          │ │ Selected: Table     │  │
│ │ □ Cover    │ │  Title                                   │ │ x: 0.32"  y: 0.75"  │  │
│ │ □ Section1 │ │ ────────────────────────────────────     │ │ w: 12.70" h: auto   │  │
│ │ □ S1 Body  │ │ │  Table  (full width, sortable)    │    │ │ [Bring Forward]     │  │
│ │ □ S1 Imgs  │ │ ────────────────────────────────────     │ │ [Send Backward]     │  │
│ │ □ Section2 │ │                                          │ │                     │  │
│ │ □ S2 Body  │ │         ┌──────────────┐                 │ │ Sort: Overall ▾ desc│  │
│ │ □ Cumul.   │ │         │              │ ┌─────────┐     │ │                     │  │
│ │            │ │         │ Radar chart  │ │ Props   │     │ │                     │  │
│ │            │ │         │  (centered)  │ │ textbox │     │ │                     │  │
│ │            │ │         │              │ └─────────┘     │ │                     │  │
│ │            │ │         └──────────────┘                 │ │                     │  │
│ ├────────────┤ │                                          │ ├─────────────────────┤  │
│ │ Samples    │ │  Alignment guides + snap-to-grid live    │ │ Toolbar             │  │
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

Default layout matches the current sensory report **exactly** — title on
top, table full-width below the title, radar chart centered below the
table filling the remaining vertical space (aspect-locked, square),
properties textbox in the bottom-right corner. The preview just makes
those elements moveable/resizable; nothing about the default layout
changes.

### Mode adapter: `IReportSource`

Mode-specific differences are encapsulated behind an interface so the
dialog never touches mode-specific schema or PPTX-writer arguments.
v1 implements one concrete: `SensoryReportSource`.

```cpp
class IReportSource {
public:
    virtual ~IReportSource() = default;

    // Identity
    virtual QString modeId() const = 0;          // "sensory" (only mode in v1)
    virtual QString sourceLabel() const = 0;     // shown in dialog title bar

    // Slides
    virtual int slideCount() const = 0;
    virtual SlideKind slideKind(int idx) const = 0;     // Cover|Divider|Content|Image|Cumulative
    virtual ReportSlideSpec buildSlide(int idx, const ReportLayout&,
                                        const QSet<QString>& excludedSamples) const = 0;

    // Sample enumeration (for the checkbox panel)
    virtual QVector<SampleRef> allSamples() const = 0;  // {sessionId, sampleId, displayName}

    // Persistence
    virtual ReportLayout loadLayout() const = 0;
    virtual void saveLayout(const ReportLayout&) = 0;

    // Final PPTX write
    virtual bool writePptx(const QString& outPath, const ReportLayout&,
                            const QSet<QString>& excludedSamples,
                            QString* errorOut) = 0;
};
```

`SensoryReportSource` wraps one or more `SensorySession` (single-tester
report or cumulative across testers).

### Slide types and editability

| Slide kind | Editable elements |
|---|---|
| Cover | Title text, subtitle/date text |
| Section divider | Title text |
| Content (per session) | Title, Table, Radar chart, Properties textbox |
| Image (per session) | Image positions, sizes, crops |
| Cumulative summary | Title, Table, Radar chart, Properties textbox |

Notes:
- Sensory radar charts have their **legend baked into the chart pixmap**
  (rendered by `RadarChartWidget` then exported as a single PNG). There's
  no separate `LegendItem` — moving the chart moves its legend.
- The **Properties textbox** is the existing white-background callout
  carrying Media / Control / Blind? / Primary Difference(s) / Highest
  Rated / Lowest Rated lines (see `SensoryPanel::generateCombinedPptx`,
  the `extraXml` block). It's a `TextItem` with multi-line support.
- Logos, colour theme, background graphics on cover/divider slides are
  **not** editable — they remain template-driven so the brand stays
  consistent.

### Editable canvas items

Each editable element is a subclass of `ResizableSlideItem` (new),
generalising `ImageViewDialog::ResizableImageItem`:

| Item | Aspect ratio | Notes |
|---|---|---|
| `PlotItem` | Locked (1:1 for radar; legend baked in) | Re-renders pixmap on resize |
| `TableItem` | Free | Column headers clickable for sort |
| `TextItem` | Free | Double-click to edit. Used for slide titles, cover subtitles, and the properties textbox |
| `ImageItem` (existing) | Locked | Crop mode preserved |

Drag/resize/move from `ResizableImageItem`. Snap-to-grid + alignment
guides layer on top.

No `LegendItem` in v1 — the existing sensory chart pipeline renders the
legend inside the radar pixmap. If we ever decide to surface the legend
as a movable element, that's a chart-pipeline change separate from this
preview.

### Sample checkbox panel

Below the thumbnail strip; scrollable. Grouped by session, one
`QCheckBox` per sample, defaults **all checked**.

Toggling re-renders affected slides immediately:

- Content slides drop the row from the table and the polygon from the
  radar plot.
- Image slides are unaffected (images aren't tied to specific samples).
- Cumulative slides recompute: a sample appears in cumulative IFF at
  least one of its host sessions has it checked.

Excluded set lives **outside** the layout JSON (it's per-report-staging,
not persisted). Holds in `ReportPreviewDialog::m_excludedSamples` for
the duration of the dialog.

### Sortable column headers

Click sequence on a column header in any `TableItem`:

1. First click → sort descending by that column
2. Second click → sort ascending
3. Third click → revert to insertion order

Sort persists into the layout JSON as `tableSort: { column, order }`.

### Snap-to-grid + alignment guides

- Grid spacing: 0.1" default; toggleable in the toolbar.
- Snap targets: grid lines + every other item's edges + slide
  centerlines.
- While dragging, magenta dashed lines show active guides; snap-to
  within 6 px (image space) of a guide.
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
    mode_id     TEXT NOT NULL,           -- 'sensory' (only mode in v1)
    name        TEXT NOT NULL,
    layout_json TEXT NOT NULL,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(mode_id, name)
);
```

The `mode_id` column is forward-compatible — when detailed-sensory is
added later, presets stay scoped to their mode without schema change.

Toolbar UI:

- **Preset dropdown** — lists presets for the current mode
- **Save as Preset…** — prompts for name, writes current layout
- **Manage Presets…** — small dialog to delete/rename

Presets store slide layouts, default sort, and Z-order. They do **not**
store excluded samples or image-specific layouts (those are per-source;
presets are reusable across sources).

### Layout import / export

- **Export Layout…** — file dialog → writes JSON file with the current
  layout + a `mode` field for cross-validation on import
- **Import Layout…** — file dialog → validates schema → applies to
  current source

Import refuses files whose `mode` doesn't match the current source.
Same JSON shape as the DB column.

### Persistence

#### Per-session layouts

Layout JSON is anchored to the `SensorySession` row in the DB AND to
the source `.xlsx` file:

| Storage | Location |
|---|---|
| DB | new column `sensory_sessions.layout_json TEXT` (nullable; NULL = use defaults) |
| Excel | workbook custom property `dve_layout` (JSON string) |

JSON shape:

```json
{
  "version": 1,
  "mode": "sensory",
  "tableSort": { "column": "Overall Liking", "order": "desc" },
  "slides": {
    "cover":     { "title": [x,y,w,h], "subtitle": [x,y,w,h] },
    "divider_<sessionId>": { "title": [x,y,w,h] },
    "content_<sessionId>": {
      "title":         [x,y,w,h],
      "table":         [x,y,w,h],
      "radar":         [x,y,w,h],
      "propertiesBox": { "rect": [x,y,w,h], "text": "Media: …\nControl: …" }
    },
    "image_<sessionId>": {
      "imageLayouts": [[x,y,w,h], ...],
      "imageCrops":   [[x,y,w,h], ...]
    },
    "cumulative": {
      "title":         [x,y,w,h],
      "table":         [x,y,w,h],
      "radar":         [x,y,w,h],
      "propertiesBox": { "rect": [x,y,w,h], "text": "…" }
    }
  },
  "zOrder": ["table", "radar", "propertiesBox", "title"]
}
```

The existing per-sample `imageLayouts` / `imageCrops` fields on
`SensorySample` are folded into this JSON. Backwards compat: if
`layout_json` is NULL, fall back to the legacy per-sample fields.

Coordinates in inches.

#### Defaults

When `layout_json` is NULL (fresh session, no preview state yet), the
`SensoryReportSource::loadLayout()` method computes defaults that match
the current report's positioning logic **exactly** — table at
`(0.32, 0.75)` with `w=12.7"` and computed height; radar centered
horizontally below the table filling the remaining vertical space with
0.10" gaps top and bottom (square aspect-locked); properties textbox
anchored bottom-right at `(slideW - tbW - 0.05, slideH - tbH - 0.05)`
with dynamic height from line count.

This is captured by porting the existing positioning math from
`SensoryPanel::generateCombinedPptx` (the block around the
`tableBottom` / `chartY` / `chartH` calculations) into a single
`SensoryReportSource::computeDefaultLayout(const SensorySession&)`
function that the dialog calls when no saved layout exists. The
`generateCombinedPptx` code path then becomes a thin wrapper that calls
the same default-computation function and feeds the result to the
PPTX writer — guaranteeing default-vs-no-preview parity.

#### Cumulative layout (cross-session)

A cumulative report spans multiple `SensorySession`s, so per-session
storage doesn't fit the cumulative slide. Stored globally:

| Storage | Location |
|---|---|
| DB | `settings` table, key `sensory.cumulative_layout`, value = JSON |
| Excel | not stored (cumulative reports span multiple files; no natural home) |

Stored as a single global preference. The DB is authoritative.

#### Write timing

- **Auto-save** on every edit, debounced 500ms (matches existing
  `MainWindow::m_excelWriteTimer` pattern).
- On dialog close, flush pending writes immediately before destruction.
- Auto-save: DB synchronously, Excel via the existing Python subprocess
  pipeline.

#### Conflict resolution

- **On load:** if both DB and Excel hold a layout for the session, take
  the one with the newer timestamp. Same last-write-wins convention
  used elsewhere for sensory data.
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
- `src/reporting/ReportLayout.{h,cpp}` — JSON model: serialize/deserialize, defaults, version migration
- `src/reporting/LayoutCommand.{h,cpp}` — undo/redo command base + concrete commands
- `src/reporting/PresetStore.{h,cpp}` — DB access for `layout_presets`

### Modified files

- `src/database/DatabaseManager.{h,cpp}` — schema migration adds:
  - `sensory_sessions.layout_json`
  - `layout_presets` table
- `src/ExcelReader.{h,cpp}` and the writeback Python helper — read/write
  `dve_layout` custom property
- `src/ui/SensoryPanel.{h,cpp}` — `generateFullReport()` and
  `generateCombinedPptx()` route through `ReportPreviewDialog` via
  `SensoryReportSource` (instead of going straight to PPTX)
- `src/reporting/PptxWriter.{h,cpp}` — `addContentSlide(...)`,
  `addCoverSlide(...)`, `addSectionDividerSlide(...)` accept layout
  overrides for element positions and editable text
- `src/reporting/ReportGenerator.{h,cpp}` — accepts a `ReportLayout`
  parameter; uses overrides where present, kPlotLayout defaults otherwise

(`MainWindow.cpp` is **not** touched — TPM report flow stays as-is.)

## Data flow

```
User clicks a sensory "Report" button
              │
              ▼
SensoryPanel creates SensoryReportSource, opens ReportPreviewDialog
              │
              ▼
Dialog loads layout (DB > Excel > defaults via SensoryReportSource)
Dialog populates thumbnails, canvas, sample checkboxes
              │
              ▼
User edits → LayoutCommand → ReportLayout updated → debounced auto-save
                                                  → SensoryReportSource.saveLayout()
                                                  → DB + Excel writeback
              │
              ▼
User clicks "Create Report" → SensoryReportSource.writePptx(outPath, layout, excluded)
                            → existing PptxWriter with overrides applied
              │
              ▼
Dialog closes; existing onReportFinished plumbing handles success/failure
```

## Testing

New tests:

- `tests/tst_reportlayout/` — round-trip JSON serialization, version
  migration, defaults computation, cross-mode rejection on import
- `tests/tst_layoutcommand/` — apply/undo/redo invariants for every
  concrete command
- `tests/tst_presetstore/` — DB CRUD with the new table

Extended tests:

- `tests/tst_databasemanager/` — `sensory_sessions.layout_json` column
  persistence, `layout_presets` table, settings table for cumulative
  layout, schema migration from pre-layout DB

Manual deployment self-test additions
(`tests/deployment/Test-Deployment.ps1`):

- Open sensory report preview, edit one element, close, reopen,
  verify layout restored.
- Save a preset, change layout, re-apply preset, verify restored.
- Export a layout JSON, import it on a fresh source, verify applied.

## Risks & mitigations

| Risk | Mitigation |
|---|---|
| Excel custom-property write fails on Synology share | DB authoritative; log warning; report still generates |
| Layout JSON shape drift across releases | `version` field; migration helpers; never delete unknown fields on load |
| Plot rerender on resize is slow | Render at low DPI in canvas (~96), high DPI only for final PPTX |
| Cover/divider title editing risks brand inconsistency | Title text only — logos and template assets stay locked |
| Undo stack memory grows unboundedly | Cap at 100 commands |
| Preset name collisions | Allow overwrite-on-save with confirmation prompt |
| Cross-mode layout import causes confusion | Reject import if `mode` field doesn't match current source |
| User unchecks all samples → empty content slide | Inline warning on the slide; prompt before "Create Report" |
| Stale layout when session data changes (sample renamed/deleted) | Layout references samples by name; missing names ignored, new samples default to checked |

## Recommended phasing

The implementation plan (`writing-plans` skill output, next step) will
break the work into shippable phases so we can stop early or cut a
phase if it proves unnecessary in practice:

1. **Sensory minimum** — `IReportSource` interface, `SensoryReportSource`,
   `ReportPreviewDialog` shell with canvas + drag/resize for table + plot
   + legend + title, sample checkboxes, sort, DB-only persistence on
   `sensory_sessions.layout_json` and `settings.sensory.cumulative_layout`.
2. **Excel custom-property round-trip** for sensory.
3. **Snap-to-grid + alignment guides.**
4. **Undo/redo.**
5. **Presets** (DB table + UI).
6. **Layout JSON import/export.**
7. **Cover/divider title-text editing.**

Each phase is independently revertible — if (5) feels unnecessary in
practice, we drop it without disturbing 6–7.

## Out of scope (this project)

- TPM mode report preview. The TPM data hierarchy and report flow stay
  exactly as they are today. If TPM ever needs a preview, it'll be a
  separate project that can reuse `ReportPreviewDialog`,
  `SlideCanvasItems`, `ReportLayout`, `LayoutCommand`, and `PresetStore`
  by writing a `TpmReportSource` adapter.

## Follow-on (small project after v1 ships)

- **Detailed Sensory port.** Sensory and Detailed Sensory share most of
  their data shape. Implementation: write `DetailedSensoryReportSource`,
  add `detailed_sensory_sessions.layout_json` column, route the Detailed
  Sensory report buttons through `ReportPreviewDialog`. No dialog or
  canvas changes expected. Estimated 1–2 days once sensory is shipped.
