# DataViewer Enterprise — UI Polish Design

**Date:** 2026-05-21
**Status:** Approved (sections 1–3)
**Target:** v2.0.9
**Branch:** feature/ui-polish-v2.0.9 (to be cut off main after v2.0.8 ships)

## Goals

Bring the UI to a "modern industrial/engineering" polish level (Autodesk Fusion / SolidWorks / JetBrains direction) without changing the underlying widget paradigm. Three concrete user-visible goals:

1. **Visual polish** — refined spacing, alignment, hover/focus states; lighter overall feel.
2. **Readability** — black text by default; the dark-blue status bar that masks colored text gets replaced.
3. **Responsive layout** — when the window snaps to half a screen via `Win+→` / `Win+←`, the UI reflows logically and stays usable.

## Non-goals (out of scope this session)

- **Plot / chart styling.** TPM plot colors, sensory radar chart rendering — handled in a separate session per user's stated preference.
- **Dark mode.** Stay light-only.
- **Custom keyboard shortcuts.** Rely on Windows' native `Win+→/←/↑/↓` snapping; no in-app hotkeys for window placement.
- **Major UX rework.** No reorganizing widget hierarchies. Layout polish only (alignment, padding, grouping).
- **Reordering ribbon buttons.** Already done as a separate v2.0.8 change.

## Scope summary

| Area | In | Out |
|---|---|---|
| 6 main screens (TPM, Sensory, Detailed Sensory, Database Browser, + Home/Reports/Tools workspace chrome) | ✅ | |
| Modal dialogs (NewFile, HeaderEdit, DataCleanup, Sop, ImageInbox, ImageView, ReportPreview) | ✅ | |
| Ribbon icons replaced (Qt standard pixmaps → Lucide SVG) | ✅ | |
| Status bar redesign (light bg, breadcrumb separators, semantic dots) | ✅ | |
| Responsive layout (compact mode at < 1100px) | ✅ | |
| Plot / radar chart styling | | ✅ excluded |
| Custom window-snap hotkeys | | ✅ excluded |
| Dark mode | | ✅ excluded |

## Phase 1 — Design System Foundation

**Agent F** owns this phase. Sequential — Phase 2 cannot begin until Phase 1 is merged.

### Files owned

- `src/utils/AppTheme.h` / `AppTheme.cpp` — refactor to named tokens
- `src/widgets/RibbonWidget.h` / `RibbonWidget.cpp` — icons-only compact mode, icon-loading helper
- `src/MainWindow.cpp` — **only** the status bar setup code and the `ResponsiveLayout` instantiation in the constructor. All other MainWindow polish is Agent A's territory.
- `src/utils/ResponsiveLayout.h` / `ResponsiveLayout.cpp` — **new file**: breakpoint detection + signal
- `resources/icons/*.svg` — **new directory**: bundled Lucide SVGs (≈16 files)

> **File-sharing rule with Agent A:** Phase 1 must commit + merge to the feature branch before Phase 2 starts. Agent A reads the post-Phase 1 state and never edits MainWindow.cpp in parallel with Phase 1. If Agent A finds itself wanting to change Phase 1's status-bar code, it must stop and surface the conflict.

### Color tokens

Replace the ad-hoc hex values in `AppTheme` with named accessors. Both the C++ accessors (`AppTheme::surfaceApp()`) and the QSS string need to reference the same source of truth (e.g., generate QSS from token map, or define hex constants once and use in both places).

| Token | Current | New | Notes |
|---|---|---|---|
| `surface-app` | `#F0F0F0` | `#F5F6F8` | slightly lighter |
| `surface-panel` | `#FFFFFF` | `#FFFFFF` | unchanged |
| `surface-statusbar` | `#1F4E79` | `#ECEEF1` | **major change — fixes readability** |
| `border-subtle` | `#DCDCDC` | `#E4E6EA` | softer divider |
| `border-default` | `#BCBCBC` | `#CFD3D8` | cooler gray |
| `accent` | `#0066CC` | `#0066CC` | kept identical |
| `accent-subtle` | (none) | `#E8F2FC` | new — for tinted backgrounds |
| `text-primary` | `#1A1A1A` | `#1A1D21` | slightly cooler |
| `text-secondary` | `#555555` | `#5C636A` | — |
| `table-header` | `#1F4E79` | `#2C3E50` | flatter slate, less navy-dominant |
| `success` | `#28A745` | `#16A34A` | — |
| `warning` | `#ED8B00` | `#D97706` | — |

### Other token scales

- **Spacing:** `4 · 8 · 12 · 16 · 24 · 32 px` — expose as `AppTheme::space(n)` where n is the scale index, or as named constants (`SPACE_SM`, `SPACE_MD`, …).
- **Radius:** `4px` (controls), `6px` (cards/panels), `8px` (dialogs).
- **Type:** Segoe UI — `9pt body` · `8pt small` · `10pt label` · `11pt section` · `12pt page-title`. Add `AppTheme::fontLabel()` and `AppTheme::fontPageTitle()`.
- **Elevation:** dialog shadow — `box-shadow: 0 4px 16px rgba(0,0,0,0.08);` (apply via `setGraphicsEffect(QGraphicsDropShadowEffect)` since Qt QSS doesn't honor `box-shadow`).

### Status bar redesign

Replace the current `QStatusBar` background-color + white text approach with a custom widget arrangement:

- Background: `surface-statusbar` (`#ECEEF1`), 1px top border in `border-default`.
- Text: `text-primary` (`#1A1D21`) by default.
- Status dot: `●` colored by state — `success` green for clean/online, `warning` orange for modified, `error` red for disconnected.
- Segments separated by `|` rendered in `border-default` color.
- Breadcrumb in the right segment uses `›` (U+203A) instead of `|` to differentiate path from status.

Specifically, the current static format `"File closed. Local DB: sensory modified (Ctrl+U) | file.xlsx | sheet | sample"` becomes a structured arrangement:

```
[●] File status                | [●] Local DB status                | filename › sheet › sample
```

with the indicator dots colored semantically and the breadcrumb's middle segments truncatable in compact mode.

### Lucide icon bundle

Download these ~16 SVGs from https://lucide.dev (MIT licensed) into `resources/icons/`:

| File | Lucide name | Used by |
|---|---|---|
| `file-plus.svg` | file-plus | New File, New Session |
| `folder-open.svg` | folder-open | Load File, Load Excel |
| `x.svg` | x | Close |
| `info.svg` | info | SOPs |
| `database.svg` | database | Database |
| `refresh-cw.svg` | refresh-cw | Refresh Snapshot |
| `sparkles.svg` | sparkles | Sensory |
| `list-checks.svg` | list-checks | Detailed Sensory |
| `image.svg` | image | Images |
| `file-text.svg` | file-text | Test Report |
| `files.svg` | files | Full Report |
| `eraser.svg` | eraser | Clean Data |
| `rotate-ccw.svg` | rotate-ccw | Reset Cleanup |
| `languages.svg` | languages | Translator |
| `save.svg` | save | Save (sensory modes) |
| `menu.svg` | menu | Compact mode sidebar toggle |

`AppTheme::icon(QString name)` returns a `QIcon` loaded from `resourcePath()/icons/<name>.svg`. Use `QIcon::setIsMask(true)` or tint via `QPixmap` recoloring so the icons inherit `text-primary` and re-color on `:hover` / `:checked` states.

Resources are loaded from disk (not Qt resource system) per the project's `RESOURCES` workaround in `.pro` (see CLAUDE.md).

### Done when

1. App builds clean (`-Werror -Wall -Wextra`).
2. All 34 Qt tests pass (`tests\run-tests.ps1`).
3. Every screen renders without missing styles or broken visuals — verified by launching the app and clicking through each mode (TPM, Sensory, Detailed Sensory, Database Browser, each modal dialog).
4. Status bar reads correctly with new text colors at all states (clean / modified / disconnected / multi-segment).
5. Ribbon shows new icons in place of generic Windows icons.

## Phase 2 — Per-screen polish (4 parallel agents)

Each agent owns disjoint files. No coordination needed during execution. Each agent's work is independently mergeable.

### Agent A — Main workspace + chrome

**Files owned:**
- `src/MainWindow.cpp` (non-ribbon UI code only — ribbon already handled in Phase 1)
- `src/ui/PropertiesPanel.cpp` / `PropertiesPanel.h`

**Polish targets:**
- TPM data table padding/alignment, header column widths, row height tightening.
- Plot control row (`Plot Type:` combo + save button): consistent spacing, button styling.
- Navigator panel (file tree): spacing, indent rhythm, selected-state visual weight.
- Sheet selector header ("No file loaded" with prev/next arrows): better visual treatment as a single composite control.
- "Add Row" / "Remove Row" buttons: align to table, primary/secondary distinction.
- Image dock bar at bottom: align Load Images / View Images buttons, lighter visual weight.
- Sample Properties panel: clearer key/value alignment, section heading hierarchy (`Session Info`, `Sample Properties` etc. as proper section labels).

**Responsive additions:**
- Sidebar collapse-to-32px-icon-strip below 1100px.
- Pop-out overlay panel when user clicks an icon in the strip; dismiss on outside-click or repeat click.
- Status bar middle breadcrumb segments truncate with `…` below 1100px (keep first + last visible).
- TPM plot panel: maintain minimum height of 200px for readability; allow taller-than-wide aspect at narrow widths.

### Agent B — Sensory workspace

**Files owned:**
- `src/ui/SensoryPanel.h` / `SensoryPanel.cpp`

**Polish targets:**
- Sample cards: verify V/R/HT/P fields line up vertically across cards; normalize any drift with a fixed-width label column.
- Field grouping: V/R/HT/P/PT in one logical row; Burnt Taste / Vapor Volume / Overall Flavor / Smoothness / Overall Liking in a second group.
- "Save Test Headers" button: looks too much like a label currently — give it proper button affordance + state (disabled when nothing to save, primary styling when ready).
- Test Averages panel: cleaner readout with proper section label.
- "+ Add Sample" button: matches new design system, primary styling.
- **Do NOT touch** the radar chart rendering or colors.

**Responsive additions:**
- Sample card grid: 3-up (standard) → 2-up below 1100px → 1-up below 700px.
- Radar chart container: maintains 1:1 aspect, minimum 240px square; below that, hide legend first, then chart.

### Agent C — Detailed Sensory workspace

**Files owned:**
- `src/ui/DetailedSensoryPanel.h` / `DetailedSensoryPanel.cpp`

**Polish targets:**
- 14-question grid: consistent vertical rhythm, column widths aligned across rows, label/control gap normalized.
- Sample selector dropdown: visually more prominent — it's the primary navigation control inside this panel.
- Comments field: proper textarea height (currently looks empty/floating), placeholder/hint text.
- Header bar (Test Title / Assessor / Tester / Media / Date / Save / Add Sample / Remove): more even spacing, consistent button styling.
- **Do NOT touch** the dual radar chart rendering or colors.

**Responsive additions:**
- 14-question grid: 2-column form → 1-column below 800px.
- Dual radar charts: side-by-side at standard width → stacked vertically below 1000px so each chart gets enough room.

### Agent D — Database Browser

**Files owned:**
- `src/ui/DatabaseBrowserDialog.h` / `DatabaseBrowserDialog.cpp`

**Polish targets:**
- Tab styling (TPM Data / Sensory Data / Detailed Sensory Data): match new accent-blue underline pattern from QSS.
- Search bar: bigger input, clearer placeholder treatment, optional clear-button icon.
- File table: alternating rows tinted with `accent-subtle`, hover state, sort indicator on `Loaded At`, row icons by template type ("new" vs "old").
- Bottom action bar: visually separated from table (subtle top border + lighter background), button hierarchy clarified:
  - **Load Selected** (primary) — `accent` color
  - **Load All Visible** (secondary) — default
  - **Delete Selected** (destructive) — red-tinted hover/focus
  - **Cleanup Duplicates** (secondary) — default
- Footer "Close" button: keep both X (top-right) and Close (bottom-right) since this is a modal — they serve different muscle memory.

**Responsive:** nothing — modal dialog, resizing the underlying main window doesn't affect it. Qt layouts handle in-dialog resize already.

## Phase 3 — Modal dialog sweep (1 agent, parallel with Phase 2)

**Agent M** owns this. Runs concurrently with Phase 2 agents.

**Files owned (each dialog's `.h` + `.cpp` pair):**
- `src/ui/NewFileDialog`
- `src/ui/HeaderEditDialog`
- `src/ui/DataCleanupDialog`
- `src/ui/SopDialog`
- `src/ui/ImageInboxDialog`
- `src/ui/ImageViewDialog`
- `src/ui/ReportPreviewDialog`

**Polish targets (apply to all):**
- Consistent 16px dialog padding.
- Drop shadow elevation (via `QGraphicsDropShadowEffect`, not QSS).
- Form-row 12px vertical rhythm.
- Label/control alignment (labels right-aligned where applicable).
- Button bar pattern: `[Cancel · Primary]` right-aligned at bottom.
- Inline error/warning rendering with `error` / `warning` semantic colors.

**Per-dialog one-offs as encountered.** Agent has authority to fix small visible issues without escalation, where "small" means: alignment fixes, padding adjustments, button text typos, missing tooltips, redundant separators, missing primary-action styling. Anything that changes behavior (signals, slots, dialog modality, save logic) is out of scope — surface it as a follow-up.

## Phase 4 — Verification

Sequential, after all Phase 2 + 3 agents complete:

1. `python tools/decrypt_via_copy.py --apply` (precondition).
2. Clean rebuild: `cd build && mingw32-make clean && mingw32-make -j8 release`.
3. Full test suite: `tests\run-tests.ps1` — must report 34/34 pass.
4. Manual visual smoke test:
   - Standard width (≥ 1100px): each of the 6 screens looks polished and consistent.
   - Compact width (snap with `Win+→`): each screen reflows correctly per the rules in Phase 2.
   - Toggle through TPM ↔ Sensory ↔ Detailed Sensory modes; verify ribbon transitions and panel switches work cleanly.
   - Open each modal dialog (New File, Load File → existing workbook, Database Browser, SOPs, etc.) — verify shadow, padding, button bar all match.
5. Self-test passes: `DataViewer.exe --self-test`.
6. Build installer: `build_installer.bat`.

## Dependency Graph

```
Phase 1                       Phase 2 + 3 (parallel)
─────────                     ───────────────────────
                              ┌─→  Agent A  (MainWindow + PropertiesPanel)
                              │
Agent F  ────────────────────▶├─→  Agent B  (SensoryPanel)
(Tokens, icons, status bar,   │
 ResponsiveLayout helper,     ├─→  Agent C  (DetailedSensoryPanel)
 ribbon icon swap)            │
                              ├─→  Agent D  (DatabaseBrowserDialog)
                              │
                              └─→  Agent M  (7 modal dialogs)
```

## Estimated effort

| Phase | Wall-clock | Notes |
|---|---|---|
| Phase 1 (Agent F) | 90-150 min | Sequential, on critical path |
| Phase 2 (max of A/B/C/D) | 60-120 min | Parallel; longest one wins |
| Phase 3 (Agent M) | 90 min | Parallel with Phase 2 |
| Phase 4 (verification) | 30-60 min | Includes installer build |
| **Total** | **~3-5 hours** of agent runtime | |

## Risk register

- **MIP encryption.** Source files may pick up MIP labels at rest. Each subagent should run `python tools/decrypt_via_copy.py --apply` before building. Phase 1's foundation work especially since it touches many files.
- **QSS specificity wars.** New tokens replacing inline-set styles in C++ code (e.g., `setStyleSheet()` calls in widget constructors) may collide. Phase 2 agents should remove per-widget `setStyleSheet()` calls when the new AppTheme rules cover them.
- **Icon sizing on HiDPI.** Lucide SVG icons must render crisply at both 100% and 150% DPI. Use Qt's HiDPI scaling support; size icons via `QSize(24, 24)` at logical pixels.
- **Status bar refactor breaks existing connections.** `MainWindow` has multiple `m_statusBar` updates throughout — Phase 1 must preserve all signal connections during the refactor.
- **Subagent merge conflicts.** Mitigated by strict file ownership. If an agent finds it needs to edit another agent's file, it must stop and surface the conflict rather than editing.
