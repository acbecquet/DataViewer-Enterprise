---
date: 2026-06-29
author: acbecquet (Charlie) / Claude Code (Opus 4.8)
topic: "v2.7.0 — Responsive UI overhaul (Approach C)"
status: design-approved
branch: feature/v2.7.0-responsive-ui (off v2.6.0 / fae7fea)
tags: [v2.7.0, ui, responsive, dpi, scroll, ribbon, accessibility]
---

# v2.7.0 — Responsive UI Overhaul (Approach C)

## Problem

DataViewer's UI assumes a large, fixed desktop. It hard-codes a `setMinimumSize(1280, 800)`
window floor, sizes ~50 widgets with `setFixedSize/Width/Height`, and leaves several top-level
regions (ribbon group row, plot top-bar) with no scroll fallback. The consequences:

- The window **cannot be corner-snapped** (quarter-screen ≈ 960×540 < the 1280×800 floor) and
  barely fits a half-screen split.
- Under **OS text-scaling (125/150/200%)** and high-DPI, point-size fonts scale but fixed-pixel
  widgets do **not**, so text clips and overlaps (e.g. the ribbon **"View Raw Data"** label wraps
  to 3 lines and spills out of its 76px-tall button into the Navigator).
- At any size below the content's natural extent, content is simply **cut off with no scrollbar**,
  so controls become unreachable.

## Goals / Success Criteria

1. The window shrinks to a **~480×360 floor at any aspect ratio** with **nothing inaccessible** —
   horizontal *and* vertical scrollbars appear **as-needed** wherever content exceeds its viewport.
2. **Split-screen (half)** and **corner-snap (quarter)** on common monitors work: content reflows
   where it can (existing breakpoints), scrolls where it can't.
3. **OS text-scaling and varied DPI**: no text clipping, no widget overlap — text-bearing
   fixed-height widgets grow with the font.
4. Concrete: the ribbon **"View Raw Data"** (and any long label) wraps to **≤2 lines**, never
   overflows its button, never spills into the Navigator.
5. **No visual regression** at today's standard sizes (≥1100px wide, 100% scale).

## Non-Goals (YAGNI)

- No data-model, pipeline, DB, save-path, plotting-engine, or reporting changes. **View/layout only.**
- No new theming system, no user-facing font-size preference, no per-widget DPI overrides beyond the
  shared `AppTheme` helpers.
- No redesign of any panel's information architecture — only its sizing/scroll behavior.

## Existing infrastructure to build on

- **`ResponsiveLayout`** singleton (`src/utils/ResponsiveLayout.{h,cpp}`): already installs a resize
  event filter on `MainWindow`, debounces (50ms), and emits `breakpointChanged(Breakpoint, width)` +
  `widthChanged(int)`. Current breakpoints: Compact <1100 (ribbon icons-only, sidebar→32px strip),
  sensory-narrow <700, detailed-narrow <800, detailed-stack-charts <1000. We **extend** this, not
  replace it.
- **`AppTheme`** (`src/utils/AppTheme.{h,cpp}`): the central font authority (`fontDefault()` 9pt,
  `fontSmall()` 8pt, etc.). We **add** scale-aware sizing helpers here.
- **`RibbonWidget`** (`src/widgets/RibbonWidget.{h,cpp}`): hand-built ribbon; large buttons fixed
  80×76 with a single-newline wrap heuristic; group/tab fixed heights 98/108; compact-mode path
  already exists.

## Design — Components

### 1. `ScrollHost` — reusable universal scroll wrapper (new: `src/widgets/ScrollHost.{h,cpp}`)

A thin `QScrollArea` factory/subclass:
- `setWidgetResizable(true)` — the content widget expands to fill the viewport when there's room.
- Both scrollbar policies = `Qt::ScrollBarAsNeeded` — invisible when content fits, scrollbars the
  instant it doesn't, in either direction.
- `setFrameShape(QFrame::NoFrame)`, transparent background, zero viewport margins — visually a no-op
  until it scrolls.
- Helper `ScrollHost::wrap(QWidget* content) -> ScrollHost*` so call sites read cleanly.

**Where it wraps (per-region — the correct Qt pattern for a `QMainWindow` with docks; a single
whole-window scroll is not feasible because the dock manager owns the dock areas):**
- Each `m_centralStack` page: the TPM splitter, `SensoryPanel`, `DetailedSensoryPanel`.
- The Navigator dock's full sidebar panel (file tree already scrolls, but the Properties + Test
  Averages + image-button stack below it can overflow — wrap the whole `m_sidebarFullPanel`).
- The Notes dock content (ensure both-direction as-needed).
- The ribbon's group row (horizontal scroll — see §4).
- Dialog content panels (see §6).

The content widgets keep a sensible `minimumSizeHint`; when the viewport is smaller, `ScrollHost`
scrolls instead of clipping. **This is the structural guarantee that nothing is ever inaccessible**,
and it makes the whole overhaul robust to any fixed-size site we miss.

### 2. Window floor + DPI policy (`MainWindow`, `src/main.cpp`)

- `MainWindow`: `setMinimumSize(1280, 800)` → **`setMinimumSize(480, 360)`**.
- `main.cpp`: `QApplication::setHighDpiScaleFactorRoundingPolicy(Qt::HighDpiScaleFactorRoundingPolicy::PassThrough)`
  **before** the `QApplication` is constructed, for smooth fractional text scaling at 125/150%
  (Qt 6 default rounds, which causes abrupt jumps). High-DPI scaling itself is already on by default
  in Qt 6.

### 3. `AppTheme` font-metric helpers (extend `src/utils/AppTheme.{h,cpp}`)

One home for scale-aware sizing so clip-prone dimensions track the active font:
- `int AppTheme::lineUnit(const QFont& f = fontDefault())` — `QFontMetrics(f).height()` (one text line).
- `int AppTheme::controlHeight(const QFont& f = fontDefault(), int vPad = 6)` — `lineUnit(f) + vPad`,
  the standard height for single-line controls (replaces the hard-coded 20/22/24px heights).
- `int AppTheme::em(qreal n, const QFont& f = fontDefault())` — `n * QFontMetrics(f).averageCharWidth()`,
  for width minimums.

These are **pure functions of the current font** — no caching, no state, so they re-evaluate
correctly when the OS scale changes (Qt re-polishes fonts on DPI change).

### 4. Ribbon fixes (`src/widgets/RibbonWidget.cpp`)

- **Balanced ≤2-line wrap** (`addLargeButton`): compute the **real** available text width (button
  width − border − padding), choose the single split point that yields two lines that both fit, and
  set `btn->setText(line1 + "\n" + line2)`. Critically, prevent `QToolButton` from re-wrapping a line
  into a 3rd row (the current bug): if even the best 2-line split doesn't fit at the active font,
  the **button grows in width** (see next bullet) rather than spawning a 3rd line.
- **Large buttons become min-size, not fixed**: `setMinimumSize(80, controlHeight-derived 2-line
  height + icon)` and a `Preferred`/`Minimum` size policy instead of `setFixedSize(80, 76)`. At
  standard scale the rendered size stays 80×76 (no visual regression); under text-scaling the button
  grows instead of clipping.
- **Group/tab fixed heights (98/108) → font-derived minimums** so the group title row ("Data" in the
  bug screenshot) is never clipped when text scales. `RibbonGroup` height = content + separator +
  title, each from `lineUnit`.
- **Horizontal scroll on the group row**: wrap the `RibbonTab`'s group container in a horizontal
  `ScrollHost` so groups never clip off-screen; compact-mode (icons-only) remains the first-line
  response, scroll is the fallback after that.

### 5. `ResponsiveLayout` extension (`src/utils/ResponsiveLayout.{h,cpp}`)

- Keep the existing breakpoints; add a **very-narrow** state that auto-collapses both side docks
  (Navigator + Notes) so the central `ScrollHost` gets maximum room before scrolling.
- Optional aspect-ratio hook (only if needed during implementation): pick splitter orientation
  (vertical vs horizontal) for the TPM/Sensory splits when the window is very tall-narrow vs
  wide-short. Kept minimal — add only if a panel demonstrably needs it.

### 6. Fixed→min sweep + dialogs

- Systematic pass over the ~50 offender sites mapped in the session's Explore report:
  `setFixedWidth/Height/Size` → `setMinimumWidth/Height` + appropriate size policy, **except**
  genuinely-fixed elements (icon sizes, 1px separators, square icon-strip buttons). Width caps that
  exist only for cross-card alignment become `minimumWidth` (grow-allowed); the `ScrollHost` catches
  any residual overflow.
- Dialogs (`DatabaseBrowserDialog`, `DataCleanupDialog`, `ImageViewDialog`, `ImageInboxDialog`,
  `HeaderEditDialog`, …): lower their `setMinimumSize` floors and wrap their content in a `ScrollHost`
  so they fit small screens.

## Data flow / behavior

No model or signal changes beyond `ResponsiveLayout`'s existing emissions. `ScrollHost`s react
automatically to resize; `AppTheme` helpers are pure; the ribbon wrap runs at button-construction
and on font-change. Risk class: **view/layout only** — the lowest-risk change category in this app.

## Testing / Verification (closed-loop harnesses)

Per the owner's standing preference for self-runnable closed-loop harnesses:

1. **Extend + stabilize `tst_responsivelayout`** (currently pre-existing-flaky): drive `MainWindow`
   through a size sweep — 1920×1080, 1280×800, 960×540 (corner-snap), 800×600, 480×360, plus extreme
   aspect ratios 600×1200 and 1600×500 — and assert, for each major region, that either the content
   fits **or** the wrapping `ScrollHost`'s scrollbar is active in the overflow direction. I.e.
   **no region is clipped without a scrollbar.** Also assert the ribbon "View Raw Data" label renders
   in ≤2 lines within the button bounds. Stabilize the existing flakiness as part of this.

2. **`--ui-stress` screenshot harness** (new CLI flag in `main.cpp`, like `--self-test`): cycles
   `MainWindow` through a matrix of (window size × text-scale factor) presets, calls `grab()` on the
   window for each, and writes one PNG per case to a target dir (default `%TEMP%\dve_ui_stress\`),
   plus a small JSON index. Runs headless-ish (offscreen where possible), no GUI interaction, exits.
   This lets the agent run a closed verification loop and lets the owner eyeball every aspect-ratio /
   DPI case at a glance without owning every monitor. Document it in `tests/deployment/README.md`.

3. Build `-Werror -Wall -Wextra` clean; full `tests\run-tests.ps1` suite green (modulo the known
   pre-existing-flaky `tst_excelreader`/`tst_dataprocessor` bundled-Python file-load alternation).

## Phasing

- **Phase 1 — structural guarantee + the headline bug.** `ScrollHost` + wrap central pages/docks +
  lower window floor + HiDPI PassThrough + ribbon ≤2-line/grow fix. Delivers "nothing inaccessible"
  *and* the View Raw Data fix.
- **Phase 2 — text-scale safety.** `AppTheme` helpers + convert clip-prone fixed heights.
- **Phase 3 — completeness + verification.** Per-panel fixed→min sweep + dialogs + `ResponsiveLayout`
  tuning + the extended `tst_responsivelayout` + the `--ui-stress` harness.

## Risks & mitigations

- **Splitters inside scroll areas** (TPM/Sensory use `QSplitter`): with `widgetResizable(true)` the
  splitter reports its `sizeHint`; verify each split page behaves at small sizes (Phase 1 check).
- **Visual regression at standard scale**: every fixed→min conversion keeps the old value as the
  *minimum*, so the standard-size look is unchanged; the `--ui-stress` baseline at 1280×800/100%
  is the regression guard.
- **MIP at build time**: decrypt (`python tools/decrypt_via_copy.py --apply`) before each build.
