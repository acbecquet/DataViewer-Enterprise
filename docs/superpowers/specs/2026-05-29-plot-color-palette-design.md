# Distinct-color invariant for all plots

**Date:** 2026-05-29
**Status:** Approved (design) — pending spec review
**Scope:** ~5 files + a new shared helper + a project CLAUDE.md rule

## Problem

Plot colors are assigned from small fixed palettes indexed with `% N`, so they
**wrap and repeat** once the element count exceeds the palette size. The user
reported it on the full report's *Lifetime TPM Comparison* chart — with only six
base hues and per-file brightness shading, the chart reads as "only blue and
orange," and a 7th file would wrap back to blue.

The same `% N` flaw exists across the app:

| Site | Current | Symptom |
|---|---|---|
| `ReportGenerator::lifetimeBarColor` | `kFileHues[fileIdx % 6]` + brightness shade | comparison chart repeats hues at 7+ files |
| `ReportGenerator` per-sheet TPM-trend / bar | `kColors[i % 6]` | repeats after 6 samples |
| `PlotEngine` default fallback | `kDefaultColors[i % 8]` | repeats after 8 |
| `PlotEngine` per-sheet bar chart | `kCols[i % 6]` (all blues) | monochrome + repeats |
| `PlotWidget` (on-screen) | `kColors[i % 6]`, `kCbColors[i % N]` | repeats |
| `RadarChartWidget` | `kColors[i % size]` | repeats |

## The invariant (the rule)

> **Within any single plot, every series / bar / slice gets a distinct,
> well-separated color. The palette is always sized to the number of elements —
> never indexed with `% N`.**

Additional constraints (this project — branded reports shown on projectors):

- **No yellow.** Exclude the yellow/gold band (HSV hue ≈ 45°–70°) from both the
  curated palette and the generated fallback. (Removes today's `#CCAA00` gold.)
- **Projector-safe.** Every color must read clearly against a white slide under
  a projector in a lit room. That means: keep saturation reasonably high
  (≳ 0.55), keep value in a mid-dark band (≲ 0.85 at the light end), and never
  let a shade wash out toward white. No pale pastels, no near-white tints, no
  bright cyan/lime.

## Design

### 1. Shared palette helper in `AppTheme`

`AppTheme` already owns every color token, so it is the single source of truth.
Add:

```cpp
// n distinct, well-separated, projector-safe, non-yellow colors.
// Curated qualitative palette for the first ~10; golden-angle hue rotation
// beyond that. Never repeats for any n.
static QVector<QColor> seriesColors(int n);

// A distinct shade of a base hue for grouped charts. Varies value (and floors
// saturation) within a projector-safe band that keeps the hue intact, so
// shades of two different base hues can never collide.
static QColor shade(const QColor& base, int idx, int count);
```

**Curated palette** — reuse the **existing 20-color palette already defined in
`RadarChartWidget::kColors`**, which the team already vetted for exactly this
constraint (its own comment: *"NO yellow / yellow-adjacent hues — those wash out
in PowerPoint and on projector screens"*). Lifting it into `AppTheme` makes it
the single source of truth; `RadarChartWidget` then consumes it instead of
owning a private copy. Ordered most-distinct-first: blue, red, green, purple,
orange, cyan, magenta-pink, brown, gray, navy, dark-red, dark-green, indigo,
magenta, deep-sky-blue, firebrick, steel-blue, deep-pink, dark-slate-gray,
slate-blue. (No entry lies in the yellow band, hue ≈ 45°–70°.)

**Golden-angle fallback** (when `n` exceeds the curated set): start at hue 210°,
step +137.5° each color; if a hue lands in the excluded yellow band, shift it
out; clamp saturation ≈ 0.70 and value ≈ 0.75 so generated colors stay vivid,
non-yellow, and projector-safe. Hues are well-spread and never repeat, so
distinctness holds for arbitrary `n`.

**Shade band:** vary value across ≈ `[0.55, 0.85]` (never up to 1.0 — that is
what washes out on a projector) with saturation floored at ≈ 0.5. This gives
*more* within-family separation than today's `0.6→1.0` ramp while staying
readable. Because only value/saturation change and hue is fixed, no shade of
hue A can equal any shade of hue B.

### 2. Comparison chart — grouped, provably no reuse (`lifetimeBarColor`)

- Compute `fileHues = AppTheme::seriesColors(fileCount)` once — one distinct,
  well-separated base hue per file (no more `% 6`).
- Each bar = `AppTheme::shade(fileHues[fileIdx], sampleIdx, samplesInFile)`.

Distinct base hues + hue-preserving shading ⇒ **no two bars in the chart are
ever the same color**, yet each file remains one recognizable color family
(the "keep file grouping" choice). Refactor `lifetimeBarColor` to take the
file's base color (so it no longer hides its own palette / modulo).

### 3. Every other plot — flat distinct colors via the same helper

Replace each `palette[i % N]` with `AppTheme::seriesColors(count)[i]` in:
`ReportGenerator` (per-sheet trend + bar), `PlotEngine` (default fallback +
per-sheet bar chart), `PlotWidget` (on-screen), `RadarChartWidget`.

- Within a sheet, the trend line and the bar for the **same sample** keep
  matching colors — preserved by computing one palette per sheet and indexing
  it consistently (this intent already exists in the code).
- The per-sheet bar chart changes from all-blue to multi-color (consistent with
  "use lots of colors"; user approved).

### 4. The permanent rule → project `CLAUDE.md`

Add under **Conventions** (next to the other plotting / `rcc` notes):

> **Never reuse a color within a single plot.** Every series/bar/slice gets a
> distinct, well-separated color. Never index a fixed palette with `% N` — size
> the palette to the element count via `AppTheme::seriesColors(n)` (or `shade()`
> for grouped charts). No yellow (hue ≈ 45°–70°); keep colors saturated and
> mid-dark so they read on a projector against white. Beyond the curated
> palette, generate more by golden-angle rotation — never fall back to repeating.

## Files touched

- `src/utils/AppTheme.h` / `AppTheme.cpp` — new `seriesColors` / `shade`.
- `src/reporting/ReportGenerator.cpp` — comparison chart + per-sheet plots.
- `src/plotting/PlotEngine.cpp` — default fallback + per-sheet bar chart.
- `src/plotting/PlotWidget.cpp` — on-screen plots.
- `src/ui/RadarChartWidget.cpp` — radar overlays.
- `CLAUDE.md` — the rule.

## Verification

- Build (qmake + MinGW, `-Werror`).
- Run the Qt Test suite (`tests\run-tests.ps1`); add a small unit test asserting
  `seriesColors(n)` returns `n` **unique** colors for a range of `n` (e.g. 1–40),
  none in the yellow band.
- Generate a sample multi-file report and eyeball: bars distinct, files still
  grouped, nothing yellow, nothing washed out.

## Out of scope

- Semantic colors (`AppTheme::success/warning/error`), presence-avatar colors,
  and UI chrome — unchanged.
- The on-screen `kCbColors` colorblind palette stays colorblind-aware; the new
  curated palette is chosen to be reasonably colorblind-distinguishable for the
  common (≤ ~8 series) case but prioritizes the explicit no-reuse / no-yellow /
  projector mandate where they conflict at high series counts.
