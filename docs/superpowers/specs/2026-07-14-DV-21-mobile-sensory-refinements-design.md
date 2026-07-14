---
date: 2026-07-14
topic: "DV-21 - mobile sensory form refinements (Phase 3 of the v2.8 master spec)"
status: approved
approved_by: owner (Charlie), 2026-07-14
branch: feature/v2.9
base: main @ 20bedf4 (tag v2.8.0)
plane_issue: DV-21
parent_spec: docs/superpowers/specs/2026-07-08-v2.8-master-spec.md (Section 4, Phase 3)
target_app: deploy/sensory-collect/ (NAS-deployed Flask web form; owner deploys)
release: v2.9 (with DV-23)
---

# DV-21 - Mobile Sensory Form Refinements (design)

## 1. Context and goal

The phone sensory web form already shipped (v2.6.0) and lives at `deploy/sensory-collect/`.
It is a single-page Flask app: one `GET /` (renders `templates/form.html`) and one `POST /submit` (validates, then calls the `dve_append_sensory_sample` stored function as the least-privilege `sensory_web` role).
DV-21 adds four refinements on top of it without disturbing that core submit flow.

The four refinements (master spec Phase 3):

- 3a - store up to 24 hours of local submission history.
- 3b - a top-left slide-out menu listing past tests, grouped by test then tester+round; tapping an entry shows that session's first sample.
- 3c - prev/next buttons below Submit to page through a file's samples.
- 3d - a Plot button (right of Next) that opens an overlay radar page with per-file show/hide toggles and a Back button.

## 2. Confirmed decisions (owner, 2026-07-14)

These resolve the open questions the master spec flagged for Phase 3.

1. Data source: local device only (browser `localStorage`, 24 h TTL).
   No database read endpoint is added, so the anonymous endpoint still cannot enumerate the test catalog or other users' sessions (its core security property).
2. Sample view: read-only.
   Browsing history never repopulates or mutates the live input form.
3. Delete: dual-level.
   A per-sample delete (an X on the read-only sample cell) and a per-file delete (an X on each drawer row).
   Both delete from local history only.
4. Plot chart: a radar that matches the desktop sensory radar exactly (see Section 6).
5. Plot scope and toggles: one test active at a time (defaults to the first test in history, all of its files shown); within that test, each tester+round file has its own on/off toggle, and many files may be active at once.
   (This supersedes the original ticket's "per-sample checkboxes"; the owner confirmed per-file toggles - the per-sample wording was a mistake.)

Two clarifications the owner confirmed at approval (2026-07-14):

- Per-file toggling is final (see decision 5).
- The Mfused host is a pure reskin: Mode A/B/C map to rounds 1/2/3, so local history stores and the drawer labels the round number (e.g. "R1"), not the mode letter. Left as-is by design.

## 3. Data model (browser localStorage)

One key: `dve_sensory_history_v1`, holding a JSON array of submission records, newest last.
Each record is written by the submit-success callback and captures exactly what was submitted:

```
{
  ts:            <epoch ms of the successful submit>,
  test_title:    <string>,
  tester:        <string>,
  round:         <string, e.g. "1" | "2" | "N/A">,
  assessor:      <string>,
  media:         <string>,
  sample: {
    name:            <string>,
    scores:          { "Burnt Taste": n, "Vapor Volume": n, "Overall Flavor": n, "Smoothness": n, "Overall Liking": n },
    puff_length_sec: <number>,
    comments:        <string>,
    sample_uid:      <string>
  }
}
```

TTL: on every load and before every write, records with `ts < now - 24h` are pruned and the pruned array is written back.
This is the whole of 3a; there is no timer, the prune runs on page load and on each submit.

Grouping for display (3b, 3d): fold the flat array into `test_title -> (tester,round) file -> samples[]`.
Ordering: tests in first-seen order; files in first-seen order; samples in submit order (`ts`).
A "file" corresponds to one desktop sensory session (natural key test-title + tester + round), which is the unit the owner toggles in the plot.

Note on field order: the submit form's metric order is Burnt Taste, Vapor Volume, Overall Flavor, Smoothness, Overall Liking.
The desktop plot order is different (Section 6) and the plot must use the plot order, not the submit order.

## 4. Only touch to the existing submit flow

The existing `form.html` submit handler is unchanged except for one added call in its success branch:

```
if (res.ok && res.body.ok) {
  show("ok", "Saved. Ready for the next sample.");
  if (window.SensoryHistory && SensoryHistory.record)  // guarded: a failed
    SensoryHistory.record(form, res.body.sample_uid);  // script load can never
  resetSampleOnly();                                    // break the Saved flow
  regenUid();
}
```

The call is guarded so that if the external scripts fail to load, the core submit UX is unaffected.
`app.py` is not modified at all.
No new route, no new Python dependency, no schema change.

## 5. UI surfaces

All four refinements are additive DOM plus JS; the form, sliders, and submit path render and behave exactly as today when the user ignores the new controls.

### 5a. Slide-out drawer (3b)

- A fixed top-left hamburger button (`aria-label="History"`) sits above `<h1>`.
- Tapping it slides in a left drawer (CSS `transform: translateX`) over a dimmed backdrop; tapping the backdrop or the drawer's X closes it.
- Drawer body: the grouped 24 h history.
  Each test is a section header; under it, one row per file showing `tester - Rn (count)` plus a right-aligned delete X.
  Tapping the row (not the X) closes the drawer and loads that file into the read-only viewer at sample 1.
- Empty state: "No submissions in the last 24 hours."

### 5b. Read-only sample viewer + delete (3b, 3c)

- A hidden section directly below the Submit button, revealed when a file is selected.
- It renders the current sample read-only: sample name, the five scores, puff length, comments, plus a small caption "Sample k / N - <tester> Rn - <test title>".
- A delete X on the right of the sample cell removes that one sample from local history, then advances to the next sample (or closes the viewer and the file if it was the last).
- The live input form above is never touched by any of this.

### 5c. Prev/Next + Plot controls (3c, 3d)

- A control row sits directly below Submit: `[< Prev] [Next >] [Plot]`.
- The row is hidden until a file is selected from the drawer.
- Prev/Next page through the selected file's samples in the viewer and clamp at the ends (disabled at the first/last sample).
- Plot (to the right of Next) opens the overlay plot page.

### 5d. Overlay plot page (3d)

- A full-screen in-page overlay (normal flow, not `position: fixed`) with a top bar: a Back arrow (returns to the form and closes the overlay) and the title "Overlay plot".
- A test selector (native `<select>`) choosing exactly one test; it defaults to the first test in history.
- A radar canvas (Section 6) drawn from the selected test's currently-enabled files.
- A toggle list below the chart: one checkbox per tester+round file of the selected test, all checked by default; toggling re-renders the radar.
- Switching the test selector rebuilds the toggle list (all on) and redraws.

## 6. Radar rendering - match desktop exactly

Ported faithfully from `src/ui/RadarChartWidget.cpp` (the on-screen, non-report path) so the phone chart reads like the desktop sensory radar.
Drawn by hand on a `<canvas>` (no chart library, matching the app's no-framework ethos).

Geometry and scale:

- Axes: the five metrics in plot order `["Overall Liking", "Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness"]`, with Overall Liking at 12 o'clock.
- Axis i is at angle `270 + (360/5) * i` degrees in screen coordinates (y down), i.e. clockwise from the top.
- A score `v` in [1,9] maps to radius `((v - 1) / 8) * R`: score 1 sits at the center, score 9 on the outer ring. (This 5-metric widget is not inverted.)

Grid and axes:

- Integer rings 1..9. Rings 1..8 are dashed grey `rgb(165,165,165)` width ~1.4; ring 9 is solid grey width ~1.8 as the boundary (this is the DV-20 styling).
- Scale numbers 1..9 drawn along the top spoke (axis 0), grey `rgb(80,80,80)`, bold.
- Spokes from center to each score-9 vertex, grey `rgb(150,150,150)` width 1.
- Axis labels (the metric names) outside the polygon at each vertex; Overall Liking top, Vapor Volume and Overall Flavor lower, Burnt Taste upper-right and Smoothness upper-left.
  Exact desktop label rotation is fussy on a phone; label placement is best-effort-match (names at the correct vertices, legible, not overlapping the rings), not a pixel port of the desktop label geometry.

Samples:

- One outline-only polygon per sample (no fill), width ~2.5, closed across the five axes.
- Color is `seriesColor(globalSampleIndex)`, a faithful JS port of `AppTheme::seriesColor`:
  a curated 20-color RGB palette first (indices 0..19), then golden-angle HSV for the rest:
  `hue = floor((210 + (k+1) * 137.508) mod 360)` with `k = idx - 20`, remapping the yellow band `[45,70]` to `hue+30`, and `val = 196 - (lap%3)*28`, `sat = 178 + (lap%2)*40` where `lap = floor(k/8)`.
  The curated 20 (RGB) are: (31,119,180) (214,39,40) (44,160,44) (148,103,189) (255,127,14) (23,190,207) (227,119,194) (140,86,75) (127,127,127) (0,0,139) (139,0,0) (0,100,0) (75,0,130) (255,0,255) (0,191,255) (178,34,34) (70,130,180) (255,20,147) (47,79,79) (106,90,205).
- Global sample index is assigned once across the selected test's samples in (file order, then sample order) and is stable, so toggling a file off and on does not recolor anything.
  A file's toggle controls the visibility of all of that file's sample polygons together (per-file toggle over per-sample rendering).

Interaction difference from desktop, by owner design: desktop toggles per sample in its legend; the web plot toggles per file (tester+round).
The rendering (geometry, scale, rings, outline-only polygons, color sequence) is identical; only the toggle granularity differs.

Background is white, matching the desktop chart.

## 7. Code structure and files touched

New (all under `deploy/sensory-collect/`):

- `static/sensory_history.js` - pure, side-effect-free functions, usable from both the browser and a Node/headless test:
  `prune(records, nowMs)`, `group(records)` (-> tests/files/samples), `seriesColorHex(idx)`, and radar geometry helpers `axisPointXY(i, score, n, cx, cy, r)`.
  Exposes both a browser global (`window.SensoryHistory`) and `module.exports` for tests.
- `static/sensory_ui.js` - the DOM wiring that consumes the pure module and `localStorage`: `record(form, uid)` (reads the submitted form fields and writes the pruned history back), drawer open/close and rendering, the read-only viewer, prev/next, both delete paths, and the plot overlay including the canvas radar draw.
  The impure `record` is attached to the shared `window.SensoryHistory` global so the one-line submit-callback change can call `SensoryHistory.record(...)`; all side-effecting code stays in this file, the pure math stays in `sensory_history.js`.

Modified:

- `templates/form.html` - add the hamburger button, drawer markup, the read-only viewer section, the prev/next/plot control row, the plot overlay markup, the two `<script src=...>` includes, and the single `SensoryHistory.record(...)` line in the existing submit-success branch.
  New CSS lives in the existing `<style>` block (drawer transform, backdrop, overlay, toggle rows).

The Mfused host variant is unaffected: the new features read only what was actually submitted, so they work identically for both the default and Mfused forms.

## 8. Non-goals and guardrails

- No change to `POST /submit`, `dve_append_sensory_sample`, the DB schema, or the `sensory_web` role.
- No server-side session listing, so no new way to enumerate the catalog.
- No new Python dependency; the container image is unchanged.
- History is per-device and ephemeral (24 h); it is explicitly not a sync feature or a source of truth (the DB remains the record).
- Do not touch Synology or the NAS; the owner deploys this web change.

## 9. Testing and verification

- Pure logic (`sensory_history.js`): a headless assertion harness covering 24 h TTL prune (boundary at exactly 24 h), grouping into tests/files/samples with correct ordering, the `seriesColorHex` port matching the C++ values for indices 0..19 and a couple of golden-angle indices, and radar point math (score 1 at center, score 9 at the axis-9 vertex).
- Flask render test: extend the existing `tests/test_mfused_form.py` pattern to assert the new elements and script includes render for both the default and Mfused hosts, and that the core form fields are still present (regression guard on "submit flow untouched").
- Local browser walkthrough (closed-loop visual probe): run the Flask dev server, seed `localStorage` with representative history, then drive drawer -> viewer -> prev/next -> per-sample delete -> per-file delete -> plot (test switch and file toggles), capturing screenshots. `GET /` needs no database, so this runs without the NAS.
- The existing server-side `pytest` (`tests/test_append.py`) stays green because the server is unchanged.

## 10. Acceptance criteria

From the master spec, plus the confirmed refinements:

- The core submit flow is unchanged and still writes correctly to Postgres (verified by the render regression test and by the untouched `test_append.py`).
- A successful submit appends to local history; records persist for 24 h and are pruned after.
- The drawer lists the 24 h history grouped by test then tester+round; tapping a file opens its first sample read-only.
- Prev/Next pages the selected file's samples and clamps at the ends.
- A per-sample X and a per-file X each delete from local history at their level and update the drawer/viewer immediately.
- The Plot page shows one test at a time (default first), with per-file toggles all on by default, rendering a desktop-matching radar; toggling files and switching tests redraw correctly; Back returns to the form.
- The headless logic harness and the browser walkthrough both pass; the full existing web test set is green.
