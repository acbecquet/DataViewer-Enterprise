# DV-21 Mobile Sensory Form Refinements - Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add four additive refinements to the existing phone sensory web form (24h local history, a slide-out session list, prev/next sample navigation, and an overlay radar plot) without touching the core submit flow.

**Architecture:** All new behavior is client-side. A pure, side-effect-free JS module (`sensory_history.js`) holds the 24h-TTL pruning, grouping, the desktop-matching color sequence, and radar geometry; it is unit-tested under Node. Two impure UI modules wire the DOM: `sensory_ui.js` (localStorage, the record hook, the drawer, the read-only viewer, prev/next, and both delete paths) and `sensory_plot.js` (the overlay page and the hand-drawn `<canvas>` radar). `form.html` gains the markup, CSS, three script includes, and a single guarded `SensoryHistory.record(...)` call in its existing submit-success branch. `app.py`, `POST /submit`, the DB, and the `sensory_web` role are untouched.

**Tech Stack:** Flask/Jinja (existing), vanilla ES5-ish JS (no framework, matching the current file), HTML `<canvas>` 2D, Node (pure-logic tests), pytest + Flask test client (render tests). Design source: `docs/superpowers/specs/2026-07-14-DV-21-mobile-sensory-refinements-design.md`.

---

## File Structure

New (all under `deploy/sensory-collect/`):

- `static/sensory_history.js` - pure functions: `prune`, `group`, `fileKey`, `fileLabel`, `seriesColorHex`, `axisPointXY`, plus constants `TTL_MS` and `PLOT_METRICS`. Browser global `window.SensoryHistory` and Node `module.exports`.
- `static/sensory_ui.js` - impure: `load`/`save` (localStorage), `record` (form -> history), the drawer, the read-only viewer, prev/next, and per-sample + per-file delete. Attaches `load`/`save`/`record` to the shared `window.SensoryHistory` global.
- `static/sensory_plot.js` - impure: the overlay plot page and the `<canvas>` radar (test selector, per-file toggles, draw). Reads history via `SensoryHistory.load()`.
- `tests/test_history.js` - Node assertion harness for `sensory_history.js` (no framework; exits non-zero on failure).
- `tests/test_form_render.py` - pytest render assertions for the new elements (default + Mfused hosts) and a regression guard that the core submit fields are still present.

Modified:

- `templates/form.html` - add the hamburger, drawer, viewer, prev/next/plot control row, plot overlay markup, new CSS in the existing `<style>` block, three `<script>` includes, and the one guarded `SensoryHistory.record(...)` line.

Notes for the implementer:

- Write every file with the Write/Edit tools (they persist plaintext on this machine). Do NOT create these `.js`/`.py` via Python - Python-written files pick up MIP encryption at rest on this machine. No C++ is built in this plan, so no decrypt step is needed.
- Metric axis order for the plot is `["Overall Liking", "Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness"]` (desktop `kSensoryMetricsPlot`), which is deliberately different from the submit-form field order.
- Do not touch Synology or the NAS; the owner deploys this web change.

---

## Task 1: Pure logic module `sensory_history.js` + Node harness

**Files:**
- Create: `deploy/sensory-collect/static/sensory_history.js`
- Test: `deploy/sensory-collect/tests/test_history.js`

- [ ] **Step 1: Write the failing Node test**

Create `deploy/sensory-collect/tests/test_history.js`:

```javascript
// Node assertion harness for sensory_history.js. No framework: prints ok lines
// and exits non-zero on the first failed assertion (throw). Run with `node`.
const assert = require("assert");
const H = require("../static/sensory_history.js");

let pass = 0;
function t(name, fn) { fn(); pass++; console.log("ok - " + name); }

const HOUR = 3600 * 1000, DAY = 24 * HOUR, NOW = 1000000000000;

t("prune keeps <24h, drops >=24h and malformed", () => {
  const recs = [
    { ts: NOW - 1 },            // fresh -> keep
    { ts: NOW - DAY + 1 },      // just under 24h -> keep
    { ts: NOW - DAY },          // exactly 24h -> drop
    { ts: NOW - DAY - 1 },      // older -> drop
    { ts: "x" },                // malformed -> drop
    null                        // malformed -> drop
  ];
  assert.strictEqual(H.prune(recs, NOW).length, 2);
  assert.deepStrictEqual(H.prune("nope", NOW), []);
});

t("group folds into tests->files->samples with correct order + labels", () => {
  const recs = [
    { ts: NOW - 5, test_title: "T1", tester: "Al", round: "1", sample: {} },
    { ts: NOW - 4, test_title: "T1", tester: "Al", round: "1", sample: {} },
    { ts: NOW - 3, test_title: "T1", tester: "Bo", round: "2", sample: {} },
    { ts: NOW - 2, test_title: "T2", tester: "Al", round: "N/A", sample: {} }
  ];
  const g = H.group(recs);
  assert.strictEqual(g.length, 2);
  assert.strictEqual(g[0].test, "T1");
  assert.strictEqual(g[0].files.length, 2);
  assert.strictEqual(g[0].files[0].label, "Al R1");
  assert.strictEqual(g[0].files[0].samples.length, 2);
  assert.strictEqual(g[0].files[1].label, "Bo R2");
  assert.strictEqual(g[1].files[0].label, "Al");   // N/A round -> no suffix
});

t("seriesColorHex matches the curated 0..19 palette exactly", () => {
  assert.strictEqual(H.seriesColorHex(0), "#1f77b4");
  assert.strictEqual(H.seriesColorHex(1), "#d62728");
  assert.strictEqual(H.seriesColorHex(8), "#7f7f7f");
  assert.strictEqual(H.seriesColorHex(19), "#6a5acd");
});

t("seriesColorHex golden-angle beyond curated is a valid, distinct hex", () => {
  const a = H.seriesColorHex(20), b = H.seriesColorHex(21);
  assert.ok(/^#[0-9a-f]{6}$/.test(a) && /^#[0-9a-f]{6}$/.test(b));
  assert.notStrictEqual(a, b);
});

t("axisPointXY: score 1 at center, score 9 at the i=0 top vertex", () => {
  const c = H.axisPointXY(0, 1, 5, 100, 100, 80);
  assert.ok(Math.abs(c.x - 100) < 1e-9 && Math.abs(c.y - 100) < 1e-9);
  const top = H.axisPointXY(0, 9, 5, 100, 100, 80);   // 270deg -> straight up
  assert.ok(Math.abs(top.x - 100) < 1e-6 && Math.abs(top.y - 20) < 1e-6);
});

console.log("\n" + pass + " passed");
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node deploy/sensory-collect/tests/test_history.js`
Expected: FAIL - `Cannot find module '../static/sensory_history.js'`.

- [ ] **Step 3: Write the module**

Create `deploy/sensory-collect/static/sensory_history.js`:

```javascript
// DV-21 pure helpers for the phone sensory form: 24h-TTL pruning, grouping,
// the desktop-matching series-color sequence, and radar geometry. No DOM, no
// localStorage, no side effects -- unit-testable under Node and reused by the
// browser UI modules. Color sequence + geometry mirror the desktop
// RadarChartWidget / AppTheme::seriesColor (see the DV-21 design spec).
(function (factory) {
  var api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  if (typeof window !== "undefined") window.SensoryHistory = api;
  else if (typeof globalThis !== "undefined") globalThis.SensoryHistory = api;
})(function () {
  "use strict";

  var TTL_MS = 24 * 60 * 60 * 1000;

  // Plot axis order (desktop kSensoryMetricsPlot): Overall Liking at 12 o'clock.
  // Deliberately different from the submit-form field order.
  var PLOT_METRICS = ["Overall Liking", "Burnt Taste", "Vapor Volume",
                      "Overall Flavor", "Smoothness"];

  // Drop records older than 24h (exactly 24h expires); ignore malformed entries.
  function prune(records, nowMs) {
    if (!Array.isArray(records)) return [];
    var cutoff = nowMs - TTL_MS;
    return records.filter(function (r) {
      return r && typeof r.ts === "number" && r.ts > cutoff;
    });
  }

  function fileLabel(tester, round) {
    var t = (tester || "").trim();
    return (round && round !== "N/A") ? (t + " R" + round) : t;
  }
  function fileKey(rec) {
    return [rec.test_title || "", rec.tester || "", rec.round || ""].join("");
  }

  // Fold a flat record array into ordered tests -> files -> samples. Tests and
  // files preserve first-seen (earliest-ts) order; samples keep submit order.
  function group(records) {
    var tests = [], testIdx = {}, fileIdx = {};
    var ordered = (Array.isArray(records) ? records.slice() : [])
      .sort(function (a, b) { return a.ts - b.ts; });
    ordered.forEach(function (r) {
      var title = r.test_title || "";
      if (!(title in testIdx)) { testIdx[title] = tests.length; tests.push({ test: title, files: [] }); }
      var ti = testIdx[title], k = fileKey(r);
      if (!(k in fileIdx)) {
        fileIdx[k] = { t: ti, f: tests[ti].files.length };
        tests[ti].files.push({ key: k, tester: r.tester || "", round: r.round || "",
                               label: fileLabel(r.tester, r.round), samples: [] });
      }
      var loc = fileIdx[k];
      tests[loc.t].files[loc.f].samples.push(r);
    });
    return tests;
  }

  // Port of AppTheme::seriesColor -> "#rrggbb": curated 20 first, then
  // golden-angle HSV. Distinct, projector-safe, stable per index.
  var CURATED = [
    [31,119,180],[214,39,40],[44,160,44],[148,103,189],[255,127,14],
    [23,190,207],[227,119,194],[140,86,75],[127,127,127],[0,0,139],
    [139,0,0],[0,100,0],[75,0,130],[255,0,255],[0,191,255],
    [178,34,34],[70,130,180],[255,20,147],[47,79,79],[106,90,205]
  ];
  function hx(n) { var s = n.toString(16); return s.length < 2 ? "0" + s : s; }
  function rgbHex(r, g, b) { return "#" + hx(r) + hx(g) + hx(b); }

  // HSV (h:0-359, s:0-255, v:0-255) -> [r,g,b], matching Qt QColor::fromHsv.
  function hsvToRgb(h, s, v) {
    var S = s / 255, V = v / 255, c = V * S, hp = h / 60.0;
    var x = c * (1 - Math.abs((hp % 2) - 1)), r1 = 0, g1 = 0, b1 = 0;
    if (hp < 1) { r1 = c; g1 = x; }
    else if (hp < 2) { r1 = x; g1 = c; }
    else if (hp < 3) { g1 = c; b1 = x; }
    else if (hp < 4) { g1 = x; b1 = c; }
    else if (hp < 5) { r1 = x; b1 = c; }
    else { r1 = c; b1 = x; }
    var m = V - c;
    return [Math.round((r1 + m) * 255), Math.round((g1 + m) * 255), Math.round((b1 + m) * 255)];
  }

  function seriesColorHex(idx) {
    if (idx < 0) idx = 0;
    if (idx < CURATED.length) { var c = CURATED[idx]; return rgbHex(c[0], c[1], c[2]); }
    var k = idx - CURATED.length;
    var hue = Math.floor((210.0 + (k + 1) * 137.508) % 360.0);
    if (hue >= 45 && hue <= 70) hue = (hue + 30) % 360;
    var lap = Math.floor(k / 8);
    var val = 196 - (lap % 3) * 28, sat = 178 + (lap % 2) * 40;
    var rgb = hsvToRgb(hue, Math.max(0, Math.min(255, sat)), Math.max(0, Math.min(255, val)));
    return rgbHex(rgb[0], rgb[1], rgb[2]);
  }

  // Axis i at (270 + 360/n * i) deg (screen coords, clockwise from top); score 1
  // at center, 9 at radius. Matches RadarChartWidget::axisPoint.
  function axisPointXY(i, score, n, cx, cy, r) {
    var a = (270.0 + (360.0 / n) * i) * Math.PI / 180.0;
    var rr = ((score - 1) / 8.0) * r;
    return { x: cx + rr * Math.cos(a), y: cy + rr * Math.sin(a) };
  }

  return { TTL_MS: TTL_MS, PLOT_METRICS: PLOT_METRICS, prune: prune, group: group,
           fileKey: fileKey, fileLabel: fileLabel, seriesColorHex: seriesColorHex,
           axisPointXY: axisPointXY };
});
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `node deploy/sensory-collect/tests/test_history.js`
Expected: 5 `ok - ...` lines then `5 passed`, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add deploy/sensory-collect/static/sensory_history.js deploy/sensory-collect/tests/test_history.js
git commit -m "feat(DV-21): pure history/color/radar-geometry module + Node tests"
```

---

## Task 2: `form.html` scaffolding + Flask render test

**Files:**
- Modify: `deploy/sensory-collect/templates/form.html`
- Test: `deploy/sensory-collect/tests/test_form_render.py`

- [ ] **Step 1: Write the failing render test**

Create `deploy/sensory-collect/tests/test_form_render.py`:

```python
"""DV-21 -- render tests for the mobile sensory form refinements.

GET / must include the new history/nav/plot affordances and the three script
includes for BOTH the default and the Mfused host, while the core submit fields
remain present (regression guard: the submit flow is untouched). DB-free.
"""
import importlib
import os
import sys

import pytest

pytest.importorskip("flask")
pytest.importorskip("flask_limiter")

HERE = os.path.dirname(os.path.abspath(__file__))
SERVICE_DIR = os.path.dirname(HERE)
MFUSED_HOST = "mfused-sensory.ccell-sdr.com"
PLAIN_HOST = "sensory.example.com"


@pytest.fixture()
def client(monkeypatch):
    monkeypatch.setenv("DVE_MFUSED_HOSTS", MFUSED_HOST)
    sys.path.insert(0, SERVICE_DIR)
    try:
        app_module = importlib.import_module("app")
        importlib.reload(app_module)
        app_module.app.config["TESTING"] = True
        app_module.app.config["RATELIMIT_ENABLED"] = False
        yield app_module.app.test_client()
    finally:
        if SERVICE_DIR in sys.path:
            sys.path.remove(SERVICE_DIR)


def _html(client, host):
    r = client.get("/", headers={"Host": host})
    assert r.status_code == 200, r.get_data(as_text=True)
    return r.get_data(as_text=True)


@pytest.mark.parametrize("host", [PLAIN_HOST, MFUSED_HOST])
def test_new_affordances_present(client, host):
    html = _html(client, host)
    for token in ('id="hist-btn"', 'id="hist-drawer"', 'id="hist-list"',
                  'id="hist-viewer"', 'id="hist-nav"',
                  'id="nav-prev"', 'id="nav-next"', 'id="nav-plot"',
                  'id="plot-page"', 'id="plot-canvas"', 'id="plot-test"',
                  'id="plot-toggles"', 'id="plot-back"',
                  "sensory_history.js", "sensory_ui.js", "sensory_plot.js"):
        assert token in html, token


@pytest.mark.parametrize("host", [PLAIN_HOST, MFUSED_HOST])
def test_core_submit_fields_untouched(client, host):
    html = _html(client, host)
    for token in ('id="f"', 'id="submit-btn"', 'name="test_title"', 'name="tester"',
                  'name="Burnt Taste"', 'name="Vapor Volume"', 'name="Overall Flavor"',
                  'name="Smoothness"', 'name="Overall Liking"'):
        assert token in html, token
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python -m pytest deploy/sensory-collect/tests/test_form_render.py -q`
Expected: `test_new_affordances_present` FAILS (e.g. `AssertionError: id="hist-btn"`); `test_core_submit_fields_untouched` PASSES (those fields already exist).

- [ ] **Step 3: Add the new CSS to the `<style>` block**

In `deploy/sensory-collect/templates/form.html`, add these rules immediately before the closing `</style>` (after the `#status.err` rule):

```css
    /* DV-21 history + nav + plot */
    #hist-btn { position:fixed; top:10px; left:10px; width:42px; height:42px; z-index:20;
                display:flex; align-items:center; justify-content:center; font-size:20px;
                margin:0; padding:0; background:var(--card); color:var(--fg);
                border:1px solid var(--line); border-radius:10px; }
    #hist-backdrop { position:fixed; inset:0; background:rgba(0,0,0,0.5); opacity:0;
                     pointer-events:none; transition:opacity .2s; z-index:30; }
    #hist-backdrop.open { opacity:1; pointer-events:auto; }
    #hist-drawer { position:fixed; top:0; left:0; bottom:0; width:min(86vw,340px); z-index:40;
                   background:var(--card); border-right:1px solid var(--line);
                   transform:translateX(-100%); transition:transform .2s;
                   overflow-y:auto; padding:14px; }
    #hist-drawer.open { transform:translateX(0); }
    .hist-drawer-head { display:flex; align-items:center; gap:10px; margin-bottom:8px; }
    .hist-test { color:var(--muted); font-size:0.72rem; text-transform:uppercase;
                 letter-spacing:0.05em; margin:14px 0 6px; }
    .hist-row { display:flex; align-items:center; justify-content:space-between; gap:8px;
                background:#0d1015; border:1px solid var(--line); border-radius:9px;
                padding:10px 12px; margin-bottom:6px; }
    .hist-row-label { flex:1; overflow-wrap:anywhere; }
    .hist-x { flex:0 0 auto; width:30px; height:30px; padding:0; margin:0; font-size:15px;
              background:transparent; color:var(--err); border:none; border-radius:8px; }
    .hist-empty { color:var(--muted); }
    #hist-viewer { display:none; background:var(--card); border:1px solid var(--line);
                   border-radius:10px; padding:10px 12px; margin-top:10px; }
    .viewer-head { display:flex; align-items:center; justify-content:space-between; gap:8px; }
    .viewer-cap { color:var(--muted); font-size:0.82rem; overflow-wrap:anywhere; }
    .viewer-body { margin-top:6px; font-size:0.9rem; overflow-wrap:anywhere; }
    #hist-nav { display:none; gap:8px; margin-top:10px; }
    #hist-nav button { margin-top:0; }
    #hist-nav .nav-btn { flex:1; }
    #hist-nav .nav-plot { flex:0 0 auto; width:auto; padding-left:18px; padding-right:18px; }
    #plot-page { position:fixed; inset:0; z-index:50; background:var(--bg); display:none;
                 flex-direction:column; padding:14px; overflow-y:auto; }
    #plot-page.open { display:flex; }
    .plot-head { display:flex; align-items:center; gap:10px; margin-bottom:12px; }
    #plot-canvas { width:100%; max-width:360px; height:auto; align-self:center;
                   background:#fff; border-radius:10px; }
    .toggle-row { display:flex; align-items:center; gap:10px; padding:9px 4px; font-size:0.92rem; }
    .toggle-row input { width:auto; margin:0; }
```

- [ ] **Step 4: Add the drawer + hamburger markup**

Immediately after the opening `<body>` tag (before `<main>`), insert:

```html
<button type="button" id="hist-btn" aria-label="History">&#9776;</button>
<div id="hist-backdrop"></div>
<aside id="hist-drawer" aria-label="Past tests">
  <div class="hist-drawer-head">
    <button type="button" id="hist-close" class="hist-x" aria-label="Close">&#10005;</button>
    <strong>Past tests (this device)</strong>
  </div>
  <div id="hist-list"></div>
</aside>
```

- [ ] **Step 5: Add the nav row + viewer after the form**

Immediately after the closing `</form>` tag (still inside `<main>`), insert:

```html
    <div id="hist-nav">
      <button type="button" id="nav-prev" class="secondary nav-btn">&#8249; Prev</button>
      <button type="button" id="nav-next" class="secondary nav-btn">Next &#8250;</button>
      <button type="button" id="nav-plot" class="secondary nav-plot">Plot</button>
    </div>
    <div id="hist-viewer"></div>
```

- [ ] **Step 6: Add the plot overlay after `</main>`**

Immediately after the closing `</main>` tag (before the `<script>` block), insert:

```html
<div id="plot-page" aria-label="Overlay plot">
  <div class="plot-head">
    <button type="button" id="plot-back" class="hist-x" aria-label="Back">&#8592;</button>
    <strong>Overlay plot</strong>
  </div>
  <label for="plot-test">Test</label>
  <select id="plot-test"></select>
  <canvas id="plot-canvas" width="340" height="320"></canvas>
  <div id="plot-toggles"></div>
</div>
```

- [ ] **Step 7: Add the three script includes**

Immediately before the existing `<script>` line (the inline block that starts with `// Per-sample fields ...`), insert:

```html
<script src="{{ url_for('static', filename='sensory_history.js') }}"></script>
<script src="{{ url_for('static', filename='sensory_ui.js') }}"></script>
<script src="{{ url_for('static', filename='sensory_plot.js') }}"></script>
```

- [ ] **Step 8: Add the guarded record hook in the submit-success branch**

In the existing inline `<script>`, replace this block:

```javascript
        if (res.ok && res.body.ok) {
          show("ok", "Saved. Ready for the next sample.");
          resetSampleOnly();   // keep headers, clear the sample
          regenUid();          // new uid ONLY after a successful submit
        } else {
```

with:

```javascript
        if (res.ok && res.body.ok) {
          show("ok", "Saved. Ready for the next sample.");
          if (window.SensoryHistory && SensoryHistory.record)   // DV-21: record to
            SensoryHistory.record(form, res.body.sample_uid);   // 24h local history;
          resetSampleOnly();   // keep headers, clear the sample  // guarded so a failed
          regenUid();          // new uid ONLY after a successful submit // load can't break submit
        } else {
```

- [ ] **Step 9: Run the render test to verify it passes**

Run: `python -m pytest deploy/sensory-collect/tests/test_form_render.py -q`
Expected: 4 passed (2 tests x 2 hosts).

- [ ] **Step 10: Verify the existing Mfused render test still passes**

Run: `python -m pytest deploy/sensory-collect/tests/test_mfused_form.py -q`
Expected: all passed (the form still renders Round/Mode correctly).

- [ ] **Step 11: Commit**

```bash
git add deploy/sensory-collect/templates/form.html deploy/sensory-collect/tests/test_form_render.py
git commit -m "feat(DV-21): form.html scaffolding (drawer, viewer, nav, plot overlay) + render test"
```

---

## Task 3: `sensory_ui.js` - storage, record, drawer, viewer, prev/next, dual delete

**Files:**
- Create: `deploy/sensory-collect/static/sensory_ui.js`

This module is verified by a scripted local-browser walkthrough (Step 3) rather than a unit test, because it is DOM/localStorage wiring. The pure math it relies on is already unit-tested in Task 1.

- [ ] **Step 1: Write the module**

Create `deploy/sensory-collect/static/sensory_ui.js`:

```javascript
// DV-21 mobile sensory form: local 24h history storage, the submit-success
// record hook, the slide-out drawer, a read-only sample viewer with prev/next,
// and dual-level delete (per-sample + per-file). All additive; the core submit
// flow in form.html is untouched except one guarded SensoryHistory.record call.
// Pure math/geometry lives in sensory_history.js (window.SensoryHistory).
(function () {
  "use strict";
  var H = window.SensoryHistory;
  if (!H) return;   // pure module missing -> disable extras; core form unaffected

  var KEY = "dve_sensory_history_v1";
  var METRICS = ["Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness", "Overall Liking"];

  function $(id) { return document.getElementById(id); }

  // ---- storage (impure; exposed on the shared global for sensory_plot.js) ----
  function load() {
    var raw;
    try { raw = JSON.parse(localStorage.getItem(KEY) || "[]"); } catch (e) { raw = []; }
    var pruned = H.prune(raw, Date.now());
    if (!Array.isArray(raw) || pruned.length !== raw.length) save(pruned);
    return pruned;
  }
  function save(records) {
    try { localStorage.setItem(KEY, JSON.stringify(records)); } catch (e) {}
  }

  // Build a record from the just-submitted form and append it. Called (guarded)
  // from the form.html submit-success branch.
  function record(form, uid) {
    var el = form.elements;
    function v(name) { return el[name] ? String(el[name].value || "").trim() : ""; }
    function num(name, d) { var n = parseFloat(el[name] ? el[name].value : ""); return isNaN(n) ? d : n; }
    var scores = {};
    METRICS.forEach(function (m) { scores[m] = num(m, 5); });
    var rec = {
      ts: Date.now(),
      test_title: v("test_title"), tester: v("tester"), round: v("round") || "N/A",
      assessor: v("assessor"), media: v("media"),
      sample: { name: v("sample_name"), scores: scores,
                puff_length_sec: num("puff_length_sec", 3.0),
                comments: v("comments"), sample_uid: uid || "" }
    };
    var recs = load();
    recs.push(rec);
    save(recs);
  }

  // ---- element refs ----
  var drawer = $("hist-drawer"), backdrop = $("hist-backdrop"), list = $("hist-list");
  var viewer = $("hist-viewer"), navRow = $("hist-nav");

  // ---- view state ----
  var state = { fileKey: null, sampleIdx: 0 };

  // ---- drawer ----
  function openDrawer() { renderDrawer(); drawer.classList.add("open"); backdrop.classList.add("open"); }
  function closeDrawer() { drawer.classList.remove("open"); backdrop.classList.remove("open"); }

  function renderDrawer() {
    var tests = H.group(load());
    list.innerHTML = "";
    if (!tests.length) {
      var empty = document.createElement("p");
      empty.className = "hist-empty";
      empty.textContent = "No submissions in the last 24 hours.";
      list.appendChild(empty);
      return;
    }
    tests.forEach(function (t) {
      var h = document.createElement("p"); h.className = "hist-test"; h.textContent = t.test;
      list.appendChild(h);
      t.files.forEach(function (f) {
        var row = document.createElement("div"); row.className = "hist-row";
        var lab = document.createElement("span"); lab.className = "hist-row-label";
        lab.textContent = f.label + " (" + f.samples.length + ")";
        lab.addEventListener("click", function () { closeDrawer(); openFile(f.key); });
        var del = document.createElement("button");
        del.className = "hist-x"; del.setAttribute("aria-label", "Delete " + f.label);
        del.textContent = "✕";
        del.addEventListener("click", function (ev) { ev.stopPropagation(); deleteFile(f.key); });
        row.appendChild(lab); row.appendChild(del);
        list.appendChild(row);
      });
    });
  }

  function deleteFile(key) {
    save(load().filter(function (r) { return H.fileKey(r) !== key; }));
    if (state.fileKey === key) hideViewer();
    renderDrawer();
  }

  // ---- read-only viewer + prev/next ----
  function samplesFor(key) {
    return load().filter(function (r) { return H.fileKey(r) === key; })
                 .sort(function (a, b) { return a.ts - b.ts; });
  }
  function openFile(key) { state.fileKey = key; state.sampleIdx = 0; renderViewer(); }
  function hideViewer() { state.fileKey = null; viewer.style.display = "none"; navRow.style.display = "none"; }

  function abbr(metric) { return metric.split(" ").map(function (w) { return w.charAt(0); }).join(""); }

  function renderViewer() {
    var recs = samplesFor(state.fileKey);
    if (!recs.length) { hideViewer(); return; }
    if (state.sampleIdx >= recs.length) state.sampleIdx = recs.length - 1;
    if (state.sampleIdx < 0) state.sampleIdx = 0;
    var rec = recs[state.sampleIdx], s = rec.sample || {}, sc = s.scores || {};
    var scoreStr = H.PLOT_METRICS.map(function (m) { return abbr(m) + " " + sc[m]; }).join(" · ");

    viewer.innerHTML = "";
    var head = document.createElement("div"); head.className = "viewer-head";
    var cap = document.createElement("span"); cap.className = "viewer-cap";
    cap.textContent = "Sample " + (state.sampleIdx + 1) + " / " + recs.length + " · " +
                      H.fileLabel(rec.tester, rec.round) + " · " + rec.test_title;
    var del = document.createElement("button"); del.className = "hist-x";
    del.setAttribute("aria-label", "Delete this sample"); del.textContent = "✕";
    del.addEventListener("click", function () { deleteSample(rec); });
    head.appendChild(cap); head.appendChild(del);

    var body = document.createElement("div"); body.className = "viewer-body";
    body.textContent = (s.name || "(no name)") + "  —  " + scoreStr +
                       "  —  puff " + s.puff_length_sec + "s" +
                       (s.comments ? ("  —  " + s.comments) : "");
    viewer.appendChild(head); viewer.appendChild(body);

    viewer.style.display = "block"; navRow.style.display = "flex";
    $("nav-prev").disabled = (state.sampleIdx === 0);
    $("nav-next").disabled = (state.sampleIdx === recs.length - 1);
  }

  function deleteSample(rec) {
    save(load().filter(function (r) {
      return !(H.fileKey(r) === H.fileKey(rec) && r.ts === rec.ts &&
               (r.sample || {}).sample_uid === (rec.sample || {}).sample_uid);
    }));
    renderViewer();   // re-reads; hides the viewer if the file is now empty
    if (drawer.classList.contains("open")) renderDrawer();
  }

  // ---- wire up ----
  var hb = $("hist-btn"); if (hb) hb.addEventListener("click", openDrawer);
  var hc = $("hist-close"); if (hc) hc.addEventListener("click", closeDrawer);
  if (backdrop) backdrop.addEventListener("click", closeDrawer);
  var np = $("nav-prev"); if (np) np.addEventListener("click", function () { state.sampleIdx--; renderViewer(); });
  var nn = $("nav-next"); if (nn) nn.addEventListener("click", function () { state.sampleIdx++; renderViewer(); });

  // Expose storage + record for form.html and sensory_plot.js.
  H.load = load; H.save = save; H.record = record;
})();
```

- [ ] **Step 2: Start the local Flask server**

Run (background): `cd deploy/sensory-collect && DVE_MFUSED_HOSTS=mfused-sensory.ccell-sdr.com python -m flask --app app run --port 5057`
Expected: `Running on http://127.0.0.1:5057`. `GET /` needs no DB.

- [ ] **Step 3: Browser walkthrough (closed-loop verify)**

Using the in-app Browser (mcp__Claude_Browser__*):
1. `navigate` to `http://127.0.0.1:5057/`.
2. Seed history via `javascript_tool`:

```javascript
(function () {
  var now = Date.now();
  function rec(t, tester, round, name, o, l, dt) {
    return { ts: now - dt, test_title: t, tester: tester, round: round, assessor: "", media: "",
             sample: { name: name, scores: { "Burnt Taste": o, "Vapor Volume": 7, "Overall Flavor": 6,
                       "Smoothness": 5, "Overall Liking": l }, puff_length_sec: 3, comments: "", sample_uid: name } };
  }
  var recs = [ rec("Live Resin LVL 1","Isabel","1","P-1",3,7,5000), rec("Live Resin LVL 1","Isabel","1","P-2",4,6,4000),
               rec("Live Resin LVL 1","Isabel","2","P-3",2,8,3000), rec("Live Resin KO","NA","1","K-1",5,5,2000) ];
  localStorage.setItem("dve_sensory_history_v1", JSON.stringify(recs));
  return "seeded " + recs.length;
})();
```

3. Reload, then `computer` left_click the `#hist-btn` hamburger. `read_page`/screenshot: the drawer shows "Live Resin LVL 1" with rows "Isabel R1 (2)" and "Isabel R2 (1)", and "Live Resin KO" with "NA R1 (1)", each with an X.
4. Click the "Isabel R1 (2)" label. Confirm the drawer closes, `#hist-viewer` shows "Sample 1 / 2 ... Isabel R1 ... Live Resin LVL 1" and the P-1 scores, and `#hist-nav` shows Prev/Next/Plot with Prev disabled.
5. Click `#nav-next`: viewer shows "Sample 2 / 2" (P-2), Next now disabled.
6. Click the viewer's X: viewer drops to "Sample 1 / 1" (only P-1 left in R1... note P-2 deleted). Re-open the drawer: "Isabel R1 (1)".
7. In the drawer, click the X on "Isabel R2 (1)": that row disappears.
8. Take a screenshot of the drawer + viewer for the record.

Expected: every assertion above holds; the core form above the viewer is unchanged.

- [ ] **Step 4: Commit**

```bash
git add deploy/sensory-collect/static/sensory_ui.js
git commit -m "feat(DV-21): history storage, record hook, slide-out drawer, read-only viewer, prev/next, dual delete"
```

---

## Task 4: `sensory_plot.js` - overlay page + canvas radar

**Files:**
- Create: `deploy/sensory-collect/static/sensory_plot.js`

- [ ] **Step 1: Write the module**

Create `deploy/sensory-collect/static/sensory_plot.js`:

```javascript
// DV-21 overlay plot page: a hand-drawn <canvas> radar matching the desktop
// sensory RadarChartWidget. One test active at a time (default = first test in
// history), per-file toggles (all on by default), one outline-only polygon per
// sample colored by a stable global index. Reads history via SensoryHistory.load.
(function () {
  "use strict";
  var H = window.SensoryHistory;
  if (!H) return;

  function $(id) { return document.getElementById(id); }
  var page = $("plot-page"), canvas = $("plot-canvas"),
      testSel = $("plot-test"), toggleList = $("plot-toggles");

  var pstate = { testIdx: 0, hidden: {} };   // hidden: fileKey -> true

  function tests() { return H.group(H.load()); }
  function currentTest(ts) { return ts[pstate.testIdx] || ts[0]; }

  function openPlot() {
    var ts = tests();
    if (!ts.length) return;
    testSel.innerHTML = "";
    ts.forEach(function (t, i) {
      var o = document.createElement("option"); o.value = String(i); o.textContent = t.test;
      testSel.appendChild(o);
    });
    pstate.testIdx = 0; pstate.hidden = {};   // default: first test, all files shown
    testSel.value = "0";
    renderToggles(ts); drawRadar(ts);
    page.classList.add("open");
  }
  function closePlot() { page.classList.remove("open"); }

  function renderToggles(ts) {
    var t = currentTest(ts); toggleList.innerHTML = "";
    if (!t) return;
    t.files.forEach(function (f) {
      var row = document.createElement("label"); row.className = "toggle-row";
      var cb = document.createElement("input"); cb.type = "checkbox";
      cb.checked = !pstate.hidden[f.key];
      cb.addEventListener("change", function () {
        if (cb.checked) delete pstate.hidden[f.key]; else pstate.hidden[f.key] = true;
        drawRadar(tests());
      });
      var span = document.createElement("span");
      span.textContent = f.label + " (" + f.samples.length + ")";
      row.appendChild(cb); row.appendChild(span); toggleList.appendChild(row);
    });
  }

  function drawRadar(ts) {
    var t = currentTest(ts); if (!t) return;
    var ctx = canvas.getContext("2d");
    var W = canvas.width, Ht = canvas.height, n = H.PLOT_METRICS.length;
    var cx = W / 2, cy = Ht / 2 + 4, radius = Math.min(W, Ht) / 2 - 46;
    ctx.clearRect(0, 0, W, Ht); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, W, Ht);

    // rings 1..9 (dashed grey; solid heavier ring 9 as the boundary)
    for (var s = 1; s <= 9; s++) {
      ctx.beginPath();
      for (var i = 0; i < n; i++) {
        var p = H.axisPointXY(i, s, n, cx, cy, radius);
        if (i === 0) ctx.moveTo(p.x, p.y); else ctx.lineTo(p.x, p.y);
      }
      ctx.closePath();
      ctx.strokeStyle = "rgb(165,165,165)";
      if (s === 9) { ctx.lineWidth = 1.8; ctx.setLineDash([]); }
      else { ctx.lineWidth = 1.4; ctx.setLineDash([4, 4]); }
      ctx.stroke();
    }
    ctx.setLineDash([]);

    // spokes + axis labels
    ctx.strokeStyle = "rgb(150,150,150)"; ctx.lineWidth = 1;
    ctx.fillStyle = "rgb(40,40,40)"; ctx.font = "bold 11px system-ui";
    ctx.textAlign = "center"; ctx.textBaseline = "middle";
    for (var j = 0; j < n; j++) {
      var tip = H.axisPointXY(j, 9, n, cx, cy, radius);
      ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(tip.x, tip.y); ctx.stroke();
      var lp = H.axisPointXY(j, 9, n, cx, cy, radius + 22);
      ctx.fillText(H.PLOT_METRICS[j].replace("Overall ", ""), lp.x, lp.y);
    }

    // scale numbers along axis 0 (top spoke)
    ctx.fillStyle = "rgb(80,80,80)"; ctx.font = "bold 9px system-ui"; ctx.textAlign = "right";
    for (var sc = 1; sc <= 9; sc++) {
      var pt = H.axisPointXY(0, sc, n, cx, cy, radius);
      ctx.fillText(String(sc), pt.x - 3, pt.y);
    }

    // samples: stable global color index across the test's files/samples;
    // toggled-off files are skipped but still consume an index (stable colors).
    var gi = 0;
    t.files.forEach(function (f) {
      var visible = !pstate.hidden[f.key];
      f.samples.forEach(function (rec) {
        var color = H.seriesColorHex(gi); gi++;
        if (!visible) return;
        var sco = (rec.sample || {}).scores || {};
        ctx.beginPath();
        for (var i2 = 0; i2 < n; i2++) {
          var val = sco[H.PLOT_METRICS[i2]]; if (typeof val !== "number") val = 5;
          var pp = H.axisPointXY(i2, val, n, cx, cy, radius);
          if (i2 === 0) ctx.moveTo(pp.x, pp.y); else ctx.lineTo(pp.x, pp.y);
        }
        ctx.closePath(); ctx.strokeStyle = color; ctx.lineWidth = 2.5; ctx.stroke();
      });
    });
  }

  var pl = $("nav-plot"); if (pl) pl.addEventListener("click", openPlot);
  var pb = $("plot-back"); if (pb) pb.addEventListener("click", closePlot);
  if (testSel) testSel.addEventListener("change", function () {
    pstate.testIdx = parseInt(testSel.value, 10) || 0; pstate.hidden = {};
    var ts = tests(); renderToggles(ts); drawRadar(ts);
  });
})();
```

- [ ] **Step 2: Browser walkthrough (closed-loop verify)**

With the Flask server from Task 3 still running and history seeded (re-seed via the Task 3 Step 3 snippet if needed):
1. `navigate` to `http://127.0.0.1:5057/`, open the drawer, open "Isabel R1", so the nav row is visible.
2. Click `#nav-plot`. Confirm `#plot-page` is shown, the test `<select>` defaults to "Live Resin LVL 1", `#plot-toggles` lists "Isabel R1" and "Isabel R2" both checked, and `#plot-canvas` shows a radar with grey rings, the five metric labels (Liking / Burnt Taste / Vapor Volume / Flavor / Smoothness), and multiple colored polygons. Screenshot it.
3. Uncheck "Isabel R2": the radar redraws with fewer polygons (and the remaining polygons keep their colors). Screenshot.
4. Change the test `<select>` to "Live Resin KO": toggles rebuild to just "NA R1" (checked), the radar redraws for that test.
5. Click `#plot-back`: the overlay closes and the form + viewer are shown again unchanged.

Expected: all assertions hold; polygon colors are distinct and stable under toggling.

- [ ] **Step 3: Stop the Flask server**

Stop the background server started in Task 3 (Ctrl-C / kill the background job).

- [ ] **Step 4: Commit**

```bash
git add deploy/sensory-collect/static/sensory_plot.js
git commit -m "feat(DV-21): overlay plot page + desktop-matching canvas radar (test select, per-file toggles)"
```

---

## Task 5: End-to-end verification, Phase Log update, and wrap

**Files:**
- Modify: `docs/superpowers/specs/2026-07-08-v2.8-master-spec.md` (Phase Log)

- [ ] **Step 1: Full local end-to-end run**

Start the Flask server (Task 3 Step 2). In the in-app Browser, with a CLEARED `localStorage`:
1. Fill Test Title + Tester + move a couple of sliders, tap Submit. Confirm "Saved" and that `localStorage["dve_sensory_history_v1"]` now has 1 record (check via `javascript_tool`).
2. Submit two more samples (same header, then a different Round). Confirm the drawer groups them correctly.
3. Exercise drawer -> viewer -> prev/next -> per-sample delete -> per-file delete -> Plot (toggle + test switch) -> Back once more, end to end.
4. Confirm the submit form still works after all of the above (submit one more; "Saved").
Stop the server.

- [ ] **Step 2: Run the whole web test set**

Run:
```bash
node deploy/sensory-collect/tests/test_history.js
python -m pytest deploy/sensory-collect/tests/test_form_render.py deploy/sensory-collect/tests/test_mfused_form.py -q
```
Expected: the Node harness prints `5 passed`; pytest reports all passed. (`tests/test_append.py` needs a live DB and is out of scope here - it is unchanged and still skips cleanly without `DVE_TEST_PG_CONN`.)

- [ ] **Step 3: Update the master-spec Phase Log**

In `docs/superpowers/specs/2026-07-08-v2.8-master-spec.md`, replace the line:

```
- Phase 3 (DV-21, v2.9): not started.
```

with a CODE-COMPLETE entry recording: the four refinements shipped client-side (localStorage 24h history, slide-out drawer, read-only viewer + prev/next, overlay radar), the confirmed decisions (localStorage-only, dual-level delete, per-file toggles, desktop-matching radar), that `/submit`/`app.py`/schema are untouched, the new files, and the verification (Node harness + render tests + browser walkthrough). Note any deviation from this plan here.

- [ ] **Step 4: Commit**

```bash
git add docs/superpowers/specs/2026-07-08-v2.8-master-spec.md
git commit -m "docs(DV-21): master-spec Phase 3 log - mobile form refinements code-complete"
```

- [ ] **Step 5: Report status to the owner**

DV-21 is code-complete on `feature/v2.9` and verified locally. It does not ride the desktop installer; the owner deploys the updated `deploy/sensory-collect/` to the NAS. Do NOT touch the NAS/Synology. Next sprint item: DV-23 (NOTIFY storm).

---

## Self-review notes

- Spec coverage: 3a (Task 1 prune + Task 3 load/record), 3b drawer (Task 3), read-only view + dual delete (Task 3), 3c prev/next (Task 3), 3d overlay radar (Task 4). Core-flow-untouched guard: Task 2 render test + the guarded record hook. Desktop-match radar: Task 1 geometry/color + Task 4 draw. Testing strategy: Node harness (Task 1), render tests (Task 2), browser walkthroughs (Tasks 3-4), full E2E (Task 5).
- Type/name consistency: `SensoryHistory` exposes `prune/group/fileKey/fileLabel/seriesColorHex/axisPointXY/PLOT_METRICS/TTL_MS` (Task 1) plus `load/save/record` attached in Task 3; `sensory_plot.js` (Task 4) uses only those names; element ids (`hist-btn/hist-drawer/hist-list/hist-close/hist-backdrop/hist-viewer/hist-nav/nav-prev/nav-next/nav-plot/plot-page/plot-canvas/plot-test/plot-toggles/plot-back`) match across form.html, ui.js, plot.js, and the render test.
- localStorage key `dve_sensory_history_v1` is identical in ui.js and the browser-walkthrough seed snippets.
