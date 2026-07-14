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
  // Collision-free identity for a (test, tester, round) file. JSON.stringify of
  // the triple is unambiguous regardless of what the free-text fields contain.
  function fileKey(rec) {
    return JSON.stringify([rec.test_title || "", rec.tester || "", rec.round || ""]);
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
