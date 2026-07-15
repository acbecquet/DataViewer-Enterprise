// DV-21 overlay radar, rendered INLINE in the review page (below prev/next).
// Keeps the desktop RadarChartWidget shape (5 axes, 1..9 rings, one outline-only
// polygon per sample) but themed to the dark site palette (dark panel, gray grid,
// light labels, blue-family data). Overlays every sample from the files the drawer
// currently has SHOWN (across all tests), with a per-sample checkbox legend above
// the canvas. Driven by SensoryHistory.renderPlot() (called from the review page).
(function () {
  "use strict";
  var H = window.SensoryHistory;
  if (!H) return;

  function $(id) { return document.getElementById(id); }
  var canvas = $("plot-canvas"), legend = $("plot-toggles");

  var samples = [];            // [{ key, color, label, scores }]
  var hiddenSamples = {};      // sampleKey -> true (persists across renders)

  function sampleKey(f, rec) {
    var s = rec.sample || {};
    return f.key + "::" + (s.sample_uid ? s.sample_uid : String(rec.ts));
  }

  // Blue-family series color that reads on the dark panel: single sample = the
  // site accent; multiple = a hue ramp across the blue band (cyan-blue..indigo)
  // at constant lightness so each sample is distinct but on-theme.
  function hslHex(h, s, l) {
    var c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs((h / 60) % 2 - 1)), m = l - c / 2;
    var r = 0, g = 0, b = 0;
    if (h < 60) { r = c; g = x; } else if (h < 120) { r = x; g = c; }
    else if (h < 180) { g = c; b = x; } else if (h < 240) { g = x; b = c; }
    else if (h < 300) { r = x; b = c; } else { r = c; b = x; }
    function hx(v) { var t = Math.round((v + m) * 255).toString(16); return t.length < 2 ? "0" + t : t; }
    return "#" + hx(r) + hx(g) + hx(b);
  }
  function blueish(i, n) {
    if (n <= 1) return "#3b82f6";                 // the site accent
    return hslHex(203 + (i / (n - 1)) * 34, 0.78, 0.62);   // cyan-blue .. blue (no purple)
  }

  // Collect samples from every SHOWN file, across all tests, in a stable order.
  function collect() {
    var tests = H.group(H.load()), out = [];
    tests.forEach(function (t) {
      t.files.forEach(function (f) {
        if (H.fileHidden(f.key)) return;               // drawer-level file filter
        f.samples.forEach(function (rec) {
          var s = rec.sample || {};
          out.push({ key: sampleKey(f, rec),
                     label: (s.name || "(no name)") + " · " + t.test + " · " + f.label,
                     scores: s.scores || {} });
        });
      });
    });
    out.forEach(function (sm, i) { sm.color = blueish(i, out.length); });
    return out;
  }

  // Public entry: the review page calls this when it opens or its data changes.
  function renderPlot() {
    samples = collect();
    renderLegend();
    drawRadar();
  }

  function renderLegend() {
    legend.innerHTML = "";
    if (!samples.length) {
      var empty = document.createElement("p");
      empty.className = "hist-empty";
      empty.textContent = "No shown files. Enable some in the menu.";
      legend.appendChild(empty);
      return;
    }
    samples.forEach(function (sm) {
      var row = document.createElement("label"); row.className = "toggle-row";
      var cb = document.createElement("input"); cb.type = "checkbox";
      cb.checked = !hiddenSamples[sm.key];
      cb.addEventListener("change", function () {
        if (cb.checked) delete hiddenSamples[sm.key]; else hiddenSamples[sm.key] = true;
        drawRadar();
      });
      var sw = document.createElement("span"); sw.className = "swatch"; sw.style.background = sm.color;
      var lab = document.createElement("span"); lab.className = "lgd-label"; lab.textContent = sm.label;
      row.appendChild(cb); row.appendChild(sw); row.appendChild(lab);
      legend.appendChild(row);
    });
  }

  function drawRadar() {
    var ctx = canvas.getContext("2d");
    var W = canvas.width, Ht = canvas.height, n = H.PLOT_METRICS.length;
    var cx = W / 2, cy = Ht / 2 + 4, radius = Math.min(W, Ht) / 2 - 46;
    ctx.clearRect(0, 0, W, Ht); ctx.fillStyle = "#181d25"; ctx.fillRect(0, 0, W, Ht);   // dark panel

    // rings 1..9 (subtle gray; heavier solid ring 9 as the boundary)
    for (var s = 1; s <= 9; s++) {
      ctx.beginPath();
      for (var i = 0; i < n; i++) {
        var p = H.axisPointXY(i, s, n, cx, cy, radius);
        if (i === 0) ctx.moveTo(p.x, p.y); else ctx.lineTo(p.x, p.y);
      }
      ctx.closePath();
      if (s === 9) { ctx.strokeStyle = "rgba(154,166,178,0.55)"; ctx.lineWidth = 1.8; ctx.setLineDash([]); }
      else { ctx.strokeStyle = "rgba(154,166,178,0.28)"; ctx.lineWidth = 1.4; ctx.setLineDash([4, 4]); }
      ctx.stroke();
    }
    ctx.setLineDash([]);

    // spokes + axis labels (light, on-theme)
    ctx.strokeStyle = "rgba(154,166,178,0.4)"; ctx.lineWidth = 1;
    ctx.fillStyle = "#e9edf2"; ctx.font = "bold 11px system-ui";
    ctx.textAlign = "center"; ctx.textBaseline = "middle";
    for (var j = 0; j < n; j++) {
      var tip = H.axisPointXY(j, 9, n, cx, cy, radius);
      ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(tip.x, tip.y); ctx.stroke();
      var lp = H.axisPointXY(j, 9, n, cx, cy, radius + 22);
      ctx.fillText(H.PLOT_METRICS[j].replace("Overall ", ""), lp.x, lp.y);
    }

    // scale numbers along axis 0 (top spoke)
    ctx.fillStyle = "#9aa6b2"; ctx.font = "bold 9px system-ui"; ctx.textAlign = "right";
    for (var sc = 1; sc <= 9; sc++) {
      var pt = H.axisPointXY(0, sc, n, cx, cy, radius);
      ctx.fillText(String(sc), pt.x - 3, pt.y);
    }

    // one outline-only polygon per shown, non-hidden sample (blue-family)
    samples.forEach(function (sm) {
      if (hiddenSamples[sm.key]) return;
      ctx.beginPath();
      for (var i2 = 0; i2 < n; i2++) {
        var val = sm.scores[H.PLOT_METRICS[i2]]; if (typeof val !== "number") val = 5;
        var pp = H.axisPointXY(i2, val, n, cx, cy, radius);
        if (i2 === 0) ctx.moveTo(pp.x, pp.y); else ctx.lineTo(pp.x, pp.y);
      }
      ctx.closePath(); ctx.strokeStyle = sm.color; ctx.lineWidth = 2.5; ctx.stroke();
    });
  }

  H.renderPlot = renderPlot;   // review page drives rendering
})();
