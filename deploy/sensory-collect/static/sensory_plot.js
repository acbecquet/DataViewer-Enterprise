// DV-21 review-page radar (inline, below prev/next). Shows JUST the samples of
// the file being reviewed - one outline-only polygon per sample on the desktop
// pentagon shape, themed to the dark site panel, with a per-sample checkbox
// legend above it. Rendered at devicePixelRatio so it stays crisp (not blurry).
// Driven by SensoryHistory.renderPlot(records) from the review page.
(function () {
  "use strict";
  var H = window.SensoryHistory;
  if (!H) return;

  function $(id) { return document.getElementById(id); }
  var canvas = $("plot-canvas"), legend = $("plot-toggles");

  var LW = 340, LH = 320;            // logical drawing size (CSS units)
  var samples = [];                  // [{ key, color, label, scores }]
  var hiddenSamples = {};            // sampleKey -> true (persists across renders)

  function sampleKey(rec) {
    var s = rec.sample || {};
    return s.sample_uid ? s.sample_uid : String(rec.ts);
  }

  function hslHex(h, s, l) {
    var c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs((h / 60) % 2 - 1)), m = l - c / 2;
    var r = 0, g = 0, b = 0;
    if (h < 60) { r = c; g = x; } else if (h < 120) { r = x; g = c; }
    else if (h < 180) { g = c; b = x; } else if (h < 240) { g = x; b = c; }
    else if (h < 300) { r = x; b = c; } else { r = c; b = x; }
    function hx(v) { var t = Math.round((v + m) * 255).toString(16); return t.length < 2 ? "0" + t : t; }
    return "#" + hx(r) + hx(g) + hx(b);
  }

  // Wide, distinct, high-contrast categorical colors for the dark panel, NO
  // yellow/lime. Golden-angle hue spread (well-separated, unique for many
  // samples) that skips the yellow band; lightness/saturation vary per lap.
  function seriesColor(i) {
    var hue = (210 + i * 137.508) % 360;
    if (hue >= 38 && hue <= 95) hue = (hue + 62) % 360;    // skip yellow / lime
    var lap = Math.floor(i / 6);
    var l = 0.64 - (lap % 3) * 0.07;                        // 0.64 / 0.57 / 0.50
    var s = 0.74 + (lap % 2) * 0.12;                        // 0.74 / 0.86
    return hslHex(hue, Math.min(s, 1), l);
  }

  // White text with a dark outline (readable over grid, polygons, or panel).
  function outlinedText(ctx, txt, x, y, lw) {
    ctx.lineJoin = "round";
    ctx.lineWidth = lw; ctx.strokeStyle = "rgba(15,18,23,0.92)"; ctx.strokeText(txt, x, y);
    ctx.fillStyle = "#ffffff"; ctx.fillText(txt, x, y);
  }

  // Build the plot set from the review file's sample records (in order).
  function renderPlot(records) {
    records = records || [];
    samples = records.map(function (rec, i) {
      var s = rec.sample || {};
      return { key: sampleKey(rec), color: seriesColor(i),
               label: (s.name || ("Sample " + (i + 1))), scores: s.scores || {} };
    });
    renderLegend();
    drawRadar();
  }

  function renderLegend() {
    legend.innerHTML = "";
    if (!samples.length) return;
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
    // Crisp on hi-DPI: back the canvas at devicePixelRatio, draw in logical units.
    var dpr = window.devicePixelRatio || 1;
    var SCALE = Math.max(2, Math.min(3, Math.round(dpr)));
    if (canvas.width !== LW * SCALE) { canvas.width = LW * SCALE; canvas.height = LH * SCALE; }
    var ctx = canvas.getContext("2d");
    ctx.setTransform(SCALE, 0, 0, SCALE, 0, 0);

    var W = LW, Ht = LH, n = H.PLOT_METRICS.length;
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

    // spokes
    ctx.strokeStyle = "rgba(154,166,178,0.4)"; ctx.lineWidth = 1;
    for (var j = 0; j < n; j++) {
      var tip = H.axisPointXY(j, 9, n, cx, cy, radius);
      ctx.beginPath(); ctx.moveTo(cx, cy); ctx.lineTo(tip.x, tip.y); ctx.stroke();
    }

    // one outline-only polygon per shown, non-hidden sample
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

    // axis labels (white + dark outline). The two upper-side labels (Burnt Taste
    // i=1, Smoothness i=4) tilt +/-36 deg to run parallel to the pentagon edge.
    ctx.font = "bold 12px system-ui"; ctx.textAlign = "center"; ctx.textBaseline = "middle";
    for (var j2 = 0; j2 < n; j2++) {
      var rot = (j2 === 1) ? 36 : (j2 === 4) ? -36 : 0;
      var lp = H.axisPointXY(j2, 9, n, cx, cy, radius + (rot ? 18 : 22));
      var label = H.PLOT_METRICS[j2].replace("Overall ", "");
      if (rot) {
        ctx.save(); ctx.translate(lp.x, lp.y); ctx.rotate(rot * Math.PI / 180);
        outlinedText(ctx, label, 0, 0, 3.5);
        ctx.restore();
      } else {
        outlinedText(ctx, label, lp.x, lp.y, 3.5);
      }
    }
    // scale numbers along axis 0 (top spoke)
    ctx.font = "bold 10px system-ui"; ctx.textAlign = "right";
    for (var sc = 1; sc <= 9; sc++) {
      var pt = H.axisPointXY(0, sc, n, cx, cy, radius);
      outlinedText(ctx, String(sc), pt.x - 3, pt.y, 2.5);
    }
  }

  H.renderPlot = renderPlot;   // review page drives rendering (passes this file's records)
})();
