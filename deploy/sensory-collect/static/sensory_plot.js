// DV-21 overlay plot page: a hand-drawn <canvas> radar matching the desktop
// sensory RadarChartWidget. Overlays every sample from the files the drawer
// currently has SHOWN (across all tests at once), one outline-only polygon per
// sample colored by a stable index, with a per-sample legend (checkbox + swatch)
// to hide individual samples. Reads history via SensoryHistory.load/.fileHidden.
(function () {
  "use strict";
  var H = window.SensoryHistory;
  if (!H) return;

  function $(id) { return document.getElementById(id); }
  var page = $("plot-page"), canvas = $("plot-canvas"), legend = $("plot-toggles");

  // Active plot set, rebuilt each openPlot: [{ key, color, label, scores }].
  var samples = [];
  var hiddenSamples = {};   // sampleKey -> true

  function sampleKey(f, rec) {
    var s = rec.sample || {};
    return f.key + "::" + (s.sample_uid ? s.sample_uid : String(rec.ts));
  }

  // Collect samples from every SHOWN file, across all tests, in a stable order;
  // each gets a stable series color by running index (matches the desktop).
  function collect() {
    var tests = H.group(H.load());
    var out = [], gi = 0;
    tests.forEach(function (t) {
      t.files.forEach(function (f) {
        if (H.fileHidden(f.key)) return;               // drawer-level file filter
        f.samples.forEach(function (rec) {
          var s = rec.sample || {};
          out.push({
            key: sampleKey(f, rec),
            color: H.seriesColorHex(gi),
            label: (s.name || "(no name)") + " · " + t.test + " · " + f.label,
            scores: s.scores || {}
          });
          gi++;
        });
      });
    });
    return out;
  }

  function openPlot() {
    samples = collect();
    hiddenSamples = {};                                 // default: all samples shown
    renderLegend();
    drawRadar();
    page.classList.add("open");
  }
  function closePlot() { page.classList.remove("open"); }

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
      var sw = document.createElement("span"); sw.className = "swatch";
      sw.style.background = sm.color;
      var lab = document.createElement("span"); lab.className = "lgd-label";
      lab.textContent = sm.label;
      row.appendChild(cb); row.appendChild(sw); row.appendChild(lab);
      legend.appendChild(row);
    });
  }

  function drawRadar() {
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
  }

  var pl = $("review-plot"); if (pl) pl.addEventListener("click", openPlot);
  var pb = $("plot-back"); if (pb) pb.addEventListener("click", closePlot);
})();
