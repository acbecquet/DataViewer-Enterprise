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
