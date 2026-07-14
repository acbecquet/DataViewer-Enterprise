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
