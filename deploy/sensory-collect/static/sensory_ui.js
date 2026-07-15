// DV-21 mobile sensory form: local 24h history storage, the submit-success
// record hook, the slide-out drawer (with a per-file show/hide filter for the
// plot), and a dedicated review page (results table + prev/next + per-sample
// delete) reached by tapping a file. All additive; the core submit flow in
// form.html is untouched except one guarded SensoryHistory.record call.
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
  var reviewPage = $("review-page"), reviewTable = $("review-table"), reviewTitle = $("review-title");

  // ---- view state ----
  var state = { fileKey: null, sampleIdx: 0 };
  // Files (tester+round) hidden from the plot; default all shown.
  var hiddenFiles = {};

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
        var chk = document.createElement("input");
        chk.type = "checkbox"; chk.className = "hist-check";
        chk.checked = !hiddenFiles[f.key];
        chk.setAttribute("aria-label", "Show " + f.label + " in plot");
        chk.addEventListener("click", function (ev) { ev.stopPropagation(); });   // don't open the review
        chk.addEventListener("change", function () {
          if (chk.checked) delete hiddenFiles[f.key]; else hiddenFiles[f.key] = true;
        });
        var lab = document.createElement("span"); lab.className = "hist-row-label";
        lab.textContent = f.label + " (" + f.samples.length + ")";
        lab.addEventListener("click", function () { closeDrawer(); openFile(f.key); });
        var del = document.createElement("button");
        del.className = "hist-x"; del.setAttribute("aria-label", "Delete " + f.label);
        del.textContent = "✕";
        del.addEventListener("click", function (ev) { ev.stopPropagation(); deleteFile(f.key); });
        row.appendChild(chk); row.appendChild(lab); row.appendChild(del);
        list.appendChild(row);
      });
    });
  }

  function deleteFile(key) {
    save(load().filter(function (r) { return H.fileKey(r) !== key; }));
    delete hiddenFiles[key];
    if (state.fileKey === key) closeReview();
    renderDrawer();
  }

  // ---- review page (opened by tapping a file in the drawer) ----
  function samplesFor(key) {
    return load().filter(function (r) { return H.fileKey(r) === key; })
                 .sort(function (a, b) { return a.ts - b.ts; });
  }
  function openFile(key) { state.fileKey = key; state.sampleIdx = 0; openReview(); }
  function openReview() { renderReview(); reviewPage.classList.add("open"); }
  function closeReview() { reviewPage.classList.remove("open"); state.fileKey = null; }

  function addInfo(box, label, val) {
    if (val === undefined || val === null || val === "") return;
    var p = document.createElement("div"); p.className = "rev-info";
    var k = document.createElement("span"); k.className = "rev-info-k"; k.textContent = label + ": ";
    var v = document.createElement("span"); v.textContent = val;
    p.appendChild(k); p.appendChild(v); box.appendChild(p);
  }

  function renderReview() {
    var recs = samplesFor(state.fileKey);
    if (!recs.length) { closeReview(); return; }
    if (state.sampleIdx >= recs.length) state.sampleIdx = recs.length - 1;
    if (state.sampleIdx < 0) state.sampleIdx = 0;
    var rec = recs[state.sampleIdx], s = rec.sample || {}, sc = s.scores || {};

    reviewTitle.textContent = H.fileLabel(rec.tester, rec.round) + " · " + rec.test_title;

    reviewTable.innerHTML = "";
    var cap = document.createElement("p"); cap.className = "rev-cap";
    cap.textContent = "Sample " + (state.sampleIdx + 1) + " / " + recs.length;
    reviewTable.appendChild(cap);
    addInfo(reviewTable, "Sample", s.name);
    addInfo(reviewTable, "Assessor", rec.assessor);
    addInfo(reviewTable, "Media", rec.media);
    addInfo(reviewTable, "Puff length", (typeof s.puff_length_sec === "number" ? s.puff_length_sec + " s" : ""));

    var table = document.createElement("table"); table.className = "rev-table";
    var hr = document.createElement("tr");
    ["Metric", "Score"].forEach(function (htxt) {
      var th = document.createElement("th"); th.textContent = htxt; hr.appendChild(th);
    });
    table.appendChild(hr);
    H.PLOT_METRICS.forEach(function (m) {
      var tr = document.createElement("tr");
      var td1 = document.createElement("td"); td1.textContent = m;
      var td2 = document.createElement("td"); td2.className = "rev-score";
      td2.textContent = (typeof sc[m] === "number" ? sc[m] : "-");
      tr.appendChild(td1); tr.appendChild(td2); table.appendChild(tr);
    });
    reviewTable.appendChild(table);

    if (s.comments) {
      var c = document.createElement("div"); c.className = "rev-comments";
      var ck = document.createElement("div"); ck.className = "rev-info-k"; ck.textContent = "Comments";
      var cv = document.createElement("div"); cv.textContent = s.comments;
      c.appendChild(ck); c.appendChild(cv); reviewTable.appendChild(c);
    }

    $("rev-prev").disabled = (state.sampleIdx === 0);
    $("rev-next").disabled = (state.sampleIdx === recs.length - 1);
  }

  function deleteCurrentSample() {
    var recs = samplesFor(state.fileKey);
    var rec = recs[state.sampleIdx];
    if (!rec) return;
    save(load().filter(function (r) {
      return !(H.fileKey(r) === H.fileKey(rec) && r.ts === rec.ts &&
               (r.sample || {}).sample_uid === (rec.sample || {}).sample_uid);
    }));
    renderReview();   // re-reads; closes the page if the file is now empty
  }

  // ---- wire up ----
  var hb = $("hist-btn"); if (hb) hb.addEventListener("click", openDrawer);
  var hc = $("hist-close"); if (hc) hc.addEventListener("click", closeDrawer);
  if (backdrop) backdrop.addEventListener("click", closeDrawer);
  var rp = $("rev-prev"); if (rp) rp.addEventListener("click", function () { state.sampleIdx--; renderReview(); });
  var rn = $("rev-next"); if (rn) rn.addEventListener("click", function () { state.sampleIdx++; renderReview(); });
  var rb = $("review-back"); if (rb) rb.addEventListener("click", closeReview);
  var rd = $("review-del"); if (rd) rd.addEventListener("click", deleteCurrentSample);

  // Expose storage + record for form.html and sensory_plot.js.
  H.load = load; H.save = save; H.record = record;
  // The plot reads which files the drawer currently has shown.
  H.fileHidden = function (key) { return !!hiddenFiles[key]; };
})();
