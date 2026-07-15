// DV-21 in-progress form draft: mirror the live submit form to localStorage on
// every edit (and after each successful submit's reset), then restore it on
// load, so closing or reloading the page before submitting never loses the
// entry. 24h TTL, matching the sample history. All guarded so a missing script
// can never break the core submit flow. Restore runs on DOMContentLoaded, i.e.
// AFTER the inline script's synchronous regenUid(), so the restored uid wins.
(function () {
  "use strict";
  var KEY = "dve_sensory_draft_v1";
  var TTL_MS = (window.SensoryHistory && SensoryHistory.TTL_MS) || 24 * 60 * 60 * 1000;

  function form() { return document.getElementById("f"); }

  // Every named data control in the form (skip the submit button); this captures
  // the headers, sample_name, the metric sliders, puff length, comments, and the
  // hidden sample_uid without hard-coding the (server-templated) metric names.
  function fields(f) {
    var out = [];
    for (var i = 0; i < f.elements.length; i++) {
      var el = f.elements[i];
      if (el.name && el.type !== "submit") out.push(el);
    }
    return out;
  }

  function save() {
    var f = form(); if (!f) return;
    var data = {};
    fields(f).forEach(function (el) { data[el.name] = el.value; });
    try { localStorage.setItem(KEY, JSON.stringify({ ts: Date.now(), fields: data })); } catch (e) {}
  }

  function clear() { try { localStorage.removeItem(KEY); } catch (e) {} }

  function restore() {
    var f = form(); if (!f) return;
    var rec;
    try { rec = JSON.parse(localStorage.getItem(KEY) || "null"); } catch (e) { rec = null; }
    if (!rec || typeof rec.ts !== "number" || !rec.fields) return;
    if (Date.now() - rec.ts > TTL_MS) { clear(); return; }   // stale draft -> drop it
    Object.keys(rec.fields).forEach(function (name) {
      var el = f.elements[name];
      if (!el || el.name !== name) return;                    // skip missing / collections
      el.value = rec.fields[name];
      if (el.type === "range") el.dispatchEvent(new Event("input"));   // refresh the readout
    });
  }

  function wire() {
    var f = form(); if (!f) return;
    restore();                       // repopulate BEFORE listening, so it doesn't re-save itself
    f.addEventListener("input", save);
    f.addEventListener("change", save);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", wire);
  else wire();

  // save() is also called from the submit-success branch (mirrors the post-reset
  // state); clear() is available for a hard reset if ever needed.
  window.SensoryDraft = { save: save, clear: clear };
})();
