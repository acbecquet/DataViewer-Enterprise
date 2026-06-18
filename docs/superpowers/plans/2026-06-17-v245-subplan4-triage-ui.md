# v2.4.2 Sub-plan 4 — Triage UI (version/health filter + classifier + repair/delete) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`). Each task: implementer → spec review → quality review. The UI tasks (T3/T4) are additionally **verified through the visual feedback loop** (launch the app against the sandbox test container, screenshot, hand to the owner) — unit tests cover only the pure logic + data layer.

**Goal:** Give the owner a Database-Browser version/health multi-select filter (the feature originally requested) to see which rows came from which v2-era version and which are unhealthy, plus manual repair/delete — so the v-era data can be triaged: pushed up to date (lossless repair) or cleaned up/removed.

**Architecture:** Tier-A (backwards-compat features A3/A4/A5) of the v2.4.2 batch (spec `docs/superpowers/specs/2026-06-11-v242-backcompat-resilience-design.md`). The DB stamping (A1) + normalizer (A2) shipped in SP1/SP2; the `app_version` column is on all 3 tables + the offline snapshot (SP3-T1). This sub-plan is the **client-side** consumer: a pure-C++ `CompatClassifier` (era from a release-date table + health from data shape), the browser filter UI, and manual repair/delete. Wraps to **v2.4.5** internal; the whole batch then wraps to deployable **v2.5.0** on the owner's install-test approval.

**Tech Stack:** C++17 / Qt 6.10 (QTreeWidget — the browser has NO model/proxy layer; filtering is a full re-populate), PostgreSQL 16 (jsonb), qmake + MinGW.

**Key learnings carried from SP1–SP3 (apply to every task):**
- MIP: before edits/builds `python tools/decrypt_via_copy.py --apply` (its own step; ignore the known exit-2 re-label). Create NEW files (CompatClassifier.{h,cpp}, the new test) via Python round-trip + immediate `git add`; verify blobs `git show HEAD:<path> | head`. Commit trailer `Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>`.
- Test run-PATH gotcha: a DB-container test exe needs `vendor/libpq-16/*.dll` next to it + Qt+MinGW on PATH or it silently exits.
- The adversarial review caught a real Critical/Important on the SP2 keystone, SP3 R5, and SP3 R6 that green unit suites masked. **The browser UI has no automated UI test in this repo — so T3/T4 lean on the visual loop + careful review; do NOT claim the UI works on a green build alone.**
- Branch `feature/v2.4.0-bugfix-batch`; do NOT push/merge; never touch Synology. Build is `-Werror -Wall -Wextra -Wpedantic`.

---

## Test environment

Container `dve-test-pg` on :5433 (conn `host=127.0.0.1 port=5433 dbname=dve_test user=test password=test`), seeded with representative versioned/health demo rows (5 files / 7 sensory / 3 detailed across eras 2.0.2→2.4.1). Visual-loop sandbox: launch `release/DataViewer.exe` with `LOCALAPPDATA=C:\Users\S1134987\AppData\Local\Temp\dve_vis` (a sandbox `db.conf` there points at the test container; prod never contacted).

**Build & run ONE suite:** `cd tests/<suite> && cmd.exe //c "set PATH=C:\Qt\Tools\mingw1310_64\bin;C:\Qt\6.10.1\mingw_64\bin;%PATH% && qmake.exe <suite>.pro && mingw32-make.exe -j8"` then `cp ../../vendor/libpq-16/*.dll ./release/ ; DVE_TEST_PG_CONN='...' PATH="<repo>/vendor/libpq-16:/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH" ./release/<suite>.exe -o res.txt,txt ; grep -E "PASS|FAIL|Totals" res.txt ; rm -f res.txt`. Incremental app compile: `cd build && cmd.exe //c "set PATH=...;%PATH% && qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro CONFIG+=release && mingw32-make.exe -j8 release"`.

---

## File Structure

- `src/database/CompatClassifier.{h,cpp}` *(new)* — pure C++ free functions: era bucket (release-date table + `QVersionNumber`) + health flags (from record/struct shape). Mirror `VersionLookup.{h,cpp}`.
- `tests/tst_compatclassifier/` *(new)* — clone `tests/tst_versionlookup/`; era-boundary + health + NULL-stamp tests. Add to `tests/tests.pro` SUBDIRS.
- `src/database/DatabaseManager.{h,cpp}` — extend `FileRecord`/`SensoryRecord`/`DetailedSensoryRecord` + the 3 `list*` SELECTs with `app_version` + a server-side `has_legacy_string_scores` bool; mirror in the snapshot read path (`OfflineSnapshot::list*`).
- `src/ui/DatabaseBrowserDialog.{h,cpp}` — per-tab "Version ▾" checkable dropdown + Health column + AND-combined filter predicate + live per-bucket counts; Repair button; confirm-dialog that LISTS rows.
- `DataViewerEnterprise.pro` — add CompatClassifier sources; VERSION 2.4.4 → 2.4.5.

---

## Task 1 — A3: CompatClassifier (pure C++, unit-tested)

**Files:** `src/database/CompatClassifier.{h,cpp}` (new), `tests/tst_compatclassifier/` (new, clone `tst_versionlookup`), `DataViewerEnterprise.pro`, `tests/tests.pro`.

- [ ] **Step 1: Create CompatClassifier.h** (via Python round-trip + git add). Free functions in `namespace DVE`, no QObject. Sketch:
```cpp
#pragma once
#include <QString>
#include <QDate>
namespace DVE {
enum class EraBucket { V2_0, V2_1, V2_2, V2_3, V2_4, PreV242_Unknown };  // coarse era groups
struct CompatClass {
    QString eraLabel;        // e.g. "v2.2.x", "v2.4.1", "pre-v2.4.2 (unstamped)"; "(approx.)" suffix when inferred from date
    bool    approx = false;  // era inferred from creation date (app_version was NULL)
    QStringList health;      // {"Legacy string scores","Junk candidate","No samples","Missing puff regimes"} or empty == Healthy
    bool isHealthy() const { return health.isEmpty(); }
};
// app_version like "DataViewer/2.2.5" or "" / null; creation = loaded_at/date/timestamp.
CompatClass classifyEra(const QString& appVersion, const QDate& creation);
// Health builders (data-shape, NOT app_version):
//   sensory/detailed: hasLegacyStringScores (server-provided) -> "Legacy string scores";
//                     placeholder/unnamed/zero-samples -> "Junk candidate".
//   TPM files: sampleCount==0 -> "No samples"; missingRegimes -> "Missing puff regimes".
QStringList sensoryHealth(bool hasLegacyStringScores, bool isPlaceholder, int sampleCount);
QStringList fileHealth(int sampleCount, bool missingRegimes);
}
```
- [ ] **Step 2: Write the failing test** `tests/tst_compatclassifier/tst_compatclassifier.cpp` (clone `tst_versionlookup` structure + its `.pro`). Cover: (a) `classifyEra("DataViewer/2.2.5", …)` → "v2.2.x", approx=false; (b) `classifyEra("", date-before-v2.1.0-release)` → "v2.0.x (approx.)", approx=true; (c) `classifyEra("", invalid date)` → "pre-v2.4.2 (unknown)"; (d) era-boundary edges (a creation date exactly on a release date buckets to that release); (e) `sensoryHealth(true,false,1)` → contains "Legacy string scores"; (f) `sensoryHealth(false,true,0)` → "Junk candidate"; (g) `fileHealth(0,false)` → "No samples". Add `tst_compatclassifier` to `tests/tests.pro` SUBDIRS.
- [ ] **Step 3: Run → FAIL** (functions absent). Build+run `tst_compatclassifier`.
- [ ] **Step 4: Implement CompatClassifier.cpp.** The release-date table — the 13 v2 releases (spec line 13: v2.0.0, 2.0.2, 2.0.5, 2.0.8, 2.0.10, 2.1.0, 2.2.0, 2.2.1, 2.2.3, 2.2.4, 2.2.5, 2.3.1, 2.4.1) with their dates DERIVED FROM GIT (`git log --format='%ad %s' --date=short | grep -iE "release.*2\.(0|1|2|3|4)"` or the tag dates). Parse `appVersion` with `QVersionNumber::fromString` after stripping the `DataViewer/` prefix + a leading `v` (reuse the idiom from `UpdateChecker.cpp:83-88`). If a version parses → coarse era label from major.minor. If NULL/empty → find the newest release whose date ≤ `creation` → that era + `approx=true`. If no usable date → "pre-v2.4.2 (unknown)". Health builders: reuse `isPlaceholderSession` (`SensoryData.h:141`) semantics for the junk predicate. NOTE: legacy-string-scores is NOT derivable in C++ (the reader coerces strings→doubles) — it's passed IN as a bool from the server query (Task 2).
- [ ] **Step 5: Run → PASS.** Add CompatClassifier.{cpp,h} to `DataViewerEnterprise.pro` SOURCES/HEADERS.
- [ ] **Step 6: Commit** (CompatClassifier + test + both .pro).

---

## Task 2 — A4 data: surface app_version + legacy-string flag in the record queries

**Files:** `src/database/DatabaseManager.{h,cpp}`, `src/database/OfflineSnapshot.cpp` (snapshot read path), `tests/tst_databasemanager/tst_databasemanager.cpp`.

- [ ] **Step 1: Write the failing test** in `tst_databasemanager.cpp`: seed a sensory_sessions row with `app_version='DataViewer/2.2.5'` and a STRING-typed score (`"7.5"`), and a clean numeric one; call `listSensoryRecords()`; assert the returned record carries `appVersion=="DataViewer/2.2.5"` and `hasLegacyStringScores==true` for the string row, `false` for the numeric. Same shape for files (`appVersion`). RED: the fields don't exist on the record yet.
- [ ] **Step 2: Run → FAIL.**
- [ ] **Step 3: Extend the record structs** (`DatabaseManager.h:41-76`): add `QString appVersion;` to `FileRecord`/`SensoryRecord`/`DetailedSensoryRecord`, and `bool hasLegacyStringScores = false;` to the two sensory records (files: add `bool missingRegimes` only if cheaply queryable — else defer the missing-regime flag to a fuller load and document it).
- [ ] **Step 4: Extend the 3 `list*` SELECTs** (`DatabaseManager.cpp:2070`, `2975`, `3495`): add `app_version`, and for the sensory/detailed ones add a server-side `EXISTS (SELECT 1 FROM jsonb_array_elements(json_data->'samples') s, jsonb_each(s.value) kv WHERE kv.key = ANY($scoreKeys) AND jsonb_typeof(kv.value)='string') AS has_legacy_string_scores` (the score-key list = the 15 keys, same as the normalizer). Bind the keys array. Read the new columns into the structs.
- [ ] **Step 5: Offline path.** Mirror in `OfflineSnapshot::listFiles/listSensoryRecords/listDetailedSensoryRecords`: select `app_version` (now in the snapshot per SP3-T1). SQLite has no `jsonb_typeof` — so `has_legacy_string_scores` can't be computed offline the same way; degrade to `false` offline (the era filter still works; the legacy-score health flag is online-only — document this, and the spec's "filter degrades gracefully" philosophy covers it).
- [ ] **Step 6: Run → PASS** (app_version + legacy flag surfaced online). Build+run `tst_databasemanager`.
- [ ] **Step 7: Commit** (DatabaseManager + OfflineSnapshot + test).

---

## Task 3 — A4 UI: Version ▾ multi-select filter + Health column (VISUAL-VERIFIED)

**Files:** `src/ui/DatabaseBrowserDialog.{h,cpp}`. No automated UI test exists; verified by incremental app compile + the **visual loop**.

- [ ] **Step 1: Classify on load.** In `onRefresh()` (`DatabaseBrowserDialog.cpp:387`), after the 3 vectors load, run `CompatClassifier` over each record → store the `CompatClass` alongside each record (parallel vector or a field). Compute per-(era,health)-bucket counts here.
- [ ] **Step 2: Add a Health column** to each tab's tree (`setColumnCount(6)` + header; set `item->setText(5, klass.health.join(", ") | "Healthy")`, and optionally a Version column showing `klass.eraLabel`). Locations: `.cpp:59-60/153-154/263-264` (headers), `.cpp:453-461/535-545/838-848` (per-item set).
- [ ] **Step 2b: Build the "Version ▾" dropdown** per tab — a `QToolButton` (text "Version ▾", `setPopupMode(InstantPopup)`) with a `QMenu` of checkable `QAction`s (or `QWidgetAction`+`QCheckBox`): one per era bucket present + "pre-v2.4.2 (unstamped)" + one per health flag, each showing a live count (e.g. "v2.2.x (3)"). Place it in each filter row beside the existing `QLineEdit` (`.cpp:46-55/140-149/250-259`). Store checked-state in a per-tab `QSet<QString>`.
- [ ] **Step 3: AND-combine the predicate.** In each `populate*` (`.cpp:421-429/496-503/799-806`), keep the existing text `.contains` check AND add: if any version/health checkboxes are checked, the row must match at least one checked era bucket OR health flag (define the AND/OR semantics per the spec: text-filter AND (selected-buckets); within buckets it's OR). Re-fire `populate*` on any checkbox toggle (connect like the text filter at `.cpp:119/210-212/320-322`). Update the per-bucket counts + the status label.
- [ ] **Step 4: Incremental app compile** → clean under `-Werror`.
- [ ] **Step 5: VISUAL VERIFICATION (loop).** Launch the sandbox app (test container), open the Database Browser, screenshot each tab; toggle the Version ▾ dropdown, check/uncheck a couple of eras + a health flag, screenshot the filtered result + the live counts. Hand the annotated screenshots to the owner to confirm: the dropdown reads naturally, counts are right, the filter AND-combines with the search box, the Health column is legible, no color/wrap/truncation issues (respect the project's no-truncation + color conventions). Iterate on any visual issue the owner flags BEFORE committing.
- [ ] **Step 6: Commit** (DatabaseBrowserDialog) once the owner signs off on the screenshots.

---

## Task 4 — A5: manual repair / delete on filtered rows (VISUAL-VERIFIED)

**Files:** `src/ui/DatabaseBrowserDialog.{h,cpp}` (+ possibly a small DatabaseManager repair entry point).

- [ ] **Step 1: Confirm-dialog lists rows.** Extend the existing per-tab delete confirms (`onDelete` `.cpp:654`, `onSensoryDelete` `.cpp:695`, `onDetailedSensoryDelete` `.cpp:876`) so the `QMessageBox` ENUMERATES the exact rows to be removed (name + era), not just a count (spec A5). Reuse `collectSensoryIds` (`.cpp:678`). Deletion stays manual + confirmed (existing `removeFile`/`removeSensorySession`/`removeDetailedSensorySession`).
- [ ] **Step 2: Add a Repair button** per tab (mirror the `m_deleteBtn` styled-button construction `.cpp:98-102`). Repair semantics by health flag:
  - **Legacy string scores** (sensory/detailed) → lossless: run the server normalizer on the selected rows. Simplest correct path: `SELECT dve_normalize_legacy_json()` (normalizes ALL damaged rows) or a per-row variant; then `onRefresh()`. (The nightly cron already converges these; the button is "normalize now".)
  - **TPM incomplete** (missing data_rows / regimes) → `DbRepair::runDbRepair`/`backfillFile` (needs the source `.xlsx` + bundled Python — `DbRepair.h:45-55`). Only enable when the source file is locatable; else disable + tooltip.
  - **Junk candidate / no samples** → no lossless repair; Delete only (Repair disabled for those rows).
- [ ] **Step 3: Incremental app compile** → clean.
- [ ] **Step 4: Behavior test (test container).** Against the sandbox: select a legacy-string row → Repair → confirm scores became numeric (re-query) + the row's "Legacy string scores" health flag clears on refresh. Select a junk row → Delete → confirm it's gone + the confirm dialog listed it. (These are exercised live against the throwaway container, NOT prod.)
- [ ] **Step 5: VISUAL VERIFICATION (loop).** Screenshot the Repair confirm + the row-listing Delete confirm + the post-repair refreshed view; hand to the owner. Confirm the destructive-action UX is clear (lists what's removed, no accidental-delete footguns). Iterate on owner feedback.
- [ ] **Step 6: Commit** (DatabaseBrowserDialog + any repair entry point) once the owner signs off.

---

## Task 5 — v2.4.5 bump + clean rebuild + verify + installer

**Files:** `DataViewerEnterprise.pro` (VERSION 2.4.4 → 2.4.5), `release_overview/release_overview_v_2_4_5.txt` (new).

- [ ] **Step 1: Bump** VERSION → 2.4.5; confirm CompatClassifier sources are in the `.pro`.
- [ ] **Step 2: MIP decrypt + clean rebuild** (`make clean` on the bump). Clean under `-Werror`.
- [ ] **Step 3: Run all affected suites fresh** — `tst_compatclassifier`, `tst_databasemanager`, plus a regression sweep (`tst_offlinesnapshot`, `tst_livesync`, `tst_saveintegrity_e2e`). All green.
- [ ] **Step 4: Build the installer** (rebuild-dataviewer flow); verify `release\DataViewer.exe` + `dist\DataViewer-setup.exe` = 2.4.5. **No Synology.**
- [ ] **Step 5: Final visual pass (loop)** over the whole triage workflow + a regression glance at the main TPM/sensory views; screenshots to the owner.
- [ ] **Step 6: Write + commit `release_overview_v_2_4_5.txt`** (customer-readable: filter the database by app version + health; spot and clean up or repair old/broken records). Commit the bump + overview.
- [ ] **Step 7: Hand off.** SP4 done → the v2.4.2 batch (SP1–SP4) is feature-complete. On the owner's install-test approval of the stacked builds, wrap to deployable **v2.5.0** (the single minor release dropped on Synology by the owner).

---

## Self-Review

**Spec coverage:** A3 CompatClassifier → T1. A4 surface app_version+legacy flag → T2; filter UI + counts + Health column (3 tabs) → T3. A5 manual repair/delete (deletion always manual + confirm-listed; lossless repair) → T4. ✓

**SP1–SP3 learnings applied:** pure-logic classifier is unit-tested (T1) so the only UI-only parts (T3/T4) are the ones leaning on the visual loop; the legacy-string flag is server-side (the C++ reader can't see it post-coercion — flagged); offline degrades gracefully (no `jsonb_typeof` in SQLite); MIP Python round-trip for new files; the browser has no model/proxy so the filter is a re-populate predicate (map-confirmed). The visual loop is the belt-and-suspenders for the UI that no automated test covers.

**Risk notes:** T3/T4 are UI with no automated UI test in this repo → the visual loop + owner sign-off is the verification gate; do not merge on a green build alone. The Version ▾ checkable-dropdown widget is built from QToolButton+QMenu (no existing reusable widget). Repair semantics differ per health flag (normalize vs DbRepair) — confirm the mapping with the owner during T4.

**Type/name consistency:** `CompatClass`/`classifyEra`/`sensoryHealth`/`fileHealth`, `appVersion`/`hasLegacyStringScores` record fields, the 15 score keys (shared with the normalizer), the era labels — consistent across T1/T2/T3.
