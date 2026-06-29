# DATAVIEWER-6 / MS-4 Option B — Remove the TPM data table + Notes panel + View Raw Data + Cleanup hardening: Design Spike

**Date:** 2026-06-24 · **Status:** SPIKE COMPLETE — awaiting owner sign-off on §9 · **Parent:** MS-4 in `docs/superpowers/specs/2026-06-23-backlog-master-spec.md` · **Plane:** DATAVIEWER-6 (+ two new items to create, see §11).

> Output of the MS-4 Option-B design spike (5-agent code-grounded pass: teardown surface + Notes panel + View-Raw-Data + a Cleanup→plots→reports robustness audit, then an adversarial completeness critic). **Owner decided Option B** (remove the TPM table). The critic caught real coupling the naive teardown would have broken and **three live cleanup bugs**; this sub-spec is the authoritative plan for MS-4. Where it and the master-spec MS-4 prose disagree, this wins.

---

## 1. Decision & scope (owner, 2026-06-24)

**Option B:** the TPM-mode data table (`m_dataTable`) is **removed entirely** — display + per-cell edit + live multi-user sync + add/remove-row + in-app raw/SOP-sheet rendering. Rationale (owner): TPM data is authored in a separate Excel and uploaded only to view plots/overview; in-app TPM editing/sync was never used. **TPM mode only** — Sensory/Detailed input + live-sync **stay** (separate panels, not `m_dataTable`).

The TPM center becomes: **PlotWidget fills the center + a read-only right-side Notes panel** (per-sample, `Note N: Puff <puffs>, Current TPM <tpm>, Average TPM <averageTPM>, <notes>`). Plus two additions:
- **"View Raw Data" button in the Tools ribbon tab** — opens the **Navigator-selected active file's** source Excel (TPM `.xlsx`; Sensory/Detailed source Excel, not the JSON); disabled when there is no source Excel (DB-loaded session).
- **Keep + harden the plot Cleanup** — robustly remove out-of-constraint data; plots **and** reports reflect **only** cleaned data. (Cleanup is fundamental to report generation.)

---

## 2. Phasing (the spike's recommended order — cleanup fixes FIRST)

> The critic's key sequencing insight: the cleanup-correctness bugs (§3) are **independent of the teardown** and are **live today**, and reports must reflect cleaned data — so fix them **first, on their own, with tests**, before touching the table.

- **Phase 2b.0 — Cleanup correctness fixes** (GAP-A, GAP-B, GAP-F) + tests. No teardown. Internal patch.
- **Phase 2b.1 — Build the new center additively**, table kept compiled-but-hidden so behavior can be diffed: add the read-only Notes `QListWidget`, re-orient `m_centralSplitter` to horizontal `Plot | Notes`, re-parent the sample-nav bar, populate Notes from the **cleaned** sample.
- **Phase 2b.2 — Remove the table surface in one atomic commit** (§5), `-Werror` clean after each deletion.
- **Phase 2b.3 — Add "View Raw Data"** (§6).
- **Phase 2b.4 — Raw/SOP empty-state** (§7).
- **Phase 2b.5 — Verify** (§10): build + full suite + new Notes test + manual smoke + visual sign-off.

This replaces the master-spec plan's single "Phase 2b" with a 6-step sub-sequence. MS-4 was always "its own multi-phase plan" — this is it.

---

## 3. Cleanup robustness "full pass" (the owner's explicit ask)

**What is already correct** (verified end-to-end): exclusion semantics for a single file/sheet (the dialog stores **full** `sr.rows` indices, `buildCleanedSample` removes by the same index, `calculateMetrics`/`computeSheetAggregates` genuinely recompute on survivors — `DataCleanupDialog.cpp:277-284,392-398`, `MainWindow.cpp:6878-6904`, `SheetProcessors.cpp:225-341`); **the UI plots DO reflect cleaned data live** (cleaned sheet → `PlotWidget::setSheetData` → `m_currentSheet`, every render branch reads it incl. the `tpmTrend`/`puffCounts` fast path which `buildCleanedSheet` rebuilds; `refreshPlotRegimes` re-renders off the stored cleaned copy — `MainWindow.cpp:3954-3961`, `PlotWidget.cpp:157-159,435-469`, `MainWindow.cpp:6649-6652`); **reports use cleaned data for the current file** (`onGenerateTestReport`/`onGenerateFullReport` feed `buildCleanedFile` → `ReportGenerator`, which also appends the `cleanupNote` audit — `MainWindow.cpp:4113-4115,4171-4175`, `ReportGenerator.cpp:330-341`); all-rows-excluded degrades to zeroed aggregates (no crash); cross-sheet + re-include are correct.

**Live bugs to FIX (Phase 2b.0):**
- **GAP-A (HIGH, mis-applies constraints across files):** `buildCleanedFile(const FileResult&)` hardcodes `exclusionsFor(m_currentFileIndex, …)` (`MainWindow.cpp:6906-6917`, the lookup at `:6914`) but `onGenerateFullReport` calls it **per selected file** (`:4171-4175`). So in a **multi-file Full report or the Combined report, every file ≠ the current one gets the WRONG file's exclusions (or none)** → reports do **not** reflect each file's own cleanup. **Fix:** pass `fileIdx` into `buildCleanedFile` (derive the index from the passed file), or key exclusions by a stable file identity. **Test:** multi-file Full report asserts file B's exclusions apply to file B.
- **GAP-B (HIGH, corrupts which rows are excluded):** `closeFile` removes only the closed file's keys then `m_loadedFiles.removeAt(idx)` shifts every higher file down one — surviving keys `"k:…"` for `k>idx` now **alias the wrong file** (`MainWindow.cpp:2511-2518`; `m_excludedRows` is positional, `MainWindow.h:324`; files are append-only). **Fix:** re-key surviving entries after `removeAt`, or switch `m_excludedRows` to a stable key (normalized file path). **Test:** close a lower-index file, assert a higher file's exclusions still target it.
- **GAP-F (HIGH, discards manual exclusions on reopen):** re-opening `DataCleanupDialog` **wipes prior manual per-row choices** — the ctor loops `applyThresholdToSample(i)` (`DataCleanupDialog.cpp:205-208`) which `m_exclusions.remove(si)` **first** then re-applies the default TPM∈[1.0,50.0] threshold, ignoring the passed `initExclusions`. **Fix:** seed from `initExclusions`; auto-apply the threshold only on **first** open (or when no exclusions exist yet). **Test:** set manual exclusions, reopen the dialog, assert they survive.

**Decisions / lower-priority (§9 + hardening):**
- **GAP-C (durability):** exclusions are **in-memory only** — never serialized to the recovery snapshot or Postgres (`MainWindow.cpp:2511-2516`, no serializer). On close/reload or recovery restore, a reopened file **silently reverts to raw** with no warning. **Owner decision (§9):** persist exclusions (so reports authored in a later session honor cleanup) or accept session-only.
- **GAP-D (Notes/cleanup consistency — must fix with the panel):** the new Notes panel must be fed the **SAME cleaned sample** the plot uses, not the raw one — else with cleanup active it would still list excluded rows and print the raw `averageTPM`, disagreeing with the plot + report. Build the cleaned sheet **once** and feed plot + properties + Notes from it; renumber `Note N` over cleaned visible rows (`MainWindow.cpp:3954-3961`, `6860-6871`).
- **GAP-E (cosmetic):** `buildCleanedSample`'s `++rowNum` is computed but unused; the `cleanupNote` identifies rows by puff# not visible-row index (`:6862-6871`). Reconcile only if the Notes panel + cleanup note must cross-reference by the same index.

---

## 4. The new center — PlotWidget + right-side Notes panel

- **Layout:** the TPM page is the single vertical `m_centralSplitter` holding `m_tablePanel` (top) + `m_plotWidget` (bottom), added at `m_centralStack` index 0 (`MainWindow.cpp:860-1017`, `MainWindow.h:248`). **Re-orient `m_centralSplitter` to horizontal** holding `PlotWidget` (dominant stretch) + the new Notes panel. Reusing `m_centralSplitter` means the mode-return lines `m_centralStack->setCurrentWidget(m_centralSplitter)` (`:4307/:4348`) need no change (critic: keep the TPM page identity).
- **Re-parent (do NOT delete) the sample-nav bar** — `m_prevBtn`/`m_nextBtn`/`m_sampleCountLabel`/`m_sampleNavBar` live inside `m_tablePanel` (`:871-932`) but sample navigation **survives** (samples still have plots+notes). Re-parent them into the new Plot|Notes container; `onPrevSample`/`onNextSample` (`:3237-3249`), `updateSampleNav` (`:3982-3992`) and the Ctrl+Left/Right shortcuts (`:1564-1569`) stay.
- **Notes data source (zero new plumbing):** `DataRow.puffs/.tpm/.notes` + `SampleResult.averageTPM` (`ReportData.h:15,23,24,57`). Populate in `displayCurrentSample` over the **visible-row filter** (skip `beforeWeight==0 || afterWeight==0`, `:3829-3853`), one entry per visible row, **from the cleaned sample** (GAP-D). `ReportGenerator` already aggregates per-row notes (`:330-341`) so the panel is a pure on-screen echo.
- **Styling:** echo the left-dock header recipe (`propHeader` at `MainWindow.cpp:1309-1313`: `AppTheme::tableHeader()` bg, white bold 8pt) over a word-wrapped `QListWidget` (font 8pt, item padding — `m_testAvgList` at `:1256-1261` is the template). Word-wrap mechanism (`setWordWrap` vs item widgets) chosen at implementation.
- **Empty-state clears:** the empty-sample branch (`:3785-3797`), and the all-files-closed reset (`:2536-2538`) must also clear the Notes panel where they clear `m_dataTable` today.

---

## 5. Teardown surface (Phase 2b.2) — verified removable vs must-keep + the critic's catches

**~99 `m_dataTable` references.** Removable set: the table-panel block `setupCentralWidget:862-1004`; the 6 connects `:952-992`; `onTableCellChanged` (`:1933-2052`); `onAddRow`/`onRemoveRow` (`:6939-7035`); `onDataTableItemChanged`/`onDataTableItemClicked`/`clearRemoteDecoration`/`findTableRowForDataRowId` (`:3425-3511`); `onRemoteCellChanged`/`Focused`/`Blurred` (`:3653-3728`) + their 3 connects (`:372-377`); the raw/SOP in-table render + strikethrough in `displayCurrentSample` (`:3740-3811, 3909-3948`); both delegates `CellFocusDelegate` + `RegimeComboDelegate` (`.cpp/.h`); orphaned members/helpers `m_applyingRemote` (`MainWindow.h:321`), `dataTableHeaders` (`:7341`), `liveColumnForDataCol`/`dataColForLiveColumn` (`:6627-6641`), the file-static `kDataTableColumns`/`columnNameForDataTableColumn`/`dataTableColumnForColumnName` (`:85-112`).

**Critic's catches the grounding missed (the dangerous ones):**
- **`onPropCellChanged` TPM branch writes to `m_dataTable` — THE BIGGEST MISS.** The Sample-Properties handler is **MUST-KEEP** (sensory uses it), but its TPM branch (`MainWindow.cpp:2100-2199`) pushes calc columns into `m_dataTable` at `:2168-2187`. **Rework** it: drop the `m_dataTable` writes, refresh the Notes panel instead, **keep** the cleanup-aware plot push at `:2189-2194`.
- **`buildViewTab` / the "View" tab is ALREADY dead code.** `setupRibbon` (`:601-604`) registers only Home/Reports/Tools/Settings — there is **no** `addTab("View")`. So `buildViewTab` (`:790-803`), `onViewDataTable`/`onViewPlots`/`onViewBoth` (`:4270-4272`), `onZoomIn`/`Out`/`FitToWindow` (`:4273-4275`, empty stubs) are dead and reference `m_tablePanel` — **delete them** (the master-spec acceptance "retire the Layout group" is **moot** — there is no live group).
- **Orphan members for `-Werror`:** `m_applyingRemote` (read/set only by removed handlers); `flushPendingEdits`/`struct PendingEdit`/`m_pendingEdits` — the only producer is `onTableCellChanged:2039`, but `flushPendingEdits` is still **called** from `onConnectionCameOnline:6227`, so either delete the whole unit (struct + member + fn + 3 call sites) **together** or risk `-Wunused-private-field` on `m_pendingEdits`.
- **`handleRemoteRowChange` must be SPLIT, not deleted:** keep the T19 row-deleted-banner head (`:3563-3585`, drives `m_rowDeletedBanner` for file/sensory/detailed — **not** table-specific); drop the T18 `data_rows` tail (`:3587-3648`). `m_liveSync` stays (sensory uses it independently).
- **MUST-KEEP:** `m_liveSync`, `onPropCellChanged` (reworked), `m_propTable`/Sample-Properties panel, the sample-nav widgets (re-parented), the whole-file DB save of `data_rows` (`DatabaseManager.cpp:1079-1117` + `tryWriteFile` — the "upload .xlsx → view plots" round-trip is preserved; TPM data still persists/reloads).

---

## 6. "View Raw Data" Tools-tab button (Phase 2b.3)

- **Where:** `buildToolsTab` (`MainWindow.cpp:805-818`) — copy the existing `addGroup` + `addLargeButton` pattern. Store the button as a new member `m_viewRawDataBtn` so `updateRibbonForMode` (`:772`) can toggle its enabled/tooltip on mode/selection change. API: `RibbonGroup::addLargeButton(text, icon, tooltip) → QToolButton*` (`RibbonWidget.h:30-38`; disabled styling already exists `:159-161`).
- **Resolve the active file (Navigator-selected, per owner):** branch on `currentReportMode()` (`:2423-2428`):
  - **TPM →** `m_loadedFiles[m_currentFileIndex].filePath` (`MainWindow.h:313,325`, `ReportData.h:134`). The Navigator selection drives `m_currentFileIndex` — **verify it stays updated on file/plot selection once the table's tree-driven selection is gone** (open gap).
  - **Sensory →** `m_sensoryPanel->currentSession()->sourceFilePath` (`SensoryPanel.h:179`, `SensoryData.h:113` "empty if DB-only").
  - **Detailed →** `DetailedSensorySession` has **NO `sourceFilePath`** field (`DetailedSensoryData.h:145-187`). To open Excel in Detailed mode (which the owner wants), **add the field + set it in the Detailed `.xlsx` loaders** (`DetailedSensoryPanel.cpp loadFile/loadFiles`). See §9.
- **Open:** `QDesktopServices::openUrl(QUrl::fromLocalFile(path))` (already included `MainWindow.cpp:52-53`; pattern used at `:4246,4262`).
- **Disable when no source Excel:** path empty (DB-loaded; `SensoryPanel::loadSessions` never sets `sourceFilePath`, `:1479`) or `!QFileInfo(path).exists()` → `setEnabled(false)` + tooltip "No source Excel for this session (loaded from the database)". Recompute in `updateRibbonForMode`.

---

## 7. Raw/SOP sheet empty-state (Phase 2b.4)

`displayCurrentSample:3740-3779` is the **only** in-app view of `isRawTable` sheets (rendered into `m_dataTable`). After Option B, selecting a raw/SOP sheet leaves a **blank center** (`m_plotWidget->hide()` at `:3745`, no table). Owner intentionally drops in-app raw rendering → **spec the blank-center UX:** when a raw/SOP sheet (or no samples) is selected, clear the Notes panel and show a hint ("Use **View Raw Data** to open this sheet in Excel") in the Notes panel or a center placeholder.

---

## 8. Test impact

- **`tst_cellfocusdelegate`** compiles `CellFocusDelegate.cpp` directly (`tst_cellfocusdelegate.pro`); `RegimeComboDelegate` inherits from it. Both delegates are removed → **remove the test + its `tests/tests.pro:4` SUBDIRS entry**, else the build breaks.
- **`tst_mainwindow_remotecell`** tests `applyRemoteValueToCell` on a bare `QTableWidget`; its sole production caller (`onRemoteCellChanged:3670`) is removed. `RemoteCellHelpers.{cpp,h}` + the test + `tests/tests.pro:40` are a **delete-together unit** OR keep the helper+test as harmless — **owner/impl decision (§9)**.
- **`tst_twoclient_e2e` / `tst_saveintegrity_e2e`** exercise `data_rows` at the DB/LiveSync layer (no `MainWindow`) → **stay green**.
- **New:** a `MainWindow` Notes-population test (one entry per visible row, **cleaned** values matching `ReportGenerator`). Plus the GAP-A/B/F regression tests (§3).

---

## 9. Owner decisions needed (sign-off gate)

a. **Detailed-mode "View Raw Data":** you said all 3 modes open the source Excel. Detailed has no `sourceFilePath` today → **add the field to `DetailedSensorySession` + set it in the Detailed loaders (recommended, matches your ask)**, vs. ship the button **disabled** in Detailed mode. Confirm: add the field?
b. **GAP-C — persist cleanup exclusions?** Today they vanish on close/reload (a report authored next session silently loses cleanup). **Persist to the recovery snapshot + DB (recommended, since cleanup is critical to reports)** vs. accept session-only. Your call.
c. **`RemoteCellHelpers` / `tst_mainwindow_remotecell`:** delete with the table (cleaner) vs. keep (harmless). Default: delete.

---

## 10. Acceptance criteria

- Build `-Werror -Wall -Wextra -Wpedantic` clean (no `-Wunused` from `m_applyingRemote`/`m_pendingEdits` orphans); full Qt suite green (accounting for removed/kept `tst_cellfocusdelegate` + `tst_mainwindow_remotecell`; known-flaky `tst_responsivelayout` excluded).
- **Cleanup (Phase 2b.0):** multi-file Full report applies **each file's own** exclusions (GAP-A); closing a lower file doesn't corrupt a higher file's exclusions (GAP-B); reopening the dialog preserves manual exclusions (GAP-F) — all with regression tests.
- **TPM mode:** PlotWidget fills the center, Notes panel on the right echoing the Navigator styling; one entry per **visible, cleaned** row with values matching `ReportGenerator` (incl. the zero-weight filter); `Note N` numbering matches the cleaned plot/report when cleanup is active (GAP-D).
- Mode switches intact (`m_centralStack->setCurrentWidget(m_centralSplitter)` still valid); sample-nav (Prev/Next + Ctrl+arrows) works; raw/SOP sheet shows the empty-state hint.
- **View Raw Data:** opens the **Navigator-selected active file's** Excel for the active mode; disabled (with tooltip) for a DB-loaded session with no source Excel.
- New Notes-population test green; `tasks/lessons.md` appended: "the TPM data grid was the editing + live-collaboration + offline-capture + raw-sheet-render + cleanup-strikethrough surface, and `onPropCellChanged` also wrote it — removing it touches all five."

---

## 11. Plane items to create (deferred — MCP connector down)

1. **DATAVIEWER-6** → record Option B + this sub-spec + the 6-phase plan (2b.0–2b.5).
2. **New: "View Raw Data" Tools-tab button** (`type:feature`, `area:ui-ribbon`, all 3 modes; opens Navigator-selected active file's Excel; disable when none; Detailed needs a new `sourceFilePath` field).
3. **New: Cleanup robustness pass** (`type:bug` for GAP-A/B/F — live correctness defects in multi-file reports / file-close / dialog-reopen — `area:data`/`area:reports`; + the GAP-C persistence decision). These are independent of the teardown and should land first.
