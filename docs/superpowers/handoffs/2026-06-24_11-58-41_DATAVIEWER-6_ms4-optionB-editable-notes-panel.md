---
date: 2026-06-24T11:58:41-0700
researcher: Charlie (becquetcharlie) + Claude Opus 4.8
git_commit: 9811fd0
branch: feature/v2.6.0-backlog-batch
repository: DataViewer-Enterprise
topic: "v2.5.x→v2.6.0 backlog batch — MS-4 Option B (editable Notes panel), Phase-1 done, DV-11/12/13/15"
tags: [implementation, strategy, ms4, notes-panel, tpm, cleanup, dataviewer-6, dataviewer-11, plane]
status: complete
last_updated: 2026-06-24
last_updated_by: Claude Opus 4.8
type: implementation_strategy
---

# Handoff: DATAVIEWER-6 / MS-4 Option B — the Notes panel must be EDITABLE (reverses the spike), + Phase-1 done

## 🚨 CRITICAL DIRECTIVES — READ THESE FIRST (owner, 2026-06-24)

### 0. Knowledge foundation (use these as your basis for ALL work in the new session)
- **ponytail** — https://github.com/DietrichGebert/ponytail (over-engineering review / "what to delete" discipline; this repo's CLAUDE.md already gates non-trivial diffs on `/ponytail-review`).
- **andrej-karpathy-skills** — https://github.com/multica-ai/andrej-karpathy-skills (the engineering principles this project runs on: think-before-coding, simplicity-first, surgical-changes, goal-driven verification).
Read/internalize both before touching code.

### 1. ⚠️ MS-4 reversal: the Notes panel is FULLY EDITABLE — NOT read-only
This **overrides** the just-written MS-4 Option-B sub-spec (`docs/superpowers/specs/2026-06-24-DATAVIEWER-6-MS4-optionB-spike.md`), which specced a **read-only** Notes panel and **removing** the TPM editing + live-sync machinery. **That is now WRONG.** The owner's words:

> "the notes side panel should be **FULLY EDITABLE** and have the **exact same live edit functionality as the TPM table**. the only difference is that it's presented in a better way; there should still be a **1:1 link between the file and this section**. NOT READ ONLY!!!"

So Option B = **replace the TPM grid's PRESENTATION with a better, editable Notes panel — but KEEP all of its behavior**: per-cell edit → `DataRow` mutation → recalc → **Excel write-back** + **DB persistence** + **per-cell LiveSync** + remote-edit painting. The teardown is a **re-presentation, not a capability removal.**
- **Do NOT delete** `onTableCellChanged`, `onDataTableItemChanged`→`LiveSync::commitCell`, the remote-cell handlers, `RemoteCellHelpers`, or the delegates as the spike's §5 proposed — **repurpose** them to drive editing in the Notes panel (this flips decision #3 below). The `m_dataTable` `QTableWidget` widget goes away; the edit/sync/write-back **plumbing behind it stays and re-points at the new panel.**
- The **1:1 file↔panel link** means edits in the panel write back to the source `.xlsx` (the existing `queueExcelWrite`/`writeCellsToExcel` path) and to Postgres `data_rows` (the existing `commitCell` path), exactly as the table did.

### 2. ⭐ The Notes-panel VISION (this is the actual goal — "turn the useless table into something valuable")
The panel should **tell the story of the device's lifetime, puff 1 → last puff**:
- Notes are **not** entered on every row. **Only show detailed (editable) data on rows that HAVE notes.**
- For the stretches **in between** notes, just summarize: **number of puffs, average TPM, and maybe variation** — a compact "puffs 5–18: avg TPM 3.4, var 0.2" style roll-up.
- Rows with notes get the **detailed, editable** treatment.
- **Bidirectional plot ↔ note linking (exploratory, owner wants to fine-tune):** click a note → draw a **highlighted circle at that point on the plot**; click a plot point → **highlight the corresponding note**. Owner: "I am not sure how this will work too" — so this needs design iteration (brainstorm with the owner; prototype).
- Expect **fine-tuning** — get it in front of the owner early and iterate on the look/behavior.

### 3. The three decisions (owner-resolved)
1. **"View Raw Data" in Detailed Sensory:** **ship it DISABLED in Detailed for now** — revisit later (Excel files *should* eventually be created from Detailed sessions, but not yet). TPM + Sensory open the source Excel; Detailed button stays disabled.
2. **Persist cleanup exclusions: YES** — but there **must be a one-click "Undo all cleanup" button** that clears every cleanup change at once (easy full reset).
3. **`RemoteCellHelpers`:** **do NOT delete** — "we may need to repurpose this to work with the notes panel on the side." Keep it for the editable Notes panel's live-sync.

### 4. Plane is NOT updated yet (do this in the new session)
The Plane MCP connector dropped mid-session and would not re-expose its tools to the running session, so **the Plane writes are still pending**. A fresh session should reconnect Plane and:
- **DATAVIEWER-6:** record Option B + the **editable-notes** reversal + the 6-phase plan; move as the owner wants.
- **New issue — "View Raw Data" Tools button** (`type:feature`, `area:ui-ribbon`; opens the Navigator-selected active file's Excel; disabled in Detailed for now).
- **New issue — Cleanup robustness pass** (`type:bug`; GAP-A/B/F live bugs + persist + the Undo-all button; `area:data`/`area:reports`).
- (DATAVIEWER-15 "unify the 3 sensory `session_name` derivations" was already created.)

---

## Task(s)

Sprint **v2.5.x internal → v2.6.0 deployable**, branch `feature/v2.6.0-backlog-batch` (off `main` @ `v2.5.0`). Anchored by the master spec + multi-phase plan (see Critical References). Per-item status:

- **DATAVIEWER-12** (MS-3, "Save plot" labels) — **triaged, in Plane Todo. Phase-1 CODE DONE** (built + compiles clean). Awaiting owner **visual sign-off**.
- **DATAVIEWER-13** (MS-5, sensory stopwatch) — triaged, in Plane Todo. Not yet implemented (Phase 2a).
- **DATAVIEWER-11** (MS-8, phone sensory web form) — **spiked + in Plane Todo.** Decisions locked: anonymous one-way web UX (no user password; all controls backend), **create-or-append by natural key**, dedicated least-privilege `sensory_web` DB role, **+ a REQUIRED desktop merge-safety fix** (so a desktop save can't drop a phone-appended sample). Sub-spec written. Not built.
- **DATAVIEWER-15** — created (unify the 3 sensory `session_name` derivations; a DV-11 prerequisite). Not started.
- **MS-7** (DATAVIEWER-9, LiveSync authority) — **resolved: MINIMAL** (keep Ctrl+U + both write paths; only fix the `oil_smell_liking` session-field merge gap). Not implemented.
- **MS-4 / DATAVIEWER-6** (this handoff's focus) — **Option B, spiked.** BUT the sub-spec must now be **revised for the EDITABLE Notes panel** (§1 above) before any code.
- **Phase-1 (MS-1 plot y-axis @0, MS-2 ribbon/image clipping, MS-3 Save-plot labels)** — **CODE DONE, built clean `-Werror`, MS-1 tests green** (`tst_plotengine` 14/0 incl. new `autoRange_anchorsYAtZero`; `tst_reportgenerator` 21/0 incl. new `yMax_nonLongPuff_peakAboveCeilingAvgBelow`). **UNCOMMITTED — awaiting owner visual sign-off**, then it commits separately.

## Critical References
- `docs/superpowers/specs/2026-06-23-backlog-master-spec.md` — THE anchor (MS-1..MS-8; spec wins over plan).
- `docs/superpowers/plans/2026-06-23-backlog-resolution-plan.md` — the multi-phase plan.
- `docs/superpowers/specs/2026-06-24-DATAVIEWER-6-MS4-optionB-spike.md` — the MS-4 Option-B spike. **⚠ Read it for the teardown/cleanup grounding, but treat its "read-only Notes panel" + "remove editing/live-sync" as SUPERSEDED by §1 of this handoff (editable).**
- `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md` — DV-11 sub-spec.

## Recent changes
- Doc-decision batch committed (docs only): `9811fd0` "owner decisions round 2 + MS-4 Option-B spike" on `feature/v2.6.0-backlog-batch` (NOT pushed). Prior: `7d2e027` (Phase-0 planning docs).
- **Phase-1 code (UNCOMMITTED on disk):** `src/plotting/PlotEngine.cpp` + `.h` (autoRange yMin=0 + dual-axis right axis + test accessor), `src/reporting/ReportGenerator.cpp` (`computeTpmYMax` → `qMax(7.0, maxTPM+1.0)`), `src/MainWindow.cpp` (ribbon label "Detailed Sensory Path"; `imgBar` height 32→40), `src/widgets/RibbonWidget.h` (stale 56×70→80×76 comment), `src/plotting/PlotWidget.cpp` + `src/ui/SensoryPanel.cpp` (`tr("Save plot")` QLabel by each save button), `tests/tst_plotengine/*` + `tests/tst_reportgenerator/*` (new/updated assertions).

## Learnings
- **The TPM `m_dataTable` is the editing + live-collaboration + offline-capture + raw-sheet-render + cleanup-strikethrough surface, AND `onPropCellChanged` also writes it** (`src/MainWindow.cpp:2168-2187`). With the editable-Notes reversal, do **not** rip out that behavior — re-point it at the panel. (~99 `m_dataTable` references; `onPropCellChanged` is the easy miss.)
- The "View/Layout" ribbon tab is **already dead code** (`setupRibbon` registers only Home/Reports/Tools/Settings; no `addTab("View")`). `buildViewTab`/`onViewDataTable/Plots/Both`/`onZoomIn/Out/Fit` are dead — delete them; "retire the Layout group" was moot.
- **Three LIVE cleanup bugs** (independent of the table; fix FIRST in Phase 2b.0): **GAP-A** `buildCleanedFile` hardcodes `m_currentFileIndex` (`MainWindow.cpp:6914`) → multi-file/Combined reports apply the WRONG file's exclusions. **GAP-B** `closeFile` (`:2511-2518`) shifts file indices and corrupts surviving exclusion keys. **GAP-F** reopening `DataCleanupDialog` (`DataCleanupDialog.cpp:205-208`) discards manual exclusions. Single-file cleanup is correct and UI plots already reflect cleaned data live.
- Notes data source needs no new plumbing: `DataRow.puffs/.tpm/.notes` + `SampleResult.averageTPM` (`src/pipeline/ReportData.h:15,23,24,57`); `ReportGenerator.cpp:330-341` already aggregates notes. The panel must use the **cleaned** sample when cleanup is active (was GAP-D), and with editing it must write to the **raw** `DataRow` while presenting cleaned (reconcile during design).
- "View Raw Data": `QDesktopServices::openUrl(QUrl::fromLocalFile(path))` (already included `MainWindow.cpp:52-53`); TPM path `m_loadedFiles[m_currentFileIndex].filePath`; Sensory `m_sensoryPanel->currentSession()->sourceFilePath`; **Detailed has no `sourceFilePath` field** (`DetailedSensoryData.h:145-187`) → button disabled in Detailed (decision #1).
- MIP: run `python tools/decrypt_via_copy.py --apply` before any C++ build/edit (one stubborn unrelated file `tests/excel/test_atomic_delete.py` re-labels — ignore). Build: `qmake CONFIG+=release` + `mingw32-make -j8` under `C:/Qt/6.10.1`. Test exes need the Qt `bin` on PATH + `QT_QPA_PLATFORM=offscreen`; they're windows-subsystem (no console output) so use `-o file,txt`.

## Artifacts
- This handoff: `docs/superpowers/handoffs/2026-06-24_11-58-41_DATAVIEWER-6_ms4-optionB-editable-notes-panel.md`
- `docs/superpowers/specs/2026-06-24-DATAVIEWER-6-MS4-optionB-spike.md` (revise for editable Notes)
- `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md`
- `docs/superpowers/specs/2026-06-23-backlog-master-spec.md` (MS-4 banner @ §MS-4; MS-8 banner)
- `docs/superpowers/plans/2026-06-23-backlog-resolution-plan.md` (Phase 2b → 6-step Option-B sequence)
- Raw spike output (full grounding + critic): `C:/Users/S1134987/AppData/Local/Temp/claude/.../tasks/wy9a3x1ez.output` (may be transient).

## Action Items & Next Steps
1. **Revise the MS-4 Option-B sub-spec for the EDITABLE Notes panel** (§1): the panel keeps full per-cell edit + Excel write-back + DB + LiveSync (repurpose, don't delete the plumbing/`RemoteCellHelpers`); design the "device-lifetime story" presentation (notes-rows detailed+editable, in-between rows summarized: puffs/avgTPM/variation) and the bidirectional plot↔note highlighting (§2). Brainstorm the plot-linking with the owner. Run `/ponytail-review` on the design.
2. **Update Plane** (reconnect first): DATAVIEWER-6 → Option B + editable-notes; new "View Raw Data" issue; new "Cleanup robustness" issue (DV-15 already exists).
3. **Phase 2b.0 — cleanup fixes FIRST:** GAP-A, GAP-B, GAP-F + persist exclusions + the **one-click "Undo all cleanup"** button (decision #2), each with tests. Independent of the teardown; reports depend on it.
4. **Phase-1 visual sign-off → commit** Phase-1 code (MS-1/2/3) separately; then **push** `feature/v2.6.0-backlog-batch`.
5. Then Phase 2b.1+ (build the editable Notes panel additively behind the hidden table → swap → View Raw Data → raw empty-state → verify), per the (revised) sub-spec.

## Other Notes
- **Never touch the Synology drive / release channel** (user-level standing rule). Build the installer in-repo only; the owner does the drop. Don't push to `main` or deploy without owner approval; the branch is for owner install-test.
- Sprint scoping: only the current minor's plans are active; `v2.5.0` already shipped. This batch wraps into deployable **v2.6.0**; **MS-8's web service ships on its own NAS track** (but its required desktop merge fix rides in v2.6.0).
- DV-11/12/13 are in Plane **Todo**; DV-15 in Backlog. Phase-1 (MS-1/2/3) is built+green but **uncommitted** pending visual sign-off.
- Memory file updated: `~/.claude/.../memory/dataviewer-11-phone-sensory-collection.md`.
