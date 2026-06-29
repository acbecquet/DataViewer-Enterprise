# DATAVIEWER-6 / MS-4 — Editable TPM "Story" Notes panel: Design

**Date:** 2026-06-24 · **Status:** DESIGN APPROVED (owner, 2026-06-24, via brainstorm + mockups) · **Parent:** MS-4 in `docs/superpowers/specs/2026-06-23-backlog-master-spec.md` · **Plane:** DATAVIEWER-6 (this) + DATAVIEWER-16 (cleanup robustness) + DATAVIEWER-17 (View Raw Data) + DATAVIEWER-18 (Stage-2 plot→note linking).

> **Supersedes the read-only parts of** `2026-06-24-DATAVIEWER-6-MS4-optionB-spike.md`. That spike specced a **read-only** notes echo and proposed deleting the TPM edit + live-sync machinery. **The owner REVERSED that:** the panel is **fully editable** with the same live-edit behaviour as the old table — a *re-presentation, not a capability removal*. Where this doc and the spike disagree on editability or panel design, **this wins**. The spike still governs the teardown-surface inventory (its §5) and the three live cleanup bugs (its §3 → now DATAVIEWER-16).

> **Knowledge foundation (owner directive):** build on **ponytail** (over-engineering "what to delete" discipline — gate the diff on `/ponytail-review`: reuse existing plumbing, no speculative abstractions, no new deps; never trim validation) and **andrej-karpathy-skills** (think-before-coding, simplicity-first, surgical-changes, goal-driven verification).

---

## 1. Goal

Turn the "useless" TPM data grid into a per-sample **story of the device's lifetime, puff 1 → last puff** — and keep it **fully editable**, live-synced **1:1 to the source Excel + Postgres**. Notes aren't entered on every row, so only note-bearing rows get detail; the stretches between are summarised.

## 2. Scope & the "re-presentation, not removal" rule

- **TPM mode only.** Sensory / Detailed-Sensory are untouched (separate panels).
- The `m_dataTable` `QTableWidget` **widget** goes away. **All of its behaviour stays and re-points at the new panel:** per-cell edit → `DataRow` mutation → recalc → **Excel write-back** (`queueExcelWrite`/`writeCellsToExcel`) + **Postgres `data_rows` persistence** + **per-cell LiveSync** (`commitCell`) + **remote-edit painting**.
- **Do NOT delete** `onTableCellChanged`, `onDataTableItemChanged`→`LiveSync::commitCell`, the remote-cell handlers, `RemoteCellHelpers`, or the delegates — **repurpose** them to drive the panel's editing. Watch the easy miss: `onPropCellChanged` also writes `m_dataTable` (`MainWindow.cpp:2168-2187`) — rework it to refresh the panel, keep the cleanup-aware plot push. (Teardown inventory: spike §5.)

## 3. Layout

- **Plot fills the centre; the editable Story panel sits on the right**, mirroring the left Navigator dock. Re-orient `m_centralSplitter` to horizontal `Plot | Panel` (keeps the TPM-page identity so the mode-return lines need no change).
- **Re-parent (do not delete) the sample-nav bar** — Prev/Next + `1/N` + Ctrl+Left/Right cycle samples (samples still have plots + notes).
- **The panel is a `QScrollArea`: vertical scroll, width-constrained.** **HARD owner constraint — zero horizontal overflow off the panel or screen, ever.** Cards word-wrap; chip rows wrap; the expanded mini-table uses fixed/min column widths and small (≥11px) fonts.

## 4. Per-sample story content

For the current sample (prev/next cycles samples), walk the rows in order:

- **Rows that HAVE a note → a detailed, editable note card.**
- **Contiguous runs of non-note rows between notes → one small-font summary bar.** Clicking its chevron expands it into a **compact editable mini-table** of those individual rows ("compacted table; highlights shown by default").

### 4.1 Note card

- **Header:** `Puff N` pill + a "live" indicator.
- **Context line (read-only):** `TPM <tpm> · avg <averageTPM> · var <variationTPM>`.
- **Editable note text** — multi-line, wraps.
- **Field chips:**
  - **Read-only CONTEXT (quantitative):** `TPM before` (row above's `tpm`), `TPM after` (row below's `tpm`) — shows the change the note records — and `Draw` pressure.
  - **EDITABLE (qualitative):** `Smell` (0–4 scale), `Clog` (Y/N), `Puffing regime`.
- **The editability rule (owner):** qualitative fields are editable — they can be annotated *after* a test (`notes`, `smell`, `clog`, `puffingRegime`, all `QString`). Quantitative values (`tpm`, `drawPressure`, both `double`/calculated) are **read-only context** sourced from the test template — to correct a measurement, use **View Raw Data** (§7).

### 4.2 Summary bar + expansion

- **Collapsed:** `Puffs A–B · avg <avgTPM> · var <var>` (small font).
- **Expanded:** a compact editable grid of the individual rows — columns `Puff | TPM (context) | Draw (context) | Smell (edit) | Clog (edit) | + add-note`. **Adding a note promotes that row to its own card.** All within the panel width.

### 4.3 Editable = live, 1:1 to the file

Every editable change routes through the **existing** plumbing (§2): `DataRow` mutation → recalc → Excel write-back + `data_rows` DB persist + LiveSync + remote-edit painting. The panel is simply the new presentation of the same per-cell edit surface.

## 5. Cleanup / excluded rows (GAP-D)

- The panel **still shows excluded rows, clearly MARKED as excluded** (subtle tint + an "excluded" tag) **but fully READABLE** — the owner wants to read the note + data because the note usually records **why** the row was excluded.
- **Summary aggregates (avg TPM, var) compute over INCLUDED rows only**, matching the plot + report. Build the cleaned sheet **once** and feed plot + panel from it.
- Depends on the cleanup-correctness fixes (DATAVIEWER-16, Phase 2b.0) landing first.

## 6. Plot ↔ note linking (STAGED — owner approved)

- **v1 (this release):** `PlotEngine` always draws a **ring** at each note-row's `(puffs, tpm)`; clicking a note **card** emphasises its point (re-render with a selected-puff ring + guide line) and lights up the card. `PlotEngine` gains: a set of note-puffs to ring + an optional selected puff. No mouse/inverse-transform needed.
- **Stage 2 (next sprint — DATAVIEWER-18):** **plot → note.** Clicking near a point on the plot image finds the nearest note and scrolls/highlights its card. Needs `PlotEngine` to expose its axis→pixel transform and `PlotWidget` to hit-test clicks. Owner wants to fine-tune this — intentionally decoupled.

## 7. "View Raw Data" (DATAVIEWER-17)

Tools-ribbon button opens the active file's **source Excel** (TPM `m_loadedFiles[m_currentFileIndex].filePath`; Sensory `currentSession()->sourceFilePath`; **disabled in Detailed** — no `sourceFilePath` yet). This is the path to correct the quantitative/measurement data the panel intentionally does not edit.

## 8. Data source (no new plumbing)

`DataRow.puffs / .tpm / .variationTPM / .drawPressure / .smell / .clog / .puffingRegime / .notes` + `SampleResult.averageTPM` (`src/pipeline/ReportData.h`). `ReportGenerator.cpp:330-341` already aggregates per-row notes — the panel is a presentation over data that already exists.

## 9. Phasing

- **2b.0 — Cleanup fixes** (GAP-A/B/F) + persist exclusions + one-click "Undo all cleanup" — **DATAVIEWER-16** (in progress, background agent).
- **2b.1 — Build the panel additively** behind the still-compiled, hidden table so behaviour can be diffed.
- **2b.2 — Remove the table widget** in one atomic commit; keep the plumbing; `-Werror` clean after each deletion.
- **2b.3 — Wire "View Raw Data"** — DATAVIEWER-17.
- **2b.4 — Raw/SOP sheet empty-state** (selecting a raw/SOP sheet → "Use View Raw Data to open this sheet in Excel").
- **2b.5 — Plot note-rings + note→plot emphasis; verify.**
- **Stage 2 — plot→note** — next sprint (DATAVIEWER-18).

## 10. Tests

- **Notes-population:** one card per note-row; summary bars over the in-between runs; excluded rows shown + marked; aggregates over included rows match `ReportGenerator`.
- **Edit→sync:** editing note / smell / clog / regime mutates `DataRow` and fires the Excel + LiveSync paths; quantitative fields are not editable.
- **No-overflow / fit:** panel content stays width-bounded (where unit-testable).
- Plus the inherited teardown test impacts (spike §8): `tst_cellfocusdelegate` (delegates removed), `tst_mainwindow_remotecell` / `RemoteCellHelpers` (KEPT per owner decision #3 — repurposed for the panel's live-sync).

## 11. Acceptance criteria

- Build `-Werror -Wall -Wextra` clean; full Qt suite green (known-flaky `tst_responsivelayout` excluded).
- TPM mode: plot fills the centre; editable Story panel on the right; one detailed editable card per note-row; in-between runs summarised and expandable into an editable mini-table; **no horizontal overflow at any panel width**; vertical scroll works.
- Editable = note/smell/clog/regime, live-synced to Excel + DB 1:1; TPM + draw pressure read-only context.
- Excluded rows shown, marked, and readable; aggregates over included rows match plot + report.
- Plot rings note-points; clicking a note emphasises its point (v1).
- Sample-nav (Prev/Next + Ctrl+arrows) intact; mode switches intact; raw/SOP sheet shows the empty-state hint.

## 12. Open / to confirm during build

- Smell 0–4 input control (spin box vs combo).
- Puffing-regime editor (combo of known regimes vs free text).
- Exact promote-on-add-note flow from an expanded summary row.
