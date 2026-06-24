# DataViewer Enterprise — v2.5.x → v2.6.0 Backlog Master Spec

**Date:** 2026-06-23 · **Status:** FINAL — single source of truth for the v2.5.x → v2.6.0 sprint · **Branch (to cut):** `feature/v2.6.0-backlog-batch` off `main` (current tag `v2.5.0`, `.pro` VERSION = 2.5.0).
**Sources:** the seven code-grounded per-item analyses (DATAVIEWER-5/6/7/9/10/11/12/13), the v2.4.x/v2.5.0 save-sync regression history, and the project engineering principles.

> **Note on the branch name.** The deployable target of this sprint is **v2.6.0** (the `y` in `x.y.z` delimits the sprint). `main` already sits at `v2.5.0`; internal patches are v2.5.x and wrap into the deployable v2.6.0. The branch is therefore named for its *deployable target* — `feature/v2.6.0-backlog-batch` — not for the internal-patch series. Do NOT name it `feature/v2.5.0-*` (that minor already shipped).

---

## How to use this spec

This document is **THE anchor**. The companion plan (`docs/superpowers/plans/2026-06-23-backlog-resolution-plan.md`) breaks the work into ordered phases, and **every phase and task in the plan references a section ID here** (e.g. "Task 1.2 → MS-1"), or is explicitly tagged `(wrap)` / `(setup)` for cross-cutting steps that serve the whole batch. When the two disagree, *this spec wins* — update it first, then re-derive the plan.

- Each work item has a stable **MS-N** ID mapped to its **DATAVIEWER-N** Plane ticket. The mapping is not 1:1 numerically — see the table below.
- Read the cross-cutting sections (build/test standards, versioning, MIP hygiene, data-safety invariants, principles) **once**; they bind every item and are not repeated per-item.
- Items carry a **priority** and **effort**. The sequencing rationale at the end explains the order; the plan operationalizes it.
- Three items are **already-shipped / verify-only** (MS-6) or **blocked on a product decision** (MS-4, MS-7) or **gated on a design spike + network decision** (MS-8). These are flagged loudly — do not write code against them until the gate clears.

### ID map

| MS-ID | DATAVIEWER | Title | Effort | Priority | Status |
|---|---|---|---|---|---|
| MS-1 | DV-5 | Anchor plot Y-axis at 0 + stop clipping report TPM-trend peaks | S | P1 | Ready |
| MS-2 | DV-7 | Ribbon label + image-button clipping fixes | S | P1 | Ready (1 wording fork) |
| MS-3 | DV-12 | "Save plot/chart" labels on icon-only save buttons | S | P1 | Ready (Detailed = separate item) |
| MS-4 | DV-6 | Remove TPM data tables → right-side per-sample Notes panel | L (XL if option B) | P2 | **BLOCKED on product fork** |
| MS-5 | DV-13 | Per-sample stopwatch in Sensory mode | S | P2 | Ready |
| MS-6 | DV-10 | One-shot JSONB string-score backfill | S (verify-only) | P3 | **ALREADY SHIPPED — verify + close** |
| MS-7 | DV-9 | LiveSync-authoritative sensory/detailed persistence | L (XL if literal) | P3 | **BLOCKED on product fork** |
| MS-8 | DV-11 | Remote phone sensory web form → Postgres | L | P3 | Spike DONE (2026-06-24); service ships independently (+1 required desktop C++ fix in v2.6.0) |

---

## Vision / goal

Resolve the **entire** open DataViewer Enterprise backlog as one coherent sprint that:
1. Ships fast, low-risk **UI-polish** fixes (plot axis, clipping, labels) that users feel immediately.
2. Delivers the **medium UI** changes (Notes panel, stopwatch) once their product forks are resolved.
3. Closes out the **already-shipped** JSONB backfill as a verified-and-closed ticket rather than re-implementing it.
4. Touches the **risky persistence architecture** (LiveSync authority) *only* behind heavy verification and an explicit product decision — never re-litigating the deliberate v2.5.0 dual-path design by accident.
5. Stands up the **new remote-collection feature** as a standalone NAS web service (**near-zero C++** — but the spike found ONE required desktop-side merge-safety fix that ships in the v2.6.0 train; see MS-8) gated by a design spike that produces its own sub-spec, **shipped on its own independent versioning track** (the NAS Docker stack is not part of the DataViewer.exe minor).

The DataViewer.exe work ships as consecutive **internal patches (v2.5.x)** that wrap into **one deployable minor (v2.6.0)** the owner drops on Synology manually. MS-8's **web service** ships independently and is **not** part of the v2.6.0 wrap — but the spike's one required desktop-side merge fix (Phase 5.0 / sub-spec §4.1) **does** ship in v2.6.0.

---

## Cross-cutting standards (bind every item)

### Engineering principles (the floor)
1. **Think before coding** — resolve ambiguity first; name every entry point of a save/sync/dedup/load path before touching it (the v2.4.11→14 dup saga cost 4 round-trips patching the wrong path).
2. **Simplicity first** — minimum code for the *stated* problem; no speculative config, no single-use abstraction. Prefer the lowest ponytail rung that works (already-in-codebase → Qt/stdlib → platform → one line → new code).
3. **Surgical changes** — every changed line traces to the request; don't refactor unbroken code; match DVE namespace, ribbon button sizing, the no-color-reuse plotting rule.
4. **Goal-driven, verifiable** — write the reproducing test/observation first; cite the green/red count; a `VERSION` bump requires a clean rebuild.
5. **Ponytail delete-list gate** — run `/ponytail-review` on any non-trivial diff before commit; cut reinvented stdlib/Qt, needless deps, speculative abstractions, dead flexibility — **never** cut validation, security, accessibility, or **data-safety**.

**Conflict rule:** when principles collide with each other or with a detailed rule, **data-safety and correctness win** — this app writes to a shared Postgres DB and edits users' source `.xlsx` files; silent data loss is the cardinal sin.

### Build / test / verification standards
- **Build:** qmake + MinGW only (Qt 6.10.1, MinGW 13.1.0). Must be clean under `-Werror -Wall -Wextra -Wpedantic`. No warning-level downgrades to silence a single warning.
- **VERSION-bump hygiene:** `main.o` embeds the version via `-DDVE_APP_VERSION`; qmake does **not** detect `.pro` VERSION changes. Every VERSION bump → `mingw32-make clean && mingw32-make -j8`. `build_installer.bat` enforces the FileVersion match.
- **Unit/integration tests:** `tests\run-tests.ps1` (~34 Qt Test suites). DB-dependent suites need the ephemeral `postgres:16` via `tests\start-test-postgres.ps1` (auto-applies `deploy/postgres/migrations/*` after `init.sql`; OCC overloads live only in migrations). Suites skip cleanly when `DVE_TEST_PG_CONN` is unset. If Docker Desktop is off, **ask the owner to start it** — do not try to start it autonomously.
- **Known-flaky:** `tst_responsivelayout` is pre-existing flaky (tracked separately) — not a regression gate.
- **Deployment self-test:** `tests\deployment\Test-Deployment.ps1` (3 phases: install tree, in-proc `--self-test`, independent Synology probe) — run on the work machine before any deployable wrap. Add a `TestResult test*()` in `src/utils/SelfTest.cpp` for any new deployment-sensitive path.
- **UI verification:** the ribbon/browser UI has **no automated UI test** in this repo. Label/layout/visual changes (MS-1 visual sign-off, MS-2, MS-3, MS-4, MS-5) lean on the **visual loop + owner sign-off** — never claimed "done" on a green build alone.

### MIP / AIP build hygiene (work machine)
- Source `.cpp`/`.h`/`.py` are MIP-encrypted at rest (raw bytes start `%TSD-Header-###%`). **Before any C++ build:** `python tools/decrypt_via_copy.py --apply` from repo root.
- **Create new source files via the Python delete-and-rewrite convention** (remove-then-write UTF-8, `newline="\n"`) so they don't inherit a MIP label; `git add` immediately and verify the blob (`git show HEAD:<path> | head`).
- Forgetting the decrypt yields bogus "does not name a type" compile errors on the very files being edited.

### Data-safety invariants (NEVER violate)
- **Never lose a user edit.** The whole-session save path is **load-bearing**: it is the only INSERT path for new (id≤0) sessions, it persists fresh-from-disk TPM files, and it is the safety net when LiveSync is gapped. Do not delete it naively (MS-7).
- **LiveSync is per-cell + OCC; whole-session save is read-merge-write under `FOR UPDATE` (row-level last-writer-wins by design).** Changing which path owns a field changes cross-client conflict semantics — requires explicit decision + e2e proof.
- **Scores are JSON numbers (1–9), never strings.** String scores are the DATAVIEWER-4 "reset-to-5" data-loss class. Any new writer (MS-8 web form) emits numbers.
- **Excel write-back must be atomic** (tmp + `os.replace`); a mid-write kill must not truncate the user's workbook.
- **Offline snapshot** mirrors `json_data` verbatim; once the live DB is correct, the next clean-online-close regen carries it along. Any new write path must keep the snapshot-regen trigger + parity.
- **DB-Browser deletion is always manual + confirmed-with-a-listing.** No automatic row destruction, ever.

### Versioning / sprint rules
- Semantic `x.y.z`: **patch z = internal** (not deployed, verified locally), **minor y = deployable** (the version dropped on Synology), **major x = fundamental**.
- A batch of internal patches wraps into **one** deployable minor. This sprint: **v2.5.x internal → v2.6.0 deployable**.
- The `y` delimits the sprint; only the current minor's plans are active.
- **MS-8 (web service) is exempt** — it is a separate NAS Docker artifact versioned independently of DataViewer.exe; it does NOT participate in the v2.6.0 wrap.
- **Never auto-drop the installer on Synology.** Build `dist\DataViewer-setup.exe` in-repo, surface the path, let the owner transfer it. (User-level standing rule; absolute.)

---

## MS-1 — Anchor plot Y-axis at 0 + stop report TPM-trend peak clipping (DATAVIEWER-5)

**Priority:** P1 · **Effort:** S · **Status:** Ready (negative-value + dual-axis policy needs a one-line owner call; safe defaults proposed).

### Validated scope
Two confirmed render-only defects:
- **Floating yMin.** `PlotEngine::autoRange` (`src/plotting/PlotEngine.cpp:53-82`) derives yMin from data, pads 5%, snaps down to a nice tick — never forces 0. Reached only via `cfg.autoScale=true`: UI TPM Trend (`PlotWidget.cpp:482`), UI Power Density (`:591`), single-series `renderTPMTrend` fallback (`PlotEngine.cpp:954` via `PlotWidget.cpp:439`), and both axes of `renderLinePlotDualAxis` (`PlotEngine.cpp:758`, `:774-786`, the oil-overlay path `PlotWidget.cpp:519`). For TPM ~3.3–3.6 mg this floors the axis near 3.0 — the reported symptom.
- **Report TPM-trend peak clip.** `computeTpmYMax` (`ReportGenerator.cpp:684-702`) returns a flat **7.0** when sample-average TPM ≤ 7.0, ignoring per-row peaks. The trend **line** uses this yMax (`:146`) and `renderLinePlot` hard-clips to the rect (`PlotEngine.cpp:370`), so individual rows > 7 mg are silently cut.

**Already correct (regression guards, no change):** `renderBarChart` (yMin=0 always), all three report plots (`cfg.yMin=0.0`), UI Draw Pressure (`PlotWidget.cpp:638`). **Corrections to the original brief:** draw-pressure is *already* 0-anchored (the "floor at 2" is on the *upper* bound `drawPressureYMax`, not yMin — no fork). **Out of scope:** radar charts (fixed 1–9 ring, no data-derived axis); there is no Power Density *report* plot.

### Minimal approach
1. In `autoRange`, clamp the lower bound to 0 before the nice-tick snap (`yMin = 0.0`, or `qMin(0.0, dataMin)` if negatives must survive). Fixes UI TPM Trend, Power Density, and the single-series fallback in one place. Keep padding + nice-tick on the **top** only; leave X untouched.
2. Apply the same yMin=0 clamp explicitly to the dual-axis **right/secondary** axis (`:774-786`, computed inline, not via autoRange) — default to anchoring the oil axis to 0 for consistency.
3. In `computeTpmYMax`, in the non-long-puff branch return `max(7.0, maxTPM + headroom)` (maxTPM already computed at `:693`) so the trend line's peaks fit; bars still plot the average ≤ new max. Preserve the long-puff (15–25) branch and the >7 branch.
4. **No new config knob, no axis-policy enum** — unconditional clamp is the ponytail-minimum.

### Files
`src/plotting/PlotEngine.cpp`, `src/plotting/PlotEngine.h` (test accessor on autoRange), `src/reporting/ReportGenerator.cpp`, `tests/tst_plotengine/tst_plotengine.cpp`, `tests/tst_reportgenerator/tst_reportgenerator.cpp`, `tasks/lessons.md`.

### Dependencies
None blocking — render-only, no DB/LiveSync/snapshot/schema. Standard MIP decrypt before build. Tests are non-DB.

### Risks
- **Visual regression** — anchoring 0 compresses tightly-clustered high-baseline trends; changes the look of every existing trend → owner visual sign-off required.
- **Dual-axis alignment** — anchoring both left (TPM) and right (oil) axes independently changes relative visual alignment; deliberate.
- **Negative data** — a hard 0 clips a genuinely negative (bad-sensor) row; for non-negative metrics fine, but pick `0.0` vs `qMin(0.0,dataMin)` explicitly.
- **`computeTpmYMax` feeds 3 call sites** — the 5 assertions in `tst_reportgenerator` (lines 242–270) must update in lockstep; must not shrink the bar axis or change the long-puff case.
- **Scope creep** — do NOT add an `anchorYToZero` setting.

### Acceptance criteria
- UI TPM Trend / Power Density / single-series fallback: Y-axis bottom tick reads 0; lowest point not pinned to the edge.
- UI oil-overlay: left axis 0; right axis per agreed rule (0 by default).
- UI Draw Pressure + all bar charts: unchanged (still 0).
- Report TPM Trend: with rows > 7 but average < 7, no point/segment clipped at top; yMin still 0.
- `tst_plotengine` asserts auto-scaled line plots produce yMin==0 (via a test accessor); `tst_reportgenerator` `computeTpmYMax` assertions updated incl. a new max>avg-over-7 anti-clip case.
- `tests/run-tests.ps1` green; build clean under `-Werror -Wall -Wextra -Wpedantic`; `/ponytail-review` finds no new axis-policy abstraction; `tasks/lessons.md` updated.

### Open questions
- Negative-value policy: hard `0.0` vs `qMin(0.0, dataMin)`? (Recommend hard 0 for these non-negative metrics.)
- Anchor the dual-axis right (oil) axis to 0? (Recommend yes.)
- `computeTpmYMax` headroom: reuse the existing `+1.0` convention or bar-chart ~10%? (Pick one.)
- Visual sign-off on the compression trade-off across all existing decks.

---

## MS-2 — Ribbon label + image-button clipping fixes (DATAVIEWER-7)

**Priority:** P1 · **Effort:** S · **Status:** Ready (one wording fork on Defect A).

### Validated scope
Two independent presentation-only clipping defects (brief line numbers were stale):
- **Defect A — ribbon path-button label clipped.** "Set Detailed Sensory Output Path" (`MainWindow.cpp:830`, buildSettingsTab → "Output Paths" group) is a 4-word label in an 80×76 large button (`RibbonWidget.cpp:134`; the header doc-comment "56×70" at `RibbonWidget.h:29` is stale). The word-wrap helper (`RibbonWidget.cpp:116-131`) inserts at most **one** newline against a 74px budget; 4 words can't fit even after one split → the longest run overflows/elides. Siblings "Set TPM Output Path" / "Set Sensory Output Path" wrap cleanly.
- **Defect B — Load/View Images buttons clipped at bottom.** `m_loadImagesBtn`/`m_viewImagesBtn` (`MainWindow.cpp:1339-1340`) have no per-button stylesheet, inheriting global `QPushButton { padding:5px 14px; min-height:24px }` (`AppTheme.cpp:146-148`) → intrinsic ~36px, but sit in `imgBar` forced to `setFixedHeight(32)` (`:1348`) with 4px top/bottom margins → bottom crushed. Same class as the documented nav-button workaround (`MainWindow.cpp:887-891`).

### Minimal approach
- **Defect A:** edit the single string literal at `:830` to a label whose longest post-wrap run fits 74px — recommend "Detailed Sensory Path" (wraps "Detailed Sensory" / "Path") or "Set Detailed Path" to keep the "Set …" parallelism. **Do NOT widen the button** (`setFixedSize(80,76)` is hard-coded at 3 sites incl. compact mode — ripples everywhere). Optional hygiene: fix the stale "56×70" comment in `RibbonWidget.h:29-32`.
- **Defect B:** change `imgBar->setFixedHeight(32)` → `40` at `:1348` so the global-QSS buttons get their natural ~36px (matches every other default button). Reject "zero the padding" — `min-height:24px` still fights and it adds a per-button stylesheet that looks inconsistent.

### Files
`src/MainWindow.cpp` (label `:830`, `imgBar` height `:1348`), `src/widgets/RibbonWidget.h` (optional stale-comment fix), `tasks/lessons.md`, `DataViewerEnterprise.pro` (VERSION only at batch-wrap, not in this item).

### Dependencies
None blocking. MIP decrypt before build. Batch into the v2.5.x → v2.6.0 series.

### Risks
- Defect B `imgBar` height is hand-tuned vs the left splitter sizes (`leftSplitter->setSizes({210,80,310})` `:1361`); +8px should be eyeballed in both TPM and sensory modes (panel shared).
- Defect A residual clip — even "Set Detailed Sensory Path" may still clip after one auto-wrap; **must be visually verified on the work machine**, shorten if needed.
- Zero data/concurrency surface.

### Acceptance criteria
- Defect A: third path button's full label legible, no clipping/elision at default + normal ribbon widths; `getExistingDirectory` flow + tooltip for `ReportMode::DetailedSensory` unchanged.
- Defect B: "Load Images" / "View Images (N)" render full text + bottom edge in both modes; the count update + enable/disable still track count>0.
- Build clean under `-Werror …`; no layout regression beyond the intended ~8px taller image bar; `tasks/lessons.md` updated (2-line/80px label cap; `RibbonWidget.cpp` is the size source of truth, not the header comment).

### Open questions
- Exact Defect A wording (parallelism vs guaranteed fit).
- Include the stale-comment fix in scope?
- Audit the other path labels for the same cap (proactive vs out-of-scope)?

---

## MS-3 — "Save plot/chart" labels on icon-only save buttons (DATAVIEWER-12)

**Priority:** P1 · **Effort:** S · **Status:** Ready. **Owner decision (2026-06-24): Detailed Sensory is OUT of scope entirely — NOT a separate item.** Label only TPM + Sensory (both save buttons already exist); Detailed has no save-chart button and will not get one here. Plane DATAVIEWER-12 moved to Todo.

### Validated scope
- **TPM mode: button EXISTS.** `m_saveBtn = makeBtn("", "Save image")` (`PlotWidget.cpp:64`), icon-only 32×26, added at `:88`, wired to `onSaveImage` (`:148`). `makeBtn` already accepts a `text` arg.
- **Sensory mode: button EXISTS.** `saveChartBtn` (`SensoryPanel.cpp:908-914`), icon-only 24×24 flat, wired to `onSaveChart` (`:1627`). It is a **local** variable (not in the header) → label addable purely inline.
- **Detailed Sensory: NO button, NO `onSaveChart` slot.** Only the PowerPoint export path exists (`generateCombinedPptx`, `DetailedSensoryPanel.cpp:1979-1998`). "Apply to Detailed too" is therefore a **new feature** (new button + slot grabbing two radar widgets) — split into its own backlog item; not part of this labeling fix.

### Minimal approach
- TPM: pass a caption to `makeBtn` at `:64`, or add an adjacent `QLabel` before `:88` (prefer the label if the caption clips the fixed-size button). Keep the icon.
- Sensory: add `layout->addWidget(new QLabel(tr("Save chart")))` before the button at `:914` (cleaner than relaxing the 24×24 button). Keep icon + tooltip.
- Use **one** wording across both modes (`tr()`-wrapped). Do nothing for Detailed; document the gap.

### Files
`src/plotting/PlotWidget.cpp`, `src/ui/SensoryPanel.cpp`, `tasks/lessons.md`. (Detailed `.h`/`.cpp` only if the separate feature is approved — out of scope here.)

### Dependencies
None hard. Batch with MS-1/MS-2 into the same internal patch series.

### Risks
- TPM top-bar is `setFixedHeight(36)`; the button is fixed 32×26 — prefer a separate QLabel over enlarging the button to avoid clip/reflow at narrow responsive widths.
- A bare QLabel must be parented to avoid `-Werror` new-without-parent/unused warnings.
- Zero data path.

### Acceptance criteria
- TPM: visible "Save plot/chart" caption adjacent to the save button; icon + `onSaveImage` unchanged; top-bar doesn't clip/grow past 36px.
- Sensory: visible matching caption; `onSaveChart` still opens the Save Chart dialog; header row doesn't reflow awkwardly.
- Wording consistent across modes and clearly "saves the chart, not the file."
- Detailed: explicitly decided — left as-is with the no-button gap documented in `tasks/lessons.md` + flagged to owner (default), unless the separate feature is approved.
- Build clean; eyeball-verified on an installed build; lessons.md appended.

### Open questions
- Detailed save-chart button — separate item (default) or pull in now?
- If pulled in: one combined image, two files, or a chooser for the two radar charts?
- Caption wording: "Save plot" vs "Save chart" (pick one).
- Caption on the button vs adjacent QLabel.

---

## MS-4 — Remove TPM data tables → right-side per-sample Notes panel (DATAVIEWER-6)

**Priority:** P2 · **Effort:** L (option A) / **XL (option B)** · **Status:** **BLOCKED on a product fork — no code until resolved.**

### Validated scope (with a major correction)
`m_dataTable` (`MainWindow.cpp:936`) is **not** a passive grid — it is the **only editing surface** for all 7 per-row TPM input fields (puffs, before/after weight, draw pressure, resistance/regime, smell, clog, notes) **and** the live-collaboration surface. Edits flow `QTableWidgetItem → onTableCellChanged (:1933) → DataRow mutation → recalculateSampleMetrics → queueExcelWrite + DB`; `onDataTableItemChanged (:3425)` pushes each cell to Postgres via `m_liveSync->commitCell`; remote edits paint back as T18 yellow decoration (`:3587-3644`) keyed by `data_rows.id`; `onDataTableItemClicked (:3464)` is "take their value"; `onAddRow/onRemoveRow (:6939/:6998)` + `findTableRowForDataRowId (:3502)` all operate on it; `onTableCellChanged (:2014-2046)` queues offline `PendingEdit`s; `displayCurrentSample (:3740-3779)` reuses it to render raw/SOP sheets.

So **"remove the table" = "remove all per-row TPM data entry, live-sync of measurement cells, offline edit capture, add/remove-row, and raw-sheet rendering"** unless each is re-implemented. The Sample-Properties grid (`m_propTable`, `:1315`) is **separate** (sample-level fields only). The Notes data the panel wants already exists per-row (`DataRow.notes/.puffs/.tpm`, `SampleResult.averageTPM`); `ReportGenerator.cpp:330-341` already aggregates per-row notes.

**Product fork:**
- **Option A (safe default):** Notes panel is **additive**, the grid stays (relocated/collapsible/secondary tab). Effort L. No subsystem deleted.
- **Option B (literal brief):** in-app per-row editing is gone (data entry moves to source `.xlsx` only) → tear out editing + live-sync decoration + add/remove-row + delegates + offline queue + raw-sheet rendering + re-spec multi-user TPM behavior. Effort XL, high data-loss/collaboration regression risk. **If chosen, option B becomes its own multi-phase plan** — it is not implemented inline in this batch.

### Minimal approach (option A)
1. Add a right-side Notes panel as a **read-only `QListWidget` with word-wrap** (NOT a re-implemented table or hand-rolled flow layout). Echo the Navigator's section-header style (reuse `propHeader` `:1309-1313` + AppTheme tokens).
2. In `setupCentralWidget`, reorganize so `PlotWidget` gets center stretch and the Notes panel sits right (reuse the existing `QSplitter` — no new layout abstraction).
3. Populate in `displayCurrentSample` using the **same** visible-row filter (skip zero-weight rows), format `"Note N: Puff <puffs>, Current TPM <dr.tpm>, Average TPM <sample.averageTPM>, <dr.notes>"`; reuse the `ReportGenerator` aggregation pattern.
4. Retire the Layout ribbon group (Table/Plots/Both, built `buildViewTab :790-803`, handlers `:4270-4272`) or repurpose to a Notes show/hide.
5. **Keep the grid as the editor** (collapsible/secondary tab) so no data-entry/live-sync capability is lost.
6. **Ponytail cuts:** no custom flow/per-note QFrame widget; no new persistence (notes already persist); no "panel width" setting; do not mirror the full left-dock splitter.

### Files
`src/MainWindow.cpp` (setupCentralWidget `:857-1017`, displayCurrentSample `:3730-3980`, buildViewTab `:790-803`, view handlers `:4270-4272`, mode setCurrentWidget `:4307/:4348`; **option B only:** delete the editing/live-sync/add-row machinery above), `src/MainWindow.h` (members `:251-252`, UserRole comments `:467-483`, slot decls `:139-141`, new `m_notesPanel`), `src/plotting/PlotWidget.{h,cpp}` (verify center-fill; likely reparent only), `src/widgets/CellFocusDelegate.{h,cpp}` + `RegimeComboDelegate.{h,cpp}` (dead under option B — retire cleanly for `-Werror`), `tasks/lessons.md`, a new MainWindow-level Notes-population test.

### Dependencies
- **Product fork (blocking).** Everything downstream forks on A vs B.
- Option B ties into the Excel-sidecar template initiative + the future native-data-collection-UI item; removes the only `data_rows` cell-commit producer/consumer → multi-user TPM behavior must be re-specced/accepted.
- No schema/migration dependency either way; reports unchanged.

### Risks
- **Data-entry loss (highest)** if B is chosen by mistake — silent capability regression.
- **Multi-user/live-sync regression** — removing `onDataTableItemChanged` + the T18 block kills live TPM-cell collaboration.
- **Offline pending-edit loss** — `onTableCellChanged` is where offline TPM edits queue.
- **Raw/SOP sheet display breaks** unless a separate raw-table widget is kept.
- **Empty-row filter coupling** — the panel must replicate the visible-row indexing so "Note N"/"Puff #" line up with plot/reports.
- `-Werror` — orphaned delegates/UserRole machinery must be removed, not commented out.
- Scope creep — keep it one right-side panel; don't refactor the left dock.

### Acceptance criteria
- Build clean (`-Werror …`), no `-Wunused` from orphaned members.
- TPM mode: plot fills center, Notes panel on the right echoing the Navigator styling.
- Panel lists one entry per **visible** row, format above, cycling per sample, wrapped (no truncation); values match the grid + report (verified vs `ReportGenerator` output incl. the zero-weight filter).
- Raw/SOP sheet display still correct after the restructure.
- Layout ribbon group retired/repurposed cleanly; no dangling buttons/slots.
- Mode switches intact; Notes panel TPM-only.
- Per chosen option: (A) per-row editing + add/remove + live-sync still work; or (B) removal documented + multi-user/offline implications owner-signed-off.
- New MainWindow Notes test green; full suite green; `tasks/lessons.md` updated (the "data grid" is the editing + live-collaboration surface).

### Open questions
- **Blocking fork:** option A (additive) vs option B (drop in-app editing)? (Recommend A.)
- If B: is loss of real-time multi-user sync on TPM cells acceptable?
- Right-side QDockWidget (movable, symmetric) vs fixed central pane?
- Panel read-only vs editable notes (if grid goes, read-only loses note editing)?
- "Current TPM" = per-row `DataRow.tpm`, "Average TPM" = `SampleResult.averageTPM` — confirm.
- Samples with zero notes: hide entry, or keep the metric header line?

---

## MS-5 — Per-sample stopwatch in Sensory mode (DATAVIEWER-13)

**Priority:** P2 · **Effort:** S · **Status:** Ready — **owner-approved 2026-06-24** (Plane DATAVIEWER-13 moved to Todo).

### Validated scope (narrowed to a pure UI addition)
The per-sample puff field already exists end-to-end and needs no schema/serializer/sync work:
- `SensorySample::puffLengthSec` (double, default 3.0) at `SensoryData.h:70`; serialized as `"puff_length_sec"` both ways (`SensoryData.cpp:45,91-92`) with tolerant string→double read; round-trip + coercion already unit-tested (`tst_sensorydataplaceholder` lines 292/322-323/352-353/376-377/419/437).
- Editor widget already exists: `SampleCard::m_puffLengthSpin` (NoWheelDoubleSpinBox 0.1–60.0s, 0.5 step, 1-decimal, " s" suffix) in `puffRow` (`SensoryPanel.cpp:499-523`); its `valueChanged` already emits `cellCommitted("puff_length_sec", v)` + `changed()`.
- Persistence routes through the per-card handler (`SensoryPanel.cpp:959-976`) → `LiveSync::commitCell` → `dve_commit_cell_json` (stores numeric-looking values as JSON number via generic regex, `init.sql:466-467`). Whole-session merge only arbitrates score keys → puff is in-memory-authoritative, not DB-revert-prone. Report already prints it (`SensoryReportSource.cpp:185,649`).

**The work:** a Start/Stop button per `SampleCard` that, on Stop, calls `m_puffLengthSpin->setValue(elapsedSeconds)` — `setValue()` fires `valueChanged`, reusing the entire existing persistence/LiveSync/dirty/report chain. **Detailed Sensory has no puff field** → adding it there is a separate, larger item (struct field + both serializers + migration + UI). TPM out of scope.

### Minimal approach
Add a small Start/Stop `QPushButton` to `SampleCard` inside `puffRow` (~`:500-516`), holding timing in a `QElapsedTimer` **value member** (Qt monotonic — do not hand-roll with QDateTime). Click toggles: first → `start()`, relabel "Stop", optionally a ~200ms `QTimer` for a live readout; second → `elapsed = m_stopwatch.elapsed()/1000.0`, `qBound` to the spin range, `setValue(elapsed)`, relabel "Start", stop the tick timer. Nothing else changes. **Cut:** no new struct/JSON field, no migration/stored-fn change, no LiveSync edit, no separate Stopwatch class, no sub-100ms precision, no pause/resume/lap, no raw timestamp persistence.

### Files
`src/ui/SensoryPanel.h` (SampleCard: `QPushButton* m_stopwatchBtn`, `QElapsedTimer m_stopwatch`, optional tick `QTimer*`; includes), `src/ui/SensoryPanel.cpp` (ctor wiring), `tests/tst_sensorydataplaceholder/` (optional — only if a pure rounding/clamp helper is extracted).

### Dependencies
None hard — field/serializer/LiveSync/numeric-commit/report all exist + tested. Detailed parity = separate item if scoped in.

### Risks
- Range clamp: spin caps at 60.0s / floors 0.1s; a >60s real puff or a left-running timer is silently clamped on `setValue` — decide raise-cap vs clamp+warn; never write out-of-range.
- Programmatic `setValue` emits `cellCommitted` exactly like a manual edit → streams a per-cell update on a persisted (id>0) session — correct/desired; confirm in test.
- A live tick `QTimer` must be stopped on Stop, on card destruction/removal (`onRemoveCard`/`remapDirtyCellsAfterSampleRemoval` handle teardown — the timer must not outlive the card), and reset in `fromSample()`. Keep `QElapsedTimer` as a value member.
- `-Werror` — no unused member/capture.
- Card is fixed-size in a FlowLayout — the button must not break wrapping (visual sign-off).

### Acceptance criteria
- Each SampleCard shows a Start/Stop button next to "Puff length:"; card layout/wrapping unchanged otherwise (owner eyeball).
- Start→Stop sets the spin to elapsed, rounded to 1-decimal, ±0.2s.
- Stopped value flows the same path as manual entry: survives whole-session save (Ctrl+U) without revert; appears in the sensory report.
- On a persisted (id>0) LiveSync-connected session, Stop streams a single `json_path:samples[i].puff_length_sec`; reload shows a JSON number (not reverted to 3.0).
- Out-of-range elapsed clamped to [0.1, 60.0]; >60s behavior decided + documented.
- Removing a card mid-run / loading another session via `fromSample()` stops/clears the timer — no dangling timer, no spurious commit, no crash.
- Build clean; `tests\run-tests.ps1` green (existing puff round-trip still passes).
- VERSION bumped as internal patch with clean rebuild; not auto-dropped to Synology.

### Open questions
- Detailed Sensory parity — separate item (recommend) or now?
- Precision (1-decimal vs integer seconds) + 60s cap (raise vs clamp)?
- Live ticking readout vs plain Start/Stop?
- Overwrite-on-stop semantics acceptable (no "lock")?
- Button placement/label (text vs icon).

---

## MS-6 — One-shot JSONB string-score backfill (DATAVIEWER-10)

**Priority:** P3 · **Effort:** S (verify-only) · **Status:** **ALREADY SHIPPED in v2.4.2 R4 (commit `1a31c11`) → rode into v2.5.0. Do NOT re-implement.**

### Validated scope
The requested mechanism already exists across four coordinated paths:
1. `dve_normalize_legacy_json()` — losslessly rewrites numeric-string scores → JSON numbers in both session tables; idempotent (only string values matching `^-?[0-9]+(\.[0-9]+)?$`), order-preserving (`ORDER BY elem.ord`), `dve.maintenance` GUC to suppress NOTIFY storms, `SET LOCAL statement_timeout=0`. Byte-identical in `deploy/postgres/init.sql:500`, `deploy/postgres/migrations/2026-06-11-legacy-score-normalizer.sql:19`, and the `ensureSchema` CREATE OR REPLACE heal (`DatabaseManager.cpp:512-566`).
2. The exact `schema_meta`-gated one-shot the item asks for — `ensureSchema` (`DatabaseManager.cpp:568-614`) wraps it in a transaction, takes `pg_advisory_xact_lock(4242002)`, probes `schema_meta WHERE key='v242_legacy_score_normalize'`, runs once, INSERTs the marker `ON CONFLICT DO NOTHING`.
3. Nightly `pg_cron` job `dve_legacy_score_normalize` (`init.sql:583`, 03:17 daily).
4. Manual "Repair Legacy Scores" button (`DatabaseBrowserDialog.cpp:1235` → `DatabaseManager::normalizeLegacyScores` `:1336`) + a "Legacy string scores" badge (`kLegacyScoreScanExists` `:2202`).

Already-tested: `tst_databasemanager.cpp:3055/:3224/:3171`. The backfill is correctness-neutral (the v2.4.1 tolerant reader already coerces strings on load on both DB and snapshot paths) — exactly as the item states. **Nothing functional is missing.**

### Minimal approach — VERIFY-AND-CLOSE (no code)
1. Confirm shipped state (done in analysis).
2. **One owner-approved read-only live query** (per the live-query guardrail — do not run unprompted): (a) `SELECT value FROM schema_meta WHERE key='v242_legacy_score_normalize'` — non-null timestamp proves the one-shot fired; (b) residual string-score count = 0 in both session tables. If >0, owner clicks Repair once or it self-heals at 03:17. **If the owner is unavailable, this verification does not block the rest of the sprint** — the feature is correctness-neutral (tolerant reader is the backstop); close the ticket on code-evidence and leave the live-DB check as a follow-up chore.
3. Transition DATAVIEWER-10 → Done/Duplicate-of-v2.4.2-R4 (link `1a31c11`); append `tasks/lessons.md` retiring the "pending backfill" assumption.

### Files
None for implementation. Reference: `deploy/postgres/init.sql`, the migration, `src/database/DatabaseManager.cpp`, `src/ui/DatabaseBrowserDialog.cpp`, `tests/tst_databasemanager/`. `tasks/lessons.md` for the closing note.

### Dependencies
DV-2 (the schema_meta-gated precedent) + DV-4 (parent) — both shipped, not blocking. No schema/migration/build/test work.

### Risks
- **Primary risk = redundant re-implementation.** A new backfill collides with the shipped one: a second marker key needlessly re-scans; a second CREATE OR REPLACE risks three-way drift in the deliberately-identical triplet. **Do nothing in code.**
- A naive re-do that skipped `dve.maintenance` would storm every connected client + bump version/updated_at → OCC conflict dialogs. Dropping `ORDER BY elem.ord` would silently reorder samples.
- Offline snapshot needs no separate work (verbatim TEXT copy + tolerant reader).

### Acceptance criteria
- Confirmed in code: function present in all 3 locations; one-shot schema_meta-gated under `pg_advisory_xact_lock`; nightly cron + manual Repair wired; existing tests cover it — PASS.
- Owner-run live verification: marker holds a timestamp AND residual string-score count = 0 in both tables (outstanding, gated; non-blocking — see approach step 2).
- DATAVIEWER-10 → Done/Duplicate with link to `1a31c11`; `tasks/lessons.md` updated.
- **No new SQL function, migration, schema_meta key, C++, or UI added** — any such diff is rejected at `/ponytail-review`.

### Open questions
- Did the one-shot run on the live NAS DB post-v2.5.0? (Owner-confirmed via the marker; ask before querying.)
- Close as Duplicate (recommend) vs keep open as a verify chore?

---

## MS-7 — LiveSync-authoritative sensory/detailed persistence (DATAVIEWER-9)

**Priority:** P3 · **Effort:** L (minimal scope) / **XL (literal "sole writer")** · **Status:** **BLOCKED on a product fork — most of the original justification is already shipped.**

### Validated scope (with major correction)
The brief's premise (whole-session save is a removable batch overwrite) is **largely obsolete**. v2.5.0 already fixed the DATAVIEWER-4 data-loss class and made the dual-path model **intentional** (explicit DESIGN note at `DatabaseManager.cpp:1767-1778`).
- `onUpdateDatabase` (`MainWindow.cpp:5131-5415`) is the convergence point of **three** triggers: Ctrl+U (`:1577-1589`, flush=true), the 5s debounced auto-save timer (`:500-503`, kicked at `:5742`, flush=false), and program-close (`:6007`, flush=true) — plus `save*BeforeClose` (`:2889/:2991`).
- `tryWriteSensoryCore` (`DatabaseManager.cpp:1741`) / `tryWriteDetailedSensoryCore` (`:2345`) are **already read-merge-write under `FOR UPDATE`** via `mergeSensory/DetailedPreservingDbScores` with a panel-supplied dirtyCells set (v2.5.0 Task 3).
- **The whole-session path is load-bearing and cannot be naively deleted:** (1) it is the **only INSERT path** for new (id≤0) sessions — per-cell `commitCell` is gated on `activeSessionId()>0` + an existing row; (2) it persists fresh-from-disk TPM files; (3) it is the **safety net when LiveSync is gapped** (the exact failure behind 23 silent data-loss events in 3 days).
- **The real residual:** the `oil_smell_liking` asymmetry is confirmed — it streams per-cell (`DetailedSensoryPanel.cpp:547-552`) but the whole-session detailed merge (`DetailedSensoryData.cpp:153-179`) treats only per-sample scores as DB-authoritative and keeps all session-level fields (oil_smell_liking, clog, mouthpiece_notes, test_title) from the in-memory blob → a stale in-memory value can overwrite a newer per-cell/remote one.

### Minimal approach (the ponytail-correct, recommended scope — assuming "keep dual path, close the gap, simplify UX")
**Step 0 (REQUIRED before code):** get a product decision on the fork (open questions). Do NOT do the literal removal.
1. **Close the `oil_smell_liking`/session-level asymmetry IN THE MERGE** (not by deleting the write path): extend `mergeDetailedSensoryPreservingDbScores` (and the sensory twin if needed) to arbitrate session-level fields the same way as per-sample scores — a field touched this run (tracked in dirtyCells) stays in-memory-authoritative; an untouched one takes the DB value. Mark oil_smell_liking/clog/mouthpiece dirty in `commitSessionField` alongside the existing per-cell commit. Reuses the existing dirty-aware pattern — no new architecture.
2. **Retire the user-facing Ctrl+U action as redundant** (LiveSync + 5s auto-save + close-flush already persist everything), but **keep `onUpdateDatabase` callable internally** from the auto-save timer + close paths. Update comments + the DB-sync-indicator tooltip that say "press Ctrl+U."
3. **Do NOT delete** `tryWrite*Core`, the merge helpers, dirtyCells, or `flushNowAndWait` — they remain the INSERT path, the gapped-LiveSync safety net, and the reconnect catch-up (`reloadOpenResourceAfterReconnect`, tst_saveintegrity_e2e scenario 9).
4. TDD: add a session-level-field two-writer e2e scenario (mirror scenario 9) **before** the merge change; add a `tst_databasemanager` unit test for session-field arbitration.

If the decision is instead the literal "sole writer" (XL): a per-cell INSERT/auto-create path for id≤0 sessions, re-home TPM file persistence, define the cross-client conflict model explicitly, then remove the batch — each as its own phase with its own e2e proof. **Recommend against** unless the owner specifically wants it; this becomes its own multi-phase plan outside this batch.

### Files
`src/database/DatabaseManager.cpp` (cores `:1741/:2345`, wrappers `:1890-2090/:2477-2650`), `src/pipeline/DetailedSensoryData.cpp` (merge `:153-179`), `src/pipeline/SensoryData.cpp` (merge `:103-126`), `src/ui/DetailedSensoryPanel.cpp` (`commitSessionField :2064`, dirtyCells `:2107-2135`), `src/ui/SensoryPanel.cpp` (`:959-976`), `src/MainWindow.cpp` (Ctrl+U `:1577-1589` if removing the action; `onUpdateDatabase :5131-5415`; timer `:500-503`; close `:5990-6008`; `save*BeforeClose :2889/:2991`), `src/MainWindow.h`, `tests/tst_saveintegrity_e2e/`, `tests/tst_databasemanager/`, `tests/tst_livesync/`, `tasks/lessons.md`, `docs/superpowers/specs/2026-06-05-bugfix-batch-design.md` §8.

### Dependencies
DV-4 (this is its deferred follow-up; phase 1 done removes most of DV-9's justification) · DV-8 (unnamed/unkeyed save-on-close + Name-It-Now/Discard/Cancel live INSIDE `onUpdateDatabase`/`save*BeforeClose` — must be preserved) · the v2.5.0 RC1–RC5 fix set (constraints, not legacy). No schema/migration dependency (stored fns already exist). **Soft prerequisite: a product decision on the cross-client conflict model.**

### Risks
- **Data-loss (HIGH)** — removing the whole-session write deletes the INSERT path for new sessions + the gapped-LiveSync safety net (same class as 23 silent discards).
- **Architecture regression (HIGH)** — contradicts a deliberate, tested v2.5.0 DESIGN decision; proceeding without a fresh product decision re-litigates settled work.
- **Multi-user/OCC (HIGH, hard to verify here)** — changing which path owns session-level fields changes cross-client conflict semantics; the work machine can't reproduce true concurrency — only the ephemeral-PG simulations + a manual two-client test.
- **Over-removal** — the dirty-cell ecosystem is also used by reconnect catch-up; mis-identifying dead code breaks catch-up.
- **Offline snapshot** — keep the debounced regen trigger (`:5380-5382`) + parity (`OfflineSnapshot.cpp:2244`).
- **Trigger ripple** — `onUpdateDatabase` is shared by 4 callers; the auto-save is silent/async, close paths flush/block — easy to break one fixing another.
- `-Werror` — a removed QAction/member leaving a dangling `connect` is a hard build break. MIP decrypt before build.
- **UX (LOW)** — removing Ctrl+U changes muscle memory; many comments + the sync-indicator tooltip reference it — must be updated.

### Acceptance criteria
- New e2e (tst_saveintegrity_e2e): after another client updates a detailed session's `oil_smell_liking`, a whole-session save from a client whose in-memory value is **stale-and-untouched** does NOT clobber the newer DB value; a client who DID edit it this run keeps its local value. FAILS on current code, PASSES after the merge change.
- Integrity: whole-session save preserves every json_data key it doesn't arbitrate (samples, comments, structure, layout, images, id/version).
- **New-session path intact:** a brand-new (id≤0) sensory AND detailed session still persists in full on first save (INSERT).
- **Gapped-LiveSync safety intact:** existing RC1/RC2/Task-3 scenarios still pass.
- If Ctrl+U removed: shortcut no longer triggers a manual flush; comments + sync-indicator tooltip no longer mention it; edits still persist via 5s auto-save + close-flush (manual close-without-Ctrl+U test); no dangling QAction/connect; `-Werror` clean.
- Multi-user: tst_twoclient_e2e + tst_saveintegrity_e2e green; manual two-client session-field edit behaves per the chosen conflict model.
- Build clean; full suite green except the pre-existing flaky `tst_responsivelayout`.
- Docs: bugfix-batch §8 marked resolved/re-scoped; `tasks/lessons.md` records the whole-session path is load-bearing and must not be removed naively.

### Open questions
- **Gating fork:** literal "sole writer" (XL, new per-cell INSERT path + re-derived conflict model) vs minimal (close the asymmetry + retire redundant Ctrl+U, keep the whole-session path as INSERT + safety net — recommended)?
- Cross-client conflict model for session-level fields: field-level last-writer-wins via dirty-aware merge (recommended, low effort) vs a conflict dialog (heavier)?
- Ctrl+U: remove entirely vs keep as a redundant "force save now"?
- If removed: is 5s auto-save + close-flush a sufficient deliberate-persist story, or add an explicit "Save" affordance?

---

## MS-8 — Remote phone sensory web form → Postgres (DATAVIEWER-11)

**Priority:** P3 (High in Plane) · **Effort:** L · **Status:** **SPIKE COMPLETE (2026-06-24)** — awaiting owner sign-off. The web service ships on its own NAS track, but DV-11 now also carries **ONE required desktop-side C++ fix** in the v2.6.0 train (see the banner).

> **⚠ AUTHORITATIVE UPDATE — see the sub-spec `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md`.** The Phase-5a spike (owner decisions + a 12-agent code-grounded adversarial pass — every robustness claim against the naive design failed at HIGH severity) **supersedes the "Validated scope / MVP / Minimal approach / Risks / Acceptance / Open questions" notes below.** Three corrections those notes get wrong:
> 1. **Natural key:** `session_name` = test-title **ONLY**; round is in `tester_name` (`" R1"`/`" R2"`); `date` is a local-TZ TEXT string (`SensoryPanel.cpp:1029-1041`, `TesterRound.h:31-37`, `init.sql:145-146`). It is **NOT** "title-tester-R#". So the write is **create-or-append by natural key** (resolve-then-append under `FOR UPDATE` via a new `dve_append_sensory_sample`), **NOT** "INSERT-only / auto-suffix-fork". Date must be pinned to the **office TZ server-side** (an offsite phone near midnight would otherwise fork the key).
> 2. **Desktop data-loss:** a whole-session save **silently drops** a phone-appended sample — `mergeSensoryPreservingDbScores` truncates `samples[]` to the in-memory count (`SensoryData.cpp:107-124`), last-writer-wins by design (`DatabaseManager.cpp:1767-1797`). → A **required desktop C++ merge fix** (preserve DB-only tail samples) ships in the v2.6.0 train; **DV-11 is NOT zero-C++.**
> 3. **Privilege:** the only DB role is the **superuser** `dataviewer_app`; a shared passcode over it gates full-DB write. → A **dedicated least-privilege `sensory_web` role** (INSERT/UPDATE/SELECT on `sensory_sessions` only; `REVOKE` on `dve_commit_cell*`) + per-tester revocable tokens + rate-limit.
>
> Owner-locked: reachable from **anywhere** (internet + TLS + tokens — office-WiFi REJECTED), **online-only**, **5-metric Sensory**. Plus add a stable per-sample `sample_uid` for idempotency. See sub-spec §10 for the remaining sign-off questions.

### Validated scope
Feasible **without touching the C++ app** for the MVP — a sensory session is just a Postgres row + JSONB blob, so a small web service can INSERT it and every desktop sees it live via existing NOTIFY. `DataViewer.exe` has no HTTP server and should not get one (a desktop can't be a 24/7 target). Canonical shape: `sensorySessionToJson` (`SensoryData.cpp:11-50`); row = the INSERT branch of `tryWriteSensoryCore` (`DatabaseManager.cpp:1859-1885`); natural-key unique index on (session_name, tester_name, date) at `init.sql:129-146`; the five metric keys (Burnt Taste, Vapor Volume, Overall Flavor, Smoothness, Overall Liking) at `SensoryData.h:48-50`; **scores must be JSON numbers 1–9** (string scores = DATAVIEWER-4 data loss). Proven precedent: `bug-form-app` (gunicorn Python form on the NAS writing into Plane) → mirror as a new NAS Docker stack.

**MVP:** mobile-first 5-metric Sensory form; thin Flask backend (SUPERSEDED — per sub-spec §3 it is **create-or-append by natural key**, NOT INSERT-a-fresh-row); deployed reachable by the phone; multi-sample yes; images deferred. **Hard blocker:** prod DB port 5433 is firewalled to the office VLAN (not internet-exposed, README step 7) — the co-located service reaches the DB but the phone must reach the service → fork: office-WiFi-only vs the existing Plane reverse proxy (8080/8443) vs full internet exposure (needs auth + TLS).

### Minimal approach
**Design spike (gate)** → resolve mode (5-metric vs 14-question vs both), network-exposure model, auth-or-none; see the (DONE) sub-spec `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md` — authoritative. Then:
1. Build the service in `deploy/sensory-collect/` as a single-file Flask app matching the gunicorn bug-form-app precedent: one GET serving a mobile-first form (5 sliders 1–9, per-sample name + comments, header of test title/tester/assessor/media/date, "add another sample"); one POST that validates + enforces DATAVIEWER-8 (non-empty test title + tester else 400), builds `json_data` exactly per `sensorySessionToJson` with **numeric** scores + the five literal keys, runs one parameterized **INSERT** with id=−1 semantics + `updated_by="web/<tester>"`, on unique violation auto-suffix (like `tryWriteSensorySessionAutoSuffix`) or "already exists" per the spike — reusing the DB version bump + NOTIFY so desktops update for free.
2. Dockerfile + compose mirroring `deploy/postgres` + bug-form-app, a new stack on the same Docker network (reaches the DB by **service name**, not the firewalled host port), password from the NAS env file (never committed — repo is public), phone reachability per the chosen fork.
3. Verify with a pytest that POSTs then reads the row back (numeric scores) + a manual e2e where a desktop on the test container shows the session, renders the radar, re-saves cleanly.
**Cut:** in-app C++ web server; per-cell streaming from the phone; offline PWA (unless the spike proves no-WiFi panels); image upload in v1; any ORM.

### Files
New: `deploy/sensory-collect/app.py`, `templates/form.html`, `Dockerfile`, `docker-compose.yml`, `tests/test_submit.py`; the sub-spec; `tasks/lessons.md`. **Required desktop C++ change (per the spike, ships in v2.6.0):** `src/pipeline/SensoryData.cpp` (`mergeSensoryPreservingDbScores` tail-preserving fix, §4.1), `src/ui/SensoryPanel.cpp` (grow cards on apply), `tests/tst_sensorydataplaceholder/` (append regression). New DB migration: `dve_append_sensory_sample` + the `sensory_web` least-privilege role. Reference: `src/pipeline/SensoryData.{h,cpp}`, `src/database/DatabaseManager.cpp`, `deploy/postgres/init.sql`.

### Dependencies
NAS access + existing NAS Docker workflow (prod password from the NAS env file, never committed) · the network-exposure decision (gates deployment) · DV-8 (savable rule → writer requires non-empty test title + tester) · DV-4 (numeric scores) · a REQUIRED desktop-merge dependency (the §4.1 fix so a desktop save can't drop a phone-appended sample); the write is create-or-append by natural key, NOT INSERT-only · Detailed variant only if the 14-question form is needed · images deferred.

### Risks
- **Data-loss (the real one): the DESKTOP whole-session save drops a phone-appended sample (R-M)** — fixed by the §4.1 tail-preserving merge. The web write is create-or-append by natural key under FOR UPDATE (append-to-tail only, never mutate existing samples) + a per-sample `sample_uid` for idempotency.
- **Malformed blob → silent reset-to-5** if scores are strings — mitigated by writing numbers + a pytest round-trip.
- **Security** — repo is public → password + any token live only in a NAS env file; internet exposure needs auth + TLS.
- **No identity** — phones have no per-install UUID → stamp `updated_by="web/<tester>"`; leave presence + cell-focus untouched.
- `-Werror` + MIP apply to the one required desktop C++ fix (Phase 5.0); the web service itself is Python.
- Timeline: the ~2-week ask is tight; scope is narrow (one 5-metric form, per-sample create-or-append). Internet-reachable (office-WiFi rejected); the desktop §4.1 fix is the gating C++ work.
- Desktop regression minimal (NOTIFY already handles new sessions) — verify open/render/re-save.

### Acceptance criteria
- A phone browser on the target network loads the mobile form; a non-developer completes + submits a 5-metric session.
- The stored row keeps every score as a JSON number 1–9 (not a string) with top-level columns populated — verified by an automated pytest.
- A running desktop receives the new session live via NOTIFY (no manual refresh), opens it, renders the radar with scores unchanged, then re-opens + re-saves with no version mismatch / merge error.
- Empty test title or tester is rejected with a clear message.
- *(SUPERSEDED — see sub-spec §11.)* Create-or-append: a submit whose natural key matches an existing session **appends a sample** to that row (no fork, no UniqueViolation); a desktop save of an open session **retains** the appended sample (the §4.1 desktop merge fix); a re-POSTed sample (same `sample_uid`) does not duplicate.
- Password + any token exist only in a NAS env file; a repo grep finds nothing.
- The one required desktop C++ change (§4.1 merge fix + its regression test) is green; the C++ build + full suite stay green under `-Werror`.
- The service runs as a NAS Docker stack that survives a restart, reaches the DB by internal service name, and is documented.

### Open questions
- Which mode in ~2 weeks: 5-metric, 14-question, or both?
- *(RESOLVED — sub-spec §1/§10.)* Internet-reachable via reverse proxy + TLS + per-tester tokens (office-WiFi rejected).
- Auth required, or WiFi-only unauthenticated acceptable?
- *(RESOLVED — sub-spec §3.)* Header match = **append the sample to the existing session** (create-or-append), never fork/auto-suffix.
- Per-submission tester+assessor name, or one shared device per session?
- Any no-WiFi panels (would justify an offline PWA, else over-engineering)?
- Photo capture now or deferred?
- Throwaway for one event vs durable internal tool (sets the polish bar)?

---

## Risk register (initiative-wide)

| # | Risk | Items | Severity | Likelihood | Mitigation |
|---|---|---|---|---|---|
| R-A | Silent data loss from naive removal of the whole-session write | MS-7 | Critical | Med (if literal scope chosen) | Product gate before code; keep the path as INSERT + safety net; e2e proof of new-session + gapped-LiveSync |
| R-B | Re-implementing already-shipped backfill → three-way function drift / NOTIFY storm | MS-6 | High | Med | Verify-and-close; reject any new SQL/migration at `/ponytail-review` |
| R-C | Removing the TPM grid kills per-row editing + live-sync + offline capture | MS-4 | Critical | Med (option B) | Default to option A (additive); owner sign-off + separate plan for B |
| R-D | Multi-user conflict semantics regress (cannot reproduce concurrency locally) | MS-7, MS-8 | High | Med | Ephemeral-PG two-client e2e + manual work-machine two-client test |
| R-E | Web form writes string scores → reset-to-5 | MS-8 | High | Low | Emit JSON numbers; pytest round-trip; tolerant reader is backstop |
| R-F | Secrets committed to the public repo | MS-8 | High | Low | NAS env file only; repo-grep acceptance gate |
| R-G | Visual regressions from y-anchor / layout changes | MS-1, MS-2, MS-4, MS-5 | Med | Med | Visual loop + owner sign-off before deployable wrap |
| R-H | `-Werror` break from orphaned members/delegates/dangling connects | MS-4, MS-7 | Med | Med | Remove (not comment) dead code; clean rebuild |
| R-I | VERSION bump ships stale `main.o` | all (at wrap) | Med | Low | `mingw32-make clean`; `build_installer.bat` FileVersion gate |
| R-J | MIP ciphertext breaks build mid-task | all C++ | Low | High | `decrypt_via_copy.py --apply` before every build; Python delete-and-rewrite for new files |
| R-K | Network-exposure decision stalls MS-8 | MS-8 | Med | Low | Spike DONE — network decided: internet via reverse proxy + TLS + per-tester tokens; office-WiFi rejected |
| R-L | Scope creep (config knobs, abstractions) | MS-1, MS-3, MS-4, MS-5 | Low | Med | `/ponytail-review` gate per diff |
| R-M | Desktop whole-session save silently DROPS a phone-appended sample (samples[] truncated to in-memory count) | MS-8 | Critical | High (without fix) | Required §4.1 tail-preserving merge fix + per-sample `sample_uid` + regression test; ships in v2.6.0 before the service goes live |
| R-N | Web service holding the superuser `dataviewer_app` role → a single service bug = full prod-DB compromise | MS-8 | Critical | Med | Dedicated least-privilege `sensory_web` role (INSERT/UPDATE/SELECT on sensory_sessions only; REVOKE on dve_commit_cell*); per-tester tokens + rate-limit; never call generic commit fns |

---

## Sequencing rationale

Ordered by **(dependency → risk → value)**:
1. **Quick wins first (MS-1, MS-2, MS-3)** — S-effort, render/label-only, zero data surface, immediate user-visible value, no blocking forks. They warm up the branch + installer loop with the lowest blast radius and need no product decision.
2. **Medium UI next (MS-5, then MS-4)** — MS-5 is S and unblocked (the field already exists). MS-4 is L and **gated on a product fork** (additive vs destructive) with real multi-user/data-entry risk, so it trails MS-5 and only proceeds after the fork resolves (option B spins out into its own plan).
3. **Schema-touching verify (MS-6)** — placed in a dedicated verify-and-close step so the live-DB verification + ticket closeout happen with DB context, behind an explicit "do not re-implement" guard. Non-blocking on the rest of the sprint if the owner is unavailable.
4. **Risky architecture (MS-7)** — last among the existing-app items, behind the heaviest verification gate (two-client e2e + manual concurrency test) and a mandatory product decision; it reopens the exact surface v2.5.0 hardened.
5. **New feature (MS-8)** — largest new surface, gated by a design spike producing its own sub-spec; runs in parallel-friendly isolation (separate Docker stack; near-zero C++ — one required desktop merge fix ships in v2.6.0, Phase 5.0) on an independent versioning track, so it proceeds independently of the C++ branch once the network decision lands.

This puts **value early, risk late, and every product-fork item behind an explicit gate** so the deployable v2.6.0 never absorbs an un-decided architectural change by accident.

---

## Definition of Done (whole initiative)

- Every **Ready** item (MS-1, MS-2, MS-3, MS-5) implemented, visually signed off, and merged into the batch branch.
- Every **gated** item (MS-4, MS-7) has its product fork resolved by the owner **before** code; MS-6 verified-and-closed (no code); MS-8 has its spike + sub-spec approved before any service code.
- Full `tests\run-tests.ps1` green except the pre-existing flaky `tst_responsivelayout`; new tests added per item (plot y-anchor, Notes-panel, session-field arbitration e2e, web-form pytest).
- Build clean under `-Werror -Wall -Wextra -Wpedantic`; `/ponytail-review` clean on each non-trivial diff.
- All internal v2.5.x DataViewer.exe patches wrapped into a single deployable **v2.6.0** with a clean rebuild (`mingw32-make clean`), FileVersion-verified `dist\DataViewer-setup.exe`. (MS-8 ships separately and is not part of the wrap.)
- `tests\deployment\Test-Deployment.ps1` all phases pass on the work machine.
- `tasks/lessons.md` updated for every corrected UI/architectural assumption (grid-is-editor, ribbon label cap, whole-session-path-is-load-bearing, backfill-already-shipped).
- This spec updated to reflect any fork decisions; the plan re-derived to match.
- Installer **left in-repo** for the owner's manual Synology drop — never auto-dropped.
