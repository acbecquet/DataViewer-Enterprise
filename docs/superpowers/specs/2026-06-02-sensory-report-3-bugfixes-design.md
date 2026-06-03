# Design: Three Sensory-Report Bug Fixes

- **Date:** 2026-06-02
- **Branch:** `dev` (current `VERSION` 2.2.1)
- **Mode affected:** simple Sensory Report mode (`SensoryPanel` + `RadarChartWidget` + `SensoryReportSource`). *Not* Detailed Sensory.
- **Status:** approved design, ready for implementation plan

## Overview

Three independent bug fixes in the simple Sensory report workflow:

1. **Full WYSIWYG** — element positions arranged in the report Preview now reach the generated `.pptx` for *all* movable elements (today only the table and radar chart translate).
2. **"Round" field** — a new `1 / 2 / N/A` dropdown beside Tester that separates the round marker out of the Tester text, recombining it only for the output filename/title. UI-only; no DB/schema change.
3. **Radar label nudge** — raise the *Smoothness* and *Burnt Taste* axis labels ~3 px so their gap from the pentagon matches the bottom labels.

All three are independent and can be implemented/reviewed separately. No database schema change, no `ReportLayout` struct change.

Decisions locked in with the user before writing this spec:

- Bug 1 scope: **full WYSIWYG** — wire every movable element, not just the content slide.
- Bug 1 default behavior: **preview is authoritative** — unmodified reports may shift slightly so the pptx matches what the canvas shows (kept deliberately; not aligning defaults back to the old pptx numbers).
- Bug 2 storage: **split/recombine in the UI only** — `testerName` keeps storing the combined `"Charlie R1"` string; no `round` column.
- Bug 2 default value: **"1"** for a fresh session.

---

## Bug 1 — Full WYSIWYG: report-element positions reach the pptx

### Root cause

The report Preview canvas (`ReportPreviewDialog`) renders every element from the live `ReportLayout` and lets the user drag/resize each one (title, table, radar, properties box, cover/divider titles). On export, `ReportPreviewDialog::onCreateReport()` passes that same edited `m_layout` to `SensoryReportSource::writePptx` -> `writeSensoryPptx`. The pptx writer, however, only applies the moved geometry for **two** element types:

- Content-slide **table** rect — `PptxWriter::addContentSlide` (5-arg), `src/reporting/PptxWriter.cpp:189`
- Content-slide **radar** rect — same function, `src/reporting/PptxWriter.cpp:201`

Every other element is re-derived from hardcoded math, ignoring the layout rect:

- Content-slide **title** — `buildContentSlideXml` hardcodes `(0.4, 0.1, 11.0, titleH)`, `src/reporting/PptxWriter.cpp:1253`. The 4-arg `addContentSlide` even notes the title rect is "not currently wired through buildContentSlideXml" (`:168`).
- **Properties box** — `writeSensoryPptx` builds its `extraXml` with a hardcoded bottom-right anchor `tbX/tbY/tbW/tbH`, `src/reporting/SensoryReportSource.cpp:743`. Only `propertiesBox.fontPt` is read from the layout (`:724`).
- **Cover** title + date/subtitle — `buildCoverSlideXml` hardcodes `(0.5, 2.0, 12.3, 1.5)` and `(0.5, 4.6, 12.3, 0.7)`, `src/reporting/PptxWriter.cpp:1204`/`:1214`. `addCoverSlide` only forwards font sizes, `src/reporting/SensoryReportSource.cpp:513`.
- **Divider** titles — sensory section dividers reuse `addCoverSlide(groupTitle, groupDate, dividerFontPt, 0)`, `src/reporting/SensoryReportSource.cpp:541` — same omission.

Fonts already propagate everywhere (titleFontPt, tableFontPt, propertiesBox.fontPt, coverTitleFontPt, coverSubtitleFontPt, dividerTitleFontPt). The gap is **positions only**, for everything except table + radar.

### Fix — extend the existing override pattern

The table/radar already use a clean convention in the 5-arg `addContentSlide`: *a null `QRectF` means "keep the position baked into the input"; a non-null rect overrides it* (`src/reporting/PptxWriter.cpp:186-206`). Apply the same convention to the remaining elements.

| Element | File / function | Change |
|---|---|---|
| Content-slide **title** | `PptxWriter::buildContentSlideXml` (`src/reporting/PptxWriter.cpp:1232`) | Add a `const QRectF& titleRect` parameter. When non-null, emit the title text box at its x/y/w/h instead of the hardcoded `(0.4, 0.1, 11.0, titleH)`. Thread `layout.title` from the 5-arg `addContentSlide` (`:226`). Keep `titleFontPt` behavior unchanged. OOXML text boxes auto-expand vertically at render, so a fixed rect height is safe for wrapped titles. |
| Content-slide **properties box** | `writeSensoryPptx` extraXml (`src/reporting/SensoryReportSource.cpp:719-776`) | When `slideLayout.propertiesBox.rect` is non-null, use its x/y/w/h for the textbox `xfrm` offset/extent instead of the hardcoded bottom-right anchor. Falls back to current math when null. |
| **Cumulative** slide title | cumulative emit path (`src/reporting/SensoryReportSource.cpp:945`, `addContentSlide(cumTitle, cumTable, cumPlots, layout.cumulative)`) | Title flows automatically once `addContentSlide` honors the title rect (table/radar already honored). The cumulative slide passes **no** `extraXml`, so it renders no properties box — `layout.cumulative.propertiesBox` is not applicable here. |
| **Cover** title + date/subtitle | `PptxWriter::buildCoverSlideXml` (`src/reporting/PptxWriter.cpp:1190`) + `addCoverSlide` (`:142`) | Add `const QRectF& titleRect` and `const QRectF& subtitleRect` parameters (default null for back-compat). When non-null, position the title/date boxes from them. `writeSensoryPptx` passes `layout.coverTitle` / `layout.coverSubtitle` (`:513`). |
| **Divider** titles | cover path reuse (`src/reporting/SensoryReportSource.cpp:541`) | When emitting a divider via `addCoverSlide`, pass `layout.dividerTitles.value(dividerKey)` as the title rect and a null subtitle rect. |

**Unchanged on purpose:**
- Content **table** and **radar** — already correct.
- **Image-slide** positions — already flow from `sess.imageLayouts` (`src/reporting/SensoryReportSource.cpp:797`); `ReportLayout.imageSlides` is out of scope for this fix.
- The Preview canvas / `ReportPreviewDialog` / `LayoutCommand` / `computeDefaultLayout` — all already correct. We are making the pptx match the canvas, not the other way around.
- The legacy `generateCombinedPptx` empty-layout path (DatabaseBrowser) — passes a default-constructed `ReportLayout` whose rects are null, so it keeps the hardcoded fallbacks and stays bit-for-bit unchanged.

### Deliberate behavior change (accepted)

`computeDefaultLayout` gives the canvas slightly different default positions than the pptx's hardcoded numbers — e.g. default title is full-width `(x=0.32, w=12.7)` on the canvas (`src/reporting/SensoryReportSource.cpp:1041`) vs the pptx's `(x=0.4, w=11.0)`; cover title default y is 2.5 vs 2.0; the props box default height is a fixed 2.70 vs a content-derived height. Once the pptx honors the layout rects, **unmodified reports adopt the canvas defaults** — i.e. the title widens, etc. This is the intended meaning of "preview is authoritative" and was explicitly approved. (If this ever needs reverting, the alternative is to set `computeDefaultLayout`'s defaults equal to the old hardcoded pptx values so unmodified output is pixel-identical while moved elements still translate.)

### Risk

Low-moderate. The change is additive (new optional rect params, null = old behavior) and isolated to the pptx emit path. Main verification: open a report, move the title/props/cover/divider in the Preview, export, and confirm the pptx matches; and confirm an *unmodified* report still renders cleanly (with the accepted default shift).

---

## Bug 2 — "Round" dropdown (1 / 2 / N/A), UI-only

### Goal

Add a `Round` dropdown immediately right of Tester in the simple Sensory header. Today the round is typed into the Tester field itself (`"Charlie R1"`); this separates it into its own control and recombines it only when forming the stored `testerName` (which drives the filename, navigator label, session key, and slide title). Round has no other effect on output.

### Approach — split/recombine at two chokepoints; `testerName` stays combined

No new stored field. `SensorySession::testerName` continues to hold the combined `"Charlie R1"` string, so `sessionLabel()` (`src/ui/SensoryPanel.cpp:1284`), the natural key, the DB columns, and the slide title (`src/reporting/SensoryReportSource.cpp:589`) all behave exactly as today. The header fields are plain `QLineEdit`s read only at `buildSession()` time (no per-keystroke LiveSync), so the new combo needs no signal plumbing for correctness.

**UI** — `SensoryPanel::buildHeaderRow` (`src/ui/SensoryPanel.cpp:703`):
- Add a `QComboBox*` member `m_roundCombo` (use a wheel-ignoring combo to avoid accidental scroll changes; if no such helper exists in this file, subclass `QComboBox` to `ignore()` wheel events, mirroring `DetailedSensoryPanel`'s `NoWheelComboBox`).
- Items: `"1"`, `"2"`, `"N/A"`. Default selection: **"1"**.
- Insert it into the `QHBoxLayout` right after the Tester field (`:718`) and before Media (`:719`). Because the layout is a left-to-right `QHBoxLayout` of `addWidget` calls, inserting here shifts Media / Save Test Headers / Date rightward automatically. Add a `new QLabel("Round:")` before the combo to match the other fields. Width ~80 px to match the compact options.

**Combine** — `SensoryPanel::buildSession` (`src/ui/SensoryPanel.cpp:891`), replace the plain read with a helper:

```cpp
QString uiTester = m_testerEdit->text().trimmed();
sess.testerName  = combineTesterRound(uiTester, m_roundCombo->currentText());
```

`combineTesterRound(tester, round)`:
- if `tester` is empty -> return `tester` (round is meaningless without a tester; preserves the empty-tester -> assessor fallback in `sessionLabel`).
- round `"1"` -> `tester + " R1"`
- round `"2"` -> `tester + " R2"`
- round `"N/A"` (or anything else) -> `tester` (no suffix)

**Split** — `SensoryPanel::applySession` (`src/ui/SensoryPanel.cpp:947`), replace the plain set:

```cpp
const auto tr = splitTesterRound(session.testerName);
m_testerEdit->setText(tr.tester);
m_roundCombo->setCurrentText(tr.round);   // "1" / "2" / "N/A"
```

`splitTesterRound(stored)` using `QRegularExpression(R"(^(.*\S)\s+R([12])$)")`:
- on match -> `{ captured(1), captured(2) }` (round = `"1"` or `"2"`)
- no match -> `{ stored, "N/A" }`

This auto-migrates every existing `"... R1"/"... R2"` session on load. Round-trip holds: `split(combine(t, r)) == (t, r)` for any `t` that doesn't itself end in `" R1"/" R2"`.

**Resets** — in the two new-session clears that call `m_testerEdit->clear()` (`src/ui/SensoryPanel.cpp:1173` and `:1200`), also reset `m_roundCombo` to `"1"`.

**Remote updates** — verify the incoming LiveSync `cellChanged` handler for `sensory_sessions` (`src/ui/SensoryPanel.cpp:2293`). If a remote `tester_name` change writes `m_testerEdit` directly (rather than going through `applySession`), route it through `splitTesterRound` too so the combo stays consistent. If it already re-applies the whole session, no change needed.

**`isDefaultState`** (`src/ui/SensoryPanel.cpp:1373`) — unchanged; it keys off the text fields being empty, and the combo's default does not make a blank session look non-default.

### Edge case (accepted)

A Tester name that genuinely ends in `" R1"` or `" R2"` would be parsed as a round on load. This is vanishingly unlikely in practice; trailing `R1`/`R2` is treated as the round marker by design.

### Risk

Low. UI-local; storage shape unchanged; existing data parses transparently. No schema/serialization/key changes.

---

## Bug 3 — Raise Smoothness & Burnt Taste labels ~3 px

### Root cause

In `RadarChartWidget::drawAxisLabels` (`src/ui/RadarChartWidget.cpp`), the two upper-side labels — Burnt Taste (`i==1`) and Smoothness (`i==4`) — use a special anchor computed along the upper polygon edge, finishing with a *downward* offset of `+ labelHalfH + 2.0` (`src/ui/RadarChartWidget.cpp:213`). Combined with the upper-half branch that sets the text-rect top so the **bottom** edge sits just above the anchor (`:232-233`), these two read tighter to the pentagon than the top/edge-anchored labels (Overall Liking, Vapor Volume, Overall Flavor). Qt's `+y` is downward, so reducing this anchor's y raises the labels.

### Fix

In the `i == 1 || i == 4` branch, change the anchor's y term:

```cpp
// before:
center.y() + labelCenter.y() + labelHalfH + 2.0
// after (raise ~3 px):
center.y() + labelCenter.y() + labelHalfH - 1.0
```

Only those two labels are affected (the `else` branch and the other three vertices are untouched). The `+/-36 deg` rotation applied afterward (`:263`) is unaffected because it rotates around the final text-rect center.

### Verification

This is a sub-pixel-sensitive visual tweak that can't be eyeballed headlessly. Implement the 3 px raise, then confirm in the running app that the Smoothness/Burnt-Taste gap visually matches Overall Flavor/Vapor Volume; nudge by +/-1 px if needed. The same widget renders both the on-screen chart and the pptx radar image, so the report inherits the change automatically.

### Risk

Very low. Single isolated arithmetic change to two labels' vertical anchor.

---

## Testing

Run via `tests\run-tests.ps1` (Qt Test suite). Per `test-dataviewer` skill, decrypt MIP files before building.

- **Bug 1 (reporting suite):** with a `ReportLayout` whose `title`, `propertiesBox.rect`, `coverTitle`, `coverSubtitle`, and a `dividerTitles[...]` entry are set to known non-default rects, assert the emitted slide XML contains `<a:off>`/`<a:ext>` (EMU) values matching those rects for the corresponding shapes. Add a companion case asserting that a default-constructed (empty) `ReportLayout` still produces valid slides using the fallback positions (guards the legacy DatabaseBrowser path).
- **Bug 2 (unit):** round-trip test for `combineTesterRound` / `splitTesterRound` across `{"1","2","N/A"}` and an empty tester; plus parse of legacy `"Charlie R1"` / `"Mary Jane R2"` / a plain `"Charlie"`.
- **Bug 3:** no automated assertion (visual). Confirm the radar suite still builds/passes; manual visual check in-app.

## Scope / non-goals

- Simple Sensory mode only. Detailed Sensory (`DetailedSensoryPanel`) is out of scope (though the same Bug-1 pattern likely applies there and could be a follow-up).
- No `ReportLayout` struct change, no DB schema/migration, no serialization change.
- Image-slide layout parameterization remains out of scope (pre-existing deferral).
- Auto-update / installer / Synology untouched; this is code-only and verifiable on the dev machine plus the Qt test suite.

## Files touched (summary)

- `src/reporting/PptxWriter.h` / `.cpp` — title rect param on `buildContentSlideXml` + 5-arg `addContentSlide`; title/subtitle rect params on `buildCoverSlideXml` + `addCoverSlide`.
- `src/reporting/SensoryReportSource.cpp` — pass `layout.title`, `coverTitle`, `coverSubtitle`, `dividerTitles[...]` through; honor `propertiesBox.rect` in content + cumulative extraXml.
- `src/ui/SensoryPanel.h` / `.cpp` — `m_roundCombo` member + label, `combineTesterRound` / `splitTesterRound`, wiring in `buildHeaderRow` / `buildSession` / `applySession` / new-session resets / remote-update path.
- `src/ui/RadarChartWidget.cpp` — one-line anchor tweak for `i==1`/`i==4`.
- `tests/` — reporting assertions for layout-honoring; unit test for tester/round split-combine.
