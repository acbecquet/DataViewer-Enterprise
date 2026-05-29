# Per-Row Puffing Regime — Template Change (v2.2.1)

- **Date:** 2026-05-29
- **Status:** Approved design, ready for implementation plan
- **Scope:** Fundamental test-template change — replace the unused per-row *Resistance* column with a per-row *Puffing Regime* column, and propagate it through the entire app (pipeline, database, data table, plots, reports, template file).

---

## 1. Goal

The standardized test template's per-row **`Resistance (Ω)`** column (Excel column **E**, the 5th per-row column) is effectively dead: it is almost always empty, no calculation reads it, and no plot or report uses it. The only resistance that matters is the **header-level** value (used to compute power).

Replace that per-row column with a per-row **`Puffing Regime`** string column (format `XXmL/YYs/ZZs`, e.g. `60mL/3s/30s`, `200mL/9s/300s`). This lets a single test session record **multiple regimes** across its puff ranges (e.g. puffs 5–30 at `60mL/3s/30s`, puffs 32–40 at `200mL/9s/300s`), enabling custom tests and far more readable, regime-labeled plots and reports.

Both **old-template** files (Resistance column, usually empty) and **new-template** files (per-row Puffing Regime, filled) must coexist everywhere — including in the shared PostgreSQL database and the offline SQLite snapshot.

---

## 2. Background — current state (from codebase exploration)

### Column parsing is 100% positional
Per-row columns are mapped by **index**, never by header name:
- `src/pipeline/SheetProcessors.cpp` — `namespace ColIdx { … RESISTANCE = 4 … }`; `buildSampleResult` reads `dr.resistance = varToDouble(cell(4))`.
- `src/pipeline/ReportData.h` — `namespace Cols { … RESISTANCE = 4 …, COUNT = 12 }` (UI index constants).
- The header string at the column-header row is read by `ExcelReader::getColumnHeaders()` **only for cosmetic display** (it normalizes `"Resistance"` → `"Resistance (Ω)"`); no code branches on it today.

### `DataRow::resistance` is a dead pass-through
`TpmCalculator` consumes only `puffs`, `beforeWeight`, `afterWeight`. `DataRow::resistance` (`ReportData.h:19`) is read into memory and stored, but never used in any calculation, plot, or report. (The *header-level* `SampleResult::resistance` drives `power = V²/(R+offset)` and is unaffected by this change.)

### A sample-level `puffingRegime` already exists (separate concern)
`SampleResult::puffingRegime` (`ReportData.h:52`) is parsed from header metadata (Excel row 2, col offset+7), stored in `samples.puffing_regime`, shown in the Properties dock, and printed in report tables. **This stays as-is.** The new field is a distinct **per-row** datum on `DataRow`.

### Template detection today is by sheet name and does NOT affect column parsing
`ExcelReader::detectTemplateVersion()` flags a file "new" if a sheet name matches a known list (`Lifetime Test`, etc.). Result stored in `SheetResult::templateVersion`/`FileResult::templateVersion`, but it has **zero** effect on per-row parsing. We need a **new, column-level** signal.

### Database — per-row data is fully columnar (no JSON/blob)
`deploy/postgres/init.sql` `data_rows` columns: `id, sample_id, sort_order, puffs, before_weight, after_weight, draw_pressure, resistance, smell, clog, notes, tpm, tpm_power_density, variation_tpm, oil_consumed, updated_at, updated_by, version`. There is **no** `data_rows.puffing_regime` (only `samples.puffing_regime`). Hierarchy: `files → tests(=sheet) → samples → data_rows`, all `ON DELETE CASCADE`.
- No runtime migration runner. Deployed DBs are migrated via hand-applied files in `deploy/postgres/migrations/`. `init.sql` is `CREATE TABLE IF NOT EXISTS` (won't add columns to existing tables). `schema_meta.schema_version` is currently `"2"`.
- `OfflineSnapshot.cpp` holds a **hardcoded** SQLite schema (`kCreateStatements`) and is regenerated from scratch on each clean online close — no SQLite migration needed, but the hardcoded schema + the positional `for (int c = 0; c < 18; …)` regenerate loop must be updated in lock-step.

### Live plot — actual state
`src/plotting/PlotWidget.cpp` plot-type combo has exactly **four** entries (no "Resistance" option exists):
`TPM Trend` · `TPM Bar Chart` · `Power Density` · `Draw Pressure`. None read `resistance`. ("Draw Resistance" in conversation = the **Draw Pressure** plot.)
- The live plot shows the **whole current sheet** (all samples), gated by per-sample checkboxes — not a single sample. Fed via `m_plotWidget->setSheetData(sheet)` from `MainWindow::displayCurrentSample()`.
- Top-bar layout (`PlotWidget` ctor): `typeLabel → m_plotTypeCombo → sep(VLine) → m_saveBtn → addStretch(1)`, inside a `QHBoxLayout` on a `fixedHeight(36)` bar.
- **Row-filtering precedent exists:** Data Cleanup (`buildCleanedSheet`/`buildCleanedSample`) builds a filtered `SheetResult` (rows removed, metrics recalculated) before `setSheetData`.

### Reports
`src/reporting/ReportGenerator.cpp` emits **one content slide per sheet** with a 3-plot horizontal strip (`kPlotLayout`, **triplicated** in `generateFullReport`, `generateTestReport`, `generateCombinedFullReport`). `buildPlots(SheetResult)` builds TPM Trend, Avg TPM Bar, Draw Pressure from `sheet.samples`/rows. Series colored via `AppTheme::seriesColors(n)` (no modulo; never reuses a color). Per-row resistance is not used anywhere in reporting.

### Template file
`resources/templates/Standardized Test Template - December 2025.xlsx` (shipped by `installer.iss`; also the report SOP source). 11 data sheets + 1 SOP sheet. **36** `Resistance (Ω)` header cells across all sample blocks (12-col stride; 2–8 blocks per sheet). Column **E** is blank in the template and **referenced by zero formulas** (TPM formulas reference columns A/B/C/I and header `F$2`). Each sheet carries its standard regime in the header `Puffing Regime:` cell (row 2):

| Sheet | Header regime |
|---|---|
| Negative Pressure Test | 60mL/3s/30s |
| Lifetime Test | 60mL/3s/30s |
| User Test Simulation | 100mL/2.5s/15s |
| Long Puff Lifetime Test | 200mL/10s/60s |
| Rapid Puff Lifetime Test | 60mL/2s/5s |
| Intense Test | 200mL/3s/30s |
| Big Headspace Serial Test | 60mL/3s/30s |
| Temperature Cycling Test #1 | *(non-standard layout — inspect; no standard Resistance headers found)* |
| Temperature Cycling Test #2 | 60mL/3s/30s |
| Viscosity Compatibility | 60mL/3s/30s |
| Various Oil Compatibility | 60mL/3s/30s |

The `.xlsm` "Automated Testing Template - DVE" is **out of scope** — the user will update it.

---

## 3. Approach

### Additive & non-destructive (chosen)
Add a new per-row `QString puffingRegime` to `DataRow` **alongside** the preserved `double resistance`. Add a new `data_rows.puffing_regime TEXT` column **alongside** the preserved `resistance`. Old files populate `resistance`; new files populate `puffingRegime`. No data, column, or behavior is removed.

*Rejected — full replacement:* dropping `resistance`/`data_rows.resistance` is destructive to the live multi-user DB (existing rows), loses old-file round-trip fidelity, and makes old/new distinction harder. Not justified for a near-empty column.

*Rejected — content sniffing per cell:* deciding "number → resistance, text → regime" per value is fragile, gives no clean old/new signal, and breaks on blanks. Rejected.

### Detection — per-sheet, header-string driven
The column-E header string is the authoritative signal:
- `"Puffing Regime"` (case-insensitive contains `puffing`/`regime`) → **new-template**.
- `"Resistance"`/`"Resistance (Ω)"` → **old-template** (default/fallback).

Cached as a new `bool SheetResult::hasPerRowRegime`:
- **Excel ingest:** set from the column-E header string in `ExcelReader::getColumnHeaders()` / `SheetProcessors`.
- **DB load:** derive from data — a sheet is new-template if any of its `data_rows.puffing_regime` is **non-NULL** (see §5 for the NULL-vs-empty persistence rule that keeps this robust even when a regime cell is blank).

One physical UI column (#4) switches its **label, editor, and read/write type** on this flag.

---

## 4. Component design

### 4.1 Pipeline & data model
- `ReportData.h`: add `QString puffingRegime;` to `DataRow`. Keep `resistance`. Add `bool hasPerRowRegime = false;` to `SheetResult`. Keep `Cols::RESISTANCE = 4` as the shared index but treat index 4 as a dual-purpose column (rename-or-comment; do **not** add a 13th column — `Cols::COUNT` stays 12).
- `SheetProcessors::buildSampleResult`: if `hasPerRowRegime`, read `cell(4)` as a **string** into `dr.puffingRegime` (leave `dr.resistance = 0`); else read as a **double** into `dr.resistance` (today's behavior). Propagate the flag into the produced `SheetResult`.
- `ExcelReader::getColumnHeaders()`: recognize a `"Puffing Regime"` header at index 4; set/return the per-sheet flag; keep the cosmetic `"Resistance"` → `"Resistance (Ω)"` normalization for old files.

### 4.2 Database (schema v2 → v3)
- `init.sql`: add `puffing_regime TEXT` to `data_rows` (after `resistance`).
- New migration `deploy/postgres/migrations/2026-05-29-v2.2.1-per-row-regime.sql`: `ALTER TABLE data_rows ADD COLUMN IF NOT EXISTS puffing_regime TEXT;` + bump `schema_meta` to `3`.
- `DatabaseManager.cpp`: add `puffing_regime` to `data_rows` INSERT/UPDATE column lists and the bulk SELECT; add the bind/read at the correct positional indices (careful: shifts subsequent indices). **Persistence rule:** bind `puffing_regime` as a real value (possibly empty string) for new-template rows, and as **NULL** for old-template rows; on read, `isNull()` → old (leave `puffingRegime` empty), non-null → new.
- `OfflineSnapshot.cpp`: add the column to the hardcoded `data_rows` schema; update the regenerate SELECT/INSERT and bump the positional loop `18 → 19`; update `loadFile` read.
- `LiveSync.cpp`: add `"puffing_regime"` to the `data_rows` cell-write allowlist.
- `MigrationTool.cpp`: add `"puffing_regime"` to `kColsDataRows`.

### 4.3 Data table (TPM mode, `MainWindow.cpp`)
- `dataTableHeaders()` (and the raw-table restore sites) become flag-aware: index 4 label = **"Puffing Regime"** when `hasPerRowRegime`, else **"Resistance"**.
- Populate loop (~`MainWindow.cpp:2955`): new-template → show `dr.puffingRegime` string; old-template → today's `(resistance==0)?empty:num`.
- `onTableCellChanged` case 4: new-template → `dr.puffingRegime = text;` old-template → `dr.resistance = text.toDouble();`.
- Write-back (`MainWindow.cpp:1748`): unchanged arithmetic (`excelCol = sampleIdx*12 + col + 1`); for new-template it queues the **string** (openpyxl writes strings fine).
- `kDataTableColumns()` (UI-index → DB-column for LiveSync repaint): index 4 maps to `"puffing_regime"` when `hasPerRowRegime`, else `"resistance"`.
- **Combo delegate (new):** a `QStyledItemDelegate` on column 4, active only for new-template sheets. Editable `QComboBox` pre-loaded with the file's unique regimes + common presets (`60mL/3s/30s`, `200mL/10s/60s`, `100mL/2.5s/15s`, `60mL/2s/5s`, `200mL/3s/30s`), free text allowed (helps keep the `XXmL/YYs/ZZs` format consistent).

### 4.4 Live plot — regime picker (`PlotWidget`)
- New `m_regimeCombo` inserted **between `m_plotTypeCombo` and `m_saveBtn`** (with its own label/separator), within the fixed-height top bar.
- Populated from the **file's** unique per-row regimes, with an **"All regimes"** default first. MainWindow supplies the list (e.g. `PlotWidget::setAvailableRegimes(QStringList)`) when the file loads/changes; for old files the list is empty and the combo is **hidden/disabled**.
- On selection (≠ "All"): MainWindow builds a **regime-filtered `SheetResult`** (keep only rows whose `dr.puffingRegime` == selection, recompute sample metrics) via the same filter+recalc path as Data Cleanup / report fan-out, and feeds it to `setSheetData`. This is required for correctness — the **TPM Bar Chart** reads precomputed per-sample averages, so a plain row-skip in `renderCurrentPlot` would render wrong bars. "All" / old files → no filtering (today's behavior). Filtering composes with the existing per-sample checkbox visibility and Data-Cleanup exclusions. Caveat: a regime present in the file but not the current sheet yields an empty plot (acceptable; documented). The picker scopes the **plot only** — the data table is unaffected.

### 4.5 Reports & plots — per-regime fan-out (`ReportGenerator.cpp`)
- For each sheet, compute the unique per-row regimes present (blank → bucket `"(unspecified)"`). For **each** regime, build a filtered `SheetResult` (keep only that regime's rows, recompute sample metrics — reuse the `buildCleanedSample` pattern) and emit a content slide titled **"{sheet} — {regime}"** with the full plot set + regime-filtered table.
  - One regime (or old file) → one slide as today (title gains the regime label; for old files, no regime suffix).
  - N regimes → N slides.
- **Enabling refactor:** extract the triplicated `kPlotLayout` into a single shared definition and a shared "emit one content slide for a (filtered) SheetResult" helper; add the regime loop once and call it from all three report paths. (Consistent with the post-v2.0.7 refactor backlog.)

### 4.6 Template `.xlsx`
- Python/openpyxl script (run via the MIP-allowlisted Python313): for every sample block on every data sheet, locate the per-row header row robustly (search for the `Resistance (Ω)` / `Resistance` cell at the block's column-E position; do **not** assume row 4), relabel it to **"Puffing Regime"** preserving the existing header cell style, and **pre-fill** that block's data rows in column E with the block's header `Puffing Regime:` value (single source of truth).
- Inspect `Temperature Cycling Test #1` (non-standard layout) and handle or explicitly skip it with a logged note.
- Verify formulas/conditional formatting are untouched (column E is referenced by none).
- Leave the `.xlsm` automated template alone.

### 4.7 Tests & version
- `tests/generate_fixtures.py`: add a **new-template** fixture (column-E header `"Puffing Regime"`, rows carrying regime strings, including a sample that mixes two regimes mid-session) alongside an existing old-template fixture.
- Extend: `tst_excelreader` (header detection sets the flag; parses column 4 as string vs double), `tst_sheetprocessors` (both branches), `tst_databasemanager` (round-trip `puffing_regime`, NULL-vs-value semantics, old/new coexistence), `tst_reportgenerator` (a sheet with 2 regimes → 2 content slides; 1 regime → 1; old file unchanged). Run via the `test-dataviewer` flow / `tests\run-tests.ps1`.
- Bump `VERSION` → `2.2.1` in the `.pro`; **clean rebuild required** (`mingw32-make clean && -j8`) so `main.o` re-embeds the version. Stop at building in-repo; the installer build + Synology transfer are the user's manual steps.

---

## 5. Backward compatibility & coexistence

- **Old `.xlsx` files:** column-E header is `Resistance` → `hasPerRowRegime=false`; column 4 stays "Resistance" (numeric, no combo), no regime picker, one slide per sheet. Identical to today.
- **Existing DB rows** (all old-template): `data_rows.puffing_regime` stays NULL after migration; loaded as old-template. Untouched.
- **New files:** `hasPerRowRegime=true` end-to-end; regime drives table, picker, and report fan-out.
- **Robust flag after DB round-trip:** persist NULL for old rows, a (possibly empty) string for new rows; derive `hasPerRowRegime` from "any non-NULL `puffing_regime` in the sheet". Edge case: a new-template sheet whose every regime cell was deliberately cleared *and* persisted as NULL would read back as old-template — pre-filled templates avoid this, and the degraded view (empty numeric column 4) is harmless. Documented, accepted.

---

## 6. Edge cases & error handling

- **Mixed regimes within one sample:** the defining new capability — per-row filter handles it; TPM Bar Chart recomputes each sample's average over the selected regime's rows only.
- **Blank/empty regime cells in a new file:** bucket as `"(unspecified)"` in pickers and report fan-out so no rows are silently dropped.
- **`Variation in TPM (%)`** is sequential row-to-row; when filtering to a regime subset, recompute over the subset (same as cleanup recalculation).
- **Regime in file but not current sheet:** live picker selection yields an empty plot (acceptable).
- **`Temperature Cycling Test #1`** non-standard layout: inspect during implementation; relabel if it has the column, else skip with a logged note.

---

## 7. Out of scope

- The `.xlsm` "Automated Testing Template - DVE" (user-owned).
- Installer build, deployment self-test, and Synology transfer (user's manual release steps).
- Surfacing per-row regime in the Properties dock (the existing sample-level header regime already lives there).
- Any new plot **types** (the four existing types are reused, now regime-filterable).

---

## 8. Implementation phases (detail to follow in the plan)

1. **Data model + pipeline** — `DataRow::puffingRegime`, `SheetResult::hasPerRowRegime`, header detection, dual-mode column-4 parsing.
2. **Database** — schema/migration/snapshot/LiveSync/MigrationTool, NULL-vs-value persistence.
3. **Data table UI** — flag-aware header/populate/edit/write-back/`kDataTableColumns`, combo delegate.
4. **Live plot** — regime picker + per-row plot filtering.
5. **Reports** — `kPlotLayout` de-duplication + shared slide helper + per-regime fan-out.
6. **Template `.xlsx`** — relabel 36 headers + pre-fill regimes (openpyxl).
7. **Tests + version bump** — fixtures, suite updates, `VERSION` → 2.2.1, clean rebuild.
