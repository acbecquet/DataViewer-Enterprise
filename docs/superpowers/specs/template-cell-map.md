# Template Cell-Position Map

**Source template:** `resources/templates/Standardized Test Template - December 2025.xlsx`
**Inspected:** 2026-05-27 (via `tools/inspect_template_cells.py`)
**Purpose:** Drive `ExcelExporter::exportToXlsm` so a written `FileResult`
round-trips perfectly back through `ExcelReader` + `DataProcessor`.

All positions below are 1-based (openpyxl convention) unless noted.
`off` = `sample_index * 12` (0-based column offset before adding 1-based
positions inside the block). Each sample block is **12 columns wide**.

## TL;DR

- All 10 standard sample-bearing sheets use the **identical** 12-col block
  layout. One sheet (`Temperature Cycling Test #1`) is a procedural
  checklist and gets handled separately.
- Header metadata lives in rows 1-3 inside each 12-col block (cols
  off+2 through off+8); no metadata exists outside the block.
- The "Did this burn?" / "Did this clog?" / "Did this leak?" answer cells
  (off+11, rows 1/2/3) are **part of the template band but the pipeline
  never reads them**. Burn/clog/leak detection is driven by keyword scan
  of the Notes column (col off+8 in data rows). Documented below under
  **DB-only fields**.
- Row 4 is the data-column-header row. Do not overwrite it.
- Several cells in the off+9..+12 band of rows 5+ are template formulas
  (TPM, TPM-power-density, variation, oil-consumed running total). Do not
  overwrite.
- Existing `kWriteHeaders` in `MainWindow.cpp` line 1590 is **correct**:
  every cell it writes matches the position `ExcelReader` reads from. The
  exporter must add: `testName` (off+1 row 1), `date` (off+4 row 1),
  `heatingTechnology` (off+8 row 1), and `power` (off+6 row 2 -- but
  power is an ArrayFormula in the template, see notes below).

## ASCII layout of one 12-col sample block

Columns are labelled with their offset inside the block (1-based, so the
first sample occupies block-cols 1..12). For sample 0 these are sheet
columns A..L. For sample 1 they are M..X. Etc.

```
              +1          +2          +3          +4          +5          +6           +7         +8           +9          +10          +11      +12
Row 1   TestName       --          Date:       <date>      SampleID:   <id>         HeatTech:  <tech>       --          Burn?        <BURN>   AvgTPM lbl
Row 2   Media:         <media>     Resist:     <ohms>      Power:      <P=formula>  PuffReg:   <regime>     UsageEff(arrow) Clog?    <CLOG>   =AVERAGE(I5:I115...)
Row 3   Viscosity:     <visc>      Tester:     <name>      Voltage:    <volts>      OilMass:   <init oil>   (formula)   Leak?        <LEAK>   =STDEV(I5:I115...)
Row 4   puffs          Before(g)   After(g)    DrawPress   Resistance  Smell        Clog       Notes        TPM         TPM/PD       Var%     OilCum(g)
Row 5+  <data>         <data>      <data>      <data>      <data>      <data>       <data>     <data>       [formula]   [formula]    [formula] [formula]
```

Labels in `<>` are where the exporter writes values. Cells marked
`[formula]` or `=...` must NOT be overwritten -- they are part of the
template and recalc when Excel opens the file.

- Row 1 col+12 = "Average TPM and Standard deviation" label (display).
- Row 2 col+12 = `=IF(I5="","",AVERAGEIFS(I5:I115,...))` (formula).
- Row 3 col+12 = ArrayFormula `=...STDEV.P(I5:I115)/...`.
- Row 2 col+6  = ArrayFormula for power: `=V^2/(R+offset)`.
- Row 3 col+9  = ArrayFormula for Usage Efficiency.

## Field-by-field cell map (the export target)

`SampleResult` fields organised by where they live in the template
(if anywhere). All row/col are **1-based for openpyxl**. Add `off` =
`sampleIndex * 12` to every column number to get the absolute sheet
column for sample N.

### Identification

| Field             | Row | Col (1-based)  | Example       | Source                           |
|-------------------|-----|----------------|---------------|----------------------------------|
| sampleName        | --  | --             | "Lifetime-1"  | Derived from sampleID; not stored separately |
| sampleID          |  1  | off + 6        | "Lifetime-1"  | ExcelReader L466 (`off + 5` 0-based); kWriteHeaders writes `row=1, column=off+6` |
| date              |  1  | off + 4        | "2026-03-16"  | ExcelReader L464 (`off + 3` 0-based). **Not in kWriteHeaders** -- exporter must add. Real files store as native datetime. |
| tester            |  3  | off + 4        | "Charlie"     | ExcelReader L471 (`off + 3` 0-based, row 2); kWriteHeaders writes `row=3, column=off+4` |

### Device parameters

| Field             | Row | Col (1-based)  | Example         | Source                           |
|-------------------|-----|----------------|-----------------|----------------------------------|
| media             |  2  | off + 2        | "D9"            | ExcelReader L467; kWriteHeaders ok |
| resistance        |  2  | off + 4        | 1.1             | ExcelReader L468; kWriteHeaders ok |
| voltage           |  3  | off + 6        | 3.0             | ExcelReader L472; kWriteHeaders ok |
| power             |  2  | off + 6        | (formula)       | **DO NOT WRITE** -- template ArrayFormula computes from V and R. ExcelReader L500-502 also recomputes in C++. Skip in exporter. |
| heatingTechnology |  1  | off + 8        | "CCELL3.0"      | ExcelReader L465 (`off + 7` 0-based). **Not in kWriteHeaders** -- exporter must add. |
| puffingRegime     |  2  | off + 8        | "60mL/3s/30s"   | ExcelReader L469; kWriteHeaders ok |
| viscosity         |  3  | off + 2        | 500000          | ExcelReader L470; kWriteHeaders ok |
| initialOilMass    |  3  | off + 8        | 1.0             | ExcelReader L473; kWriteHeaders ok |

### Calculated metrics

All calculated metrics live in formula cells the template owns or are
derived in C++. **Do NOT write any of these from the exporter.**

| Field                | Row    | Col            | Note                                                  |
|----------------------|--------|----------------|-------------------------------------------------------|
| averageTPM           | 2      | off + 12       | Formula `=IF(I5="","",AVERAGEIFS(I5:I115,...))`. Computed by template. |
| stdDevTPM            | 3      | off + 12       | ArrayFormula. Computed by template.                   |
| averagePowerDensity  | --     | --             | DB-only; not displayed in template.                   |
| efficiencyPercent    | 3      | off + 9        | ArrayFormula (the down-arrow cell). Computed.         |
| totalOilConsumed     | --     | --             | Derived from data rows; not in template.              |
| totalPuffs           | --     | --             | Derived from data rows; not in template.              |
| normalizedTPM        | --     | --             | DB-only / report-only; not in template.               |

### Status flags (burn / clog / leak)

These have header-band answer cells AND a code-path that pulls them from
data rows. **The pipeline does NOT read the header-band answer cells.**

| Field        | Header-band visual cell (NOT round-trip)  | Pipeline-read source                  |
|--------------|--------------------------------------------|---------------------------------------|
| burnStatus   | Row 1, off + 11 (label at off + 10)        | Notes column (off + 8, rows 5+) keyword scan, or inline Y/N in cols off+9..+11 of any data row |
| clogStatus   | Row 2, off + 11 (label at off + 10)        | Same as burn                          |
| leakStatus   | Row 3, off + 11 (label at off + 10)        | Same as burn                          |

`SheetProcessors.cpp` lines 173-209:

- **Strategy 1** scans `rowData[8..10]` (0-based, = cols off+9..off+11
  1-based) of each data row for "Y"/"N" cells; assigns in order burn -->
  clog --> leak. But cols off+9..off+11 are the TPM / power-density /
  variation **formula** cells in the template, so writing "Y"/"N" there
  destroys the formula. Real operator files (`test data/*.xlsx`) do not
  use this path.
- **Strategy 2** scans the notes column (`s.rows[i].notes`, col off+8)
  for the substrings BURN/CLOG/LEAK (with NO BURN / NO CLOG / NO LEAK
  inversion).

**Recommendation for exporter:**

1. Write the human-readable status strings to the header-band answer cells
   (row 1/2/3, col off+11) so the file looks right when opened in Excel.
   Use `"Y"` / `"N"` / `""` (already normalised in `SampleResult`).
2. ALSO embed a `[BURN]`/`[CLOG]`/`[LEAK]` (or `[NO BURN]`/etc.) marker
   in the notes column of the first data row, so SheetProcessors
   strategy 2 picks it up on re-read.
3. **Verify with a round-trip test** that `burnStatus`/`clogStatus`/
   `leakStatus` survive export then re-read.

### Per-row data (rows 5..N)

All 1-based. Write each `DataRow` in `SampleResult.rows` to its own Excel
row starting at row 5.

| DataRow field     | Col offset (1-based) | Notes                                   |
|-------------------|----------------------|-----------------------------------------|
| puffs             | off + 1              | Row 5 holds the **literal** initial puff count (e.g. `1`, `2`, or `20`). Rows 6+ are formulas `=A{n-1}+<interval>` in the template. **The exporter must overwrite with literals from `DataRow.puffs`** -- the formula offsets vary per sheet (10, 20, etc.). |
| beforeWeight      | off + 2              | Row 5 has no formula (operator types). Rows 6+ are `=C{n-1}` formulas. Overwrite with literal `DataRow.beforeWeight`. |
| afterWeight       | off + 3              | Literal in template. Always writeable. |
| drawPressure      | off + 4              | Literal. |
| resistance        | off + 5              | Literal. |
| smell             | off + 6              | Literal. |
| clog              | off + 7              | Literal. (Per-row clog Y/N text, separate from sample-level clogStatus.) |
| notes             | off + 8              | Literal. Long strings ok. |
| tpm               | off + 9              | **FORMULA** -- DO NOT WRITE. |
| tpmPowerDensity   | off + 10             | **FORMULA** -- DO NOT WRITE. |
| variationTPM      | off + 11             | **FORMULA** -- DO NOT WRITE. |
| oilConsumed       | off + 12             | **FORMULA** (ArrayFormula in rows 6+, `=IF(I5="",...)` in row 5) -- DO NOT WRITE. |

### Image data and DB identity

These never appear in the template:

- `imagePaths`, `imageLayouts`, `imageCrops`, `imageIds`, `imageVersions`
  -- exporter behaviour TBD (could embed as worksheet images via
  openpyxl, but out of scope for first round-trip).
- `id`, `version` (sample, sheet, file levels) -- DB-only.
- `extra` (`QMap<QString, QVariant>`) -- sheet-specific overflow; not in
  template.

### DB-only / report-only fields (do NOT export)

| Field                          | Why excluded                                |
|--------------------------------|---------------------------------------------|
| sampleName                     | Derived; same as sampleID if non-empty.     |
| power                          | Template formula; recomputed on open.       |
| averageTPM, stdDevTPM          | Template formulas; recomputed on open.      |
| averagePowerDensity            | Pure derived; not in template.              |
| efficiencyPercent              | Template formula in cell off+9 row 3.       |
| totalOilConsumed               | Pure derived; sum of `oilConsumed` rows.    |
| totalPuffs                     | Pure derived; last data row's `puffs`.      |
| normalizedTPM                  | Pure derived (TPM / W).                     |
| extra                          | Sheet-specific overflow; not in template.   |
| imagePaths and friends         | Out of scope for first-round exporter.      |
| id, version                    | Server-assigned; not visible in Excel.      |
| (per-row) id, version          | Same as sample-level identity.              |
| (per-row) tpm, tpmPowerDensity, variationTPM, oilConsumed | Template formulas. |

## Per-sheet variation

The December 2025 template ships 12 sheets. Their classifications:

| Sheet                              | Cols | Sample blocks | Layout            |
|------------------------------------|------|---------------|-------------------|
| Test SOP's                         | --   | 0             | Raw table (text). Skip in exporter. `isRawTable=true`. |
| Negative Pressure Test             | 72   | 6             | Standard 12-col   |
| Lifetime Test                      | 24   | 2             | Standard 12-col   |
| User Test Simulation               | 24   | 2             | Standard 12-col   |
| Long Puff Lifetime Test            | 24   | 2             | Standard 12-col   |
| Rapid Puff Lifetime Test           | 24   | 2             | Standard 12-col   |
| Intense Test                       | 24   | 2             | Standard 12-col   |
| Big Headspace Serial Test          | 24   | 2             | Standard 12-col   |
| Temperature Cycling Test #1        | 78   | 0             | **Procedural / checklist.** Different layout: tester at B1/B2, media at B2, voltage at D2, resistance at D3, heater at F1/G1. No `puffs` data column. No 12-col sample blocks. **Skip in exporter** (treat as `isRawTable` or pass-through). |
| Temperature Cycling Test #2        | 72   | 6             | Standard 12-col   |
| Viscosity Compatibility            | 96   | 8             | Standard 12-col (widest sheet) |
| Various Oil Compatibility          | 48   | 4             | Standard 12-col   |

The pipeline detects `Temperature Cycling Test #1` as raw via `isRawTable`
in `SheetProcessors`. The exporter should follow the same rule.

Row-5 puffs values seen in the template (informational -- exporter
overrides per `DataRow.puffs`):

- `1` for Negative Pressure, Big Headspace, Viscosity Compat, Various Oil
- `20` for Lifetime Test, Long Puff, Rapid Puff, Temperature Cycling #2
- Varies for Intense, User Sim

## Formula cells the exporter must NOT overwrite

Cataloged from the December 2025 template, sample 0 (cols 1-12). For
sample N, add `off = N*12` to each column number.

### Header band

| Cell                  | Type           | Formula (sample 0)                              |
|-----------------------|----------------|-------------------------------------------------|
| F2  (off+6 row 2)     | ArrayFormula   | `=...` power, derived from V and R.             |
| I3  (off+9 row 3)     | ArrayFormula   | Usage Efficiency.                               |
| L2  (off+12 row 2)    | regular        | `=IF(I5="","",AVERAGEIFS(I5:I115,I5:I115,"<>0",I5:I115,"<>"))` |
| L3  (off+12 row 3)    | ArrayFormula   | StdDev of column I (TPM).                       |

### Data band (rows 5..N)

For sample 0 (cols A..L). Add `off` for other samples.

| Cell        | Formula (sample 0)                                                                          |
|-------------|---------------------------------------------------------------------------------------------|
| A5          | `1` or `20` literal (puffs base). Overwrite with `DataRow.puffs`.                          |
| A6..A115    | `=A{n-1}+<interval>` -- chain. Overwrite all with literals.                                  |
| B6..B115    | `=C{n-1}` -- chain (before = previous after). Overwrite all with literals.                   |
| I5..I115    | `=IF(C{n}="","",(B{n}-C{n})*1000/A{n})` (row 5) or `(B-C)*1000/(A{n}-A{n-1})` (rows 6+). TPM. Do NOT write. |
| J5..J115    | `=IF(I{n}="","",I{n}/F$2/3)`. TPM Power Density. Do NOT write.                              |
| K6..K115    | `=IF(I{n}="","", 100*_xlfn.STDEV.P(I5:I{n})/AVERAGE(I5:I{n}))`. Variation in TPM. Do NOT write. |
| L5..L115    | `=IF(I{n}="","", IFERROR(((IF(I$5<>"", I$5*A$5, 0))/1000), ""))` (row 5) or ArrayFormula (rows 6+). Oil consumed. Do NOT write. |

### Practical exporter rule

For data rows: write only cols off+1 through off+8 (puffs through notes).
**Override** the A and B chain formulas in rows 6+ with literal values
from `DataRow.puffs` and `DataRow.beforeWeight` -- otherwise the formulas
will recompute against your literal C/D values and produce garbage. The
ExcelReader pipeline already does this fix-up on read (see SheetProcessors
"fill in missing Before Weight values" and "fill in missing Puffs values"
blocks), but the cleanest export is to commit literals.

## kWriteHeaders cross-reference (verification)

Each row below shows: target field, what `kWriteHeaders` writes, what
`ExcelReader::extractMetadata` reads, agreement check.

| Field             | kWriteHeaders position             | ExcelReader read position           | Agree? |
|-------------------|------------------------------------|-------------------------------------|--------|
| sampleID          | `row=1, column=off+6`              | `getCellString(0, off+5)` (0-based) | YES -- both -> R1 col 6 (1-based) |
| media             | `row=2, column=off+2`              | `getCellString(1, off+1)`           | YES -- R2 col 2 |
| resistance        | `row=2, column=off+4`              | `getCellDouble(1, off+3)`           | YES -- R2 col 4 |
| puffingRegime     | `row=2, column=off+8`              | `getCellString(1, off+7)`           | YES -- R2 col 8 |
| viscosity         | `row=3, column=off+2`              | `getCellDouble(2, off+1)`           | YES -- R3 col 2 |
| tester            | `row=3, column=off+4`              | `getCellString(2, off+3)`           | YES -- R3 col 4 |
| voltage           | `row=3, column=off+6`              | `getCellDouble(2, off+5)`           | YES -- R3 col 6 |
| initialOilMass    | `row=3, column=off+8`              | `getCellDouble(2, off+7)`           | YES -- R3 col 8 |

The existing kWriteHeaders is **exhaustively correct** for the eight
fields it writes; the exporter can lift its cell positions verbatim.

### Fields ExcelReader reads but kWriteHeaders does NOT write

| Field             | ExcelReader read position           | Exporter write target                |
|-------------------|-------------------------------------|--------------------------------------|
| testName          | `getCellString(0, off+0)`           | row 1, col off+1 (matches sheet name; default to sheet name in exporter, since SampleResult does not carry testName) |
| date              | `getCellString(0, off+3)`           | row 1, col off+4 |
| heatingTechnology | `getCellString(0, off+7)`           | row 1, col off+8 |

Power is read but always recomputed from V and R; do not write.

## Validation method

The agreement table above was verified by:

1. Reading `src/MainWindow.cpp` lines 1590-1614 for the kWriteHeaders Python script.
2. Reading `src/ExcelReader.cpp` lines 384-486 for the `extractMetadata`
   function (the "new" template branch matches the December 2025 template).
3. Running `tools/inspect_template_cells.py` to walk the actual template
   workbook with openpyxl, capturing every non-empty cell in rows 1-6 and
   every formula in the data band rows 5-15.
4. Running `tools/inspect_real_data.py` and `tools/inspect_real_data2.py`
   on two operator-generated files in `test data/` to confirm what
   real-world file structure looks like (in particular: K1/K2/K3 status
   answer cells are EMPTY in real files; status detection relies on the
   notes column).

Concrete check executed: opened
`test data/GTI ODM testing 3-16-2026.xlsx` and confirmed:

- Sheet `Lifetime Test`, sample 0: D1=date, F1=sampleID at col 6 row 1
  (kWriteHeaders position) -- present and correctly typed.
- Sheet `Lifetime Test`, sample 0: H1="CCELL3.0" at col 8 row 1
  (heating tech, not in kWriteHeaders) -- present, exporter must include.
- All `K1`/`K2`/`K3` answer cells are `None` even in well-used files.

No disagreements between kWriteHeaders and ExcelReader found; the gap is
that kWriteHeaders is incomplete (missing testName, date,
heatingTechnology -- which the exporter must add).

## Cell positions summary table

For convenience, the complete writable map (1-based, all relative to
`off = sample_index * 12`):

| Sample-level field   | Row | Col       | Notes                              |
|----------------------|-----|-----------|------------------------------------|
| testName             |  1  | off + 1   | Usually equals sheet name          |
| date                 |  1  | off + 4   | Native datetime or YYYY-MM-DD str  |
| sampleID             |  1  | off + 6   | String                             |
| heatingTechnology    |  1  | off + 8   | String                             |
| burnStatus visual    |  1  | off + 11  | Optional; pipeline does not read it |
| media                |  2  | off + 2   | String                             |
| resistance           |  2  | off + 4   | Float (ohms)                       |
| puffingRegime        |  2  | off + 8   | String                             |
| clogStatus visual    |  2  | off + 11  | Optional; pipeline does not read it |
| viscosity            |  3  | off + 2   | Float                              |
| tester               |  3  | off + 4   | String                             |
| voltage              |  3  | off + 6   | Float (V)                          |
| initialOilMass       |  3  | off + 8   | Float (g)                          |
| leakStatus visual    |  3  | off + 11  | Optional; pipeline does not read it |

| Per-row field      | Row    | Col       | Notes                              |
|--------------------|--------|-----------|------------------------------------|
| puffs              | 5..N   | off + 1   | Override formula chain with literal |
| beforeWeight       | 5..N   | off + 2   | Override formula chain with literal |
| afterWeight        | 5..N   | off + 3   |                                    |
| drawPressure       | 5..N   | off + 4   |                                    |
| resistance         | 5..N   | off + 5   | Per-row resistance reading         |
| smell              | 5..N   | off + 6   |                                    |
| clog               | 5..N   | off + 7   | Per-row clog Y/N text (not sample-level) |
| notes              | 5..N   | off + 8   | Encode BURN/CLOG/LEAK keywords here for round-trip |

Forbidden writes (template formulas):

| Cell                         | Why                                |
|------------------------------|------------------------------------|
| off + 12, row 1              | "Average TPM and Standard deviation" label |
| off + 6,  row 2              | Power ArrayFormula                 |
| off + 9,  row 2              | Usage Efficiency arrow label       |
| off + 12, row 2              | TPM average formula                |
| off + 9,  row 3              | Usage Efficiency ArrayFormula      |
| off + 12, row 3              | TPM stddev ArrayFormula            |
| All of row 4 (data column header labels) |                       |
| off + 9..+12, rows 5+        | TPM / TPM/PD / Var / OilCum formulas |
| off + 1, rows 6+             | Puff-chain `=A{n-1}+interval`      |
| off + 2, rows 6+             | Before-weight chain `=C{n-1}`      |

The puff and before-weight chain formulas in rows 6+ must be
**overwritten** with literal values from the `DataRow` vector (per the
recommendation above); leaving them in place causes cascade recompute
against the literal weights and breaks the round-trip on re-read.
