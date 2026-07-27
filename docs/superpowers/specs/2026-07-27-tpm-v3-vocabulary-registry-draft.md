# TPM v3 Vocabulary Registry - Workshop Draft

Date: 2026-07-27.
Status: WORKING DRAFT for the owner vocabulary workshop (pre-Phase-2 gate).
This document inventories every metric title, header label, and layout spelling observed across the real-file corpus and the current template, proposes canonical keys, and lists the decisions to ratify.
Ratified outcomes become: a new design-spec section (the naming contract), the compiled-in standard registry, the Phase 3 `metric_defs` seed, and the Phase 5 interactive template-builder palette.

Sources scanned 2026-07-27:

- `resources/templates/Standardized Test Template - December 2025.xlsx` (current authoritative template: 11 data sheets + Test SOP's).
- `tests/corpus/T58G 510 D9 Testing.xlsx` (2024, Project-variant header band, 12-col blocks).
- `tests/corpus/CPS2920 6-24 Lab Testing 6-26 copy.xlsx` (2025, standardized-old band, 12-col blocks, plus an 8-col User Test Simulation and a User Test tracking sheet).
- `tests/corpus/S26 4D, 4E, 4F Designs.xlsx` (Cart era, 13-col blocks with PV1-PV5).
- Code: `src/model/StandardSchema.cpp`, `src/model/SchemaInference.cpp`, `src/pipeline/TpmCalculator.cpp`.

Owner confirmed 2026-07-27 that these workbooks cover the historical template range.

## 1. Template eras observed

| Era | Block cols | Col 5 | Cols 10-12 | Header band | Example |
|---|---|---|---|---|---|
| E1 current (Dec 2025) | 12 | Puffing Regime | PD (mg/(W*s)) / Variation in TPM (%) / Oil Consumed (Cumulative, g) | Standard layout | all current data sheets except TC#1 |
| E1b current, Temp Cycling #1 | n/a (row 4 BLANK) | - | - | unique layout (see 3.4) | Temperature Cycling Test #1 |
| E2 standardized-old (2024-25) | 12 | Resistance | PD (mg/puff/W) / TPM Consistency / Rolling Average TPM | Standardized-old (no Heating Technology) | CPS2920 main sheets |
| E3 Project variant (2024) | 12 | Resistance | same as E2 | Project layout (Date/Tester/Project/Sample on row 1, Ri + Rf) | T58G |
| E4 Cart era | 13 | PV band instead of Draw Pressure | TPM only (no PD/Var/Oil cols) | Cart layout + design-spec labels | S26 Designs |
| E5 User Test Simulation (old) | 8 | - | TPM only | compressed band + instruction row | CPS2920 User Test Simulation |
| E6 tracking sheets | n/a | - | - | freeform ("User Test Tracking", "Distributor:") | CPS2920 User Test - Full Cycle |

## 2. Metric registry (data columns)

Observed titles are verbatim cell text; matching is by normalized form (lowercase, alphanumeric only).

| Canonical key | Observed titles (eras) | Type | Unit | Role |
|---|---|---|---|---|
| `puffs` | "puffs" (all) | number | count | measured |
| `before_weight` | "Before weight (g)" (E1); "Before weight/g" (E2/E3/E4); "Before Weight/g" (E5) | number | g | measured |
| `after_weight` | "After weight (g)" (E1); "After weight/g" (E2/E3/E4); "After Weight/g" (E5) | number | g | measured |
| `draw_pressure` | "Draw Pressure (kpa)" (E1/E2/E3/E5); absent in E4 | number | kPa | measured |
| `resistance` (per-row) | "Resistance" (E2/E3/E4) | number | ohm | measured |
| `puffing_regime` (per-row) | "Puffing Regime" (E1) | text | - | qualitative |
| `smell` | "Smell" (most); "Smell (1-4)" (E1 Negative Pressure); "Smell (0-4)" (E4) | text | scale | qualitative |
| `clog` | "Clog" (most); "Clog (Y/N)" (E1 Negative Pressure); "Clog?" (E4) | text | Y/N | qualitative |
| `notes` | "Notes" (all) | text | - | qualitative |
| `tpm` | "TPM (mg/puff)" (all) | number | mg/puff | derived |
| `tpm_power_density` | "TPM Power Density (mg/(W*s))" (E1); "TPM Power Density (mg/puff/W)" (E2/E3) | number | SEE 2.1 | derived |
| `variation_tpm` | "Variation in TPM (%)" (E1) | number | % | derived |
| `tpm_consistency` (NEW key proposed) | "TPM Consistency" (E2/E3) | number | fraction | derived |
| `oil_consumed` | "Oil Consumed (Cumulative, g)" (E1) | number | g (cell) vs mg (spec) - SEE 2.1 | derived |
| `rolling_avg_tpm` (NEW key proposed) | "Rolling Average TPM" (E2/E3) | number | mg/puff | derived |
| `pv1`..`pv5` (open) | "PV1".."PV5" (E4) - meaning TBD | number? | ? | measured? |
| `chronology` (open) | "Chronology" (E5) | text | - | qualitative |
| `failure` (open, key shortened) | "Failure? (if yes, add detailed notes)" (E5) | text | - | qualitative |

### 2.1 The columns 10-12 semantics problem (headline finding)

The positional parser has always mapped col 10 to power density, col 11 to variation, col 12 to oil consumed.
The actual cell formulas show three different vocabularies sharing those slots:

Column 10, power density:

- E2/E3 formula: `=I5/F$2` = TPM / Power. Unit mg/(puff*W). Matches the app's `TpmCalculator::powerDensity` recompute and the spec unit.
- E1 formula: `=I5/F$2/3` = TPM / Power / 3 s. Unit mg/(W*s). A genuinely different quantity (divides by the 3 s puff duration hard-coded, not read from the regime).

Column 11, variation:

- E2/E3 "TPM Consistency": `=STDEV.P(I5:I7)/AVERAGE(I5:I7)` = coefficient of variation over the 3-row session window, as a FRACTION.
- E1 "Variation in TPM (%)": `=100*STDEV.P(I6:I7)/AVERAGE(I6:I7)` = CV over a rolling 2-row window, in PERCENT.
- App recompute `TpmCalculator::variation`: `(tpm_i - tpm_0)/tpm_0 * 100` = percent deviation from the FIRST puff row. Matches neither template.

Column 12:

- E2/E3 "Rolling Average TPM": `=AVERAGE(I5:I7)` = 3-row rolling mean TPM. NOT oil consumed; the parser has been storing these values in `oilConsumed` for old files.
- E1 "Oil Consumed (Cumulative, g)": cumulative `TPM*puffs/1000`, in grams.

Consequences to ratify: old-era cols 11-12 should become their own metrics (`tpm_consistency`, `rolling_avg_tpm`) under name-first matching; the canonical definition of `variation_tpm` (template rolling CV% vs app deviation-from-first) must be picked; `tpm_power_density` needs either two era-keyed defs or one def with era-variant calculators; `oil_consumed` canonical unit (g vs the spec's mg) must be picked.

## 3. Header-field registry (per-sample header band)

### 3.1 Identity and admin

| Canonical key | Observed labels (layout) | Type | Notes |
|---|---|---|---|
| `test_name` | title text at r1c1, no label (E1/E2/E5); "Test:" (SOP sheet) | text | |
| `date` | "Date:" (all layouts) | text/date | |
| `tester` | "Tester:" (E1 r3c3, E1b r1c2, E3 r1c4, E2 r3c3, E5 r2c3) | text | |
| `sample_id` | "Sample ID:" (E1/E2/E5); "Cart #" (E4) | text | |
| `project_name` | "Project:" (E3) | text | joined with `sample_suffix` to form sample id |
| `sample_suffix` | "Sample:" (E3) | text | |
| `distributor` (open) | "Distributor:" (E6 tracking) | text | promote? |

### 3.2 Device and electrical

| Canonical key | Observed labels | Type/unit | Notes |
|---|---|---|---|
| `resistance` (initial) | "Resistance (Ω):" (E1); "Resistance (Ohms):" (E2); "Resistance:" (E1b/E5); "Ri (Ohms)" (E3/E4) | number, ohm | shares key with the per-row column - see D11 |
| `rf_ohms` | "Rf (Ohms)" (E3/E4) | number, ohm | final resistance; promote to standard? rename `final_resistance`? |
| `voltage` | "Voltage:" / "Voltage" (all) | number, V | E4 values carry unit text ("3.6V") |
| `power` | "Power:" / "Power" (all) | number, W | usually formula-derived; E4 values carry unit text ("6.9W") |
| `heating_technology` | "Heating Technology:" (E1); "Heater Technology:" (E1b) | text | two spellings in the SAME current template |

### 3.3 Media, oil, and test setup

| Canonical key | Observed labels | Type/unit | Notes |
|---|---|---|---|
| `media` | "Media:" / "Media" (all) | text | |
| `viscosity` | "Viscosity:" / "Viscosity" (all except E5) | number, cP | value chaos: "300kcp", "10kcp", 500000 |
| `initial_oil_mass` | "Initial Oil Mass:" (E1/E2/E3/E5) | number, g | |
| `fill_volume` (NEW) | "Fill Volume:" (E1b) | number, mL? | promote? |
| `number_of_samples` (NEW) | "Number of Samples" (E1b) | number | promote? |
| `puffing_regime` (header) | "Puffing Regime:" (E1/E2/E3); "Puff Regime" (E4) | text | |

### 3.4 Cart-era design-spec fields (E4 only)

| Proposed key | Observed label | Type/unit guess |
|---|---|---|
| `coil_material` | "Coil Material" | text |
| `thermal_conductivity` | "Thermal Conductivity" | number, W/(m*K)? |
| `column_inner_diameter` | "Column inner diameter" | number, mm? |
| `column_length` | "Column length" | number, mm? |
| `coil_shape` | "Coil shape" | text |
| `cotton_length` | "Cotton length (if applicable)" | number, mm? |

These are per-sample design descriptors and prime candidates for the standard header registry (and the template-builder palette).

Temperature Cycling Test #1 (E1b) unique layout for reference: Tester r1c2, Power r1c4, Heater Technology r1c6; Media r2c2, Voltage r2c4, Number of Samples r2c6; Fill Volume r3c2, Resistance r3c4; row 4 blank.
Because row 4 is blank, the current reader takes the positional standard path with a band layout that does not match any compiled variant, so this sheet's headers are likely misread today (unverified).

### 3.5 Status questions and in-band computed cells

Present in every era at fixed band positions, currently NOT parsed as header fields (the `burn_clog_leak` aggregate scans smell/clog/notes text instead):

- "Did this burn?" -> proposed `did_burn` (bool).
- "Did this clog?" -> proposed `did_clog` (bool).
- "Did this leak?" -> proposed `did_leak` (bool).

In-band computed labels (map to aggregates, not header fields): "Average TPM and Standard deviation", "Usage Efficiency", "Puff Per Day Calculator (Including the initial 50 puffs on day 1):" (E5).

## 4. Aggregates registry (current compiled set)

| Key | Calculator | Inputs |
|---|---|---|
| `average_tpm` | mean | tpm |
| `stddev_tpm` | stddev | tpm |
| `avg_power_density` | mean_over_power | tpm, header:power |
| `normalized_tpm` | mean_over_power | tpm, header:power |
| `total_puffs` | last | puffs |
| `total_oil_consumed` | last | oil_consumed |
| `efficiency_percent` | efficiency_v1 | total_oil_consumed, header:initial_oil_mass |
| `burn_clog_leak` | status_scan_v1 | smell, clog, notes |

## 5. Naming policy (proposed for ratification)

1. Canonical keys are snake_case, stable forever: never renamed, never reused; superseded keys are deprecated and kept as aliases.
2. Matching is name-first via normalized form: lowercase, strip all non-alphanumerics; "Before weight/g", "Before Weight (g)", "before weight g" all normalize equal.
3. Every metric and header field carries an open-ended alias list; any historical spelling in this document becomes a registered alias.
4. Unknown titles never break a load: they become open metrics/headers with a snake_case key derived from the title (current SchemaInference behavior, kept).
5. Collision policy: if two columns in one sheet normalize to the same alias, position within the block breaks the tie and a poka-yoke warning is surfaced (never a load failure).
6. Display names and units are presentation hints; all metrics are always included (owner directive: never exclude a metric).
7. Semantic changes (a new formula, a new unit) get a NEW key (Avro rule), even when the on-sheet title stays similar; the era-specific title becomes an alias of the correct key only.

## 6. Decisions needed (workshop items)

- D1 Variation semantics: `variation_tpm` canonical definition (template rolling CV% vs app deviation-from-first vs session CV). Which does the app compute and display going forward? Register old "TPM Consistency" as separate `tpm_consistency`?
- D2 Power density: one key with era-variant calculators, or two keys? Is the current template's /3 s intended to be actual puff duration from the regime? Canonical unit label to print.
- D3 Old col 12: confirm `rolling_avg_tpm` as its own metric so old files stop storing rolling averages in `oilConsumed`.
- D4 `oil_consumed` canonical unit: g (matches cells) vs mg (spec section 17 value). Recommendation: g.
- D5 PV1-PV5: meaning, canonical names, units.
- D6 Promote to standard headers: `final_resistance` (Rf), `fill_volume`, `number_of_samples`, `distributor`, and the Cart-era design-spec set (3.4) with types/units.
- D7 `heating_technology` vs "Heater Technology:" - alias merge (recommended) and fix the template at next revision.
- D8 Did this burn/clog/leak: promote to first-class boolean header fields?
- D9 Smell scale 1-4 vs 0-4: one metric with label aliases, or does the scale need to be recorded?
- D10 Units normalization on read: parse "6.9W", "3.6V", "300kcp" into number + unit; viscosity canonical unit (cP with k-multiplier parsing).
- D11 Resistance key namespaces: per-row column `resistance` vs initial-resistance header `resistance` vs `rf_ohms`. Proposal: header keys become `initial_resistance` + `final_resistance` (with Ri/Rf aliases); per-row column keeps `resistance`.
- D12 Temperature Cycling Test #1: is its unique layout intentional? Verify how the app reads it today; standardize at next template revision or register its layout.
- D13 Naming policy in section 5: ratify or amend.
- D14 Additional headers the owner wants to add (owner homework: review this inventory and extend).

## 7. Workshop log

- 2026-07-27: draft created from corpus + template scan; formulas for cols 9-12 verified across eras; awaiting owner review.
