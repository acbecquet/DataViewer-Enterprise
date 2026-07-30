# TPM v3 Vocabulary Registry - Workshop Draft

Date: 2026-07-27.
Status: RATIFIED (2026-07-27 evening) - all decisions D1-D14 closed and Q1-Q4 resolved (section 9).
This document is the binding naming contract for Phase 2 onward; the code fold-in (registry compilation into StandardSchema/SchemaInference) is tracked in the Phase 2 plan.
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

## 6. Decisions (owner rulings 2026-07-27; OPEN items noted)

- D1 DECIDED: the template definition is always the default for derived metrics (for now); `variation_tpm` = the current template's rolling CV in percent.
  The app's deviation-from-first recompute migrates to the template definition when the calculator registry lands.
  Old "TPM Consistency" becomes its own `tpm_consistency`.
- D2 DECIDED (final math in section 9.1): `tpm_puff_density` = TPM/P with unit mg/((n-second) puff * W); `tpm_power_density` = tpm_puff_density * (1 puff)/(n s) = TPM/(P*n) with unit mg/(W*s); n = puff length from the `puff_time` metric.
  Old-era TPM/P columns lower into `tpm_puff_density`; current-era mg/(W*s) columns lower into `tpm_power_density`.
  Both are acknowledged approximations (no transient power for different puff lengths; TCR changes instantaneous power over time).
  Era collisions on identical names resolve via the existing template-version indicators in the software - backwards compatibility only; future metric names are unique.
- D3 DECIDED: `rolling_avg_tpm` is its own metric.
  Standing rule ratified: when in doubt, own metric - "we can never have too many, but we can have too little".
- D4 DECIDED: `oil_consumed` unit = grams.
- D5 DECIDED: PV1-PV5 = per-puff draw pressures (PV = pressure value; PV1 is the draw pressure for puff 1 of the session; sessions were briefly standardized at 5 puffs).
  They lower into the new `draw_pressure_per_puff` list metric (section 8.2).
- D6 DECIDED: promote `resistance_final`, `fill_volume`, `number_of_samples`, `distributor`, and ALL Cart-era design-spec fields (3.4) as optional standard headers.
  Requirement: the customizable template must be able to reconstruct ANY previous template exactly, Cart era included.
- D7 DECIDED: "Heating Technology" is canonical; "Heater Technology" is an accepted alias; future templates default to Heating Technology.
- D8 DECIDED: `did_burn` / `did_clog` / `did_leak` become first-class header fields.
- D9 DECIDED: Smell 1-4 and 0-4 are the SAME scale (1-4 leaves blank for no event; 0-4 writes 0); label aliases only.
- D10 DECIDED: normalize unit-suffixed values on read ("6.9W", "3.6V", "300kcp"); viscosity canonical cP with k-multiplier parsing.
- D11 DECIDED: THREE independent resistance headers - `resistance` ("Resistance"), `resistance_initial` ("Resistance (Initial)", alias "Ri (Ohms)"), `resistance_final` ("Resistance (Final)", alias "Rf (Ohms)") - PLUS the per-row `resistance` metric stays. Max flexibility.
- D12 DECIDED: Temperature Cycling Test #1 is a tracking doc (headers only, no data collection); leave as-is.
  Non-standard files like it become obsolete as the customization scheme produces new files.
- D13 DECIDED: naming policy (section 5) ratified as written, all seven rules ("7 rules look good", 2026-07-27 evening).
- D14 DECIDED: additional headers/metrics accepted - registered in section 8.

## 7. Workshop log

- 2026-07-27: draft created from corpus + template scan; formulas for cols 9-12 verified across eras; awaiting owner review.
- 2026-07-27 (later): owner rulings received and captured (D1-D12, D14 + section 8); open threads Q1-Q4 + D13 in section 9.
- 2026-07-27 (evening): Q1-Q4 answered (power-density pair math, regime = 4 split metrics, identity fallback) and D13 ratified; document promoted to RATIFIED - the binding naming contract for Phase 2+.
- 2026-07-28: Phase 2a EXECUTED - registry compiled into `src/model/MetricRegistry` (single source; StandardSchema + SchemaInference consume it), RegimeParser live, PV1-PV5 assemble into draw_pressure_per_puff, inference labels canonicalized (design specs, status Q&A, resistance trio). Gates: full suite 59/0/0, corpus shadow 25/0/5 (byte identity holds), -Werror app build clean. Commits 577260a..80af58e.
- 2026-07-28 (later): Phase 2b EXECUTED - write provenance recorded at parse (both paths), CellAddress-derived write-back replaces the 12-wide math, editing re-enabled on inferred layouts, Cart/Project header-write corruption fixed. Round-trip harness green over fixtures + corpus; full suite 60/0/0; corpus shadow 25/0/5. Commits 771f6a4..25d9718. v2.10.3 combined smoke build cut.
- 2026-07-30: Phase 2c EXECUTED - `_dve_schema` manifest (grid grammar, registry inheritance, poka-yoke warnings), SchemaResolver ladder (manifest -> standard -> inference), NameFirst live on manifest sheets with slot-ordered write provenance, demo workbook proves shuffled columns + custom coil_temp parse correctly. Hardening: collision poka-yoke (rule 5), phantom trailing samples, per-sheet JSON guard. Gates: full suite 60/0/0, corpus shadow 26/0/6 (byte identity holds), corpus round-trip 38/0/0. Commits fa7efb2..f038cd4. v2.10.4 smoke build cut.

## 8. Owner rulings - new registry entries (2026-07-27)

### 8.1 Sample identity

The identity seed for each sample = `test_name` + `date` + `sample_id`.
Cart-era backwards compatibility: the "Cart #" value becomes the `sample_id` string for those files.
(Project-era files already assemble sample id from Project + Sample.)
Filename fallback (owner, 2026-07-29 post-v2.10.3 smoke): for the era where the FILENAME was the test name (e.g. "T58G 510 D9 Testing.xlsx"), any place that needs a test name falls back to the workbook's base filename when the sheet does not provide one ("always check the filename if the sheet name doesn't fit our map").
This supersedes the plain date+sample_id degradation from Q4 where a filename is available.

### 8.2 New per-row metrics

| Key | Display | Type | Unit | Notes |
|---|---|---|---|---|
| `image` | Image | image (any format) | - | per-row payload beyond text; exercises the spec section 18 open-ended index-value design |
| `voltage` (per-row) | Voltage | number OR text | V / curve name | e.g. 3.6 or "Curve 9"; groupable the way puffing regime is grouped today; long-term goal is grouping by ANY metric |
| `puff_volume` | Puff Volume | number | mL | regime part 1 ("60mL") |
| `puff_time` | Puff Time | number | s | regime part 2, always the middle value ("3s") |
| `puff_rest_time` | Puff Rest Time | number | s | regime part 3 ("30s") |
| `session_rest_time` | Session Rest Time | number | s | regime part 4 when present ("60mL/3s/30s/5minute" -> 300 s); DEFAULTS to 0 when the regime has only 3 parts |
| `draw_pressure_per_puff` | Draw Pressure (per puff) | list of numbers | kPa | data-logger series per session row; list length = puffs in that session; historical PV1-PV5 columns lower into this |

The puffing-regime string split must stay backwards compatible: existing 3-part regime labels parse with session_rest_time = 0.

### 8.3 Ground truth and product direction (owner, verbatim intent)

The tester's actual measurements per session are: before weight, after weight, draw pressure (final puff when manual, or logger-collected and integrated later), and smell/clog/notes only when something notable happens.
Everything else is derived processing that exists to visualize an accurate picture of device performance.
The end goal is a universal map of device details connected to the real data (lab + sensory) to deeply understand performance and how to improve it.
UI requirement (feeds Phases 4-5): a built-in metric and header addition system; metrics support ALL data types; headers support a constrained set (string, boolean, and the common types in this registry).

### 8.4 Type-system implications for the model (Phase 2 intake)

ValueType needs three additions: mixed scalar (number-or-text, for voltage curves), image payload, and numeric list (draw_pressure_per_puff).
All three are consistent with the typed-envelope JSON already shipped for `DataRow::extra`.

## 9. Resolutions (owner answers, 2026-07-27 evening)

### 9.1 The power-density pair (final math)

`tpm_puff_density` = TPM [mg/puff] * 1/P [1/W], where P is the power in W calculated from resistance + voltage.
Its unit title is mg/((n-second) puff * W), where n is the puff length taken from the `puff_time` metric.
(The compiled registry abbreviates this unit string as "mg/((n s) puff*W)"; same meaning.)
`tpm_power_density` = tpm_puff_density * (1 puff)/(n s) = TPM/(P*n), unit mg/(W*s) - dimensionally exact because the (n-second) puff cancels, which closes the former Q2 rigor flag.
Key assignment for historical columns: old-era "TPM Power Density (mg/puff/W)" (= TPM/P) lowers into `tpm_puff_density`; current-era "TPM Power Density (mg/(W*s))" (= TPM/P/3) lowers into `tpm_power_density`, where the hard-coded 3 equals n for the standard 60mL/3s/30s regime; the canonical calculator reads n from `puff_time`.
Both remain approximations: no transient power for different puff lengths, and TCR changes instantaneous power over time.

### 9.2 Regime representation (Q3)

Canonical representation in the header is the FOUR split metrics (`puff_volume`, `puff_time`, `puff_rest_time`, `session_rest_time`) - no need to keep a combined string as a first-class metric.
The composite regime string ("60mL/3s/30s[/5minute]") is a legacy source encoding parsed on read; the owner composes representations as needed in the template builder.

### 9.3 Identity fallback (Q4)

Accepted: Cart-era files (no test_name) seed identity from date + sample_id, with the sheet for context.

### 9.4 Units audit closure

Confirmed correct: tpm mg/puff; variation_tpm %; rolling_avg_tpm mg/puff; oil_consumed g (cumulative); draw_pressure kPa; puff_volume mL; puff/rest times s; voltage V; resistance ohm; initial_oil_mass g; viscosity cP (kcp = 1000 cP).
`tpm_consistency` registered as a dimensionless fraction (old cells carry STDEV/AVG without *100).
Power-density units per 9.1.
`fill_volume` defaults to mL until stated otherwise.

### 9.5 Naming policy (D13)

Ratified as written - all seven rules in section 5 stand.
