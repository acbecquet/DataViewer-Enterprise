# TPM v3 Smoke-Fix Batch: Header-Driven Schema Inference

> **For agentic workers:** execute task-by-task with the same subagent cadence as the Phase 0+1 plan.
> The "Machine + repo rules" and "Build & test commands" sections of `docs/superpowers/plans/2026-07-23-tpm-v3-phase0-1-core-model-plan.md` apply verbatim.

**Goal:** Real historical workbooks whose block layout is NOT the standard 12-column shape parse correctly: right block geometry, known metrics mapped by name to the right fields, header label:value pairs read correctly, and unknown columns/labels PRESERVED as open data (owner rule: include all metrics, exclude none).

**Trigger (owner E2E smoke, 2026-07-24):**
- `S26 4D, 4E, 4F Designs.xlsx` (Feb 2024): 13-column Cart-era blocks (`puffs, Before weight/g, After weight/g, PV1..PV5, Resistance, Smell (0-4), Clog?, Notes, TPM (mg/puff)` repeating at offsets 0/13/26).
  Both parsers assume 12-wide blocks, so sample 2+ header reads land one column short (verified against the owner's screenshot cell-for-cell: `sample_id` at 12+5 hits "Thermal Conductivity", `media` at 12+1 hits "Cart #", `puffing_regime` at 12+7 hits "Viscosity").
- `CPS2920 ... User Test Simulation` sheet: 8-column blocks (`Chronology, puffs, Before Weight/g, After Weight/g, Draw Pressure (kpa), Failure? (...), Notes, TPM (mg/puff)`) with its own label layout (`Resistance:`/`Sample ID:`/`Initial Oil Mass:` on row 1, `Media:`/`Tester:`/`Voltage:`/`Power:` on row 2).
- Classification: legacy v2.10.0 mangles both files the same way (12-wide assumption is shared); this is PRE-EXISTING wrongness, not a Phase-1 regression - confirmed by the shadow suite run against the real files (see Verification).

**Architecture:** New `src/model/SchemaInference` builds a `TemplateSchema` from the sheet's own column-header row and header band when the standard shape does not fit.
`DataProcessor::processSheet` gains a fit check: standard-shaped sheets take the EXISTING path unchanged (byte-identity gate untouched); non-fitting sheets take the inference path, parsed by the existing `SchemaDrivenReader` and lowered by a new direct adapter (no 12-wide `SampleData` detour).
Unknown per-row columns ride in a new `DataRow::extra` map; unknown header labels ride in the existing `SampleResult::extra`.
Display of extras stays for Phase 4 (schema-driven UI); this batch makes the KNOWN fields correct and preserves the rest losslessly in memory, JSON recovery, and Excel round-trip safety.

---

## Task I: `SchemaInference` (pure, unit-tested)

**Files:** Create `src/model/SchemaInference.h/.cpp`; modify `tests/tst_v3model/*` (or new `tests/tst_v3inference` suite - implementer's choice, document it), both `.pro` registrations.

API:

```cpp
namespace DVE { namespace model {

class SchemaInference {
public:
    // Repeat distance of the first non-empty normalized header token in the
    // column-header row (row index columnHeaderRow-1). 0 = no repeat found
    // (single block spanning the row's non-empty width).
    static int inferBlockCols(const QVector<QVariant>& headerRow);

    // True when the sheet is standard-shaped: inferBlockCols is 12 (or 0 with
    // width <= 12) AND the first three header cells normalize-match the
    // standard puffs/before_weight/after_weight aliases at their standard
    // positions. Standard-shaped sheets MUST keep the existing byte-identical
    // path - this predicate is the gate.
    static bool standardFits(const QVector<QVector<QVariant>>& cells,
                             const TemplateSchema& standard);

    // Build a schema from the sheet itself (columnHeaderRow=4, dataStartRow=5
    // geometry assumed - every known historical layout shares it):
    //  - columns: each block-1 header cell matched against the KNOWN metric
    //    knowledge base (standardV1 defs by key/displayName/aliases plus the
    //    alias additions below); matched -> copy of that MetricDef (role/type/
    //    unit/calculator preserved); unmatched -> NEW MetricDef (key =
    //    snake_case of the normalized header, type Number if the first
    //    non-empty data cell in that column parses numeric else Text, role
    //    Measured/Qualitative accordingly).
    //  - headerFields: scan block-1 cells in rows 1..3; a cell is a LABEL if
    //    it matches the known label-alias table OR its trimmed text ends with
    //    ':' or '?'; the VALUE is the cell to its right; known labels map to
    //    canonical keys, unknown labels become new keys (snake_case).
    //  - aggregates: copy the standard set only when every input metric of an
    //    aggregate exists in the inferred columns.
    static TemplateSchema inferSchema(const QVector<QVector<QVariant>>& cells,
                                      const QString& sheetName);
};

}}
```

Alias additions (extend the knowledge base INSIDE SchemaInference, do not touch standardV1's shipped aliases): metric headers `"Smell (0-4)"->smell`, `"Clog?"->clog`, `"Before Weight/g"/"Before weight/g"->before_weight`, `"After Weight/g"/"After weight/g"->after_weight`; label aliases `"Cart #"->sample_id`, `"Ri (Ohms)"->resistance`, `"Rf (Ohms)"->rf_ohms (new open key)`, `"Power"->power`, `"Puff Regime"/"Puffing Regime:"->puffing_regime`, `"Viscosity"/"Viscosity:"->viscosity`, `"Voltage"/"Voltage:"->voltage`, `"Media"/"Media:"->media`, `"Date:"->date`, `"Sample ID:"->sample_id`, `"Resistance:"/"Resistance (Ohms):"->resistance`, `"Tester:"->tester`, `"Initial Oil Mass:"->initial_oil_mass`, `"Test Name"->test_name`.
When a header VALUE for power exists (UserSim row 2 carries a computed Power), the inferred header keeps it as the value; the lowering (Task III) prefers a present numeric power value over deriving.

Tests (synthetic grids, full TDD): inferBlockCols on 13-wide S26-shaped and 8-wide UserSim-shaped header rows; standardFits true for `makeStandardGrid` and false for both shapes; inferSchema on an S26-shaped grid asserts: 13 columns, `puffs/before_weight/after_weight` matched with standard roles, `pv1..pv5` created as new Number/Measured metrics, `tpm` matched as Derived with `tpm_v1`, header fields include `sample_id` (from "Cart #") and new open key for an exotic label (e.g. `coil_material`); UserSim-shaped grid asserts 8 columns incl. `chronology` (Text) and `failure_if_yes_...`-style new key, and `sample_id`/`initial_oil_mass` labels found on row 1.

Commit: `feat(v3): SchemaInference - block-width + schema inference from header row (smoke-fix batch)`

## Task II: `DataRow::extra` + JSON round-trip

**Files:** `src/pipeline/ReportData.h` (add `QMap<QString,QVariant> extra;` to DataRow with a comment: open per-row metrics from inferred schemas, preserved for recovery + Phase 4 display; NOT yet persisted to Postgres - Phase 3), `src/pipeline/ReportDataJson.cpp` (both directions - honor the header's contract), `tests/tst_reportdatajson/tst_reportdatajson.cpp` (extend the existing round-trip test with a populated extra map: number, text, and a QByteArray value).

Commit: `feat(v3): DataRow::extra open per-row values + lossless JSON recovery (smoke-fix batch)`

## Task III: Engagement + lowering + fixtures + gates

**Files:** `src/pipeline/DataProcessor.h/.cpp`, `src/model/LegacyAdapter.h/.cpp` (new `lowerInferredSheet`), `tests/generate_fixtures.py` + regenerated new fixtures, `tests/tst_v3shadow/tst_v3shadow.cpp`, the inference test suite from Task I.

1. `DataProcessor::processSheet`: after the SOP short-circuit and grid load, evaluate `SchemaInference::standardFits`.
   Fits -> EXISTING code path, character-for-character untouched.
   Does not fit -> inference path: `schema = SchemaInference::inferSchema(...)`; `mSheet = SchemaDrivenReader::parseSheet(cells, sheetName, schema, /*perRowRegime=*/false, ColumnResolution::NameFirst)`; `result = LegacyAdapter::lowerInferredSheet(mSheet, sheetName, templateVersion)`.
   Track `m_lastUsedInference` (per-file flag, public getter `bool lastFileUsedInference() const`) set by processFile when ANY sheet inferred - the shadow test consumes it.
2. `LegacyAdapter::lowerInferredSheet(const model::Sheet&, sheetName, templateVersion) -> SheetResult`: map KNOWN metric keys to DataRow fixed fields (puffs/before_weight/after_weight/draw_pressure/resistance/puffing_regime/smell/clog/notes; derived fields left for recompute); every OTHER series value at row r goes into `DataRow.extra[key]`; header fields with canonical keys map to the same SampleResult members the standard branch fills (sample_id -> both sampleID and sampleName exactly as buildSampleResult does - READ buildSampleResult first and mirror its assignment/fallback rules); power: use a present numeric header value, else derive via the HeatingTech formula when voltage+resistance exist; every unknown header key/value goes into `SampleResult.extra`; then run `GenericSheetProcessor::calculateMetrics` per sample + `computeSheetAggregates` (recomputes the tpm chain from the correctly-mapped fields); set `sheet.columnHeaders` to the actual header texts of block 1.
3. Fixtures: `gen_pv13()` (S26-shaped: 13-col, 2 blocks, Cart-era labels incl. one exotic "Coil Material" label, PV1..PV5 with a few numeric values, realistic puffs/weights) and `gen_usersim8()` (8-col, 2 blocks, UserSim labels incl. Power value) in `tests/generate_fixtures.py`; commit the two new `.xlsx`.
4. Inference E2E test (in the Task I suite): run the real `DVE::DataProcessor::processFile` on both new fixtures and assert VALUES: sample 2's sampleID/resistance/media are the block-2 values (the exact regression class the owner hit), tpm[0] recomputed correctly from the fixture's weights, `rows[0].extra` carries pv1..pv5 / chronology, `samples[0].extra` carries the exotic label, and `lastFileUsedInference()` is true.
5. `tst_v3shadow::productionMatchesLegacyParser`: after running production, `if (pNew.lastFileUsedInference()) QSKIP("inference path - correctness gated by the inference E2E tests, legacy is known-wrong here")`.
   The determinism row keeps covering ALL files including inference-path ones.
   The 9 standard fixtures MUST still hit the identity assertion (0 skipped among them) - assert count or verify in the run log.
6. Gates: inference suite green; tst_v3shadow with `DVE_TEST_CORPUS_DIR` pointing at the local real-file corpus: T58G rows byte-identical as before, S26 + CPS2920 determinism green with identity-skip only for those two; tst_reportdatajson green; FULL suite `tests/run-tests.ps1` all green; then rebuild the installer as v2.10.2 (rebuild-dataviewer skill) for the owner's re-smoke.

Commit: `feat(v3): inference path wired - non-standard block layouts parse correctly (smoke-fix batch)`
Final: tracker + spec Phase log addendum + lessons entry.

## Verification summary

| Gate | Expectation |
|---|---|
| Inference unit tests | block widths 13/8 inferred; known/unknown split correct |
| Inference E2E on new fixtures | sample-2 fields correct; extras preserved; tpm recomputed |
| tst_reportdatajson | extra map round-trips losslessly |
| tst_v3shadow (9 std fixtures) | identity rows all PASS, 0 skipped |
| tst_v3shadow (real corpus) | T58G identity PASS; S26/CPS identity SKIP-by-inference; determinism PASS |
| Full suite | all green |
| Owner re-smoke (v2.10.2) | S26 sample 2+ properties correct; UserSim columns correct |

## Explicit non-goals (this batch)

- Displaying `extra` metrics/headers in the UI (Phase 4).
- Persisting extras to Postgres (Phase 3).
- Write-back into non-standard layouts (Phase 2's CellAddressMap; qualitative-cell edits on inference-path files should be REJECTED gracefully for now - verify the current write path cannot fire with wrong coordinates on inferred sheets; if it can, disable those edits for inferred sheets and log it).
- The `User Test - Full Cycle` tracking sheet (no repeating block header; falls through to today's behavior).
