# TPM Template v3 - Free-Standing Metric Model - Design Spec

Date: 2026-07-16.
Status: approved design, pre-implementation (research branch `worktree-tpm-template-v3-research`).
Target release: v3.0.0, after the v2.10 sprint ships.
Companion doc: `docs/superpowers/specs/template-cell-map.md` (authoritative physical layout of the current template).

## 1. Problem statement

The standardized TPM template's physical layout is hardcoded into the software at roughly nineteen coupling sites across five layers.
Those layers are: the `Cols::` index constants plus the fixed struct fields in `src/pipeline/ReportData.h`, the positional arithmetic in `src/ExcelReader.cpp` (the `COLS_PER_SAMPLE=12` literal appears three times, the header row and data start row are fixed), the column-to-field mapping in `src/pipeline/SheetProcessors.cpp`, the write-back math in `src/MainWindow.cpp` (`excelCol = sampleIndex*12 + col + 1` and a hardcoded prop-table row switch), and the Postgres fixed-column schema mirrored again in hand-written SQL bind lists and in `ReportDataJson` keys.
Adding, removing, or reordering a single column requires touching all of them in lockstep.
The owner wants a moldable, generalizable template model where metrics can be reordered and added at will, and custom tests can be authored, without meticulous software updates.

## 2. Owner decisions (locked 2026-07-16)

1. Compatibility scope: round-trip everything.
   Every existing `.xlsx` keeps loading, the standard template reconstructs cell-for-cell from the new model including write-back, and existing DB data survives.
2. Schema home: in the workbook.
   A hidden manifest sheet makes each file self-describing; a compiled-in standard schema is the fallback for legacy files.
3. Delivery: phased, core first, targeting v3.0 after v2.10, tracking v2.10's incoming changes.
4. Mode scope: TPM now, sensory-ready.
   Core types stay mode-agnostic so sensory can migrate later without a second redesign.
5. Depth: full long-format.
   Both the in-memory model and Postgres go generic (a single `measurements` long table); old fixed columns become compatibility views; existing data is migrated.
   The save-path risk is accepted and managed through hard verification gates.

## 3. Conceptual model

The canonical shape is the owner's stated model.

```
Sample
├── headers: key -> value          (applies to all data of the sample equally)
└── data:    metric key -> [values] (one free-standing series per metric)
```

The Excel sheet becomes a presentation projection of this model, not its definition.
A `TemplateSchema` carries the projection rules (which metric goes in which column, which header field sits in which cell), so reconstruction matches the physical template exactly while the model itself stays layout-free.

## 4. In-memory types (new `src/model/`, namespace `DVE`)

- `MetricDef`: stable snake_case `key`, `displayName`, value type (`Number` | `Text`), `unit`, `role` (`Measured` | `Qualitative` | `Derived` | `Identity`), and for derived metrics a `calculator` name plus `inputs` (metric keys or `header:<key>` refs).
  Presentation hints ride along: `plottable`, `editable`, `precision`.
- `HeaderFieldDef`: `key`, `displayName`, value type, `unit`, block-relative cell ref (row, col), optional `calculator`/`inputs` for derived header values (e.g. power).
- `AggregateDef`: sample-level derived value with `key`, `calculator`, `inputs`.
- `TemplateSchema`: `schemaId`, `version`, block geometry (`headerRows`, `columnHeaderRow`, `dataStartRow`, `blockCols`), ordered `headerFields`, ordered `columns`, `aggregates`, and per-sheet info (initial puff seed, raw-table flag).
  Column order in the schema IS the physical column order.
- `MetricSeries`: metric key plus `QVector<QVariant>` values.
- `Sample`: `headers` map, ordered `MetricSeries` list, DB `id`/`version`.
- `Sheet` / `File` wrappers carry the schema reference, samples, aggregates, and images.

The legacy `FileResult`/`SheetResult`/`SampleResult`/`DataRow` structs survive Phases 1-3 as an adapter projection of the new model, and are deleted in Phase 4 once every consumer reads the schema-driven model directly.

## 5. Manifest format (`_dve_schema` worksheet)

The manifest is a very-hidden worksheet named `_dve_schema`, following the lockbox `_Template_NN` pattern.
Very-hidden sheets survive SaveAs to macro-free `.xlsx`, so distributed copies stay self-describing while remaining invisible to receivers, consistent with the sidecar product vision.
The format is a plain grid (not JSON-in-a-cell): Excel-native, hand-editable by power users, diffable, and parseable through the existing `ExcelReader` grid path with zero new plumbing.

Illustrative content (section marker in column A, fields across the row):

```
[schema]    | id                 | version
            | standard           | 2
[geometry]  | header_rows        | 3
            | column_header_row  | 4
            | data_start_row     | 5
            | block_cols         | 12
[header]    | key                | display            | row | col | type   | unit
            | test_name          | Test Name          | 1   | 1   | text   |
            | date               | Date               | 1   | 4   | text   |
            | ...                |
[column]    | key                | display            | type   | unit    | role        | calculator      | inputs
            | puffs              | puffs              | number | count   | measured    |                 |
            | before_weight      | Before Weight (g)  | number | g       | measured    |                 |
            | tpm                | TPM                | number | mg/puff | derived     | tpm_v1          | puffs,before_weight,after_weight
            | ...                |
[aggregate] | key                | calculator         | inputs
            | average_tpm        | mean               | tpm
            | ...                |
```

The sidecar `.xlsm` emits and maintains the manifest in every workbook it generates.
Validation on load is a poka-yoke: malformed or partial manifests produce warnings and fall back gracefully; they never block loading.

## 6. Builtin `standard-v1` schema

The compiled-in `standard-v1` schema reproduces today's template exactly and is the fallback for every legacy file.
It exists in two variants resolved by the existing column-4 header sniff (`RegimeUtils::isRegimeHeader`): the old-template variant with a `resistance` data column, and the new-template variant with a per-row `puffing_regime` column.

Data columns (block-relative position = today's `Cols::` order):

| Pos | Key | Type | Unit | Role | Calculator (inputs) |
|---|---|---|---|---|---|
| 1 | puffs | number | count | measured | |
| 2 | before_weight | number | g | measured | |
| 3 | after_weight | number | g | measured | |
| 4 | draw_pressure | number | kPa | measured | |
| 5 | resistance OR puffing_regime | number/text | ohm / - | measured/qualitative | |
| 6 | smell | text | | qualitative | |
| 7 | clog | text | | qualitative | |
| 8 | notes | text | | qualitative | |
| 9 | tpm | number | mg/puff | derived | tpm_v1(puffs, before_weight, after_weight) |
| 10 | tpm_power_density | number | mg/(puff*W) | derived | power_density_v1(tpm, header:power) |
| 11 | variation_tpm | number | % | derived | variation_v1(tpm) |
| 12 | oil_consumed | number | mg | derived | oil_consumed_v1(before_weight, after_weight) |

Header fields (block-relative 1-based cell, per `template-cell-map.md`):

| Key | Cell | Notes |
|---|---|---|
| test_name | R1 C1 | |
| date | R1 C4 | |
| sample_id | R1 C6 | |
| heating_technology | R1 C8 | |
| media | R2 C2 | |
| resistance | R2 C4 | |
| power | R2 C6 | derived: `power_v1(header:voltage, header:resistance, header:heating_technology)` = V^2 / (R + tech offset) |
| puffing_regime | R2 C8 | |
| viscosity | R3 C2 | |
| tester | R3 C4 | |
| voltage | R3 C6 | |
| initial_oil_mass | R3 C8 | |

Aggregates: `average_tpm = mean(tpm)`, `stddev_tpm = stddev(tpm)`, `avg_power_density = mean(tpm)/power`, `normalized_tpm = mean(tpm)/power`, `total_puffs = last(puffs)`, `total_oil_consumed = last(oil_consumed)`, `efficiency_percent = efficiency_v1(total_oil_consumed, header:initial_oil_mass)`, and `burn/clog/leak status = status_scan_v1(qualitative columns)` preserving today's keyword scan.
Raw-table sheets (`Test SOP's`, the Temperature Cycling checklist) keep the existing `isRawTable` handling, expressed as a schema flag rather than a sheet-name special case.

## 7. Parsing and write-back

- `SchemaResolver`: if `_dve_schema` is present, parse it; otherwise run the existing legacy detection and select the builtin `standard-v1` variant.
  Legacy files never need manifests; their parse output is bit-identical to today's.
- `SchemaDrivenReader`: sample blocks are located by schema `blockCols`; header fields are read from schema cell refs; data columns are matched **by header name first, position as fallback**.
  Name-first matching is the Avro rule (match by name, never by position), which makes column reordering safe even before a manifest exists.
- `CellAddressMap`: built during parse, it records the physical source cell for every (sample, metric, rowIdx) and every (sample, headerKey).
  Write-back looks up the recorded cell, which retires the `sampleIndex*12+col+1` arithmetic and the hardcoded prop-table row switch in `MainWindow`.
- Formula-cell protection (derived columns, puff/before-weight chains) derives from `role == derived` and schema knowledge instead of hardcoded column indices.
- The debounced write queue, openpyxl subprocess, and atomic tmp+replace mechanics stay as they are; only the coordinate computation changes.

## 8. Derived metrics

- `CalculatorRegistry` maps a name to a C++ function over input series and scalars.
- The current chain ships as registered calculators: `interval_v1`, `tpm_v1`, `power_density_v1`, `variation_v1`, `oil_consumed_v1`, `power_v1`, plus aggregates `mean`, `stddev`, `sum`, `last`, `mean_over_power`, `efficiency_v1`, `status_scan_v1`.
  Sentinel behavior (tpm > 100 -> 0, oil cap, interval reuse) moves inside the calculators unchanged.
- The manifest declares which derived columns exist and which calculators feed them; a small DAG resolves evaluation order and rejects cycles with a warning.
- There is no formula DSL in v3.0 (YAGNI).
  Custom tests compose existing calculators; genuinely new math is one registered C++ function instead of nineteen coupling sites.
- A future `expr:` calculator type can introduce a DSL without changing the manifest format.

## 9. Database (full long-format)

New v3 schema replacing the fixed metric columns:

- `metric_defs(id, key UNIQUE, display_name, value_type, unit, role, calculator, inputs JSONB, created_at, ...)`: the metric registry, seeded from `standard-v1` and extended as manifests introduce new keys.
- `template_schemas(id, schema_key, version, definition JSONB, ...)`: a registry mirror of manifests the system has seen; it drives UI and report defaults, while the workbook manifest stays authoritative on read.
- `samples`: reduced to identity plus audit (`sample_name`, `date`, `id`, `version`, `updated_at/by`); device-parameter headers move out.
- `sample_headers(sample_id, field_key, value_num, value_text, updated_at, updated_by, version)` with `UNIQUE(sample_id, field_key)`.
- `measurements(id, sample_id, metric_id FK, sort_order, value_num, value_text, updated_at, updated_by, version)` with `UNIQUE(sample_id, metric_id, sort_order)`.
- Derived values stay stored, exactly as today, because the offline snapshot and reports need them without recomputation.
- Compatibility views (`data_rows_v` and a wide `samples_v`) pivot the long tables back to the old column shape, preserving read access for any existing SQL or reporting during the transition.
- A one-shot migration transforms existing `data_rows`/`samples` columns into `measurements`/`sample_headers`.
  It is rehearsed on a prod-dump copy in the local test container first, and the production run is gated behind a `pg_dump` backup plus snapshot as the abort path.
- The SQLite offline snapshot and `ReportDataJson` adopt the same long shapes (metric key -> value arrays).

## 10. LiveSync / multi-user

- Per-cell commits target `measurements` (or `sample_headers`) rows: one uniform UPDATE path with per-measurement optimistic versioning.
- `liveColumnForDataCol` and the per-column NOTIFY mapping are deleted.
- Conflict granularity improves: today one `data_rows` row is one version counter, so edits to different metrics of the same puff row collide; per-measurement versions eliminate that class of false conflict (the DV-25 family).
- The three conflict dialogs re-point at measurement/header granularity; presence and the NOTIFY plumbing stay architecturally unchanged.

## 11. UI and reports

- The story panel / data table generates its columns from the schema (order plus editability from `MetricDef.role`), replacing the `Cols::` switches.
- The prop table generates its rows from `headerFields`, replacing the hardcoded row switch.
- Plots select plottable metrics from schema hints; the standard schema yields today's TPM trend unchanged.
  Color pinning stays keyed by sample name, compatible with DV-26.
- `ReportGenerator` reads series by key; the standard report renders identically for the standard schema, and custom metrics get generic table/plot sections.
- `DataCleanupDialog` exclusions key on `sort_order`; semantics unchanged.

## 12. Backward-compatibility matrix and evolution contract

| Input | Resolution | Result |
|---|---|---|
| Legacy `.xlsx`, no manifest | legacy detection -> builtin `standard-v1` | parses and round-trips exactly as today |
| v3 workbook, standard manifest | manifest | identical output, now reorder/add-safe |
| v3 workbook, custom manifest | manifest | fully supported end-to-end |
| Unknown metric key in DB | `metric_defs` row exists | loads generically; UI/report render from registry metadata |
| Old external SQL consumers | compatibility views | read the wide projection |

Schema evolution contract (adopted from Avro/Protobuf discipline):

1. Columns and header fields match by key/header name, never by position.
2. A key is never renamed or reused; evolution adds a new key and deprecates the old one.
3. New metrics are strictly additive; readers ignore unknown keys, and missing series read as empty with defined defaults.
4. Geometry changes travel only through the manifest.
5. The schema `version` bumps on any layout change; the `id` names the template family.

## 13. Modern data-science grounding

1. Tidy data (Wickham): store long, present wide; `measurements` is tidy data and the Excel block is the wide projection.
2. Apache Arrow / Parquet: named typed column vectors plus a schema object is exactly `MetricSeries` + `TemplateSchema`, and opens later zero-friction export to Parquet/pandas for analytics.
3. Frictionless Data Table Schema / CSVW: a machine-readable schema traveling with the data validates the manifest-sheet approach; its field descriptors map ~1:1 onto `MetricDef`.
4. Avro/Protobuf evolution rules: the contract in section 12.
5. Semantic layers / metric stores (dbt MetricFlow, LookML): derived metrics as named declarations over base measures, evaluated by an engine, instead of math buried in code and cells.
6. Pandera / Great Expectations: declarative validation-on-load surfaced as warnings, i.e. poka-yokes, never gates.
7. Units as first-class data (pint / UCUM): unit strings live in `MetricDef`, so labels and conversions are data, not code.
8. EAV done right: typed value columns, a registry FK, and uniqueness constraints, with views to project wide; the classic EAV pitfalls are mitigated by the registry and per-type columns.

## 14. Phases and gates (v3.0 track)

- Phase 0 - Spec + golden corpus (this research branch, now).
  Commit this spec; assemble a golden-file corpus (historical workbooks via the DB Data workflow plus the current templates); build the round-trip harness skeleton.
- Phase 1 - Core model + schema-driven read path, zero behavior change.
  New `src/model/` types, `SchemaResolver` with builtin `standard-v1`, `SchemaDrivenReader`, adapter to legacy `FileResult`.
  Gate: shadow-parse harness diffs old vs new parser output across the corpus; must be empty.
- Phase 2 - Write-back + manifest.
  `CellAddressMap` write-back, `_dve_schema` reader, sidecar manifest emission, poka-yoke validation.
  Gate: round-trip harness (parse -> reconstruct -> cell diff) empty over the corpus.
- Phase 3 - DB long-format (highest risk).
  New schema, migration, compat views, generic `DatabaseOps`/`DatabaseManager` I/O, LiveSync and conflict dialogs at measurement granularity, snapshot + JSON generic.
  Gates: migration rehearsal on a prod-dump copy with row counts and per-metric value checksums; DB dual-run diff; full E2E two-client tests.
- Phase 4 - Schema-driven UI + reports.
  Story panel, prop table, plots, and report sections generated from the schema; delete `Cols::`/`ColIdx` and the legacy structs.
- Phase 5 - Custom-test enablement.
  Manifest authoring via the sidecar, a template authoring guide, docs, and the v3.0.0 release train (internal patch builds, then the v3.0.0 deploy).

## 15. v2.10 coexistence rules

- DV-25/28 touch `MainWindow.cpp` autosave, `DatabaseManager` merge logic, and `SensoryPanel`; DV-26 touches `PlotWidget`/`PlotEngine`/`ReportGenerator`/`RadarChartWidget`; DV-27 touches `PlotEngine`/`SensoryPanel`.
- Phase 0 is docs and harness only: no collision.
- No v3 product code merges before v2.10 ships, even though Phase 1's new files barely overlap.
- Phases 3-4 overlap heavily with DV-25/28 territory; hard rule: rebase onto main after the v2.10 merge and re-audit the autosave/merge changes before starting Phase 3.

## 16. Verification (closed-loop harnesses)

- Round-trip harness: parse -> reconstruct -> cell-level diff over the golden corpus; must be empty for the standard schema.
- Shadow-parse harness: old parser vs new parser JSON diff across the corpus.
- DB dual-run: the same file loaded through the old fixed-column path and the new long path on migrated data, then diffed.
- Migration rehearsal: prod dump copy -> migrate -> row counts plus per-metric value checksums.
- E2E two-client LiveSync tests at measurement granularity; unit suites extend the existing test harness.

## 17. Deferred decisions (explicit, not open)

- Formula DSL: deferred; the `expr:` calculator type is the designated extension point.
- Sensory migration: deferred to a post-v3.0 effort; the core types are deliberately mode-agnostic to enable it.
- Image storage: unchanged in this redesign.
- Template designer UI inside DataViewer: deferred; custom templates are authored via the sidecar/manifest in v3.0.
- The Temperature Cycling checklist sheet stays raw-table; no schema modeling of procedural checklists in v3.0.
- Unit metadata and template header text disagree for two derived columns: `tpm_power_density` carries spec unit `mg/(puff*W)` but the real template header reads `mg/(W*s)`, and `oil_consumed` carries spec unit `mg` but the header reads "(Cumulative, g)"; the owner must reconcile `MetricDef.unit` against the header text before anything consumes `MetricDef.unit` (report labels, plot axes, or DB unit columns).

## Phase log

- Phase 0 (shadow harness) and Phase 1 (free-standing model plus schema-driven read path) both completed 2026-07-23.
- Gate evidence: tst_v3model 17 passed / 0 failed; tst_v3shadow 20 passed / 0 failed / 0 skipped; full suite via tests/run-tests.ps1 58 passed / 0 failed / 0 skipped.
- The corpus is 9 fixtures - empty, format_a, format_b, format_c, format_d, format_e, format_e_regime, multi_sheet, and old_standard - the last added in Phase 1 to lock the standardized-old Heating-Technology drop.
- The old-vs-new shadow diff was EMPTY on the first run over every fixture, so no divergence-chasing iteration was needed.
- Strangler decision: production calls SchemaDrivenReader with ColumnResolution::Positional so the read reproduces ExcelReader::extractRow byte-for-byte (fixed 12-wide-by-position), which is what byte-identity with the legacy pipeline requires.
- ColumnResolution::NameFirst (header-text matching, covered by tst_v3model) stays dormant until Phase 2, when manifest-authored templates may reorder or rename columns.
- The Cart / Project header-layout sniffing (the isCartFormat / isProjectFormat landmark cells) is PERMANENT legacy-workbook support, not strangler scaffolding.
- It migrates into the Phase 2 SchemaResolver and is never deleted with the Phase 4 removal of ExcelReader's positional getters and processFileLegacy.
