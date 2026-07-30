# TPM v3 Phase 2c (Manifest + SchemaResolver + NameFirst) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make workbooks self-describing: a very-hidden `_dve_schema` grid sheet declares the template (metrics, headers, geometry, tags), a unified `SchemaResolver` runs the ladder manifest -> compiled-standard -> inference, and manifest sheets parse with NameFirst column resolution - so reordering or adding columns in a manifest workbook Just Works, proven by a demo workbook with shuffled standard columns plus a custom column.

**Architecture:** The manifest is Excel-native grid data parsed through the existing cell path (no new plumbing); `SchemaResolver` absorbs the routing DataProcessor currently does inline (standardFits fork + per-block Cart/Project landmark sniff); manifest-resolved sheets flow through the generalized key-based lowering (the inference lowering, parametrized) because the 12-wide positional `lowerSample` cannot survive reordered columns; write provenance for NameFirst sheets records the RESOLVED physical slot order, not schema order.
Byte-identity is untouched by construction: no existing workbook contains a manifest, the standard/inference ladder behavior is preserved, and `tst_v3shadow` (fixtures + corpus) remains the referee for the resolver refactor.
Spec anchors: design spec sections 5/7/14, registry doc (RATIFIED naming contract + section 5 rule 5 collisions warn-never-fail), Phase 2c ledger in the v3 memory file.

**Tech Stack:** C++17 / Qt 6.10 (qmake + MinGW), Qt Test, openpyxl (fixture generation, `sheet_state = 'veryHidden'`).

---

## Machine + repo rules (read first)

- Create all new source files with the Write tool ONLY - never python file writes (MIP labeling), never heredoc/echo.
- Ciphertext (`%TSD-Header-###%`): read via `git show HEAD:<path>`; run `python tools/decrypt_via_copy.py --apply` from repo root before any build.
- Public repo: `tests/corpus/` gitignored; never commit real workbooks or results.txt artifacts. Synthetic fixtures under `tests/data/` are fine and ARE committed.
- Branch `worktree-tpm-template-v3-research`. Commit per task; plain dashes; NO Co-Authored-By.
- Qt Test stdout is INVISIBLE - always `-o results.txt,txt`, and `/c/Qt/6.10.1/mingw_64/bin` MUST be on PATH when running test exes (silent death otherwise).
- Suite inner loop (from suite dir): `export PATH="/c/Qt/6.10.1/mingw_64/bin:/c/Qt/Tools/mingw1310_64/bin:$PATH"`, qmake, mingw32-make, `release/<suite>.exe -o results.txt,txt`.
- Full-suite gate: `powershell -ExecutionPolicy Bypass -File tests/run-tests.ps1` from repo root. NEVER run it concurrently with an app build (-j8 CPU starvation makes python-subprocess suites flake).
- -Werror -Wall -Wextra.

## Reference: verified seams (2026-07-29)

- `SchemaDrivenReader::resolveColumns` (`src/model/SchemaDrivenReader.cpp:26-70`): NameFirst pass 1 name-matches within the block, pass 2 positional fallback with a DOCUMENTED collision gap ("real collision handling lands with the Phase 2 manifest"). `parseSheet` (`:72-116`) reads `blocks = hdrWidth / blockCols`, per-block headers at `off + col - 1`, data rows until all-empty; the resolved column map is LOCAL and currently thrown away.
- `DataProcessor::processFile` (production, `src/pipeline/DataProcessor.cpp:296+`): iterates `reader.getSheetNames()` (which WILL include veryHidden sheets - openpyxl lists them all); SOP branch = sheet name contains "SOP" (`:383` area, and `:178-180` in the legacy referee); the standardFits fork routes to inference; `processSheet` (`:436+`) does the per-block Cart/Project landmark sniff (`sniff(1,0) Cart` / `sniff(0,5) "Project:"`), swaps header variants per block, and (2b) records write provenance with `headerCells` from the sniffed layout, empty on mixed-layout sheets.
- Lowering: `LegacyAdapter::lowerSample` maps `model::Sample.data[c]` -> `SampleData.dataRows[r][c]` IN SCHEMA ORDER - correct only when schema order == physical order (positional standard). `lowerInferredSheet` is the key-based, order-independent lowering (fixed keys -> fixed fields, unknown -> extra, derived recomputed, PV assembly, resistance_initial fallback) and currently hard-sets `fromInferredSchema = true`, `hasPerRowRegime = false`.
- 2b provenance recording appends `schema.columns` order into `columnKeys` - valid because positional standard and inference schemas are both in physical order. NameFirst manifest sheets break that assumption (see Task 3).
- `ExcelReader`: `loadFile` -> `getSheetNames` -> `selectSheet(name)` -> `currentSheetCells()`.
- The registry (`MetricRegistry`) carries every canonical metric/header def; `RegimeParser` exists; `effectiveTestName` filename fallback exists in `ReportData.h`.

## Non-goals

- NO sidecar (.xlsm) emission of manifests - that is template-revision work on the sidecar track; this phase ships the C++ writer + a python injection helper used by fixtures/tools.
- NO UI for authoring manifests (Phase 5 template builder).
- NO display of open/custom metrics (Phase 4) - the demo's custom column rides `DataRow::extra`, invisible for now, verified by tests not eyeballs.
- NO DB persistence changes (Phase 3).
- NO re-keying of old-era columns 10-12 on the standard path (dormant until Phase 3/4).
- NO per-block permuted-column support (each block resolved independently) beyond recording block-0's resolution - blocks are uniform in every known and generated workbook; the harness would catch a violator.

---

### Task 1: Manifest module (parse + write, poka-yoke validation)

**Files:**
- Create: `src/model/Manifest.h`, `src/model/Manifest.cpp`
- Modify: `DataViewerEnterprise.pro`, `tests/tst_v3model/tst_v3model.pro`
- Test: `tests/tst_v3model/tst_v3model.cpp`

**Manifest grid grammar** (column A = row tag or first field; values in columns B+; blank rows ignored; everything case-insensitive on tags):

```
[schema]
id          | custom-coil-test
version     | 1
sheets      | *                          <- comma-separated sheet names, or * (default block)
block_cols  | 13
column_header_row | 4
data_start_row    | 5
header_rows | 3
[header]                                  <- one row per header field, ordered
key | display name | row | col | type | unit
[column]                                  <- one row per column, ordered = PHYSICAL order
key | display name | type | unit | role | calculator | inputs | tags
```

- `type`: number/text/bool/mixed/numberlist/image (unknown -> warn, default text). `role`: measured/qualitative/derived/identity (unknown -> warn, default measured). `inputs`: comma-separated. `tags`: semicolon-separated `k=v` pairs.
- A column row whose `key` matches a registry metric INHERITS the registry def (aliases, tags, calculator) with manifest cells overriding display/unit/type when non-empty; unknown keys become open defs. Same inheritance for headers via `MetricRegistry::headerField`.
- Multiple `[schema]` blocks allowed; each applies to the sheets its `sheets` row names; `*` is the fallback block. First matching block wins.
- Poka-yokes (registry rule 5 - warn, never fail): unknown type/role, duplicate keys within a block (second occurrence skipped + warning), `columns.size() > block_cols` (block_cols raised to fit + warning), missing `[column]` section (block ignored + warning). Warnings surface as `QStringList` for the caller to log.

- [ ] **Step 1: Write the failing tests** (append to `tst_v3model.cpp`)

```cpp
#include "Manifest.h"   // with the other model includes

void TestV3Model::manifestRoundTripsSchema()
{
    // standardV1 -> grid -> parse must reproduce keys, order, geometry,
    // header positions, and per-row-regime column.
    const TemplateSchema src = standardV1(true);
    const QVector<QVector<QVariant>> grid = Manifest::gridFor(src, {QStringLiteral("*")});
    const Manifest::ParseResult pr = Manifest::parse(grid);
    QVERIFY(pr.warnings.isEmpty());
    QCOMPARE(pr.blocks.size(), 1);
    const TemplateSchema& back = pr.blocks[0].schema;
    QCOMPARE(back.blockCols, src.blockCols);
    QCOMPARE(back.dataStartRow, src.dataStartRow);
    QCOMPARE(back.columnHeaderRow, src.columnHeaderRow);
    QCOMPARE(back.columns.size(), src.columns.size());
    for (int i = 0; i < src.columns.size(); ++i) {
        QCOMPARE(back.columns[i].key,  src.columns[i].key);
        QCOMPARE(back.columns[i].role, src.columns[i].role);
    }
    QCOMPARE(back.headerFields.size(), src.headerFields.size());
    const HeaderFieldDef* media = back.headerField(QStringLiteral("media"));
    QVERIFY(media);
    QCOMPARE(media->row, 2);
    QCOMPARE(media->col, 2);
    // Registry inheritance: known keys regain their alias sets.
    QVERIFY(back.column(QStringLiteral("before_weight"))->headerAliases.contains(
        QStringLiteral("Before weight/g")));
}

void TestV3Model::manifestParsesCustomColumnsAndTags()
{
    TemplateSchema s;
    s.schemaId = QStringLiteral("demo");
    s.version = 1;
    s.headerRows = 3; s.columnHeaderRow = 4; s.dataStartRow = 5; s.blockCols = 3;
    MetricDef custom;
    custom.key = QStringLiteral("coil_temp");
    custom.displayName = QStringLiteral("Coil Temp (C)");
    custom.type = ValueType::Number;
    custom.unit = QStringLiteral("C");
    custom.role = Role::Measured;
    custom.tags.insert(QStringLiteral("source"), QStringLiteral("thermocouple"));
    s.columns = { *MetricRegistry::metric("puffs"), *MetricRegistry::metric("tpm"), custom };
    const auto pr = Manifest::parse(Manifest::gridFor(s, {QStringLiteral("Demo Sheet")}));
    QVERIFY(pr.warnings.isEmpty());
    QCOMPARE(pr.blocks[0].sheets, QStringList{QStringLiteral("Demo Sheet")});
    const MetricDef* back = pr.blocks[0].schema.column(QStringLiteral("coil_temp"));
    QVERIFY(back);
    QCOMPARE(back->unit, QStringLiteral("C"));
    QCOMPARE(back->tags.value(QStringLiteral("source")), QStringLiteral("thermocouple"));
}

void TestV3Model::manifestPokaYokesWarnNeverFail()
{
    TemplateSchema s = standardV1(false);
    QVector<QVector<QVariant>> grid = Manifest::gridFor(s, {QStringLiteral("*")});
    // Corrupt: an unknown role and a duplicate key row.
    // Locate the [column] row for "smell" and set its role cell to "banana";
    // duplicate the "notes" row. (Walk the grid: rows after the [column] tag.)
    bool corrupted = false;
    for (int r = 0; r < grid.size(); ++r) {
        if (grid[r].value(0).toString() == QLatin1String("smell")) {
            grid[r][4] = QStringLiteral("banana");        // role cell
            grid.insert(r + 1, grid[r + 1]);              // duplicate next row once
            corrupted = true;
            break;
        }
    }
    QVERIFY(corrupted);
    const auto pr = Manifest::parse(grid);
    QVERIFY(!pr.warnings.isEmpty());                      // warned...
    QCOMPARE(pr.blocks.size(), 1);                        // ...but parsed
    QVERIFY(pr.blocks[0].schema.column(QStringLiteral("smell")));  // role defaulted
}

void TestV3Model::manifestSheetScoping()
{
    const TemplateSchema a = standardV1(false);
    const TemplateSchema b = standardV1(true);
    QVector<QVector<QVariant>> grid = Manifest::gridFor(a, {QStringLiteral("Special")});
    grid += Manifest::gridFor(b, {QStringLiteral("*")});
    const auto pr = Manifest::parse(grid);
    QCOMPARE(pr.blocks.size(), 2);
    QCOMPARE(Manifest::blockForSheet(pr, QStringLiteral("Special"))->schema.columns[4].key,
             QStringLiteral("resistance"));
    QCOMPARE(Manifest::blockForSheet(pr, QStringLiteral("Anything"))->schema.columns[4].key,
             QStringLiteral("puffing_regime"));
    QVERIFY(!Manifest::blockForSheet(Manifest::ParseResult{}, QStringLiteral("x")));
}
```

- [ ] **Step 2: Run red** (compile failure on `Manifest.h`), **then implement**

`src/model/Manifest.h`:

```cpp
#pragma once
#include "TemplateSchema.h"
#include <QVariant>
#include <QVector>

namespace DVE { namespace model {

// The _dve_schema manifest: an Excel-native grid describing the template.
// Parsed through the same cell path as every other sheet; validation is
// poka-yoke (warn + proceed, registry naming-policy rule 5) - a manifest can
// be wrong, never fatal. Known metric/header keys inherit their registry defs
// (aliases, tags, calculators); manifest cells override display/unit/type.
class Manifest {
public:
    static const QString kSheetName;   // "_dve_schema"

    struct Block {
        TemplateSchema schema;
        QStringList    sheets;         // sheet names this block applies to; "*" = default
    };
    struct ParseResult {
        QVector<Block> blocks;
        QStringList    warnings;
    };

    static ParseResult parse(const QVector<QVector<QVariant>>& grid);

    // First block naming the sheet, else the "*" block, else nullptr.
    static const Block* blockForSheet(const ParseResult& pr, const QString& sheetName);

    // Writer: serialize a schema to manifest grid rows (round-trips through
    // parse). Appendable - concatenate grids for multi-block manifests.
    static QVector<QVector<QVariant>> gridFor(const TemplateSchema& schema,
                                              const QStringList& sheets);
};

}} // namespace DVE::model
```

`src/model/Manifest.cpp` - implement per the grammar above. Structure guidance (write real code, this is the shape):
- `kSheetName = QStringLiteral("_dve_schema")`.
- `parse`: walk rows; a cell-A value `[schema]`/`[header]`/`[column]` (trimmed, case-insensitive) switches section and `[schema]` starts a new Block (flush previous); in `[schema]` section rows are `tag | value` pairs (id, version, sheets, block_cols, column_header_row, data_start_row, header_rows); in `[header]`, rows are `key|display|row|col|type|unit` -> HeaderFieldDef, inheriting `MetricRegistry::headerField(key)` type/unit when cells empty; in `[column]`, rows are `key|display|type|unit|role|calculator|inputs|tags` -> MetricDef, starting from `*MetricRegistry::metric(key)` when known else a fresh def, then overriding non-empty cells; `tags` parsed as `k=v;k2=v2`. Duplicate key in a block: skip + warning. Type/role parse helpers with warn-default. After each block: if columns empty -> drop block + warning; if `columns.size() > blockCols` -> raise blockCols + warning.
- `blockForSheet`: exact name match first (case-insensitive, trimmed), then a block whose sheets contain "*".
- `gridFor`: emit exactly the grammar (types/roles as canonical lowercase strings; tags joined `k=v;...`; skip empty optional cells but keep column positions).
- Type/role string maps live in one place (used by both directions).

Register in `DataViewerEnterprise.pro` + `tst_v3model.pro`.

- [ ] **Step 3: Green + collision test note**

All 4 new tests pass; whole tst_v3model suite green (30 expected).

- [ ] **Step 4: Commit**

```bash
git add src/model/Manifest.h src/model/Manifest.cpp DataViewerEnterprise.pro tests/tst_v3model
git commit -m "feat(v3): _dve_schema manifest module - grid parse/write with registry inheritance + poka-yoke warnings"
```

---

### Task 2: SchemaResolver (routing ladder unified, byte-identity preserved)

**Files:**
- Create: `src/model/SchemaResolver.h`, `src/model/SchemaResolver.cpp`
- Modify: `src/pipeline/DataProcessor.cpp` (processSheet slims to consume the resolver), `DataViewerEnterprise.pro`, test .pro files that link DataProcessor (add SchemaResolver.cpp + Manifest.cpp on link failure)
- Test: `tests/tst_v3model/tst_v3model.cpp` (resolver unit tests), existing suites as gates

- [ ] **Step 1: Read `DataProcessor::processSheet` FULLY** (the production one, ~:436+). Inventory what moves: standard-vs-inference fork (`standardFits`), per-block Cart/Project landmark sniff + header-variant swap, perRowRegime resolution, blockLayouts vector feeding provenance recording. The Project sampleID join (project_name + sample_suffix) stays wherever it is today (ledger says it eventually moves here - do it ONLY if it is already inside processSheet's moved region; do not chase it into SheetProcessors).

- [ ] **Step 2: Write the failing resolver unit tests** (grids built like tst_v3inference's builders - keep them small):

```cpp
void TestV3Model::resolverLadder()
{
    // 1. Manifest present + block matches -> Manifest source, NameFirst.
    // 2. No manifest, standard 12-wide grid -> Standard source, Positional.
    // 3. No manifest, 13-wide S26-style grid -> Inference source, NameFirst.
    // Build one minimal grid per case; assert SchemaResolver::resolve returns
    // the right source enum, resolution, and schema id/blockCols.
}

void TestV3Model::resolverManifestDerivesPerRowRegime()
{
    // Manifest whose columns contain puffing_regime -> perRowRegime true;
    // resistance variant -> false.
}
```

(Write these as real compilable tests with real grids - reuse the existing grid-building style in the suite; the standard grid needs the first-3 header aliases and 12 columns.)

- [ ] **Step 3: Implement**

`src/model/SchemaResolver.h`:

```cpp
#pragma once
#include "Manifest.h"
#include "SchemaDrivenReader.h"
#include <QVariant>
#include <QVector>

namespace DVE { namespace model {

// One routing ladder for every sheet: manifest -> compiled standard ->
// header-driven inference. Absorbs the standardFits fork and the per-block
// Cart/Project landmark sniff that DataProcessor carried inline.
class SchemaResolver {
public:
    enum class Source { Manifest, Standard, Inference };

    struct Resolution {
        TemplateSchema           schema;
        Source                   source = Source::Standard;
        ColumnResolution         columnResolution = ColumnResolution::Positional;
        bool                     perRowRegime = false;
        // Standard source only: the layout each block's landmark sniff chose
        // (drives header-variant swaps + the provenance headerCells map).
        QVector<HeaderLayout>    blockLayouts;
    };

    // manifest may be null (workbook has no _dve_schema sheet).
    static Resolution resolve(const Manifest::ParseResult* manifest,
                              const QString& sheetName,
                              const QVector<QVector<QVariant>>& cells,
                              const QString& templateVersion);
};

}} // namespace DVE::model
```

`SchemaResolver.cpp`:
- Manifest rung: `Manifest::blockForSheet` hit -> schema from the block, `Source::Manifest`, `ColumnResolution::NameFirst`, `perRowRegime = schema.column("puffing_regime") != nullptr`.
- Standard rung: `SchemaInference::standardFits(cells, standardV1(perRow, Standard))` - MOVE the existing perRowRegime determination logic (read how processSheet derives it today - from templateVersion/header text - and move it verbatim) and the per-block sniff loop verbatim from processSheet into here, returning `blockLayouts`. `Source::Standard`, `Positional`.
- Inference rung: `SchemaInference::inferSchema`, `Source::Inference`, `NameFirst`, perRowRegime false (unchanged).
- `DataProcessor::processSheet` then: build/receive the manifest ParseResult (Task 4 wires the actual cells; until then pass nullptr so behavior is IDENTICAL), call resolve, and use `res.blockLayouts`/`res.schema`/`res.perRowRegime` exactly where the inline code used its own - the diff should read as code MOVING, not changing. Header-variant swap per block stays in processSheet if it mutates mSheet samples (check; keep mutations where the data lives).

- [ ] **Step 4: Gates - this is the risky refactor**

1. tst_v3model green (+2).
2. tst_v3inference 22/0/1, tst_dataprocessor 11/0/0 (run standalone, never during a build).
3. **tst_v3shadow 21/0/3** - the referee. If red: the move changed behavior; fix the move.
4. tst_v3roundtrip fixtures green (provenance recording still correct).

- [ ] **Step 5: Commit**

```bash
git add src/model/SchemaResolver.h src/model/SchemaResolver.cpp src/pipeline/DataProcessor.cpp DataViewerEnterprise.pro tests/
git commit -m "refactor(v3): SchemaResolver - manifest/standard/inference ladder unified, routing moved out of DataProcessor, shadow-gate green"
```

---

### Task 3: NameFirst slot exposure + generalized lowering

**Files:**
- Modify: `src/model/MetricSample.h` (Sheet gains resolved slots), `src/model/SchemaDrivenReader.cpp` (expose block-0 resolution), `src/model/LegacyAdapter.{h,cpp}` (parametrize the schema-sheet lowering)
- Test: `tests/tst_v3inference/tst_v3inference.cpp`

- [ ] **Step 1: Failing tests**

```cpp
void TestV3Inference::parseSheetExposesResolvedSlots()
{
    // Shuffled 3-col grid: headers [TPM (mg/puff), puffs, Notes]; a schema
    // whose columns are ordered [puffs, notes, tpm]; NameFirst resolution.
    // Sheet::columnSlots must be {1, 2, 0}: schema column i lives at block-
    // relative physical slot columnSlots[i].
}

void TestV3Inference::manifestSheetLowersByKeyWithProvenanceInSlotOrder()
{
    // Same shuffled grid through lowerSchemaSheet(..., fromInference=false,
    // perRowRegime=false): DataRow.puffs/notes/tpm land correctly despite the
    // shuffle; SheetResult.fromInferredSchema == false;
    // columnKeys == {"tpm","puffs","notes"}  (PHYSICAL slot order);
    // CellAddress::dataCell(sheet, sample, "puffs", 0) points at slot 1.
}
```

(Write as real tests with real grids in the suite's style.)

- [ ] **Step 2: Implement**

- `model::Sheet` gains `QVector<int> columnSlots;` - block-relative physical slot per schema column, from block 0's `resolveColumns` map (`map[i] - off`). `parseSheet` fills it (all resolutions; for Positional it is identity). Comment: per-block permutation divergence is out of scope (Non-goals) - block 0 is authoritative and the round-trip harness would flag a violator.
- `LegacyAdapter`: rename/extend `lowerInferredSheet` into the general `lowerSchemaSheet(const Sheet&, const QString& sheetName, const QString& templateVersion, bool fromInference, bool perRowRegime)`; keep `lowerInferredSheet` as a thin forwarding wrapper (call sites + tests untouched). Changes inside: `fromInferredSchema = fromInference`; `hasPerRowRegime = perRowRegime`; per-row `puffing_regime` series maps into `dr.puffingRegime` (already handled - verify); provenance recording orders `columnKeys` by `columnSlots` (invert the mapping: `keysBySlot[slot] = schema.columns[i].key`; unclaimed slots get the open key already at that physical position - for inference schemas columnSlots is identity so output is UNCHANGED, assert via existing tests staying green).
- headerCells for manifest sheets: from `schema.headerFields` (row/col are block-relative value-cell coordinates by manifest grammar) - same recording as inference.

- [ ] **Step 3: Gates**

tst_v3inference green (+2, all existing pass unchanged - the identity-slot assertion), tst_v3roundtrip fixtures green, tst_v3shadow 21/0/3.

- [ ] **Step 4: Commit**

```bash
git add src/model/MetricSample.h src/model/SchemaDrivenReader.cpp src/model/LegacyAdapter.h src/model/LegacyAdapter.cpp tests/
git commit -m "feat(v3): resolved-slot exposure + generalized lowerSchemaSheet - NameFirst sheets lower by key with slot-ordered provenance"
```

---

### Task 4: DataProcessor wiring - manifest read, sheet exclusion, routing

**Files:**
- Modify: `src/pipeline/DataProcessor.cpp` (+ `.h` if a member is needed)
- Test: `tests/tst_v3inference/tst_v3inference.cpp` (E2E-style with an in-memory manifest), gates

- [ ] **Step 1: Implement (read the processFile loop first)**

- In `processFile` (production AND the legacy referee `processFileLegacy`): if `sheetNames` contains `Manifest::kSheetName`, remove it from `result.sheetNames` and skip it in the sheet loop (the referee just skips it - legacy never parses manifests; production additionally parses it). A veryHidden internal sheet must never appear in the navigator or as a data sheet.
- Production: before the loop, `selectSheet(_dve_schema)` + `currentSheetCells()` -> `Manifest::parse` -> hold the ParseResult; log warnings via the existing notify/qWarning pattern (poka-yoke: never abort). Pass `&manifest` into processSheet/resolver (Task 2 seam).
- processSheet: `Source::Manifest` resolutions parse with `res.schema` + NameFirst and lower via `lowerSchemaSheet(sheet, name, templateVersion, /*fromInference=*/false, res.perRowRegime)`. SOP branch precedence unchanged (name-contains-SOP wins before the resolver - manifests do not describe SOP sheets).

- [ ] **Step 2: Failing-then-green E2E test** (tst_v3inference, no xlsx needed if the suite has a grid-level path; otherwise fold into Task 5's fixture E2E and keep this task's verification at the unit gates)

- [ ] **Step 3: Gates:** tst_v3inference, tst_dataprocessor 11/0/0, tst_v3shadow 21/0/3 (the referee skips `_dve_schema` identically - a manifest-bearing FIXTURE does not exist yet, so identity is trivially preserved; Task 5 adds one and keeps it OUT of the shadow data set or verifies both parsers skip the sheet symmetrically - decide by reading how the shadow suite enumerates fixtures and document the choice in the commit).

- [ ] **Step 4: Commit**

```bash
git add src/pipeline/DataProcessor.cpp src/pipeline/DataProcessor.h tests/
git commit -m "feat(v3): manifest wiring - _dve_schema parsed + excluded from display, manifest sheets route NameFirst through the resolver"
```

---

### Task 5: Demo workbook + E2E proof

**Files:**
- Modify: `tests/generate_fixtures.py` (gen_manifest_demo), commit the generated `tests/data/manifest_demo.xlsx`
- Test: `tests/tst_v3inference/tst_v3inference.cpp` (E2E), `tests/tst_v3roundtrip` (picks the fixture up automatically)

- [ ] **Step 1: Fixture generator** (follow generate_fixtures.py's existing gen_* pattern; python is fine here - it writes to tests/data which the repo tracks, and fixtures are data, not source; openpyxl sets the manifest sheet's `sheet_state = "veryHidden"`)

`manifest_demo.xlsx`: one data sheet "Custom Coil Test" with SHUFFLED standard columns plus one custom column - physical order: `TPM (mg/puff), puffs, Before Weight (g), After Weight (g), Coil Temp (C), Draw Pressure (kpa), Puffing Regime, Smell, Clog, Notes` (10 cols), the standardized header band (Standard layout positions), 2 samples, ~6 data rows each with distinct values (puffs 10/20/..., coil temps 200+r). Plus the `_dve_schema` sheet (veryHidden) whose `[column]` order matches the PHYSICAL order above with `coil_temp` declared `number/C/measured` + tag `source=thermocouple`, `sheets | *`, `block_cols | 10`.

- [ ] **Step 2: E2E test** (tst_v3inference):

```cpp
void TestV3Inference::manifestDemoEndToEnd()
{
    // Load tests/data/manifest_demo.xlsx through DataProcessor::processFile
    // (same loading mechanics as the suite's other E2E tests; QSKIP without python).
    // Assert:
    //  - result.sheetNames does NOT contain "_dve_schema"; sheets.size() == 1.
    //  - sheet.fromInferredSchema == false; hasPerRowRegime == true.
    //  - sample 1: puffs series from PHYSICAL column 2 (values 10,20,...) landed
    //    in DataRow.puffs; tpm recomputed (Derived recompute unchanged);
    //    smell/clog/notes from their shuffled positions.
    //  - custom column: rows[r].extra["coil_temp"] == 200+r.
    //  - provenance: columnKeys[0] == "tpm" (physical slot order),
    //    columnKeys[4] == "coil_temp"; CellAddress::dataCell(...,"puffs",0).col
    //    == startColumn + 2 + 1.
    //  - sample 2 startColumn == 10.
}
```

- [ ] **Step 3: Harness + shadow interplay**

Run tst_v3roundtrip over fixtures - manifest_demo must PASS identity (mapped-domain identity works on NameFirst sheets thanks to slot-ordered provenance). Run tst_v3shadow: read how it enumerates fixtures; if manifest_demo enters its data set, production-vs-legacy will DIFFER by design (legacy parses the shuffled sheet positionally + garbage) - add it to the same inference-skip mechanism the S26-class fixtures use (lastFileUsedInference is false here, so key the skip on the manifest instead: the cleanest is skipping files whose sheetNames contained `_dve_schema` - implement whatever the suite's existing skip idiom supports and comment it). Expected: 21 passed / 0 failed / 4 skipped (one new skip).

- [ ] **Step 4: Commit**

```bash
git add tests/generate_fixtures.py tests/data/manifest_demo.xlsx tests/tst_v3inference tests/tst_v3shadow
git commit -m "feat(v3): manifest demo workbook - shuffled standard columns + custom coil_temp parse correctly via _dve_schema (E2E + harness green)"
```

---

### Task 6: Hardening (ledger items that belong to this phase)

**Files:**
- Modify: `src/model/SchemaDrivenReader.cpp` (collision poka-yoke), `src/ExcelReader.cpp` (timeout + JSON-error guards)
- Test: `tests/tst_v3model/tst_v3model.cpp`, `tests/tst_excelreader/tst_excelreader.cpp`

- [ ] **Step 1: resolveColumns pass-2 collision poka-yoke** (registry rule 5). Failing test: schema [a, b] where b's name matches physical slot 0 and a matches nothing -> old code gave a slot 0 too (both read the same column); new behavior: pass 2 assigns the NEAREST FREE slot at-or-after the default position (wrap to earlier free slots if none after), never a taken one; qWarning once per sheet. Positional resolution untouched (byte-identity). Assert a and b read different columns.
- [ ] **Step 2: Phantom trailing samples guard** - NameFirst/inference paths only: after parseSheet builds samples, drop TRAILING samples whose every header cell and every data cell is empty (over-padded grids produce them). Failing test: a 26-wide grid with one real 13-wide block + 13 empty padded columns -> 1 sample, not 2. MUST NOT touch Positional standard parsing (legacy kept phantom behavior; shadow gate verifies).
- [ ] **Step 3: ExcelReader guards** - read the python-subprocess call sites: (a) ensure a hung python cannot block forever (verify/apply a waitForFinished timeout consistent with the existing patterns - if one exists, this is a no-op, document it); (b) a sheet whose JSON fails to parse must skip that SHEET with a logged error, not abort the FILE (read the current error path first; add the narrower guard only if missing - if both already exist, record that in the commit message and move on).
- [ ] **Step 4: Gates:** touched suites + tst_v3shadow 21/0/4 + tst_v3roundtrip fixtures.
- [ ] **Step 5: Commit**

```bash
git add src/model/SchemaDrivenReader.cpp src/ExcelReader.cpp tests/
git commit -m "fix(v3): hardening - NameFirst collision poka-yoke, phantom trailing samples dropped on schema paths, ExcelReader guards verified"
```

---

### Task 7: Gates, v2.10.4, installer, docs

- [ ] Full suite (NOT concurrent with builds): expect 60/0/0. Corpus runs: tst_v3shadow 25/0/6 (one new fixture skip), tst_v3roundtrip corpus green.
- [ ] `VERSION = 2.10.4`; decrypt pass; in-tree ROOT release: qmake CONFIG+=release, `mingw32-make clean`, `mingw32-make -j8`; `MSYS_NO_PATHCONV=1 cmd /c '.\build_installer.bat' < /dev/null`; verify ProductVersion 2.10.4.
- [ ] `release_overview/release_overview_v_2_10_4.txt`: internal build; manifest capability + demo instructions (open `tests/data/manifest_demo.xlsx` from the repo checkout - shuffled columns + custom column parse correctly, edit cells and confirm placement; T58G now shows the filename in the navigator; regression eyeball on standard + S26 + UserSim). Note Synology freeze until v3.0.0.
- [ ] Copy installer + overview to main repo; registry doc Phase log line; sprint tracker; memory. Commit docs.

---

## Self-review (authoring time)

- Spec coverage: design spec 14 manifest-in-workbook (Tasks 1/4/5), section 7 name-first (Tasks 2/3), registry rule 5 collisions (Tasks 1/6), 2c ledger (Task 6 + per-block permutation consciously deferred in Non-goals), demo = the owner's original "reorder + add without breaking" acceptance (Task 5).
- Placeholders: Tasks 2/3/4 tests are specified as contracts with real grid recipes where the suite's builders must be reused - the executor is told exactly which suite idiom to mirror; every new module's API is complete. Manifest.cpp is shape-guided rather than transcribed - its grammar table above IS the spec, and the round-trip test pins it.
- Type consistency: `Manifest::{Block,ParseResult,parse,gridFor,blockForSheet,kSheetName}`, `SchemaResolver::{Source,Resolution,resolve}`, `Sheet::columnSlots`, `lowerSchemaSheet(..., fromInference, perRowRegime)` used consistently across tasks.
- Byte-identity analysis: manifest rung fires only when `_dve_schema` exists (no existing file has one); resolver refactor is a code MOVE gated by shadow; slot exposure is identity on positional/inference paths; sheet exclusion applies to both parsers symmetrically; phantom-sample guard and collision fix scoped to non-Positional paths.
