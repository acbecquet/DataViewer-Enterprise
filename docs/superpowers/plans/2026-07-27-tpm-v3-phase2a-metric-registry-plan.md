# TPM v3 Phase 2a (Metric Registry Compilation) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Compile the ratified vocabulary registry (`docs/superpowers/specs/2026-07-27-tpm-v3-vocabulary-registry-draft.md`, RATIFIED) into code as a single-source `MetricRegistry`, extend the model type system (Bool/Mixed/NumberList/Image + `MetricDef::tags`), add the regime parser, and make schema inference draw its whole knowledge base from the registry - including PV1-PV5 assembling into the `draw_pressure_per_puff` list metric - with ZERO behavior change on the standard byte-identity path.

**Architecture:** `MetricRegistry` becomes the one place canonical keys / display names / aliases / types / units live; `StandardSchema` and `SchemaInference` become consumers that copy defs out of it (layout positions and presentation flags stay layout-side).
Name-first ACTIVATION, manifests, write-back, and UI display are NOT here - they are Phases 2b/2c; this plan only makes the vocabulary real in code and on the inference path, where output is not byte-gated.
Spec anchors: registry doc sections 2/3/5/8/9 (binding), design spec section 18.

**Tech Stack:** C++17 / Qt 6.10 (qmake + MinGW), Qt Test, existing suites `tst_v3model` / `tst_v3inference` / `tst_reportdatajson` / `tst_v3shadow`.

---

## Machine + repo rules (read first)

- This machine (S1134987) MIP-labels files written by trusted Python; the Write/Edit tools do NOT label. **Create all new source files with the Write tool** - never the python delete-and-rewrite pattern from the project CLAUDE.md (that guidance is for the other machine; see the `mip-labels-python-writes` memory).
- If any existing file reads as ciphertext (`%TSD-Header-###%`), read it via `git show HEAD:<path>` instead, and run `python tools/decrypt_via_copy.py --apply` from the repo root before any build.
- The repo is PUBLIC. Never commit real test workbooks (`tests/corpus/` is gitignored). Fixtures under `tests/data/` are synthetic and fine.
- Work on branch `worktree-tpm-template-v3-research`. Commit after every task; no co-author lines in commit messages; plain dashes, no em dashes.
- Qt Test stdout is invisible in this shell - ALWAYS run suites with `-o results.txt,txt` and read the file.

## Build & test commands (referenced by all tasks)

Single suite inner loop (from the suite dir, e.g. `tests/tst_v3model`):

```bat
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
C:\Qt\6.10.1\mingw_64\bin\qmake.exe
mingw32-make
release\tst_v3model.exe -o results.txt,txt
```

(If the binary lands in `debug\`, run that one.)

Full-suite gate (from repo root): `powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1` - expect every suite PASS.
New app sources must be added to `DataViewerEnterprise.pro` (`SOURCES`/`HEADERS`) AND to the `.pro` of every test suite that links model sources (check how `tests/tst_v3model/tst_v3model.pro`, `tests/tst_v3inference/tst_v3inference.pro`, `tests/tst_v3shadow/tst_v3shadow.pro` list `src/model/*.cpp` and extend each list the same way).

## Non-goals (explicit - do NOT do these here)

- NO change to what the standard positional path emits: old-era 12-col files keep lowering cols 10/11/12 into `tpmPowerDensity`/`variationTPM`/`oilConsumed` fixed fields even though the registry now knows them as `tpm_puff_density`/`tpm_consistency`/`rolling_avg_tpm`. The honest re-keying activates with the Phase 3/4 coordinated flip. `tst_v3shadow` byte-identity is the gate.
- NO regime-split values emitted into parse output (the parser ships tested but dormant; 2b/2c consume it).
- NO calculator implementations for the new derived keys (`puff_density_v1`, `consistency_v1`, `rolling_avg_v1`, `regime_split_v1` are registered NAMES only; the CalculatorRegistry is later Phase 2 work).
- NO era-indicator fallback for bare ambiguous titles ("TPM Power Density" without a unit suffix has never been observed; normalized full titles are unique - YAGNI until a real file collides).
- NO UI display, NO DB persistence of extras, NO manifest, NO write-back changes.

---

### Task 1: RegimeParser

**Files:**
- Create: `src/model/RegimeParser.h`, `src/model/RegimeParser.cpp`
- Modify: `DataViewerEnterprise.pro` (SOURCES + HEADERS), `tests/tst_v3model/tst_v3model.pro` (add `RegimeParser.cpp` to its model sources)
- Test: `tests/tst_v3model/tst_v3model.cpp`

- [ ] **Step 1: Write the failing tests** (append slots + bodies to `tst_v3model.cpp`, register in the class's `private slots`)

```cpp
#include "RegimeParser.h"   // add near the other model includes

void TestV3Model::regimeParserStandard()
{
    const auto r = RegimeParser::parse(QStringLiteral("60mL/3s/30s"));
    QVERIFY(r.valid);
    QCOMPARE(r.puffVolumeMl, 60.0);
    QCOMPARE(r.puffTimeS, 3.0);
    QCOMPARE(r.puffRestS, 30.0);
    QCOMPARE(r.sessionRestS, 0.0);   // 3-part regime defaults session rest to 0 (registry 8.2)
}

void TestV3Model::regimeParserFourPart()
{
    const auto r = RegimeParser::parse(QStringLiteral("60mL/3s/30s/5minute"));
    QVERIFY(r.valid);
    QCOMPARE(r.sessionRestS, 300.0);
}

void TestV3Model::regimeParserRealVariants()
{
    QCOMPARE(RegimeParser::parse(QStringLiteral("100mL/2.5s/15s")).puffTimeS, 2.5);
    QCOMPARE(RegimeParser::parse(QStringLiteral("200mL/10s/60s")).puffRestS, 60.0);
    QVERIFY(RegimeParser::parse(QStringLiteral("60 ml / 3 s / 30 s")).valid);  // spacing + case tolerant
    QCOMPARE(RegimeParser::parse(QStringLiteral("60mL/3s/30s/90s")).sessionRestS, 90.0);
    QCOMPARE(RegimeParser::parse(QStringLiteral("60mL/3s/30s/2min")).sessionRestS, 120.0);
}

void TestV3Model::regimeParserRejectsGarbage()
{
    QVERIFY(!RegimeParser::parse(QString()).valid);
    QVERIFY(!RegimeParser::parse(QStringLiteral("as needed")).valid);
    QVERIFY(!RegimeParser::parse(QStringLiteral("60mL/3s")).valid);          // too few parts
    QVERIFY(!RegimeParser::parse(QStringLiteral("60mL/3s/30s/x/y")).valid);  // too many parts
}
```

- [ ] **Step 2: Run to verify it fails**

Run the tst_v3model inner loop. Expected: compile FAILURE (`RegimeParser.h: No such file`). A compile failure on the new include is the red state.

- [ ] **Step 3: Implement**

`src/model/RegimeParser.h`:

```cpp
#pragma once
#include <QString>

namespace DVE { namespace model {

// Parsed puffing-regime string ("60mL/3s/30s" or "60mL/3s/30s/5minute").
// Registry contract (RATIFIED 2026-07-27, section 8.2/9.2): the canonical
// representation is these four values; the composite string is a legacy
// source encoding. 3-part regimes default sessionRestS to 0.
struct RegimeParts {
    bool   valid        = false;
    double puffVolumeMl = 0.0;   // part 1, mL
    double puffTimeS    = 0.0;   // part 2 (always the middle value), s
    double puffRestS    = 0.0;   // part 3, s
    double sessionRestS = 0.0;   // optional part 4, s (minutes converted)
};

class RegimeParser {
public:
    static RegimeParts parse(const QString& text);
};

}} // namespace DVE::model
```

`src/model/RegimeParser.cpp`:

```cpp
#include "RegimeParser.h"
#include <QRegularExpression>
#include <QStringList>

namespace DVE { namespace model {

namespace {

// "60mL" -> 60 given suffix ml; "2.5s" -> 2.5 given suffix s; strict about the
// unit family but tolerant of spacing and case. Returns false on mismatch.
bool numberWithUnit(const QString& raw, const QStringList& units, double* out)
{
    static const QRegularExpression re(
        QStringLiteral("^\\s*([0-9]+(?:\\.[0-9]+)?)\\s*([a-zA-Z]*)\\s*$"));
    const auto m = re.match(raw);
    if (!m.hasMatch()) return false;
    const QString unit = m.captured(2).toLower();
    if (!units.contains(unit)) return false;
    *out = m.captured(1).toDouble();
    return true;
}

} // namespace

RegimeParts RegimeParser::parse(const QString& text)
{
    RegimeParts r;
    const QStringList parts = text.split(QLatin1Char('/'));
    if (parts.size() < 3 || parts.size() > 4) return r;

    if (!numberWithUnit(parts[0], {QStringLiteral("ml")}, &r.puffVolumeMl)) return r;
    if (!numberWithUnit(parts[1], {QStringLiteral("s"), QStringLiteral("sec")}, &r.puffTimeS)) return r;
    if (!numberWithUnit(parts[2], {QStringLiteral("s"), QStringLiteral("sec")}, &r.puffRestS)) return r;

    if (parts.size() == 4) {
        double v = 0.0;
        if (numberWithUnit(parts[3], {QStringLiteral("s"), QStringLiteral("sec"),
                                      QStringLiteral("second"), QStringLiteral("seconds")}, &v)) {
            r.sessionRestS = v;
        } else if (numberWithUnit(parts[3], {QStringLiteral("min"), QStringLiteral("minute"),
                                             QStringLiteral("minutes"), QStringLiteral("m")}, &v)) {
            r.sessionRestS = v * 60.0;
        } else {
            return r;
        }
    }
    r.valid = true;
    return r;
}

}} // namespace DVE::model
```

Register `src/model/RegimeParser.cpp` / `.h` in `DataViewerEnterprise.pro` next to the other `src/model/` entries, and add the `.cpp` to `tests/tst_v3model/tst_v3model.pro`'s source list.

- [ ] **Step 4: Run to verify it passes**

tst_v3model inner loop, read results.txt. Expected: all tests PASS including the 4 new ones.

- [ ] **Step 5: Commit**

```bash
git add src/model/RegimeParser.h src/model/RegimeParser.cpp DataViewerEnterprise.pro tests/tst_v3model
git commit -m "feat(v3): RegimeParser - 4-part puffing-regime split per ratified registry 8.2/9.2"
```

---

### Task 2: Type-system extensions + MetricRegistry

**Files:**
- Modify: `src/model/MetricDef.h`
- Create: `src/model/MetricRegistry.h`, `src/model/MetricRegistry.cpp`
- Modify: `DataViewerEnterprise.pro`, `tests/tst_v3model/tst_v3model.pro` (+ the other model-linking test .pro files when they fail to link - see Step 4)
- Test: `tests/tst_v3model/tst_v3model.cpp`

- [ ] **Step 1: Check no exhaustive switches exist over ValueType**

Run: `grep -rn "switch" src/ tests/ --include=*.cpp --include=*.h | grep -i valuetype`
Expected: no hits (ValueType is only used in assignments/ternaries). If a hit appears, extend that switch in Step 3.

- [ ] **Step 2: Write the failing tests** (append to `tst_v3model.cpp`)

```cpp
#include "MetricRegistry.h"   // add near the other model includes

void TestV3Model::registryResolvesEveryObservedSpelling()
{
    // Data-driven: verbatim spellings from the RATIFIED registry doc -> canonical key.
    const struct { const char* spelling; const char* key; } metricRows[] = {
        {"puffs", "puffs"},
        {"Before weight/g", "before_weight"},
        {"Before Weight/g", "before_weight"},
        {"Before weight (g)", "before_weight"},
        {"After weight/g", "after_weight"},
        {"Draw Pressure (kpa)", "draw_pressure"},
        {"Resistance", "resistance"},
        {"Puffing Regime", "puffing_regime"},
        {"Smell (1-4)", "smell"},
        {"Smell (0-4)", "smell"},
        {"Clog (Y/N)", "clog"},
        {"Clog?", "clog"},
        {"Notes", "notes"},
        {"TPM (mg/puff)", "tpm"},
        {"TPM Power Density (mg/(W*s))", "tpm_power_density"},   // current era
        {"TPM Power Density (mg/puff/W)", "tpm_puff_density"},   // old era = TPM/P (registry 9.1)
        {"Variation in TPM (%)", "variation_tpm"},
        {"TPM Consistency", "tpm_consistency"},
        {"Rolling Average TPM", "rolling_avg_tpm"},
        {"Oil Consumed (Cumulative, g)", "oil_consumed"},
        {"Chronology", "chronology"},
        {"Failure? (if yes, add detailed notes)", "failure"},
    };
    for (const auto& row : metricRows) {
        const MetricDef* m = MetricRegistry::metricByAlias(
            SchemaDrivenReader::normalizeHeader(QString::fromUtf8(row.spelling)));
        QVERIFY2(m, row.spelling);
        QCOMPARE(m->key, QString::fromUtf8(row.key));
    }

    const struct { const char* label; const char* key; } headerRows[] = {
        {"Date:", "date"},
        {"Tester:", "tester"},
        {"Sample ID:", "sample_id"},
        {"Cart #", "sample_id"},
        {"Project:", "project_name"},
        {"Sample:", "sample_suffix"},
        {"Distributor:", "distributor"},
        {"Resistance (Ω):", "resistance"},
        {"Resistance (Ohms):", "resistance"},
        {"Ri (Ohms)", "resistance_initial"},
        {"Rf (Ohms)", "resistance_final"},
        {"Voltage:", "voltage"},
        {"Power:", "power"},
        {"Heating Technology:", "heating_technology"},
        {"Heater Technology:", "heating_technology"},
        {"Media:", "media"},
        {"Viscosity:", "viscosity"},
        {"Initial Oil Mass:", "initial_oil_mass"},
        {"Fill Volume:", "fill_volume"},
        {"Number of Samples", "number_of_samples"},
        {"Puffing Regime:", "puffing_regime"},
        {"Puff Regime", "puffing_regime"},
        {"Did this burn?", "did_burn"},
        {"Did this clog?", "did_clog"},
        {"Did this leak?", "did_leak"},
        {"Coil Material", "coil_material"},
        {"Thermal Conductivity", "thermal_conductivity"},
        {"Column inner diameter", "column_inner_diameter"},
        {"Column length", "column_length"},
        {"Coil shape", "coil_shape"},
        {"Cotton length (if applicable)", "cotton_length"},
    };
    for (const auto& row : headerRows) {
        const HeaderFieldDef* h = MetricRegistry::headerByLabel(
            SchemaDrivenReader::normalizeHeader(QString::fromUtf8(row.label)));
        QVERIFY2(h, row.label);
        QCOMPARE(h->key, QString::fromUtf8(row.key));
    }
}

void TestV3Model::registryPerPuffAliases()
{
    for (int i = 1; i <= 5; ++i) {
        const auto hit = MetricRegistry::perPuffAlias(QStringLiteral("pv%1").arg(i));
        QCOMPARE(hit.targetKey, QStringLiteral("draw_pressure_per_puff"));
        QCOMPARE(hit.index, i);
    }
    QCOMPARE(MetricRegistry::perPuffAlias(QStringLiteral("puffs")).index, 0);  // miss
}

void TestV3Model::registryNamespacesAreCollisionFree()
{
    // Naming-policy rule 1/7: within a namespace no two entries may claim the
    // same normalized alias (displayName, key, and headerAliases all count).
    QSet<QString> seen;
    for (const MetricDef& m : MetricRegistry::allMetrics()) {
        QStringList names{m.displayName, m.key};
        names += m.headerAliases;
        for (const QString& n : names) {
            const QString norm = SchemaDrivenReader::normalizeHeader(n);
            if (norm.isEmpty()) continue;
            QVERIFY2(!seen.contains(norm), qPrintable(m.key + ": " + n));
            seen.insert(norm);
        }
    }
    QSet<QString> seenH;
    for (const HeaderFieldDef& h : MetricRegistry::allHeaderFields()) {
        QStringList names{h.displayName, h.key};
        names += MetricRegistry::headerAliasesFor(h.key);
        for (const QString& n : names) {
            const QString norm = SchemaDrivenReader::normalizeHeader(n);
            if (norm.isEmpty()) continue;
            QVERIFY2(!seenH.contains(norm), qPrintable(h.key + ": " + n));
            seenH.insert(norm);
        }
    }
}

void TestV3Model::registryNewTypesAndTags()
{
    const MetricDef* dp = MetricRegistry::metric(QStringLiteral("draw_pressure_per_puff"));
    QVERIFY(dp);
    QCOMPARE(dp->type, ValueType::NumberList);
    QCOMPARE(dp->unit, QStringLiteral("kPa"));
    QVERIFY(dp->tags.contains(QStringLiteral("source")));

    const MetricDef* v = MetricRegistry::metric(QStringLiteral("voltage"));
    QVERIFY(v);
    QCOMPARE(v->type, ValueType::Mixed);         // number OR curve-name string (owner D1 addendum)

    const MetricDef* img = MetricRegistry::metric(QStringLiteral("image"));
    QVERIFY(img);
    QCOMPARE(img->type, ValueType::Image);

    const HeaderFieldDef* burn = MetricRegistry::headerField(QStringLiteral("did_burn"));
    QVERIFY(burn);
    QCOMPARE(burn->type, ValueType::Bool);

    // Regime-split quartet registered with the documented default.
    const MetricDef* srt = MetricRegistry::metric(QStringLiteral("session_rest_time"));
    QVERIFY(srt);
    QCOMPARE(srt->unit, QStringLiteral("s"));
}
```

- [ ] **Step 3: Run to verify red, then implement**

Expected first: compile failure on `MetricRegistry.h`.

`src/model/MetricDef.h` - change the enum and add tags (keep everything else untouched):

```cpp
enum class ValueType { Number, Text, Bool, Mixed, NumberList, Image };
```

and inside `struct MetricDef`, after `int precision = 2;`:

```cpp
    // Open-ended key->value annotations (design spec section 18: "tag anything
    // on to the metrics"). Registry-authored today; manifest-authored in 2c.
    QMap<QString, QString> tags;
```

(add `#include <QMap>` to the header's includes.)

`src/model/MetricRegistry.h`:

```cpp
#pragma once
#include "MetricDef.h"
#include <QVector>

namespace DVE { namespace model {

// Compiled form of the RATIFIED vocabulary registry
// (docs/superpowers/specs/2026-07-27-tpm-v3-vocabulary-registry-draft.md).
// Single source of truth for canonical keys, display names, aliases, types,
// units, and tags. StandardSchema and SchemaInference copy defs OUT of here;
// layout positions and presentation flags stay layout-side.
// Naming policy (registry section 5): keys are forever; matching is by
// normalized alias (SchemaDrivenReader::normalizeHeader); metrics and header
// fields are separate namespaces.
class MetricRegistry {
public:
    static const QVector<MetricDef>& allMetrics();
    static const QVector<HeaderFieldDef>& allHeaderFields();

    static const MetricDef* metric(const QString& key);
    static const HeaderFieldDef* headerField(const QString& key);

    // Lookup by NORMALIZED alias text (caller normalizes). nullptr on miss.
    static const MetricDef* metricByAlias(const QString& normalized);
    static const HeaderFieldDef* headerByLabel(const QString& normalized);

    // Label spellings registered for a header key (normalized-alias sources).
    static QStringList headerAliasesFor(const QString& key);

    // PV1..PV5 -> element index of the draw_pressure_per_puff list metric
    // (registry 9.1/D5). index == 0 means "not a per-puff alias".
    struct PerPuffAlias { QString targetKey; int index = 0; };
    static PerPuffAlias perPuffAlias(const QString& normalized);
};

}} // namespace DVE::model
```

`src/model/MetricRegistry.cpp` - the full ratified vocabulary.
Builder helpers mirror StandardSchema's local `col()`/`hf()` but carry tags; every entry below is from the ratified doc (sections 2, 3, 8.2, 9.1, 9.4) - do not invent or drop entries:

```cpp
#include "MetricRegistry.h"
#include "SchemaDrivenReader.h"
#include <QHash>

namespace DVE { namespace model {

namespace {

MetricDef m(const QString& key, const QString& displayName, ValueType type, const QString& unit,
            Role role, const QStringList& aliases = QStringList(),
            const QString& calculator = QString(), const QStringList& inputs = QStringList(),
            const QMap<QString, QString>& tags = {})
{
    MetricDef d;
    d.key = key; d.displayName = displayName; d.headerAliases = aliases;
    d.type = type; d.unit = unit; d.role = role;
    d.calculator = calculator; d.inputs = inputs; d.tags = tags;
    return d;
}

HeaderFieldDef h(const QString& key, const QString& displayName, ValueType type,
                 const QString& unit = QString(), const QString& calculator = QString(),
                 const QStringList& inputs = QStringList())
{
    HeaderFieldDef f;
    f.key = key; f.displayName = displayName; f.type = type; f.unit = unit;
    f.calculator = calculator; f.inputs = inputs;
    return f;   // row/col stay 0: positions are layout-specific, not registry data
}

QVector<MetricDef> buildMetrics()
{
    QVector<MetricDef> v;
    // ── The standard column set (registry section 2) ──
    v << m("puffs", "puffs", ValueType::Number, "count", Role::Measured);
    v << m("before_weight", "Before Weight (g)", ValueType::Number, "g", Role::Measured,
           {"Before weight/g", "Before Weight/g", "Before weight (g)"});
    v << m("after_weight", "After Weight (g)", ValueType::Number, "g", Role::Measured,
           {"After weight/g", "After Weight/g", "After weight (g)"});
    v << m("draw_pressure", "Draw Pressure", ValueType::Number, "kPa", Role::Measured,
           {"Draw Pressure (kpa)"});
    v << m("resistance", "Resistance", ValueType::Number, "ohm", Role::Measured,
           {QString::fromUtf8("Resistance (Ω)")});
    v << m("puffing_regime", "Puffing Regime", ValueType::Text, "", Role::Qualitative, {}, {}, {},
           {{"encoding", "legacy composite string; canonical form is the 4 split metrics (9.2)"}});
    v << m("smell", "Smell", ValueType::Text, "", Role::Qualitative,
           {"Smell (1-4)", "Smell (0-4)"}, {}, {},
           {{"scale", "0-4; 1-4-era sheets leave blank for no event (D9)"}});
    v << m("clog", "Clog", ValueType::Text, "", Role::Qualitative, {"Clog (Y/N)", "Clog?"});
    v << m("notes", "Notes", ValueType::Text, "", Role::Qualitative);
    v << m("tpm", "TPM", ValueType::Number, "mg/puff", Role::Derived, {"TPM (mg/puff)"},
           "tpm_v1", {"puffs", "before_weight", "after_weight"});
    v << m("tpm_power_density", "TPM Power Density", ValueType::Number, "mg/(W*s)", Role::Derived,
           {"TPM Power Density (mg/(W*s))", "TPM/PD"},
           "power_density_v1", {"tpm", "header:power", "puff_time"},
           {{"approximation", "no transient power; TCR shifts instantaneous power (9.1)"}});
    v << m("tpm_puff_density", "TPM Puff Density", ValueType::Number, "mg/((n s) puff*W)", Role::Derived,
           {"TPM Power Density (mg/puff/W)", "TPM Puff Density (mg/(puff*W))"},
           "puff_density_v1", {"tpm", "header:power"},
           {{"approximation", "no transient power; TCR shifts instantaneous power (9.1)"},
            {"n", "puff length taken from the puff_time metric"}});
    v << m("variation_tpm", "Variation in TPM", ValueType::Number, "%", Role::Derived,
           {"Variation in TPM (%)", "Variation"},
           "variation_v1", {"tpm"},
           {{"definition", "template rolling CV in percent is canonical (D1)"}});
    v << m("tpm_consistency", "TPM Consistency", ValueType::Number, "", Role::Derived, {},
           "consistency_v1", {"tpm"},
           {{"definition", "STDEV.P/AVERAGE over the session window, FRACTION (9.4)"}});
    v << m("rolling_avg_tpm", "Rolling Average TPM", ValueType::Number, "mg/puff", Role::Derived, {},
           "rolling_avg_v1", {"tpm"});
    v << m("oil_consumed", "Oil Consumed", ValueType::Number, "g", Role::Derived,
           {"Oil Consumed (Cumulative, g)"},
           "oil_consumed_v1", {"before_weight", "after_weight"});
    // ── Ratified additions (registry section 8.2) ──
    v << m("puff_volume", "Puff Volume", ValueType::Number, "mL", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"});
    v << m("puff_time", "Puff Time", ValueType::Number, "s", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"});
    v << m("puff_rest_time", "Puff Rest Time", ValueType::Number, "s", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"});
    v << m("session_rest_time", "Session Rest Time", ValueType::Number, "s", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"},
           {{"default", "0 when the regime string has only 3 parts"}});
    v << m("draw_pressure_per_puff", "Draw Pressure (per puff)", ValueType::NumberList, "kPa",
           Role::Measured, {}, {}, {},
           {{"source", "data-logger series; list length = puffs in the session"},
            {"lowering", "historical PV1-PV5 columns assemble into this (D5)"}});
    v << m("voltage", "Voltage", ValueType::Mixed, "V", Role::Measured, {}, {}, {},
           {{"grouping", "groupable like puffing regime; string values name curves (D1)"}});
    v << m("image", "Image", ValueType::Image, "", Role::Qualitative);
    // ── Open metrics observed in the wild (registry section 2) ──
    v << m("chronology", "Chronology", ValueType::Text, "", Role::Qualitative);
    v << m("failure", "Failure", ValueType::Text, "", Role::Qualitative,
           {"Failure? (if yes, add detailed notes)", "Failure? (if yes, when)"});
    return v;
}

QVector<HeaderFieldDef> buildHeaderFields()
{
    QVector<HeaderFieldDef> v;
    v << h("test_name", "Test Name", ValueType::Text);
    v << h("date", "Date", ValueType::Text);
    v << h("tester", "Tester", ValueType::Text);
    v << h("sample_id", "Sample ID", ValueType::Text);
    v << h("project_name", "Project", ValueType::Text);
    v << h("sample_suffix", "Sample", ValueType::Text);
    v << h("distributor", "Distributor", ValueType::Text);
    v << h("resistance", "Resistance", ValueType::Number, "ohm");
    v << h("resistance_initial", "Resistance (Initial)", ValueType::Number, "ohm");
    v << h("resistance_final", "Resistance (Final)", ValueType::Number, "ohm");
    v << h("voltage", "Voltage", ValueType::Number, "V");
    v << h("power", "Power", ValueType::Number, "W", "power_v1",
           {"header:voltage", "header:resistance", "header:heating_technology"});
    v << h("heating_technology", "Heating Technology", ValueType::Text);
    v << h("media", "Media", ValueType::Text);
    v << h("viscosity", "Viscosity", ValueType::Number, "cP");
    v << h("initial_oil_mass", "Initial Oil Mass", ValueType::Number, "g");
    v << h("fill_volume", "Fill Volume", ValueType::Number, "mL");
    v << h("number_of_samples", "Number of Samples", ValueType::Number);
    v << h("puffing_regime", "Puffing Regime", ValueType::Text);
    v << h("did_burn", "Did this burn?", ValueType::Bool);
    v << h("did_clog", "Did this clog?", ValueType::Bool);
    v << h("did_leak", "Did this leak?", ValueType::Bool);
    // Cart-era design specs (registry 3.4; optional standard headers per D6)
    v << h("coil_material", "Coil Material", ValueType::Text);
    v << h("thermal_conductivity", "Thermal Conductivity", ValueType::Number);
    v << h("column_inner_diameter", "Column inner diameter", ValueType::Number, "mm");
    v << h("column_length", "Column length", ValueType::Number, "mm");
    v << h("coil_shape", "Coil shape", ValueType::Text);
    v << h("cotton_length", "Cotton length", ValueType::Number, "mm");
    return v;
}

// Label spellings per header key (registry section 3). displayName and key are
// implicit aliases; this table adds the punctuated / historical spellings.
QHash<QString, QStringList> buildHeaderAliasTable()
{
    QHash<QString, QStringList> t;
    t.insert("test_name", {"Test:", "Test"});
    t.insert("date", {"Date:"});
    t.insert("tester", {"Tester:"});
    t.insert("sample_id", {"Sample ID:", "Cart #"});
    t.insert("project_name", {"Project:"});
    t.insert("sample_suffix", {"Sample:"});
    t.insert("distributor", {"Distributor:"});
    t.insert("resistance", {"Resistance:", QString::fromUtf8("Resistance (Ω):"), "Resistance (Ohms):"});
    t.insert("resistance_initial", {"Ri (Ohms)", "Ri"});
    t.insert("resistance_final", {"Rf (Ohms)", "Rf"});
    t.insert("voltage", {"Voltage:"});
    t.insert("power", {"Power:"});
    t.insert("heating_technology", {"Heating Technology:", "Heater Technology", "Heater Technology:"});
    t.insert("media", {"Media:"});
    t.insert("viscosity", {"Viscosity:"});
    t.insert("initial_oil_mass", {"Initial Oil Mass:"});
    t.insert("fill_volume", {"Fill Volume:"});
    t.insert("puffing_regime", {"Puffing Regime:", "Puff Regime"});
    t.insert("cotton_length", {"Cotton length (if applicable)"});
    return t;
}

} // namespace

const QVector<MetricDef>& MetricRegistry::allMetrics()
{
    static const QVector<MetricDef> v = buildMetrics();
    return v;
}

const QVector<HeaderFieldDef>& MetricRegistry::allHeaderFields()
{
    static const QVector<HeaderFieldDef> v = buildHeaderFields();
    return v;
}

const MetricDef* MetricRegistry::metric(const QString& key)
{
    for (const MetricDef& d : allMetrics())
        if (d.key == key) return &d;
    return nullptr;
}

const HeaderFieldDef* MetricRegistry::headerField(const QString& key)
{
    for (const HeaderFieldDef& d : allHeaderFields())
        if (d.key == key) return &d;
    return nullptr;
}

const MetricDef* MetricRegistry::metricByAlias(const QString& normalized)
{
    static const QHash<QString, int> index = [] {
        QHash<QString, int> ix;
        const QVector<MetricDef>& all = allMetrics();
        for (int i = 0; i < all.size(); ++i) {
            QStringList names{all[i].displayName, all[i].key};
            names += all[i].headerAliases;
            for (const QString& n : names) {
                const QString norm = SchemaDrivenReader::normalizeHeader(n);
                if (!norm.isEmpty()) ix.insert(norm, i);
            }
        }
        return ix;
    }();
    const auto it = index.constFind(normalized);
    return it == index.constEnd() ? nullptr : &allMetrics()[it.value()];
}

QStringList MetricRegistry::headerAliasesFor(const QString& key)
{
    static const QHash<QString, QStringList> table = buildHeaderAliasTable();
    return table.value(key);
}

const HeaderFieldDef* MetricRegistry::headerByLabel(const QString& normalized)
{
    static const QHash<QString, int> index = [] {
        QHash<QString, int> ix;
        const QVector<HeaderFieldDef>& all = allHeaderFields();
        for (int i = 0; i < all.size(); ++i) {
            QStringList names{all[i].displayName, all[i].key};
            names += headerAliasesFor(all[i].key);
            for (const QString& n : names) {
                const QString norm = SchemaDrivenReader::normalizeHeader(n);
                if (!norm.isEmpty()) ix.insert(norm, i);
            }
        }
        return ix;
    }();
    const auto it = index.constFind(normalized);
    return it == index.constEnd() ? nullptr : &allHeaderFields()[it.value()];
}

MetricRegistry::PerPuffAlias MetricRegistry::perPuffAlias(const QString& normalized)
{
    PerPuffAlias hit;
    if (normalized.size() == 3 && normalized.startsWith(QLatin1String("pv"))) {
        const int n = normalized.mid(2).toInt();
        if (n >= 1 && n <= 5) {
            hit.targetKey = QStringLiteral("draw_pressure_per_puff");
            hit.index = n;
        }
    }
    return hit;
}

}} // namespace DVE::model
```

Register the new files in `DataViewerEnterprise.pro` and `tests/tst_v3model/tst_v3model.pro`.

- [ ] **Step 4: Run to verify it passes**

tst_v3model inner loop; expect PASS.
Then build `tst_v3inference` and `tst_v3shadow` too - if their `.pro` files list model sources explicitly they now need `MetricRegistry.cpp`/`RegimeParser.cpp` added (link errors tell you); fix and rerun.
Collision-test note: if `registryNamespacesAreCollisionFree` fails, an alias is registered twice - fix the DATA (that is the test doing its job protecting naming-policy rule 1).

- [ ] **Step 5: Commit**

```bash
git add src/model/MetricDef.h src/model/MetricRegistry.h src/model/MetricRegistry.cpp DataViewerEnterprise.pro tests/
git commit -m "feat(v3): MetricRegistry - ratified vocabulary compiled in + ValueType Bool/Mixed/NumberList/Image + MetricDef::tags"
```

---

### Task 3: StandardSchema sources its defs from the registry (zero behavior change)

**Files:**
- Modify: `src/model/StandardSchema.cpp`
- Test: existing `tests/tst_v3model` + `tests/tst_v3shadow` (the gates; one new test)

- [ ] **Step 1: Write the failing test** (append to `tst_v3model.cpp`)

```cpp
void TestV3Model::standardSchemaDrawsFromRegistry()
{
    const TemplateSchema s = standardV1(false);
    // Column KEY ORDER is the byte-identity contract - it must never change here.
    const QStringList expectedKeys{
        "puffs", "before_weight", "after_weight", "draw_pressure", "resistance",
        "smell", "clog", "notes", "tpm", "tpm_power_density", "variation_tpm", "oil_consumed"};
    QCOMPARE(s.columns.size(), expectedKeys.size());
    for (int i = 0; i < expectedKeys.size(); ++i)
        QCOMPARE(s.columns[i].key, expectedKeys[i]);
    // Registry aliases flow through (superset of the old hand-maintained sets).
    QVERIFY(s.columns[1].headerAliases.contains(QStringLiteral("Before weight/g")));
    // Layout-side flags still applied by the builder.
    QVERIFY(s.columns[8].plottable);                       // tpm
    QVERIFY(s.column("smell") && s.column("smell")->editable);
}
```

- [ ] **Step 2: Run to verify it fails**

Expected: FAIL on the alias assertion ("Before weight/g" is currently only in SchemaInference's local copies, not standardV1's).

- [ ] **Step 3: Refactor StandardSchema.cpp**

Replace the local `col(...)` data calls with registry copies.
Delete the local `col()` helper; keep `hf()`/`agg()`.
The column section of `standardV1` becomes:

```cpp
#include "MetricRegistry.h"   // add at top

namespace {
// Copy a registry def and apply the standard layout's presentation flags
// (flags are layout policy, not vocabulary - same rules as before).
MetricDef col(const QString& key)
{
    const MetricDef* d = MetricRegistry::metric(key);
    Q_ASSERT(d);
    MetricDef m = *d;
    m.editable  = (m.role == Role::Qualitative);
    m.plottable = (m.key == QLatin1String("tpm"));
    return m;
}
} // namespace
```

and the column list:

```cpp
    s.columns.append(col(QStringLiteral("puffs")));
    s.columns.append(col(QStringLiteral("before_weight")));
    s.columns.append(col(QStringLiteral("after_weight")));
    s.columns.append(col(QStringLiteral("draw_pressure")));
    s.columns.append(col(perRowRegime ? QStringLiteral("puffing_regime")
                                      : QStringLiteral("resistance")));
    s.columns.append(col(QStringLiteral("smell")));
    s.columns.append(col(QStringLiteral("clog")));
    s.columns.append(col(QStringLiteral("notes")));
    s.columns.append(col(QStringLiteral("tpm")));
    s.columns.append(col(QStringLiteral("tpm_power_density")));
    s.columns.append(col(QStringLiteral("variation_tpm")));
    s.columns.append(col(QStringLiteral("oil_consumed")));
```

Leave the header-field layout blocks (Cart/Project/Standard `hf(...)` calls with rows/cols) EXACTLY as they are - positions are layout data, not vocabulary - and leave the aggregates section untouched.
NOTE: existing tst_v3model assertions that pin old display strings (e.g. `"TPM/PD"` as displayName or exact alias-list contents) may now fail; update those assertions to the registry truth ("TPM Power Density" display, "TPM/PD" now an alias). Key order, roles, units of the 12 slots must NOT change - if a KEY assertion fails, the refactor is wrong, not the test.

- [ ] **Step 4: Run the gates**

1. tst_v3model inner loop - all PASS (after any display-string assertion updates).
2. `tests/tst_v3shadow` inner loop - expect 21 passed / 0 failed / 3 skipped (fixtures-only) - byte identity holds.
3. If shadow fails: STOP and diff - a registry def diverged from what standardV1 shipped (unit, role, or order). Fix the registry/refactor, never the shadow harness.

- [ ] **Step 5: Commit**

```bash
git add src/model/StandardSchema.cpp tests/tst_v3model
git commit -m "refactor(v3): StandardSchema columns sourced from MetricRegistry - single vocabulary source, shadow-gate green"
```

---

### Task 4: SchemaInference knowledge base = the registry

**Files:**
- Modify: `src/model/SchemaInference.cpp` (functions `knownMetricsKnowledgeBase`, `labelAliasTable`)
- Test: `tests/tst_v3inference/tst_v3inference.cpp`

- [ ] **Step 1: Write the failing tests** (append to `tst_v3inference.cpp`, matching its existing grid-building helpers)

```cpp
void TestV3Inference::registryLabelsRecognized()
{
    // Cart-era design-spec labels have NO trailing ':' and were previously
    // skipped as plain text; the registry now knows them (D6).
    // Build a minimal 13-wide Cart-style grid the way the existing S26-shape
    // tests in this suite do, with row 1 containing "Coil Material" at c3 and
    // a value at c4, and "Did this burn?" at c11 with "no" at c12.
    QVector<QVector<QVariant>> cells = makeS26StyleGrid();   // existing helper in this suite
    cells[0][2] = QStringLiteral("Coil Material");
    cells[0][3] = QStringLiteral("SS316");
    const TemplateSchema s = SchemaInference::inferSchema(cells, QStringLiteral("t"));
    const HeaderFieldDef* cm = s.headerField(QStringLiteral("coil_material"));
    QVERIFY(cm);
    QCOMPARE(cm->displayName, QStringLiteral("Coil Material"));
    // "Did this burn?" previously landed under snake key did_this_burn; the
    // registry canonicalizes it (D8).
    QVERIFY(s.headerField(QStringLiteral("did_burn")));
    QVERIFY(!s.headerField(QStringLiteral("did_this_burn")));
}

void TestV3Inference::registryRfLabelCanonicalized()
{
    // "Rf (Ohms)" used to map to rf_ohms; ratified key is resistance_final (D11).
    QVector<QVector<QVariant>> cells = makeS26StyleGrid();
    const TemplateSchema s = SchemaInference::inferSchema(cells, QStringLiteral("t"));
    QVERIFY(s.headerField(QStringLiteral("resistance_final")));
    QVERIFY(!s.headerField(QStringLiteral("rf_ohms")));
}
```

If `makeS26StyleGrid()` does not exist under that name, reuse whatever grid-builder the S26-shape tests in this suite use (the 13-col fixture builder) - read the suite first; do NOT duplicate a builder.

- [ ] **Step 2: Run to verify it fails**

tst_v3inference inner loop. Expected: FAIL - `coil_material` absent (bare label skipped), `did_this_burn`/`rf_ohms` keys present under the old table.

- [ ] **Step 3: Implement**

In `SchemaInference.cpp`:

Replace `knownMetricsKnowledgeBase()`'s body - the KB becomes the registry, minus nothing (its local alias copies and the perRowRegime special-casing disappear because the registry carries both `resistance` and `puffing_regime` as first-class metrics):

```cpp
const QVector<MetricDef>& knownMetricsKnowledgeBase()
{
    return MetricRegistry::allMetrics();
}
```

Replace `labelAliasTable()` with a registry-derived map (keep the same signature so call sites are untouched):

```cpp
const QMap<QString, QString>& labelAliasTable()
{
    static const QMap<QString, QString> table = [] {
        QMap<QString, QString> t;
        for (const HeaderFieldDef& hf : MetricRegistry::allHeaderFields()) {
            QStringList names{hf.displayName, hf.key};
            names += MetricRegistry::headerAliasesFor(hf.key);
            for (const QString& n : names) {
                const QString norm = SchemaDrivenReader::normalizeHeader(n);
                if (!norm.isEmpty()) t.insert(norm, hf.key);
            }
        }
        return t;
    }();
    return table;
}
```

Add `#include "MetricRegistry.h"` at the top.
Delete the now-unused local alias-addition lambda contents if the compiler flags unused code (-Werror).
BEHAVIOR NOTE (intended, document in the commit): inferred sheets now canonicalize labels the old table missed or misnamed - `did_this_burn` -> `did_burn`, `rf_ohms` -> `resistance_final`, design-spec labels become fields. Recovery-JSON keys for inferred files change accordingly; v2.10.2 is internal-only and recovery snapshots are transient, so no migration is needed.

- [ ] **Step 4: Run the gates**

1. tst_v3inference - all PASS (existing tests that assert `rf_ohms`/`did_this_burn` keys, if any, get updated to the canonical keys - that rename is the point of the task).
2. tst_v3shadow - 21/0/3 unchanged (standard path untouched; `standardFits`'s first-3 check reads standardV1 aliases which only grew).
3. tst_v3model - PASS.

- [ ] **Step 5: Commit**

```bash
git add src/model/SchemaInference.cpp tests/tst_v3inference
git commit -m "feat(v3): SchemaInference knowledge base = MetricRegistry - design-spec labels, status Q&A, resistance trio canonicalized"
```

---

### Task 5: PV1-PV5 assemble into draw_pressure_per_puff (list envelope)

**Files:**
- Modify: `src/pipeline/ReportDataJson.cpp` (`extraValueToJson` / `extraValueFromJson`)
- Modify: `src/model/LegacyAdapter.cpp` (`lowerInferredSheet`)
- Test: `tests/tst_reportdatajson/tst_reportdatajson.cpp`, `tests/tst_v3inference/tst_v3inference.cpp`

- [ ] **Step 1: Write the failing envelope test** (append to `tst_reportdatajson.cpp`)

```cpp
void TestReportDataJson::extraEnvelopeRoundTripsNumberList()
{
    DataRow r;
    r.extra.insert(QStringLiteral("draw_pressure_per_puff"),
                   QVariantList{1.1, 2.2, 3.3, 4.4, 5.5});
    FileResult f = makeMinimalFileResult(r);   // reuse this suite's existing builder for a 1-row file
    const QJsonObject j = fileResultToJson(f);
    const FileResult back = fileResultFromJson(j);
    const QVariant v = back.sheets[0].samples[0].rows[0].extra.value(QStringLiteral("draw_pressure_per_puff"));
    QCOMPARE(v.typeId(), QMetaType::QVariantList);
    const QVariantList l = v.toList();
    QCOMPARE(l.size(), 5);
    QCOMPARE(l[2].toDouble(), 3.3);
}
```

(Reuse the suite's existing minimal-FileResult builder; if it is inline in each test, follow that same inline pattern rather than inventing a helper.)

- [ ] **Step 2: Run to verify it fails**

tst_reportdatajson inner loop. Expected: FAIL - the list coerces to a string under the current `default:` branch (`typeId` comparison fails).

- [ ] **Step 3: Extend the envelope**

In `extraValueToJson`, add before the `default:` case:

```cpp
    case QMetaType::QVariantList: {
        // {"a": [...]} - numeric list (draw_pressure_per_puff). Elements are
        // doubles by contract (registry 8.2); anything else coerces toDouble().
        QJsonArray arr;
        for (const QVariant& e : v.toList())
            arr.append(e.toDouble());
        o["a"] = arr;
        break;
    }
```

In `extraValueFromJson`, add before the `"i"` check:

```cpp
    if (o.contains("a")) {
        QVariantList l;
        for (const QJsonValue& e : o["a"].toArray())
            l.append(e.toDouble());
        return QVariant(l);
    }
```

Add `#include <QJsonArray>` if not already present.
Also update the envelope comment block above `extraValueToJson` to list the new `{"a": [doubles]}` form.

- [ ] **Step 4: Run to verify green, then write the failing lowering test** (append to `tst_v3inference.cpp`)

```cpp
void TestV3Inference::pvColumnsAssemblePerPuffList()
{
    // 13-col Cart-era grid: PV1..PV5 at cols 4-8 with per-row values.
    QVector<QVector<QVariant>> cells = makeS26StyleGrid();
    // Row 5 (first data row): set PV values 10..14.
    for (int i = 0; i < 5; ++i)
        cells[4][3 + i] = 10.0 + i;
    const TemplateSchema schema = SchemaInference::inferSchema(cells, QStringLiteral("t"));
    const Sheet sheet = SchemaDrivenReader::parseSheet(cells, QStringLiteral("t"), schema,
                                                       false, ColumnResolution::NameFirst);
    const SheetResult sr = LegacyAdapter::lowerInferredSheet(sheet, QStringLiteral("t"),
                                                             QStringLiteral("old"));
    const DataRow& row = sr.samples[0].rows[0];
    QVERIFY(!row.extra.contains(QStringLiteral("pv1")));   // individual pvN keys replaced
    const QVariant v = row.extra.value(QStringLiteral("draw_pressure_per_puff"));
    QCOMPARE(v.typeId(), QMetaType::QVariantList);
    const QVariantList l = v.toList();
    QCOMPARE(l.size(), 5);
    QCOMPARE(l[0].toDouble(), 10.0);
    QCOMPARE(l[4].toDouble(), 14.0);
}
```

(Match the parseSheet/lowerInferredSheet call pattern already used by this suite's lowering tests - read one first and mirror its argument style exactly.)

- [ ] **Step 5: Run to verify it fails, then implement the assembly**

In `LegacyAdapter::lowerInferredSheet`, where unknown per-row metrics are copied into `DataRow::extra` (the loop that inserts by metric key), add PV collection: consult `MetricRegistry::perPuffAlias(SchemaDrivenReader::normalizeHeader(metricKey))` for each extra-bound metric; hits are buffered per row into a fixed-size `QVector<QVariant>(5)` at `index-1` instead of being inserted; after the loop, if any hit occurred, trim unused trailing slots and insert the QVariantList under `draw_pressure_per_puff`:

```cpp
        // PV1..PV5 -> draw_pressure_per_puff list (registry D5/8.2). Buffer
        // per-puff aliases by element index; everything else rides through
        // under its own key as before.
        QVector<QVariant> perPuff(5);
        bool anyPerPuff = false;
        // ...inside the existing per-metric loop, replacing the plain insert
        // for keys where perPuffAlias().index > 0:
        //     perPuff[hit.index - 1] = value; anyPerPuff = true;
        // ...after the loop:
        if (anyPerPuff) {
            int last = perPuff.size();
            while (last > 0 && !perPuff[last - 1].isValid()) --last;
            QVariantList list;
            for (int i = 0; i < last; ++i)
                list.append(perPuff[i].isValid() ? perPuff[i] : QVariant(0.0));
            row.extra.insert(QStringLiteral("draw_pressure_per_puff"), list);
        }
```

Adapt the sketch to the function's actual loop variables (read the function first - it exists at `src/model/LegacyAdapter.cpp`, the extra-insertion happens once for per-row metrics; keep the surrounding coercion rules untouched).
Add `#include "MetricRegistry.h"`.

- [ ] **Step 6: Run the gates**

1. tst_v3inference - all PASS, including the existing `realS26SampleIdentity` corpus-conditional test; if it asserts pv1..pv5 keys in extra, update it to assert the assembled list instead (S26 block 1 has empty PV data cells in the corpus copy - the test's existing expectations tell you which samples carry values).
2. tst_reportdatajson - PASS.
3. tst_v3shadow - 21/0/3 (standard path untouched).

- [ ] **Step 7: Commit**

```bash
git add src/pipeline/ReportDataJson.cpp src/model/LegacyAdapter.cpp tests/tst_reportdatajson tests/tst_v3inference
git commit -m "feat(v3): PV1-PV5 lower into draw_pressure_per_puff NumberList + {\"a\"} extra envelope"
```

---

### Task 6: Full gates + docs wrap

**Files:**
- Modify: `docs/superpowers/specs/2026-07-27-tpm-v3-vocabulary-registry-draft.md` (workshop log), `docs/sprint-tracker.html`

- [ ] **Step 1: Full-suite gate**

`powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1` - every suite PASS (DB suites skip cleanly if `DVE_TEST_PG_CONN` is unset; if the dve-test-pg container is up they must PASS).

- [ ] **Step 2: -Werror app build**

From repo root (after `python tools/decrypt_via_copy.py --apply`):

```bat
cd build
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
C:\Qt\6.10.1\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro
mingw32-make -j8
```

Expected: clean build, zero warnings.

- [ ] **Step 3: Corpus shadow run (real files, optional but preferred)**

From `tests/tst_v3shadow` with `DVE_TEST_CORPUS_DIR` set to the corpus dir: expect T58G identity PASS, S26/CPS inference-skips - same skip pattern as the v2.10.2 baseline (25 passed / 5 skipped).

- [ ] **Step 4: Docs**

Append one line to the registry doc's Workshop log: "Phase 2a: registry compiled into `src/model/MetricRegistry` (single source); inference canonicalized; PV lowering live."
Update the sprint tracker per its skill (task statuses + lastUpdated + new commits).

- [ ] **Step 5: Commit**

```bash
git add docs/
git commit -m "docs(v3): Phase 2a wrap - registry compiled, gates green"
```

---

## Self-review (done at authoring time)

- Spec coverage: registry doc sections 2/3 (Tasks 2-4), 8.2 (Tasks 1, 2, 5), 9.1 alias split (Task 2 test rows), 9.2 (Task 1 + non-goal), 5/D-rulings (Task 2 collision test + canonicalization in Task 4). Sections NOT covered here by design: 8.1 identity seed (Phase 3 DB natural key), 8.3 UI system (Phase 4/5), manifest/write-back (2b/2c).
- Placeholders: Task 5 Step 5 intentionally shows a sketch adapted to the real loop (the function's internals are not reproduced here); the executor MUST read `lowerInferredSheet` first - its insertion loop is short and the sketch names the exact seam. All other code is complete.
- Type consistency: `MetricRegistry::PerPuffAlias{targetKey,index}`, `RegimeParts` fields, `ValueType::{Bool,Mixed,NumberList,Image}` used consistently across tasks.
