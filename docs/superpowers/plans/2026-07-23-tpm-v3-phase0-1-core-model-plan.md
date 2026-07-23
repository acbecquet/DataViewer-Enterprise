# TPM v3 Phase 0+1 (Corpus Harness + Core Metric Model) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the v3 free-standing metric model (`Sample = headers map + metric->series`) and a schema-driven read path that replaces `ExcelReader`'s positional extraction on the production path with zero behavior change, gated by a shadow-parse harness over a golden-file corpus.

**Architecture:** Strangler-fig: a new `src/model/` core (MetricDef / TemplateSchema / SchemaDrivenReader) parses worksheet grids by schema, then a thin `LegacyAdapter` lowers the model back to `ExcelReader::SampleData` so the UNCHANGED `SheetProcessor` chain computes all metrics - output is byte-identical by construction and verified by a JSON shadow-diff harness.
Spec: `docs/superpowers/specs/2026-07-16-tpm-template-v3-metric-model-design.md` (sections 4, 6, 7 govern; CalculatorRegistry and manifests are Phase 2, NOT here).

**Tech Stack:** C++17 / Qt 6.10 (qmake + MinGW), Qt Test, existing openpyxl-subprocess ExcelReader, existing `fileResultToJson` canonical serializer.

---

## Machine + repo rules (read first)

- This machine (S1134987) MIP-labels files written by trusted Python; the Write/Edit tools do NOT label. **Create all new source files with the Write tool** - never the python delete-and-rewrite pattern from the project CLAUDE.md (that guidance is for the other machine; see the `mip-labels-python-writes` memory).
- If any existing file reads as ciphertext (`%TSD-Header-###%`), read it via `git show HEAD:<path>` instead, and run `python tools/decrypt_via_copy.py --apply` from the repo root before any build.
- The repo is PUBLIC. Never commit real test workbooks (corpus files stay outside the repo or gitignored). Fixtures under `tests/data/` are synthetic and fine.
- Work on branch `worktree-tpm-template-v3-research` (already rebased onto v2.10.0 main). Commit after every task; no co-author lines in commit messages.

## Build & test commands (referenced by all tasks)

Single suite inner loop (from the suite dir, e.g. `tests/tst_v3model`):

```bat
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
C:\Qt\6.10.1\mingw_64\bin\qmake.exe
mingw32-make
release\<suite>.exe
```

(If the binary lands in `debug\`, run that one - matches how the suite was configured.)

Full-suite gate (from repo root): `powershell -ExecutionPolicy Bypass -File tests\run-tests.ps1` - expect every suite PASS.

New suites must be added to `tests/tests.pro` `SUBDIRS` in the task that creates them.
New app sources must be added to `DataViewerEnterprise.pro` (`SOURCES`/`HEADERS`) in the task that creates them.

---

## Phase 0 - Corpus + harness

### Task 1: Corpus enumeration utility

**Files:**
- Create: `tests/common/CorpusUtils.h`, `tests/common/CorpusUtils.cpp`
- Create: `tests/tst_v3harness/tst_v3harness.pro`, `tests/tst_v3harness/tst_v3harness.cpp`
- Modify: `tests/tests.pro` (add `tst_v3harness` to SUBDIRS)

- [ ] **Step 1: Write the failing test**

`tests/tst_v3harness/tst_v3harness.cpp`:

```cpp
#include <QtTest>
#include <QTemporaryDir>
#include "CorpusUtils.h"

class TestV3Harness : public QObject {
    Q_OBJECT
private slots:
    void corpusIncludesFixtures();
    void corpusIncludesEnvDir();
    void corpusSkipsUnsetEnvCleanly();
};

void TestV3Harness::corpusIncludesFixtures()
{
    qunsetenv("DVE_TEST_CORPUS_DIR");
    const QStringList files = DVE::testutil::corpusFiles();
    QVERIFY(!files.isEmpty());                       // tests/data/*.xlsx always present
    for (const QString& f : files)
        QVERIFY2(f.endsWith(".xlsx"), qPrintable(f));
}

void TestV3Harness::corpusIncludesEnvDir()
{
    QTemporaryDir dir;
    QVERIFY(dir.isValid());
    QFile f(dir.filePath("real_test.xlsx"));
    QVERIFY(f.open(QIODevice::WriteOnly));
    f.write("stub");
    f.close();
    qputenv("DVE_TEST_CORPUS_DIR", dir.path().toUtf8());
    const QStringList files = DVE::testutil::corpusFiles();
    QVERIFY(files.contains(QDir(dir.path()).filePath("real_test.xlsx")));
    qunsetenv("DVE_TEST_CORPUS_DIR");
}

void TestV3Harness::corpusSkipsUnsetEnvCleanly()
{
    qunsetenv("DVE_TEST_CORPUS_DIR");
    QVERIFY(DVE::testutil::corpusDirDescription().contains("fixtures only"));
}

QTEST_MAIN(TestV3Harness)
#include "tst_v3harness.moc"
```

`tests/tst_v3harness/tst_v3harness.pro` (mirror the pattern of `tests/tst_tpmcalculator/tst_tpmcalculator.pro`):

```pro
QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
DEFINES += SRCDIR=\\\"$$PWD\\\"
INCLUDEPATH += ../../src ../../src/pipeline ../common
SOURCES += tst_v3harness.cpp ../common/CorpusUtils.cpp
HEADERS += ../common/CorpusUtils.h
```

- [ ] **Step 2: Run to verify it fails to build** (CorpusUtils.h missing) - per §Build commands.

- [ ] **Step 3: Implement CorpusUtils**

`tests/common/CorpusUtils.h`:

```cpp
#pragma once
#include <QStringList>

namespace DVE { namespace testutil {

// Workbooks for the v3 harnesses: the in-repo synthetic fixtures
// (tests/data/*.xlsx, generated by tests/generate_fixtures.py) plus every
// .xlsx under $DVE_TEST_CORPUS_DIR (recursive) when that env var is set.
// An unset/missing corpus dir is NOT an error - harnesses run on fixtures
// alone (same skip-clean philosophy as DVE_TEST_PG_CONN).
QStringList corpusFiles();
QString     corpusDirDescription();   // "fixtures only" or "<dir> (+N files)"

}} // namespace DVE::testutil
```

`tests/common/CorpusUtils.cpp`:

```cpp
#include "CorpusUtils.h"
#include <QDir>
#include <QDirIterator>
#include <QFileInfo>

namespace DVE { namespace testutil {

static QString fixturesDir()
{
    // SRCDIR is <repo>/tests/tst_<suite>; fixtures live at <repo>/tests/data.
    return QFileInfo(QStringLiteral(SRCDIR) + "/../data").absoluteFilePath();
}

QStringList corpusFiles()
{
    QStringList out;
    QDirIterator fix(fixturesDir(), {QStringLiteral("*.xlsx")}, QDir::Files);
    while (fix.hasNext()) out << fix.next();

    const QByteArray env = qgetenv("DVE_TEST_CORPUS_DIR");
    if (!env.isEmpty() && QDir(QString::fromUtf8(env)).exists()) {
        QDirIterator it(QString::fromUtf8(env), {QStringLiteral("*.xlsx")},
                        QDir::Files, QDirIterator::Subdirectories);
        while (it.hasNext()) out << it.next();
    }
    out.removeDuplicates();
    out.sort();
    return out;
}

QString corpusDirDescription()
{
    const QByteArray env = qgetenv("DVE_TEST_CORPUS_DIR");
    if (env.isEmpty() || !QDir(QString::fromUtf8(env)).exists())
        return QStringLiteral("fixtures only");
    return QString::fromUtf8(env);
}

}} // namespace DVE::testutil
```

- [ ] **Step 4: Add `tst_v3harness` to `tests/tests.pro` SUBDIRS, build, run - all 3 tests PASS.**
  If `tests/data/*.xlsx` are missing locally, run `python tests/generate_fixtures.py` once first.

- [ ] **Step 5: Commit**

```bash
git add tests/common/CorpusUtils.h tests/common/CorpusUtils.cpp tests/tst_v3harness tests/tests.pro
git commit -m "test(v3): corpus enumeration utility + tst_v3harness suite (Phase 0)"
```

### Task 2: JSON deep-diff utility

**Files:**
- Create: `tests/common/JsonDiff.h`, `tests/common/JsonDiff.cpp`
- Modify: `tests/tst_v3harness/tst_v3harness.cpp` / `.pro`

- [ ] **Step 1: Add failing tests** to `tst_v3harness.cpp` (new slots + declarations):

```cpp
void TestV3Harness::jsonDiffEqual()
{
    QJsonObject a{{"x", 1}, {"nested", QJsonObject{{"y", "z"}}}};
    QCOMPARE(DVE::testutil::diffJson(a, a), QStringList());
}

void TestV3Harness::jsonDiffReportsPath()
{
    QJsonObject a{{"s", QJsonArray{QJsonObject{{"tpm", 3.5}}}}};
    QJsonObject b{{"s", QJsonArray{QJsonObject{{"tpm", 3.6}}}}};
    const QStringList d = DVE::testutil::diffJson(a, b);
    QCOMPARE(d.size(), 1);
    QVERIFY2(d.first().contains("/s[0]/tpm"), qPrintable(d.first()));
}

void TestV3Harness::jsonDiffReportsMissingKey()
{
    QJsonObject a{{"x", 1}};
    QJsonObject b{{"x", 1}, {"extra", 2}};
    QCOMPARE(DVE::testutil::diffJson(a, b).size(), 1);
}
```

- [ ] **Step 2: Build - fails (diffJson undefined).**

- [ ] **Step 3: Implement**

`tests/common/JsonDiff.h`:

```cpp
#pragma once
#include <QJsonValue>
#include <QStringList>

namespace DVE { namespace testutil {

// Deep-compare two JSON values. Returns one line per difference in the form
// "/path[idx]/key: <a> != <b>"; empty list means identical. Numbers compare
// exactly (both sides of every v3 diff come from the same serializer, so
// float formatting is deterministic; the Task 3 determinism test proves it).
QStringList diffJson(const QJsonValue& a, const QJsonValue& b,
                     const QString& path = QString());

}} // namespace DVE::testutil
```

`tests/common/JsonDiff.cpp`:

```cpp
#include "JsonDiff.h"
#include <QJsonArray>
#include <QJsonObject>

namespace DVE { namespace testutil {

static QString shortRepr(const QJsonValue& v)
{
    switch (v.type()) {
    case QJsonValue::Object: return QStringLiteral("{object}");
    case QJsonValue::Array:  return QStringLiteral("[array:%1]").arg(v.toArray().size());
    default:                 return v.toVariant().toString().left(60);
    }
}

QStringList diffJson(const QJsonValue& a, const QJsonValue& b, const QString& path)
{
    QStringList out;
    if (a.type() != b.type()) {
        out << QStringLiteral("%1: type %2 != %3").arg(path).arg(a.type()).arg(b.type());
        return out;
    }
    if (a.isObject()) {
        const QJsonObject oa = a.toObject(), ob = b.toObject();
        QStringList keys = oa.keys() + ob.keys();
        keys.removeDuplicates();
        keys.sort();
        for (const QString& k : keys) {
            if (!oa.contains(k)) { out << QStringLiteral("%1/%2: missing on left").arg(path, k);  continue; }
            if (!ob.contains(k)) { out << QStringLiteral("%1/%2: missing on right").arg(path, k); continue; }
            out += diffJson(oa.value(k), ob.value(k), path + '/' + k);
        }
        return out;
    }
    if (a.isArray()) {
        const QJsonArray aa = a.toArray(), ab = b.toArray();
        if (aa.size() != ab.size()) {
            out << QStringLiteral("%1: array size %2 != %3").arg(path).arg(aa.size()).arg(ab.size());
            return out;
        }
        for (int i = 0; i < aa.size(); ++i)
            out += diffJson(aa.at(i), ab.at(i), QStringLiteral("%1[%2]").arg(path).arg(i));
        return out;
    }
    if (a != b)
        out << QStringLiteral("%1: %2 != %3").arg(path, shortRepr(a), shortRepr(b));
    return out;
}

}} // namespace DVE::testutil
```

Add `../common/JsonDiff.cpp` to the suite `.pro` `SOURCES` (and the header to `HEADERS`).

- [ ] **Step 4: Build + run - all tests PASS.**

- [ ] **Step 5: Commit** - `test(v3): JSON deep-diff utility for shadow/round-trip harnesses (Phase 0)`

### Task 3: Shadow harness baseline - parser determinism

**Files:**
- Create: `tests/tst_v3shadow/tst_v3shadow.pro`, `tests/tst_v3shadow/tst_v3shadow.cpp`
- Modify: `tests/tests.pro`

This suite invokes the real Excel pipeline (Python/openpyxl subprocess), so mirror `tests/tst_dataprocessor/tst_dataprocessor.pro`'s `SOURCES`/`INCLUDEPATH` (it already links ExcelReader + DataProcessor + SheetProcessors + TpmCalculator + ReportDataJson; copy its list and add `../common/CorpusUtils.cpp ../common/JsonDiff.cpp`).

- [ ] **Step 1: Write the test**

```cpp
#include <QtTest>
#include "CorpusUtils.h"
#include "JsonDiff.h"
#include "pipeline/DataProcessor.h"
#include "pipeline/ReportDataJson.h"

class TestV3Shadow : public QObject {
    Q_OBJECT
private slots:
    void parseIsDeterministic_data();
    void parseIsDeterministic();
};

void TestV3Shadow::parseIsDeterministic_data()
{
    QTest::addColumn<QString>("path");
    const QStringList files = DVE::testutil::corpusFiles();
    QVERIFY2(!files.isEmpty(), "no corpus files; run tests/generate_fixtures.py");
    for (const QString& f : files)
        QTest::newRow(qPrintable(QFileInfo(f).fileName())) << f;
}

void TestV3Shadow::parseIsDeterministic()
{
    QFETCH(QString, path);
    DVE::DataProcessor p1, p2;
    const DVE::FileResult a = p1.processFile(path);
    if (a.sheets.isEmpty() && p1.lastError().contains("python", Qt::CaseInsensitive))
        QSKIP("bundled/system python unavailable");
    const DVE::FileResult b = p2.processFile(path);
    const QStringList diff = DVE::testutil::diffJson(DVE::fileResultToJson(a),
                                                     DVE::fileResultToJson(b));
    QVERIFY2(diff.isEmpty(), qPrintable(diff.join('\n')));
}

QTEST_MAIN(TestV3Shadow)
#include "tst_v3shadow.moc"
```

- [ ] **Step 2: Register in `tests/tests.pro`, build, run - PASS on every fixture row.**
  This proves the harness plumbing AND that `processFile` output is deterministic (the property every later shadow comparison rests on).

- [ ] **Step 3: Commit** - `test(v3): shadow harness baseline - parse determinism over corpus (Phase 0)`

### Task 4: Corpus documentation + gitignore guard

**Files:**
- Create: `tests/corpus/README.md`
- Modify: `.gitignore`

- [ ] **Step 1: Write `tests/corpus/README.md`:**

```markdown
# v3 Golden-File Corpus

Local-only stash of REAL historical workbooks used by tst_v3shadow (and, from
Phase 2, the round-trip harness). Point the harnesses at it with:

    $env:DVE_TEST_CORPUS_DIR = "<this directory or any dir of .xlsx>"

Populate it via the DB Data collection workflow (see the db-data-collection
memory topic / tools collect_db_data.py): source .xlsx resolved from the DB
files table under Weekly_Reports_Transfer. Live-DB queries need owner approval.

RULES
- This repo is PUBLIC: real workbooks must NEVER be committed. Everything in
  this directory except this README is gitignored.
- Harnesses run on the synthetic tests/data fixtures alone when
  DVE_TEST_CORPUS_DIR is unset - corpus presence widens coverage, never gates.
```

- [ ] **Step 2: Add to `.gitignore`:**

```
tests/corpus/*
!tests/corpus/README.md
```

- [ ] **Step 3: Verify `git status` shows only the README as addable; commit** - `docs(v3): corpus directory contract + public-repo gitignore guard (Phase 0)`

---

## Phase 1 - Core model + schema-driven read path

### Task 5: Model value types (`MetricDef`, `MetricSeries`, `Sample`)

**Files:**
- Create: `src/model/MetricDef.h`, `src/model/MetricSample.h` (both header-only)
- Create: `tests/tst_v3model/tst_v3model.pro`, `tests/tst_v3model/tst_v3model.cpp`
- Modify: `tests/tests.pro`, `DataViewerEnterprise.pro` (HEADERS)

- [ ] **Step 1: Write failing tests**

```cpp
#include <QtTest>
#include "model/MetricSample.h"

using namespace DVE::model;

class TestV3Model : public QObject {
    Q_OBJECT
private slots:
    void seriesLookup();
    void rowCountIsLongestSeries();
};

void TestV3Model::seriesLookup()
{
    Sample s;
    s.data.append(MetricSeries{QStringLiteral("puffs"), {10, 20}});
    s.data.append(MetricSeries{QStringLiteral("tpm"),   {3.5, 3.5}});
    QVERIFY(s.series("tpm") != nullptr);
    QCOMPARE(s.series("tpm")->values.size(), 2);
    QCOMPARE(s.series("nope"), nullptr);
}

void TestV3Model::rowCountIsLongestSeries()
{
    Sample s;
    s.data.append(MetricSeries{QStringLiteral("puffs"), {10, 20, 30}});
    s.data.append(MetricSeries{QStringLiteral("notes"), {QStringLiteral("a")}});
    QCOMPARE(s.rowCount(), 3);
}

QTEST_MAIN(TestV3Model)
#include "tst_v3model.moc"
```

`.pro`: header-only so just `INCLUDEPATH += ../../src` and the test cpp.

- [ ] **Step 2: Build - fails (headers missing).**

- [ ] **Step 3: Implement**

`src/model/MetricDef.h`:

```cpp
#pragma once
#include <QString>
#include <QStringList>

namespace DVE { namespace model {

enum class ValueType { Number, Text };
enum class Role { Measured, Qualitative, Derived, Identity };

// One free-standing metric (a data column, in template terms).
struct MetricDef {
    QString     key;             // stable snake_case id, e.g. "before_weight"
    QString     displayName;     // canonical header text
    QStringList headerAliases;   // extra header spellings matched on read
    ValueType   type = ValueType::Number;
    QString     unit;
    Role        role = Role::Measured;
    QString     calculator;      // Derived only; evaluated from Phase 2
    QStringList inputs;          // metric keys or "header:<key>"
    bool        plottable = false;
    bool        editable  = false;
    int         precision = 2;
};

// One header-band field (applies to the whole sample).
struct HeaderFieldDef {
    QString     key;
    QString     displayName;
    ValueType   type = ValueType::Text;
    QString     unit;
    int         row = 0;         // 1-based, block-relative template row
    int         col = 0;         // 1-based, block-relative template col
    QString     calculator;      // derived header values (e.g. power)
    QStringList inputs;
};

struct AggregateDef {
    QString     key;
    QString     calculator;
    QStringList inputs;
};

}} // namespace DVE::model
```

`src/model/MetricSample.h`:

```cpp
#pragma once
#include "MetricDef.h"
#include <QMap>
#include <QVariant>
#include <QVector>

namespace DVE { namespace model {

struct MetricSeries {
    QString           key;
    QVector<QVariant> values;
};

// The owner's canonical shape: headers apply to all data equally; data is a
// set of free-standing named series.
struct Sample {
    QMap<QString, QVariant> headers;
    QVector<MetricSeries>   data;      // schema column order
    qint64                  id = -1;
    int                     version = 0;

    const MetricSeries* series(const QString& key) const {
        for (const MetricSeries& m : data) if (m.key == key) return &m;
        return nullptr;
    }
    int rowCount() const {
        int n = 0;
        for (const MetricSeries& m : data) n = qMax(n, int(m.values.size()));
        return n;
    }
};

}} // namespace DVE::model
```

- [ ] **Step 4: Register suite in `tests/tests.pro`; add both headers to `DataViewerEnterprise.pro` HEADERS; build + run - PASS.**

- [ ] **Step 5: Commit** - `feat(v3): model value types - MetricDef/MetricSeries/Sample (Phase 1)`

### Task 6: `TemplateSchema` + builtin `standard-v1`

**Files:**
- Create: `src/model/TemplateSchema.h`, `src/model/TemplateSchema.cpp`
- Create: `src/model/StandardSchema.h`, `src/model/StandardSchema.cpp`
- Modify: `tests/tst_v3model/*` (tests + SOURCES), `DataViewerEnterprise.pro`

- [ ] **Step 1: Write failing tests** (add slots to `tst_v3model.cpp`):

```cpp
void TestV3Model::standardV1ColumnOrderMatchesCols()
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/false);
    const QStringList expected{"puffs","before_weight","after_weight","draw_pressure",
        "resistance","smell","clog","notes","tpm","tpm_power_density",
        "variation_tpm","oil_consumed"};
    QCOMPARE(s.columns.size(), 12);
    QCOMPARE(s.blockCols, 12);
    for (int i = 0; i < expected.size(); ++i)
        QCOMPARE(s.columns[i].key, expected[i]);
    QCOMPARE(s.columnPos("tpm"), 8);           // == DVE::Cols::TPM
    QCOMPARE(s.column("tpm")->role, Role::Derived);
    QCOMPARE(s.column("tpm")->calculator, QStringLiteral("tpm_v1"));
}

void TestV3Model::standardV1RegimeVariant()
{
    const TemplateSchema s = standardV1(/*perRowRegime=*/true);
    QCOMPARE(s.columns[4].key, QStringLiteral("puffing_regime"));
    QCOMPARE(s.columns[4].type, ValueType::Text);
    QCOMPARE(s.columns[4].role, Role::Qualitative);
}

void TestV3Model::standardV1HeaderCells()
{
    const TemplateSchema s = standardV1(false);
    const HeaderFieldDef* v = s.headerField("voltage");
    QVERIFY(v);
    QCOMPARE(v->row, 3); QCOMPARE(v->col, 6);
    const HeaderFieldDef* p = s.headerField("power");
    QVERIFY(p);
    QCOMPARE(p->calculator, QStringLiteral("power_v1"));
    QCOMPARE(s.headerFields.size(), 12);
}
```

- [ ] **Step 2: Build - fails.**

- [ ] **Step 3: Implement**

`src/model/TemplateSchema.h`:

```cpp
#pragma once
#include "MetricDef.h"
#include <QVector>

namespace DVE { namespace model {

struct TemplateSchema {
    QString schemaId;             // "standard"
    int     version = 1;

    // Block geometry, 1-based template rows (physical sheet coordinates).
    int headerRows      = 3;
    int columnHeaderRow = 4;
    int dataStartRow    = 5;
    int blockCols       = 12;

    QVector<HeaderFieldDef> headerFields;
    QVector<MetricDef>      columns;      // order == physical column order
    QVector<AggregateDef>   aggregates;

    const MetricDef*      column(const QString& key) const;
    int                   columnPos(const QString& key) const;  // 0-based, -1 absent
    const HeaderFieldDef* headerField(const QString& key) const;
};

}} // namespace DVE::model
```

`src/model/TemplateSchema.cpp`: three linear scans (column returns pointer into `columns`; columnPos returns index; headerField pointer into `headerFields`).

`src/model/StandardSchema.h`:

```cpp
#pragma once
#include "TemplateSchema.h"

namespace DVE { namespace model {

// Compiled-in schema reproducing the current standardized template exactly
// (spec section 6). perRowRegime selects the col-5 variant: the new-template
// per-row Puffing Regime text column vs the old-template Resistance column
// (same rule DataProcessor::processSheet applies via RegimeUtils today).
TemplateSchema standardV1(bool perRowRegime);

}} // namespace DVE::model
```

`src/model/StandardSchema.cpp` - transcribe spec section 6 exactly.
Pattern (abridged here only to avoid repeating the full 12+12 literal table twice in this plan; the spec tables ARE the content - copy every row):

```cpp
#include "StandardSchema.h"

namespace DVE { namespace model {

static MetricDef col(const char* key, const char* disp, ValueType t, const char* unit,
                     Role role, const char* calc = "", QStringList inputs = {},
                     QStringList aliases = {})
{
    MetricDef m;
    m.key = QLatin1String(key); m.displayName = QLatin1String(disp);
    m.type = t; m.unit = QLatin1String(unit); m.role = role;
    m.calculator = QLatin1String(calc); m.inputs = std::move(inputs);
    m.headerAliases = std::move(aliases);
    m.editable  = (role == Role::Qualitative);
    m.plottable = (m.key == QLatin1String("tpm"));
    return m;
}

static HeaderFieldDef hf(const char* key, const char* disp, int row, int colIdx,
                         ValueType t, const char* unit = "",
                         const char* calc = "", QStringList inputs = {})
{
    HeaderFieldDef h;
    h.key = QLatin1String(key); h.displayName = QLatin1String(disp);
    h.row = row; h.col = colIdx; h.type = t; h.unit = QLatin1String(unit);
    h.calculator = QLatin1String(calc); h.inputs = std::move(inputs);
    return h;
}

TemplateSchema standardV1(bool perRowRegime)
{
    TemplateSchema s;
    s.schemaId = QStringLiteral("standard");
    s.version  = 1;

    s.columns = {
        col("puffs",          "puffs",             ValueType::Number, "count",   Role::Measured),
        col("before_weight",  "Before Weight (g)", ValueType::Number, "g",       Role::Measured, "", {}, {"before(g)","before weight"}),
        col("after_weight",   "After Weight (g)",  ValueType::Number, "g",       Role::Measured, "", {}, {"after(g)","after weight"}),
        col("draw_pressure",  "Draw Pressure",     ValueType::Number, "kPa",     Role::Measured, "", {}, {"drawpress","draw pressure (kpa)"}),
        perRowRegime
          ? col("puffing_regime", "Puffing Regime", ValueType::Text,  "",        Role::Qualitative)
          : col("resistance",     "Resistance",     ValueType::Number,"ohm",     Role::Measured),
        col("smell",          "Smell",             ValueType::Text,   "",        Role::Qualitative),
        col("clog",           "Clog",              ValueType::Text,   "",        Role::Qualitative),
        col("notes",          "Notes",             ValueType::Text,   "",        Role::Qualitative),
        col("tpm",            "TPM",               ValueType::Number, "mg/puff", Role::Derived, "tpm_v1",           {"puffs","before_weight","after_weight"}),
        col("tpm_power_density","TPM/PD",          ValueType::Number, "mg/(puff*W)", Role::Derived, "power_density_v1", {"tpm","header:power"}),
        col("variation_tpm",  "Variation",         ValueType::Number, "%",       Role::Derived, "variation_v1",     {"tpm"}, {"var%"}),
        col("oil_consumed",   "Oil Consumed",      ValueType::Number, "mg",      Role::Derived, "oil_consumed_v1",  {"before_weight","after_weight"}, {"oilcum(g)","oil consumed (g)"}),
    };

    s.headerFields = {
        hf("test_name",         "Test Name",        1, 1, ValueType::Text),
        hf("date",              "Date",             1, 4, ValueType::Text),
        hf("sample_id",         "Sample ID",        1, 6, ValueType::Text),
        hf("heating_technology","Heating Technology",1, 8, ValueType::Text),
        hf("media",             "Media",            2, 2, ValueType::Text),
        hf("resistance",        "Resistance",       2, 4, ValueType::Number, "ohm"),
        hf("power",             "Power",            2, 6, ValueType::Number, "W", "power_v1", {"header:voltage","header:resistance","header:heating_technology"}),
        hf("puffing_regime",    "Puffing Regime",   2, 8, ValueType::Text),
        hf("viscosity",         "Viscosity",        3, 2, ValueType::Number, "cP"),
        hf("tester",            "Tester",           3, 4, ValueType::Text),
        hf("voltage",           "Voltage",          3, 6, ValueType::Number, "V"),
        hf("initial_oil_mass",  "Initial Oil Mass", 3, 8, ValueType::Number, "g"),
    };

    s.aggregates = {
        {"average_tpm",        "mean",          {"tpm"}},
        {"stddev_tpm",         "stddev",        {"tpm"}},
        {"avg_power_density",  "mean_over_power",{"tpm","header:power"}},
        {"normalized_tpm",     "mean_over_power",{"tpm","header:power"}},
        {"total_puffs",        "last",          {"puffs"}},
        {"total_oil_consumed", "last",          {"oil_consumed"}},
        {"efficiency_percent", "efficiency_v1", {"total_oil_consumed","header:initial_oil_mass"}},
        {"burn_clog_leak",     "status_scan_v1",{"smell","clog","notes"}},
    };
    return s;
}

}} // namespace DVE::model
```

**Verification sub-step:** before finalizing displayName/alias strings, print the REAL row-4 header texts from `tests/data` fixtures and from `resources/templates/Standardized Test Template - December 2025.xlsx` (a 5-line openpyxl print via `python -c`), and make each real string match either `displayName` or an alias (case/spacing-insensitively - Task 8 defines the normalizer).
Positional fallback covers mismatches, but Phase-1 name matching should actually engage on the standard template.

- [ ] **Step 4: Add sources to `DataViewerEnterprise.pro` + suite `.pro`; build + run - PASS.**

- [ ] **Step 5: Commit** - `feat(v3): TemplateSchema + builtin standard-v1 (both regime variants) (Phase 1)`

### Task 7: `ExcelReader::currentSheetCells()` accessor

**Files:**
- Modify: `src/ExcelReader.h`, `src/ExcelReader.cpp`
- Modify: `tests/tst_excelreader/tst_excelreader.cpp` (add one test)

- [ ] **Step 1: Add failing test** (in the existing suite, following its fixture-loading pattern):

```cpp
void TestExcelReader::currentSheetCellsExposesGrid()
{
    ExcelReader r;                       // load a fixture the suite already uses
    QVERIFY(r.loadFile(fixturePath("standard_new_template.xlsx"))); // reuse the suite's helper + fixture names
    QVERIFY(r.selectSheet(r.getSheetNames().first()));
    const auto cells = r.currentSheetCells();
    QVERIFY(!cells.isEmpty());
    QVERIFY(cells.size() >= 5);          // header band + col headers + data
    QCOMPARE(cells[3].isEmpty(), false); // row 4 (0-based 3) carries col headers
}
```

(Adapt the fixture name/helper to what `tst_excelreader` actually uses - read the suite first; its existing tests show the canonical way to load `tests/data` fixtures.)

- [ ] **Step 2: Build - fails. Step 3: Implement** - in `ExcelReader.h` public section:

```cpp
    // v3: raw typed grid of the current sheet (0-based [row][col]); empty if
    // no sheet selected. Read-only view for the schema-driven reader.
    QVector<QVector<QVariant>> currentSheetCells() const;
```

`ExcelReader.cpp`:

```cpp
QVector<QVector<QVariant>> ExcelReader::currentSheetCells() const
{
    const SheetData* sd = currentSheetData();
    return sd ? sd->cells : QVector<QVector<QVariant>>{};
}
```

- [ ] **Step 4: Build + run tst_excelreader - PASS. Step 5: Commit** - `feat(v3): ExcelReader::currentSheetCells() typed-grid accessor (Phase 1)`

### Task 8: `SchemaDrivenReader`

**Files:**
- Create: `src/model/SchemaDrivenReader.h`, `src/model/SchemaDrivenReader.cpp`
- Modify: `tests/tst_v3model/*`, `DataViewerEnterprise.pro`

- [ ] **Step 1: Write failing tests** - synthetic grids, no Excel involved:

```cpp
// Helper: build a standard-shaped grid for `blocks` samples, `rows` data rows.
static QVector<QVector<QVariant>> makeStandardGrid(int blocks, int rows,
                                                   bool swapSmellNotes = false)
{
    const TemplateSchema s = standardV1(false);
    QVector<QVector<QVariant>> g(4 + rows);
    for (int b = 0; b < blocks; ++b) {
        const int off = b * s.blockCols;
        auto set = [&](int r, int c, const QVariant& v) {
            if (g[r].size() < off + s.blockCols) g[r].resize(off + s.blockCols);
            g[r][off + c] = v;
        };
        // header band (values only where the schema looks)
        for (const HeaderFieldDef& h : s.headerFields)
            if (h.calculator.isEmpty())
                set(h.row - 1, h.col - 1, QStringLiteral("%1_%2").arg(h.key).arg(b));
        // column header row (row index 3)
        for (int c = 0; c < s.columns.size(); ++c) {
            int cc = c;
            if (swapSmellNotes && s.columns[c].key == "smell") cc = s.columnPos("notes");
            else if (swapSmellNotes && s.columns[c].key == "notes") cc = s.columnPos("smell");
            set(3, cc, s.columns[c].displayName);
        }
        // data rows
        for (int r = 0; r < rows; ++r) {
            set(4 + r, 0, (r + 1) * 10);            // puffs
            set(4 + r, 1, 25.0 - r * 0.03);         // before
            set(4 + r, 2, 25.0 - r * 0.03 - 0.035); // after
            int smellCol = swapSmellNotes ? s.columnPos("notes") : s.columnPos("smell");
            set(4 + r, smellCol, QStringLiteral("ok"));
        }
    }
    return g;
}

void TestV3Model::readerParsesTwoBlocks()
{
    const TemplateSchema s = standardV1(false);
    const Sheet sheet = SchemaDrivenReader::parseSheet(makeStandardGrid(2, 5), "Lifetime Test", s, false);
    QCOMPARE(sheet.samples.size(), 2);
    QCOMPARE(sheet.samples[0].rowCount(), 5);
    QCOMPARE(sheet.samples[0].headers.value("tester").toString(), QStringLiteral("tester_0"));
    QCOMPARE(sheet.samples[1].headers.value("tester").toString(), QStringLiteral("tester_1"));
    QCOMPARE(sheet.samples[0].series("puffs")->values[2].toInt(), 30);
}

void TestV3Model::readerMatchesReorderedColumnsByName()
{
    const TemplateSchema s = standardV1(false);
    const Sheet sheet = SchemaDrivenReader::parseSheet(makeStandardGrid(1, 3, /*swap*/true), "Lifetime Test", s, false);
    // smell data was PHYSICALLY written into the notes position, but the header
    // text moved with it - name matching must still key it as "smell".
    QCOMPARE(sheet.samples[0].series("smell")->values[0].toString(), QStringLiteral("ok"));
}

void TestV3Model::readerFallsBackPositionally()
{
    const TemplateSchema s = standardV1(false);
    auto g = makeStandardGrid(1, 3);
    g[3].fill(QVariant(), g[3].size());   // wipe every column header
    const Sheet sheet = SchemaDrivenReader::parseSheet(g, "Lifetime Test", s, false);
    QCOMPARE(sheet.samples.size(), 1);
    QCOMPARE(sheet.samples[0].series("puffs")->values.size(), 3);
}

void TestV3Model::readerStopsAtAllEmptyRow()
{
    auto g = makeStandardGrid(1, 3);
    g.append(QVector<QVariant>());                 // blank row
    g.append(QVector<QVariant>{QVariant(999)});    // stray data AFTER blank - ignored
    const Sheet sheet = SchemaDrivenReader::parseSheet(g, "Lifetime Test", standardV1(false), false);
    QCOMPARE(sheet.samples[0].rowCount(), 3);
}

void TestV3Model::normalizeHeader()
{
    QCOMPARE(SchemaDrivenReader::normalizeHeader(" Before Weight (g) "), QStringLiteral("beforeweightg"));
}
```

- [ ] **Step 2: Build - fails. Step 3: Implement**

`src/model/SchemaDrivenReader.h`:

```cpp
#pragma once
#include "MetricSample.h"
#include "TemplateSchema.h"

namespace DVE { namespace model {

struct Sheet {
    QString         sheetName;
    TemplateSchema  schema;
    bool            perRowRegime = false;
    QVector<Sample> samples;
};

struct File {
    QString        filePath;
    QString        fileName;
    QString        templateVersion;   // legacy "new"/"old" tag, adapter input
    QVector<Sheet> sheets;
};

// Parses one worksheet grid (0-based [row][col] QVariant cells, exactly as
// ExcelReader stores them) into a model::Sheet using a TemplateSchema.
// Pure function of its inputs - no I/O, fully unit-testable.
class SchemaDrivenReader {
public:
    static Sheet   parseSheet(const QVector<QVector<QVariant>>& cells,
                              const QString& sheetName,
                              const TemplateSchema& schema,
                              bool perRowRegime);
    // lowercase, [a-z0-9] only - the name-matching normalizer.
    static QString normalizeHeader(const QString& s);
};

}} // namespace DVE::model
```

(`Sheet`/`File` live here so Task 5's header stays dependency-light; keep them in this header, not MetricSample.h.)

`src/model/SchemaDrivenReader.cpp` core:

```cpp
#include "SchemaDrivenReader.h"

namespace DVE { namespace model {

QString SchemaDrivenReader::normalizeHeader(const QString& s)
{
    QString out;
    for (const QChar c : s.toLower())
        if (c.isLetterOrNumber()) out.append(c);
    return out;
}

static QVariant cellAt(const QVector<QVector<QVariant>>& g, int r, int c)
{
    if (r < 0 || r >= g.size() || c < 0 || c >= g[r].size()) return {};
    return g[r][c];
}

static bool cellEmpty(const QVariant& v)
{
    return !v.isValid() || v.toString().trimmed().isEmpty();
}

// Per metric: the ABSOLUTE grid column it reads from, name-first then position.
static QVector<int> resolveColumns(const QVector<QVector<QVariant>>& g,
                                   const TemplateSchema& s, int off)
{
    const int hdrRow = s.columnHeaderRow - 1;
    QVector<int> map(s.columns.size());
    QVector<bool> taken(s.blockCols, false);
    // pass 1: name matches
    QVector<int> unresolved;
    for (int i = 0; i < s.columns.size(); ++i) {
        const MetricDef& m = s.columns[i];
        QStringList wanted{normalizeHeader(m.displayName),
                           normalizeHeader(m.key)};
        for (const QString& a : m.headerAliases) wanted << normalizeHeader(a);
        int found = -1;
        for (int c = 0; c < s.blockCols; ++c) {
            if (taken[c]) continue;
            const QString h = normalizeHeader(cellAt(g, hdrRow, off + c).toString());
            if (!h.isEmpty() && wanted.contains(h)) { found = c; break; }
        }
        if (found >= 0) { map[i] = off + found; taken[found] = true; }
        else            { map[i] = -1; unresolved << i; }
    }
    // pass 2: positional fallback for the rest (schema position if free)
    for (int i : unresolved) {
        int c = i;                       // schema index == default position
        if (c < s.blockCols && !taken[c]) { taken[c] = true; }
        map[i] = off + c;
    }
    return map;
}

Sheet SchemaDrivenReader::parseSheet(const QVector<QVector<QVariant>>& g,
                                     const QString& sheetName,
                                     const TemplateSchema& s,
                                     bool perRowRegime)
{
    Sheet out;
    out.sheetName    = sheetName;
    out.schema       = s;
    out.perRowRegime = perRowRegime;

    const int hdrRow   = s.columnHeaderRow - 1;
    const int hdrWidth = (hdrRow < g.size()) ? g[hdrRow].size() : 0;
    const int blocks   = hdrWidth / s.blockCols;      // same floor rule as countSamples()

    for (int b = 0; b < blocks; ++b) {
        const int off = b * s.blockCols;
        Sample sample;
        for (const HeaderFieldDef& h : s.headerFields) {
            const QVariant v = cellAt(g, h.row - 1, off + h.col - 1);
            sample.headers.insert(h.key, v);          // raw; typing at lowering
        }
        const QVector<int> colMap = resolveColumns(g, s, off);
        for (const MetricDef& m : s.columns)
            sample.data.append(MetricSeries{m.key, {}});
        for (int r = s.dataStartRow - 1; r < g.size(); ++r) {
            bool allEmpty = true;
            for (int i = 0; i < colMap.size(); ++i)
                if (!cellEmpty(cellAt(g, r, colMap[i]))) { allEmpty = false; break; }
            if (allEmpty) break;                      // legacy stop rule
            for (int i = 0; i < colMap.size(); ++i)
                sample.data[i].values.append(cellAt(g, r, colMap[i]));
        }
        out.samples.append(sample);
    }
    return out;
}

}} // namespace DVE::model
```

- [ ] **Step 4: Move the `Sheet` include usage in Task 5's test if needed (`#include "model/SchemaDrivenReader.h"`), add sources to both `.pro` files, build + run - ALL tst_v3model tests PASS.**

- [ ] **Step 5: Commit** - `feat(v3): SchemaDrivenReader - name-first schema-driven block parsing (Phase 1)`

### Task 9: `LegacyAdapter` (model -> legacy raw shape)

**Files:**
- Create: `src/model/LegacyAdapter.h`, `src/model/LegacyAdapter.cpp`
- Modify: `tests/tst_v3model/*`, `DataViewerEnterprise.pro`

The adapter lowers a `model::Sample` back to `ExcelReader::SampleData` so the untouched `SheetProcessor` chain (repair rules, burn/clog/leak scan, TpmCalculator, aggregates) produces the output.
This is the Phase-1 strangler shell, deleted in Phase 4.

- [ ] **Step 1: Write failing tests:**

```cpp
void TestV3Model::adapterLowersMetadataAndGrid()
{
    const TemplateSchema s = standardV1(false);
    Sample m;
    m.headers.insert("test_name", "My Test");
    m.headers.insert("voltage", "3.7");
    m.headers.insert("resistance", "1.2 Ohm");        // tolerant parse must apply
    m.headers.insert("heating_technology", "CCELL3.0");
    for (const MetricDef& c : s.columns) m.data.append(MetricSeries{c.key, {}});
    // 2 rows of puffs/before/after
    m.data[0].values = {10, 20};
    m.data[1].values = {25.10, 25.065};
    m.data[2].values = {25.065, 25.032};

    const ExcelReader::SampleData raw = LegacyAdapter::lowerSample(m, s, /*blockIndex=*/1);
    QCOMPARE(raw.metadata.testName, QStringLiteral("My Test"));
    QCOMPARE(raw.metadata.voltage, 3.7);
    QCOMPARE(raw.metadata.resistance, 1.2);
    // power = V^2 / (R + CCELL3.0 offset 0.78) - mirrors ExcelReader::extractMetadata
    QVERIFY(qAbs(raw.metadata.power - (3.7 * 3.7) / (1.2 + 0.78)) < 1e-9);
    QCOMPARE(raw.startColumn, 12);
    QCOMPARE(raw.dataRows.size(), 2);
    QCOMPARE(raw.dataRows[0].size(), 12);             // full block width restored
    QCOMPARE(raw.dataRows[0][0].toInt(), 10);
    QCOMPARE(raw.dataRows[1][2].toDouble(), 25.032);
}
```

- [ ] **Step 2: Build - fails. Step 3: Implement**

`src/model/LegacyAdapter.h`:

```cpp
#pragma once
#include "SchemaDrivenReader.h"
#include "../ExcelReader.h"

namespace DVE { namespace model {

// Phase-1 strangler shell: lowers the v3 model to the legacy raw shape so the
// UNCHANGED SheetProcessor chain computes every metric - production output is
// byte-identical by construction. Removed in Phase 4.
class LegacyAdapter {
public:
    static ExcelReader::SampleData lowerSample(const Sample& s,
                                               const TemplateSchema& schema,
                                               int blockIndex);
};

}} // namespace DVE::model
```

`src/model/LegacyAdapter.cpp` - the metadata mapping mirrors `ExcelReader::extractMetadata` field-for-field (`ExcelReader.cpp:456-536`); numeric fields go through `ExcelReader::tolerantCellDouble`, power through the HeatingTech offset formula (`src/pipeline/HeatingTech.h`):

```cpp
#include "LegacyAdapter.h"
#include "../pipeline/HeatingTech.h"

namespace DVE { namespace model {

ExcelReader::SampleData LegacyAdapter::lowerSample(const Sample& s,
                                                   const TemplateSchema& schema,
                                                   int blockIndex)
{
    ExcelReader::SampleData out;
    auto str = [&](const char* k){ return s.headers.value(QLatin1String(k)).toString().trimmed(); };
    auto num = [&](const char* k){ return ExcelReader::tolerantCellDouble(s.headers.value(QLatin1String(k))); };

    out.metadata.testName          = str("test_name");
    out.metadata.date              = str("date");
    out.metadata.sampleID          = str("sample_id");
    out.metadata.heatingTechnology = str("heating_technology");
    out.metadata.media             = str("media");
    out.metadata.resistance        = num("resistance");
    out.metadata.puffingRegime     = str("puffing_regime");
    out.metadata.viscosity         = num("viscosity");
    out.metadata.tester            = str("tester");
    out.metadata.voltage           = num("voltage");
    out.metadata.initialOilMass    = num("initial_oil_mass");
    // Derived exactly as ExcelReader::extractMetadata does (L528-530): the
    // power CELL is ignored; P = V^2 / (R + tech offset).
    const double rTot = out.metadata.resistance
        + DVE::heatingTechResistanceOffset(out.metadata.heatingTechnology);
    out.metadata.power = (rTot > 0.0) ? (out.metadata.voltage * out.metadata.voltage) / rTot : 0.0;

    out.startColumn = blockIndex * schema.blockCols;

    const int rows = s.rowCount();
    out.dataRows.resize(rows);
    for (int r = 0; r < rows; ++r) {
        QVector<QVariant> row(schema.blockCols);
        for (int c = 0; c < s.data.size() && c < schema.blockCols; ++c)
            row[c] = (r < s.data[c].values.size()) ? s.data[c].values[r] : QVariant();
        out.dataRows[r] = row;
    }
    return out;
}

}} // namespace DVE::model
```

**Before coding, read `git show HEAD:src/ExcelReader.cpp | sed -n '419,536p'` and `git show HEAD:src/pipeline/HeatingTech.h`** - if `extractMetadata` applies any additional trimming/branching for the standardized format (e.g. the Cart/Project sub-format branches), match ONLY the standardized branch (L485-521); the shadow harness in Task 10 is the referee for anything missed.
Verify the exact offset-function name in `HeatingTech.h` and use it (the test's 0.78 CCELL3.0 constant comes from that header).

- [ ] **Step 4: Build + run - PASS. Step 5: Commit** - `feat(v3): LegacyAdapter lowering model->SampleData (Phase 1 strangler shell)`

### Task 10: `DataProcessor::processFileV3` + old-vs-new shadow gate

**Files:**
- Modify: `src/pipeline/DataProcessor.h`, `src/pipeline/DataProcessor.cpp`
- Modify: `tests/tst_v3shadow/tst_v3shadow.cpp` / `.pro` (add model sources)

- [ ] **Step 1: Declare the V3 entry points** in `DataProcessor.h` (public, below the existing ones):

```cpp
    // v3 (Phase 1): schema-driven read path. Same output contract as
    // processFile/processSheet; extraction runs through model::SchemaDrivenReader
    // + model::LegacyAdapter instead of ExcelReader's positional getters.
    FileResult  processFileV3(
        const QString& filePath,
        std::function<void(int, const QString&)> progressCallback = nullptr);
    SheetResult processSheetV3(ExcelReader& reader, const QString& sheetName);
```

- [ ] **Step 2: Implement `processSheetV3`** in `DataProcessor.cpp`.
Read the full existing `processSheet` first (`git show HEAD:src/pipeline/DataProcessor.cpp`); V3 must reproduce its structure exactly - same SOP short-circuit, same empty-sheet return, same perRowRegime detection, same try/catch envelope - with only the extraction swapped:

```cpp
SheetResult DataProcessor::processSheetV3(ExcelReader& reader, const QString& sheetName)
{
    SheetResult empty;
    empty.sheetName = sheetName;

    // SOP short-circuit: identical to processSheet. Extract the existing SOP
    // block (DataProcessor.cpp:170-193) into a private helper
    //   SheetResult processSopSheet(ExcelReader&, const QString&);
    // and call it from BOTH paths so the logic exists once.
    if (sheetName.contains(QStringLiteral("SOP"), Qt::CaseInsensitive))
        return processSopSheet(reader, sheetName);

    const auto cells = reader.currentSheetCells();
    if (cells.isEmpty()) return empty;

    const QStringList hdrs = reader.getColumnHeaders();
    const bool perRowRegime =
        (hdrs.size() > 4) && RegimeUtils::isRegimeHeader(hdrs.at(4));

    const model::TemplateSchema schema = model::standardV1(perRowRegime);
    const model::Sheet mSheet =
        model::SchemaDrivenReader::parseSheet(cells, sheetName, schema, perRowRegime);
    if (mSheet.samples.isEmpty()) return empty;      // blank sheets stay non-errors

    QVector<ExcelReader::SampleData> rawSamples;
    rawSamples.reserve(mSheet.samples.size());
    for (int i = 0; i < mSheet.samples.size(); ++i)
        rawSamples.append(model::LegacyAdapter::lowerSample(mSheet.samples[i], schema, i));

    std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));
    processor->setPerRowRegime(perRowRegime);
    const QString templateVersion = reader.detectTemplateVersion();
    SheetResult result;
    try {
        result = processor->process(rawSamples, sheetName, templateVersion);
    } catch (const std::exception& ex) {
        setError(QString("Exception processing sheet '%1': %2").arg(sheetName, ex.what()));
        return empty;
    } catch (...) {
        setError(QString("Unknown exception processing sheet '%1'").arg(sheetName));
        return empty;
    }
    return result;
}
```

`processFileV3` mirrors `processFile`'s loop verbatim (same load, sheet iteration, progress milestones, FileResult assembly) with `processSheetV3` substituted - copy the existing loop, do not restructure it.

- [ ] **Step 3: Add the shadow test** to `tst_v3shadow.cpp`:

```cpp
void TestV3Shadow::v3MatchesLegacyParser()
{
    QFETCH(QString, path);   // reuse the same _data() generator via a second _data slot
    DVE::DataProcessor pOld, pNew;
    const DVE::FileResult oldR = pOld.processFile(path);
    if (oldR.sheets.isEmpty() && pOld.lastError().contains("python", Qt::CaseInsensitive))
        QSKIP("python unavailable");
    const DVE::FileResult newR = pNew.processFileV3(path);
    const QStringList diff = DVE::testutil::diffJson(DVE::fileResultToJson(oldR),
                                                     DVE::fileResultToJson(newR));
    QVERIFY2(diff.isEmpty(), qPrintable(diff.join('\n')));
}
```

Add `../../src/model/*.cpp` sources to the suite `.pro`.

- [ ] **Step 4: Build + run.**
Expected: likely a handful of diffs on the first run (metadata trimming, empty-row edge cases, the Temperature Cycling checklist sheet).
**Fix the V3 path until the diff is empty on every fixture** - the diff output names the exact JSON path of each divergence.
Do NOT "fix" a diff by changing the legacy path or the serializer.

- [ ] **Step 5: If `DVE_TEST_CORPUS_DIR` is populated locally, run the suite against it too (widest coverage); record the file count in the commit message.**

- [ ] **Step 6: Commit** - `feat(v3): processFileV3 schema-driven path + shadow gate green over corpus (Phase 1)`

### Task 11: Flip production to the V3 path

**Files:**
- Modify: `src/pipeline/DataProcessor.h`, `src/pipeline/DataProcessor.cpp`
- Modify: `tests/tst_v3shadow/tst_v3shadow.cpp`

- [ ] **Step 1: Swap bodies.**
`processFile`/`processSheet` become the V3 implementations; the old positional implementations are renamed `processFileLegacy`/`processSheetLegacy` (public, doc-commented "shadow-harness referee only - deleted in Phase 4 with ExcelReader's positional getters").
`processFileV3`/`processSheetV3` become thin forwards to `processFile`/`processSheet` and are dropped from the header (the shadow test now diffs `processFile` vs `processFileLegacy`).

- [ ] **Step 2: Update the shadow test** to compare `processFileLegacy` (referee) against `processFile` (production).

- [ ] **Step 3: Run tst_v3shadow + tst_dataprocessor + tst_sheetprocessors + tst_excelreader - ALL PASS.**

- [ ] **Step 4: Full-suite gate:** `tests\run-tests.ps1` - every suite PASS (DB suites skip cleanly without `DVE_TEST_PG_CONN`; that is fine and expected).

- [ ] **Step 5: Manual smoke (per the verify habit):** launch the built `DataViewer.exe`, open a fixture workbook, confirm the TPM table + plot render exactly as before, edit a Notes cell, confirm write-back still lands (write-back is untouched Phase-2 territory, but the loaded structs now come from the V3 path).

- [ ] **Step 6: Commit** - `feat(v3): production read path flipped to schema-driven reader; legacy kept as shadow referee (Phase 1 complete)`

### Task 12: Wrap-up

- [ ] Update the spec (`docs/superpowers/specs/2026-07-16-tpm-template-v3-metric-model-design.md`): append a short "Phase log" section noting Phase 0 + Phase 1 complete, with the gate evidence (suite counts, corpus size used).
- [ ] Append any corrected assumptions to `tasks/lessons.md` if the shadow diff exposed a wrong belief about the legacy parser.
- [ ] Commit - `docs(v3): Phase 0+1 complete - shadow gate green, production on schema-driven path`

---

## Verification summary

| Gate | Command | Must show |
|---|---|---|
| Harness self-check | `tst_v3harness` | all PASS |
| Parser determinism | `tst_v3shadow::parseIsDeterministic` | empty diff per corpus file |
| Phase-1 shadow gate | `tst_v3shadow::v3MatchesLegacyParser` | empty diff per corpus file |
| No regression | `tests\run-tests.ps1` | every suite PASS |
| E2E smoke | open fixture in `DataViewer.exe` | identical table/plot, edits write back |

## Explicitly out of scope (later phases)

Manifest reading (`_dve_schema`), CalculatorRegistry evaluation, CellAddressMap write-back, DB long-format, UI/report schema-driving, deleting the legacy structs or `ExcelReader` positional getters.
The `calculator`/`inputs` strings recorded by `standardV1` are metadata-only in Phase 1 - nothing evaluates them yet.
