# TPM Report Overhaul + Test Runner Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the TPM Report overhaul (adaptive markers, larger fonts, fixed Y-axis scaling, three new slide types in single-file Full Reports, multi-file combined reports with section dividers + Lifetime Comparison + per-section overviews) and bring the existing 1,800-line Qt Test scaffold to life behind a unified PowerShell runner.

**Architecture:** Most logic lives in `ReportGenerator` (slide assembly, plot helpers, multi-file batch). New helper file `SopLoader` extracted from `SopDialog`. Four new slide builders added to `PptxWriter` (Test Protocol / Test Overview / Section Divider / Conclusions). Lifetime Comparison reuses `addImageSlide` with a single full-body plot. Bar-chart legend support added to `PlotEngine` as a small `legendEntries` extension on `PlotConfig`. Test runner is a new `tests\run-tests.ps1` that auto-detects Qt + MinGW, builds via `tests/tests.pro` (`SUBDIRS`), and runs each `tst_*\release\tst_*.exe`.

**Tech Stack:** Qt 6.10.x (Widgets/Sql/Network/Concurrent) + MinGW 13.1.0, qmake build, Qt Test framework, PowerShell, Inno Setup 6, QXlsx (vendored).

---

## Phase 1 — Test runner

### Task 1: Branch + plan acknowledgement

**Files:**
- No file changes; git operations only.

- [ ] **Step 1: Create feature branch off main**

```powershell
git checkout main
git pull origin main
git checkout -b feat/tpm-report-overhaul
```

- [ ] **Step 2: Verify clean state**

Run: `git status`
Expected: `nothing to commit, working tree clean` on `feat/tpm-report-overhaul`.

---

### Task 2: Write `tests\run-tests.ps1`

**Files:**
- Create: `tests/run-tests.ps1`

- [ ] **Step 1: Write the runner**

```powershell
[CmdletBinding()]
param(
    [switch] $Rebuild,
    [switch] $VerboseRun
)

$ErrorActionPreference = 'Stop'
$repoRoot  = Split-Path -Parent $PSScriptRoot
$testsRoot = $PSScriptRoot

function Find-Qt {
    $candidates = Get-ChildItem 'C:\Qt' -Directory -Filter '6.10.*' -ErrorAction SilentlyContinue |
                  Sort-Object Name -Descending
    foreach ($c in $candidates) {
        $qmake = Join-Path $c.FullName 'mingw_64\bin\qmake.exe'
        if (Test-Path $qmake) { return @{ Qmake = $qmake; Bin = (Split-Path $qmake) } }
    }
    throw 'Could not find a Qt 6.10.x install at C:\Qt\6.10.*\mingw_64\bin\qmake.exe.'
}

function Find-MinGW {
    $candidates = Get-ChildItem 'C:\Qt\Tools' -Directory -Filter 'mingw1*_64' -ErrorAction SilentlyContinue |
                  Sort-Object Name -Descending
    foreach ($c in $candidates) {
        $make = Join-Path $c.FullName 'bin\mingw32-make.exe'
        if (Test-Path $make) { return @{ Make = $make; Bin = (Split-Path $make) } }
    }
    throw 'Could not find MinGW at C:\Qt\Tools\mingw1*_64\bin\mingw32-make.exe.'
}

$qt    = Find-Qt
$mingw = Find-MinGW
Write-Host "Qt    : $($qt.Bin)"
Write-Host "MinGW : $($mingw.Bin)"

$env:PATH = "$($mingw.Bin);$($qt.Bin);$env:PATH"

Push-Location $testsRoot
try {
    if ($Rebuild -and (Test-Path 'Makefile')) {
        & $mingw.Make distclean | Out-Null
        Get-ChildItem 'Makefile*' -ErrorAction SilentlyContinue | Remove-Item -Force
    }

    if (-not (Test-Path 'Makefile')) {
        Write-Host 'Running qmake...'
        & $qt.Qmake -recursive tests.pro
        if ($LASTEXITCODE -ne 0) { throw "qmake failed (exit $LASTEXITCODE)" }
    }

    Write-Host 'Building tests...'
    $jobs = [Environment]::ProcessorCount
    & $mingw.Make "-j$jobs" release
    if ($LASTEXITCODE -ne 0) { throw "build failed (exit $LASTEXITCODE)" }
}
finally { Pop-Location }

$tstDirs = Get-ChildItem $testsRoot -Directory -Filter 'tst_*' | Sort-Object Name
$pass = 0; $fail = 0; $skip = 0
$failures = @()

Write-Host ''
Write-Host '============================================================'
Write-Host '  Running tests'
Write-Host '============================================================'

foreach ($d in $tstDirs) {
    $exeFile = Join-Path $d.FullName "release\$($d.Name).exe"
    if (-not (Test-Path $exeFile)) {
        Write-Host ("  SKIP  {0,-30} (not built)" -f $d.Name)
        $skip++
        continue
    }

    $logFile = Join-Path $env:TEMP "$($d.Name).txt"
    if (Test-Path $logFile) { Remove-Item $logFile -Force }

    $proc = Start-Process -FilePath $exeFile -ArgumentList @('-o', "`"$logFile,txt`"") `
                          -Wait -PassThru -NoNewWindow

    $totals = $null
    if (Test-Path $logFile) {
        $totals = (Select-String -Path $logFile -Pattern '^Totals:' -SimpleMatch | Select-Object -First 1).Line
    }

    if ($proc.ExitCode -eq 0) {
        $line = if ($totals) { $totals -replace '^Totals: ', '' } else { 'no totals' }
        Write-Host ("  PASS  {0,-30} {1}" -f $d.Name, $line) -ForegroundColor Green
        $pass++
    } else {
        Write-Host ("  FAIL  {0,-30} exit={1} {2}" -f $d.Name, $proc.ExitCode, $totals) -ForegroundColor Red
        $fail++
        $failures += [pscustomobject]@{ Name = $d.Name; Log = $logFile }
    }
}

Write-Host ''
Write-Host '============================================================'
Write-Host ("  Results: {0} passed, {1} failed, {2} skipped" -f $pass, $fail, $skip)
Write-Host '============================================================'

if ($failures.Count -gt 0) {
    Write-Host ''
    Write-Host 'FAILURE DETAILS:'
    foreach ($f in $failures) {
        Write-Host "--- $($f.Name) ---"
        if (Test-Path $f.Log) { Select-String -Path $f.Log -Pattern '^FAIL!' | ForEach-Object { Write-Host $_.Line } }
    }
}

exit ($(if ($fail -gt 0) { 1 } else { 0 }))
```

- [ ] **Step 2: Smoke run**

Run: `powershell -ExecutionPolicy Bypass -File .\tests\run-tests.ps1`
Expected: builds 11 test binaries; per-test PASS/FAIL line; final summary; exit code surfaced. Some pre-existing tests may currently fail under Qt 6.10.2 — those are real bugs handled in Task 3.

- [ ] **Step 3: Commit**

```powershell
git add tests/run-tests.ps1
git commit -m "feat(tests): add unified PowerShell test runner"
```

---

### Task 3: Stabilize the existing test suite

If the smoke run from Task 2 surfaces failures, fix them here. Each broken test is its own micro-task: read the test, read the production code, fix the underlying issue (do not mutate the test to silence it).

- [ ] **Step 1: Triage failures**

Run: `.\tests\run-tests.ps1` and capture FAILURE DETAILS.

Categorize each as:
- **Genuine production bug** — fix the production code, leave the test alone.
- **Test written against an older API** — update the test only if the API change was deliberate; otherwise treat as production bug.
- **Build error** — fix includes / linker line in the test's `.pro`.

- [ ] **Step 2: Fix one failure at a time**

For each failure: open the test, run only that one (`pushd tests\<dir>; & .\release\<test>.exe; popd`), fix, re-run, commit.

Commit format: `fix(<area>): <what was broken>` — one commit per logical fix.

- [ ] **Step 3: Re-run the full suite until green**

Run: `.\tests\run-tests.ps1`
Expected: `Results: N passed, 0 failed, 0 skipped`.

- [ ] **Step 4: Commit any remaining fixups**

```powershell
git status
git add -A
git commit -m "fix(tests): bring existing suite to green under Qt 6.10.2"
```

---

### Task 4: Delete legacy test scripts and update docs

**Files:**
- Delete: `tests/build_all.bat`
- Delete: `tests/run_all_tests.sh`
- Modify: `CLAUDE.md`

- [ ] **Step 1: Delete legacy scripts**

```powershell
git rm tests/build_all.bat tests/run_all_tests.sh
```

- [ ] **Step 2: Update CLAUDE.md test note**

Find:

> There are no automated unit tests yet — `test_rcc_output.cpp` at the repo root is a generated Qt resource file, and `tests/tst_*` directories are scaffolds without an active runner. Deployment is verified via `tests/deployment/Test-Deployment.ps1` (see below).

Replace with:

> Run the unit-test suite with `tests\run-tests.ps1`. 11 test classes (~1,800 lines) cover pipeline, plotting, reporting, database, and zip/xml utilities. The runner auto-detects Qt + MinGW under `C:\Qt\6.10.*` and builds incrementally via `tests/tests.pro` (a `SUBDIRS` template). Deployment is verified via `tests/deployment/Test-Deployment.ps1` after install.
>
> `test_rcc_output.cpp` at the repo root is a generated Qt resource file, not a test.

- [ ] **Step 3: Run the suite once more after cleanup**

Run: `.\tests\run-tests.ps1`
Expected: all pass.

- [ ] **Step 4: Commit**

```powershell
git add CLAUDE.md tests/build_all.bat tests/run_all_tests.sh
git commit -m "chore(tests): retire legacy scripts; document run-tests.ps1 in CLAUDE.md"
```

---

## Phase 2 — SopLoader extraction

### Task 5: Create `SopLoader.h`

**Files:**
- Create: `src/utils/SopLoader.h`

- [ ] **Step 1: Write the header**

```cpp
#pragma once

#include <QString>
#include <QVector>

namespace DVE {

struct SopEntry {
    QString test;
    QString sop;
    QString objective;
    QString passCriteria;
    QString equipment;
    QString quantity;
    QString estDuration1mL;
    QString estDuration2mL;
    QString note;
};

class SopLoader {
public:
    /// Load all SOP rows from the standardized template's "Test SOPs" sheet.
    /// Returns an empty vector and emits a qWarning if the file is missing,
    /// unreadable, or has no parseable rows. Never throws.
    static QVector<SopEntry> load(const QString& xlsxPath);
};

} // namespace DVE
```

- [ ] **Step 2: Don't commit yet** — header alone won't compile in isolation; commit happens at end of Task 6.

---

### Task 6: Implement `SopLoader.cpp` and wire into `.pro`

**Files:**
- Create: `src/utils/SopLoader.cpp`
- Modify: `DataViewerEnterprise.pro`

- [ ] **Step 1: Write the implementation**

```cpp
#include "SopLoader.h"

#include <QDebug>
#include <QFile>
#include <xlsxdocument.h>

namespace DVE {

QVector<SopEntry> SopLoader::load(const QString& xlsxPath)
{
    QVector<SopEntry> result;

    if (!QFile::exists(xlsxPath)) {
        qWarning() << "SopLoader: file not found:" << xlsxPath;
        return result;
    }

    QXlsx::Document xlsx(xlsxPath);
    if (!xlsx.load()) {
        qWarning() << "SopLoader: QXlsx failed to load:" << xlsxPath;
        return result;
    }

    if (!xlsx.selectSheet(QStringLiteral("Test SOPs"))) {
        qWarning() << "SopLoader: 'Test SOPs' sheet not found in" << xlsxPath;
        return result;
    }

    // Row 1 is headers, data starts at row 2.
    // Columns: A=Test, B=SOP, C=Objective, D=Pass Criteria, E=Equipment,
    //          F=Quantity, G=Est Duration (1mL), H=Est Duration (2mL), I=Note
    for (int row = 2; row <= 200; ++row) {
        const QString test = xlsx.read(row, 1).toString().trimmed();
        if (test.isEmpty()) break;

        SopEntry e;
        e.test           = test;
        e.sop            = xlsx.read(row, 2).toString().trimmed();
        e.objective      = xlsx.read(row, 3).toString().trimmed();
        e.passCriteria   = xlsx.read(row, 4).toString().trimmed();
        e.equipment      = xlsx.read(row, 5).toString().trimmed();
        e.quantity       = xlsx.read(row, 6).toString().trimmed();
        e.estDuration1mL = xlsx.read(row, 7).toString().trimmed();
        e.estDuration2mL = xlsx.read(row, 8).toString().trimmed();
        e.note           = xlsx.read(row, 9).toString().trimmed();
        result.append(e);
    }

    return result;
}

} // namespace DVE
```

- [ ] **Step 2: Wire into `DataViewerEnterprise.pro`**

In SOURCES append `src/utils/SopLoader.cpp`. In HEADERS append `src/utils/SopLoader.h`.

- [ ] **Step 3: Build to verify**

```powershell
$env:PATH = "C:\Qt\Tools\mingw1310_64\bin;$env:PATH"
mkdir build -ErrorAction SilentlyContinue
pushd build
& "C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++ ..\DataViewerEnterprise.pro
& mingw32-make -j8 release
popd
```

Expected: clean build under `-Werror`.

- [ ] **Step 4: Commit**

```powershell
git add src/utils/SopLoader.h src/utils/SopLoader.cpp DataViewerEnterprise.pro
git commit -m "feat(utils): add SopLoader for shared SOP-row parsing"
```

---

### Task 7: Add `tst_sopLoader` test

**Files:**
- Create: `tests/tst_sopLoader/tst_sopLoader.cpp`
- Create: `tests/tst_sopLoader/tst_sopLoader.pro`
- Modify: `tests/tests.pro`

- [ ] **Step 1: Write the test**

`tests/tst_sopLoader/tst_sopLoader.cpp`:

```cpp
#include <QtTest>
#include <QFileInfo>
#include "SopLoader.h"

class TestSopLoader : public QObject
{
    Q_OBJECT

private slots:
    void loadsKnownTemplate();
    void missingFileReturnsEmpty();
};

void TestSopLoader::loadsKnownTemplate()
{
    const QString repoRoot = QFINDTESTDATA("../../resources/templates");
    QVERIFY(!repoRoot.isEmpty());

    const QString xlsx = repoRoot + "/Standardized Test Template - December 2025.xlsx";
    QVERIFY2(QFileInfo::exists(xlsx), qPrintable("Template not found: " + xlsx));

    const QVector<DVE::SopEntry> rows = DVE::SopLoader::load(xlsx);
    QVERIFY2(rows.size() >= 3, qPrintable(QString("expected >= 3 SOP rows, got %1").arg(rows.size())));

    bool hasLifetime = false;
    for (const auto& r : rows) {
        if (r.test.compare("Lifetime Test", Qt::CaseInsensitive) == 0) {
            hasLifetime = true;
            QVERIFY2(!r.objective.isEmpty(), "Lifetime Test row has empty Objective");
            QVERIFY2(!r.passCriteria.isEmpty(), "Lifetime Test row has empty Pass Criteria");
        }
    }
    QVERIFY2(hasLifetime, "no row matched 'Lifetime Test'");
}

void TestSopLoader::missingFileReturnsEmpty()
{
    const QVector<DVE::SopEntry> rows = DVE::SopLoader::load("Z:/does/not/exist.xlsx");
    QCOMPARE(rows.size(), 0);
}

QTEST_APPLESS_MAIN(TestSopLoader)
#include "tst_sopLoader.moc"
```

- [ ] **Step 2: Write the .pro file**

`tests/tst_sopLoader/tst_sopLoader.pro`:

```pro
QT += core testlib
CONFIG += qt warn_on testcase c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_sopLoader

INCLUDEPATH += $$PWD \
               $$PWD/../../src \
               $$PWD/../../src/utils \
               $$PWD/../common

include($$PWD/../../external/QXlsx/QXlsx/QXlsx.pri)

SOURCES += tst_sopLoader.cpp \
           $$PWD/../../src/utils/SopLoader.cpp

HEADERS += $$PWD/../../src/utils/SopLoader.h

LIBS += -lz
```

- [ ] **Step 3: Add to `tests/tests.pro`**

Append `tst_sopLoader` to the SUBDIRS block.

- [ ] **Step 4: Run the test**

```powershell
.\tests\run-tests.ps1
```

Expected: `tst_sopLoader` builds and PASSes both tests.

- [ ] **Step 5: Commit**

```powershell
git add tests/tst_sopLoader tests/tests.pro
git commit -m "test(sopLoader): add tst_sopLoader covering load + missing-file cases"
```

---

### Task 8: Refactor `SopDialog` to use `SopLoader`

**Files:**
- Modify: `src/ui/SopDialog.cpp`
- Possibly: `src/ui/SopDialog.h`

- [ ] **Step 1: Refactor `SopDialog::loadFromExcel`**

Replace the body of `SopDialog::loadFromExcel(const QString& xlsxPath)` with:

```cpp
void SopDialog::loadFromExcel(const QString& xlsxPath)
{
    if (!QFile::exists(xlsxPath)) {
        QMessageBox::warning(this, "SOP Load Error",
                             "SOP file not found:\n" + xlsxPath);
        return;
    }

    const QVector<DVE::SopEntry> rows = DVE::SopLoader::load(xlsxPath);

    m_entries.clear();
    m_testList->clear();
    for (const DVE::SopEntry& e : rows) {
        m_entries.append(e);
        m_testList->addItem(e.test);
    }

    if (!m_entries.isEmpty())
        m_testList->setCurrentRow(0);
}
```

If `SopDialog.h` defines its own `SopEntry`, replace it with `using SopEntry = DVE::SopEntry;` and `#include "../utils/SopLoader.h"`.

- [ ] **Step 2: Build**

```powershell
pushd build; mingw32-make -j8 release; popd
```

Expected: clean build.

- [ ] **Step 3: Manual smoke**

Launch `build\release\DataViewer.exe`, click `Tools → SOPs`, verify the dialog still populates the test list and shows details.

- [ ] **Step 4: Commit**

```powershell
git add src/ui/SopDialog.h src/ui/SopDialog.cpp
git commit -m "refactor(ui): SopDialog uses SopLoader for SOP parsing"
```

---

## Phase 3 — Plot helpers

### Task 9: `isLongPuff` helper + tests

**Files:**
- Modify: `src/reporting/ReportGenerator.h` (private declaration + public test wrapper)
- Modify: `src/reporting/ReportGenerator.cpp`
- Modify: `tests/tst_reportgenerator/tst_reportgenerator.cpp`

- [ ] **Step 1: Add failing tests**

Append to `TestReportGenerator`:

```cpp
private slots:
    void isLongPuff_sheetNameMatch();
    void isLongPuff_regimeRegexMatch();
    void isLongPuff_neither();
```

Bodies:

```cpp
void TestReportGenerator::isLongPuff_sheetNameMatch()
{
    DVE::ReportGenerator gen;
    DVE::SheetResult sheet;
    sheet.sheetName = "Long Puff Lifetime Test";
    QVERIFY(gen.isLongPuffForTesting(sheet));
}

void TestReportGenerator::isLongPuff_regimeRegexMatch()
{
    DVE::ReportGenerator gen;
    DVE::SheetResult sheet;
    sheet.sheetName = "Lifetime Test";
    DVE::SampleResult s;
    s.puffingRegime = "200mL/10s/60s";
    sheet.samples.append(s);
    QVERIFY(gen.isLongPuffForTesting(sheet));
}

void TestReportGenerator::isLongPuff_neither()
{
    DVE::ReportGenerator gen;
    DVE::SheetResult sheet;
    sheet.sheetName = "Lifetime Test";
    DVE::SampleResult s;
    s.puffingRegime = "55mL/3s/30s";
    sheet.samples.append(s);
    QVERIFY(!gen.isLongPuffForTesting(sheet));
}
```

- [ ] **Step 2: Add header declarations**

In `ReportGenerator.h`, public:

```cpp
    bool isLongPuffForTesting(const SheetResult& s) const { return isLongPuff(s); }
```

Private:

```cpp
    bool isLongPuff(const SheetResult& sheet) const;
```

- [ ] **Step 3: Run, confirm fail**

```powershell
.\tests\run-tests.ps1
```

Expected: `tst_reportgenerator` fails to build.

- [ ] **Step 4: Implement**

In `ReportGenerator.cpp`:

```cpp
#include <QRegularExpression>

bool ReportGenerator::isLongPuff(const SheetResult& sheet) const
{
    if (sheet.sheetName.contains(QStringLiteral("Long Puff"), Qt::CaseInsensitive))
        return true;

    static const QRegularExpression kRegimeRe(
        QStringLiteral(R"(\d+\s*mL\s*/\s*10s\s*/.*)"),
        QRegularExpression::CaseInsensitiveOption);
    for (const SampleResult& s : sheet.samples) {
        if (kRegimeRe.match(s.puffingRegime).hasMatch())
            return true;
    }
    return false;
}
```

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp tests/tst_reportgenerator/tst_reportgenerator.cpp
git commit -m "feat(reporting): isLongPuff helper for fixed Y-axis scaling"
```

---

### Task 10: `computeTpmYMax` helper + tests

**Files:** same as Task 9.

- [ ] **Step 1: Failing tests**

```cpp
private slots:
    void yMax_nonLongPuff_underCeiling();
    void yMax_nonLongPuff_overCeiling();
    void yMax_longPuff_inRange();
    void yMax_longPuff_belowMin();
    void yMax_longPuff_aboveMax();
```

Helper + bodies:

```cpp
static DVE::SheetResult sheetWith(const QString& sheetName,
                                  const QString& regime,
                                  const QVector<double>& tpmValues)
{
    DVE::SheetResult sh;
    sh.sheetName = sheetName;
    DVE::SampleResult s;
    s.puffingRegime = regime;
    double sum = 0;
    for (double t : tpmValues) {
        DVE::DataRow r;
        r.tpm = t;
        s.rows.append(r);
        sum += t;
    }
    s.averageTPM = tpmValues.isEmpty() ? 0.0 : sum / tpmValues.size();
    sh.samples.append(s);
    return sh;
}

void TestReportGenerator::yMax_nonLongPuff_underCeiling()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Lifetime Test", "55mL/3s/30s", {3, 4, 5, 6, 6.5});
    QFUZZY_COMPARE(gen.computeTpmYMaxForTesting(sh), 7.0);
}

void TestReportGenerator::yMax_nonLongPuff_overCeiling()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Lifetime Test", "55mL/3s/30s", {7.5, 8.0, 8.5});
    QFUZZY_COMPARE(gen.computeTpmYMaxForTesting(sh), 9.5);
}

void TestReportGenerator::yMax_longPuff_inRange()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Long Puff Lifetime Test", "200mL/10s/60s", {15, 18, 22, 24});
    QFUZZY_COMPARE(gen.computeTpmYMaxForTesting(sh), 25.0);
}

void TestReportGenerator::yMax_longPuff_belowMin()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Long Puff Lifetime Test", "200mL/10s/60s", {8, 10, 12, 14});
    QFUZZY_COMPARE(gen.computeTpmYMaxForTesting(sh), 15.0);
}

void TestReportGenerator::yMax_longPuff_aboveMax()
{
    DVE::ReportGenerator gen;
    auto sh = sheetWith("Long Puff Lifetime Test", "200mL/10s/60s", {18, 22, 27});
    QFUZZY_COMPARE(gen.computeTpmYMaxForTesting(sh), 28.0);
}
```

- [ ] **Step 2: Add header declarations**

Public:

```cpp
    double computeTpmYMaxForTesting(const SheetResult& s) const { return computeTpmYMax(s); }
```

Private:

```cpp
    double computeTpmYMax(const SheetResult& sheet) const;
```

- [ ] **Step 3: Run, confirm fail**

- [ ] **Step 4: Implement**

```cpp
double ReportGenerator::computeTpmYMax(const SheetResult& sheet) const
{
    double maxTPM = 0.0;
    double sumAvg = 0.0;
    int    nSamples = 0;
    for (const SampleResult& s : sheet.samples) {
        ++nSamples;
        sumAvg += s.averageTPM;
        for (const DataRow& r : s.rows)
            if (r.tpm > maxTPM) maxTPM = r.tpm;
    }
    const double avgTPM = (nSamples > 0) ? (sumAvg / nSamples) : 0.0;

    if (isLongPuff(sheet)) {
        return (maxTPM >= 15.0 && maxTPM <= 25.0) ? 25.0 : maxTPM + 1.0;
    } else {
        return (avgTPM > 7.0) ? maxTPM + 1.0 : 7.0;
    }
}
```

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp tests/tst_reportgenerator/tst_reportgenerator.cpp
git commit -m "feat(reporting): computeTpmYMax enforces fixed Y-axis scaling"
```

---

### Task 11: `adaptiveDotRadius` helper + tests

- [ ] **Step 1: Failing test**

```cpp
private slots:
    void adaptiveDotRadius_cases();
```

Body:

```cpp
void TestReportGenerator::adaptiveDotRadius_cases()
{
    DVE::ReportGenerator gen;
    QCOMPARE(gen.adaptiveDotRadiusForTesting(0),    5);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(30),   5);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(150),  2);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(1000), 2);
    QCOMPARE(gen.adaptiveDotRadiusForTesting(90),   4);
}
```

Header public test wrapper:

```cpp
    int adaptiveDotRadiusForTesting(int n) const { return adaptiveDotRadius(n); }
```

Private:

```cpp
    int adaptiveDotRadius(int pointCount) const;
```

- [ ] **Step 2: Run, confirm fail**

- [ ] **Step 3: Implement**

```cpp
#include <cmath>

int ReportGenerator::adaptiveDotRadius(int pointCount) const
{
    if (pointCount <= 30)  return 5;
    if (pointCount >= 150) return 2;
    const double t = (pointCount - 30) / 120.0;
    return static_cast<int>(std::round(5.0 - 3.0 * t));
}
```

- [ ] **Step 4: Run, confirm pass**

- [ ] **Step 5: Commit**

```powershell
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp tests/tst_reportgenerator/tst_reportgenerator.cpp
git commit -m "feat(reporting): adaptive marker size scales with point density"
```

---

### Task 12: `reportPlotConfig` helper

(No test — pure factory; covered indirectly in Task 14.)

- [ ] **Step 1: Add to header (private)**

```cpp
    PlotConfig reportPlotConfig() const;
```

- [ ] **Step 2: Implement**

```cpp
PlotConfig ReportGenerator::reportPlotConfig() const
{
    PlotConfig cfg;
    cfg.titleFont = QFont("Segoe UI", 18, QFont::Bold);
    cfg.axisFont  = QFont("Segoe UI", 18);
    cfg.labelFont = QFont("Segoe UI", 14);
    return cfg;
}
```

- [ ] **Step 3: Build**

```powershell
pushd build; mingw32-make -j8 release; popd
```

- [ ] **Step 4: Commit**

```powershell
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp
git commit -m "feat(reporting): reportPlotConfig() bumps fonts for slide-embedded plots"
```

---

### Task 13: `lifetimeBarColor` helper + tests

- [ ] **Step 1: Failing tests**

```cpp
private slots:
    void lifetimeBarColor_distinctPerFile();
    void lifetimeBarColor_progressiveShading();
```

Bodies:

```cpp
void TestReportGenerator::lifetimeBarColor_distinctPerFile()
{
    QColor f0 = DVE::ReportGenerator::lifetimeBarColor(0, 0, 1);
    QColor f1 = DVE::ReportGenerator::lifetimeBarColor(1, 0, 1);
    QColor f2 = DVE::ReportGenerator::lifetimeBarColor(2, 0, 1);
    QVERIFY(f0 != f1);
    QVERIFY(f1 != f2);
    QVERIFY(f0 != f2);
}

void TestReportGenerator::lifetimeBarColor_progressiveShading()
{
    QColor a = DVE::ReportGenerator::lifetimeBarColor(0, 0, 3);
    QColor b = DVE::ReportGenerator::lifetimeBarColor(0, 1, 3);
    QColor c = DVE::ReportGenerator::lifetimeBarColor(0, 2, 3);
    QVERIFY(a != b && b != c);
    int ha, sa, va, hb, sb, vb, hc, sc, vc;
    a.getHsv(&ha, &sa, &va);
    b.getHsv(&hb, &sb, &vb);
    c.getHsv(&hc, &sc, &vc);
    QCOMPARE(ha, hb);
    QCOMPARE(hb, hc);
    QVERIFY(va != vb || sa != sb);
}
```

Header public (already static — exposed for test):

```cpp
    static QColor lifetimeBarColor(int fileIdx, int sampleIdx, int totalSamplesInFile);
```

- [ ] **Step 2: Run, confirm fail**

- [ ] **Step 3: Implement**

```cpp
QColor ReportGenerator::lifetimeBarColor(int fileIdx, int sampleIdx, int totalSamplesInFile)
{
    static const QColor kFileHues[] = {
        QColor(0x00, 0x66, 0xCC),
        QColor(0xFF, 0x73, 0x00),
        QColor(0x00, 0xAA, 0x44),
        QColor(0xCC, 0x00, 0x00),
        QColor(0x99, 0x00, 0xCC),
        QColor(0x00, 0xAA, 0xCC),
    };
    const QColor base = kFileHues[fileIdx % 6];

    if (totalSamplesInFile <= 1) return base;

    int h, s, v;
    base.getHsv(&h, &s, &v);
    const double t = static_cast<double>(sampleIdx) / (totalSamplesInFile - 1);
    const int newV = static_cast<int>(std::round(255 * (0.6 + 0.4 * t)));
    return QColor::fromHsv(h, s, qBound(0, newV, 255));
}
```

- [ ] **Step 4: Run, confirm pass**

- [ ] **Step 5: Commit**

```powershell
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp tests/tst_reportgenerator/tst_reportgenerator.cpp
git commit -m "feat(reporting): lifetimeBarColor — hue per file, value per sample"
```

---

### Task 14: Wire helpers into `buildPlots`

**Files:** `src/reporting/ReportGenerator.cpp`

- [ ] **Step 1: Rewrite the TPM Trend block**

Replace the existing TPM Trend block in `ReportGenerator::buildPlots` with:

```cpp
        QVector<PlotSeries> series;
        int colorIdx = 0;
        for (const SampleResult& sr : sheet.samples) {
            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
            ps.color     = kColors[colorIdx++ % 6];
            ps.drawLine  = true;
            ps.drawDots  = true;
            ps.lineWidth = 2;

            for (const DataRow& row : sr.rows) {
                if (row.beforeWeight == 0.0 || row.afterWeight == 0.0) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.tpm);
            }
            ps.dotRadius = adaptiveDotRadius(ps.x.size());
            if (!ps.x.isEmpty()) series.append(std::move(ps));
        }

        if (!series.isEmpty()) {
            PlotConfig cfg = reportPlotConfig();
            cfg.title      = sheet.sheetName + QStringLiteral(" – TPM Trend");
            cfg.xLabel     = "Cumulative Puffs";
            cfg.yLabel     = "TPM (mg)";
            cfg.width      = 800;
            cfg.height     = 480;
            cfg.showGrid   = true;
            cfg.showLegend = (series.size() > 1);
            cfg.autoScale  = false;
            cfg.yMin       = 0.0;
            cfg.yMax       = computeTpmYMax(sheet);

            double xMax = 0;
            for (const PlotSeries& ps : series)
                for (double xv : ps.x)
                    if (xv > xMax) xMax = xv;
            cfg.xMin = 0; cfg.xMax = (xMax > 0 ? xMax : 1);

            QPixmap pm  = PlotEngine::renderLinePlot(series, cfg);
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty()) plots.append(std::move(png));
        }
```

- [ ] **Step 2: Rewrite the Bar Chart block**

```cpp
    if (includeBarChart) {
        QVector<QString> names;
        QVector<double> avgs, sdevs;
        for (const auto& s : sheet.samples) {
            names.append(s.sampleName.isEmpty() ? s.sampleID : s.sampleName);
            avgs.append(s.averageTPM);
            sdevs.append(s.stdDevTPM);
        }
        PlotConfig cfg = reportPlotConfig();
        cfg.title      = sheet.sheetName + QStringLiteral(" – Average TPM");
        cfg.yLabel     = "Avg TPM (mg)";
        cfg.width      = 800;
        cfg.height     = 480;
        cfg.autoScale  = false;
        cfg.yMin       = 0.0;
        cfg.yMax       = computeTpmYMax(sheet);
        QPixmap pm = PlotEngine::renderBarChart(names, avgs, cfg, /*colors=*/{}, sdevs);
        QByteArray png = PlotEngine::toPng(pm, 150);
        if (!png.isEmpty()) plots.append(std::move(png));
    }
```

- [ ] **Step 3: Rewrite the Draw Pressure block (markers + fonts only — no Y-axis rule)**

```cpp
    {
        QVector<PlotSeries> series;
        int colorIdx = 0;
        for (const SampleResult& sr : sheet.samples) {
            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
            ps.color     = kColors[colorIdx++ % 6];
            ps.drawLine  = true;
            ps.drawDots  = true;
            ps.lineWidth = 2;
            for (const DataRow& row : sr.rows) {
                if (row.drawPressure == 0.0) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.drawPressure);
            }
            ps.dotRadius = adaptiveDotRadius(ps.x.size());
            if (!ps.x.isEmpty()) series.append(std::move(ps));
        }

        if (!series.isEmpty()) {
            PlotConfig cfg = reportPlotConfig();
            cfg.title      = sheet.sheetName + QStringLiteral(" – Draw Pressure");
            cfg.xLabel     = "Cumulative Puffs";
            cfg.yLabel     = "Draw Pressure (Pa)";
            cfg.width      = 800;
            cfg.height     = 480;
            cfg.autoScale  = true;
            cfg.showGrid   = true;
            cfg.showLegend = (series.size() > 1);
            QPixmap pm = PlotEngine::renderLinePlot(series, cfg);
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty()) plots.append(std::move(png));
        }
    }
```

- [ ] **Step 4: Build + run unit tests**

```powershell
pushd build; mingw32-make -j8 release; popd
.\tests\run-tests.ps1
```

Expected: clean build, all tests still green.

- [ ] **Step 5: Manual smoke**

Launch the app, load a TPM file, generate a Test Report. Open the resulting `.pptx`:
- Markers visible on every data point.
- Title font visibly larger.
- TPM Trend Y-axis fixed (0–7 normal, 0–25 Long Puff).

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/ReportGenerator.cpp
git commit -m "feat(reporting): wire adaptive markers, bumped fonts, fixed Y-axis into buildPlots"
```

---

## Phase 4 — `PlotConfig::legendEntries` for the comparison slide

### Task 15: Bar-chart legend support

The bar-chart variant has no legend support today. The Lifetime Comparison needs one. Spec said "no PlotEngine changes" — that was wrong; small extension required.

**Files:**
- Modify: `src/plotting/PlotEngine.h`
- Modify: `src/plotting/PlotEngine.cpp`
- Modify: `tests/tst_plotengine/tst_plotengine.cpp`

- [ ] **Step 1: Extend `PlotConfig`**

In `src/plotting/PlotEngine.h`, append to `PlotConfig`:

```cpp
    // Optional manual legend entries — used by bar charts (which have no
    // series to derive a legend from) and any custom plot. Empty = no legend.
    QVector<QPair<QString, QColor>> legendEntries;
```

- [ ] **Step 2: Failing test**

Append to `tests/tst_plotengine/tst_plotengine.cpp`:

```cpp
private slots:
    void barChart_drawsLegendEntries();
```

Body:

```cpp
void TestPlotEngine::barChart_drawsLegendEntries()
{
    DVE::PlotConfig cfg;
    cfg.title = "Test"; cfg.width = 600; cfg.height = 400;
    cfg.legendEntries = {
        {"File A", QColor(255, 0, 0)},
        {"File B", QColor(0, 0, 255)},
    };
    QVector<QString> labels = {"s1", "s2"};
    QVector<double>  vals   = {3.0, 4.0};
    QPixmap pm = DVE::PlotEngine::renderBarChart(labels, vals, cfg);
    QImage img = pm.toImage();

    bool sawRed = false, sawBlue = false;
    for (int y = 0; y < img.height() / 4; ++y) {
        for (int x = img.width() / 2; x < img.width(); ++x) {
            QColor c = img.pixelColor(x, y);
            if (c.red()   > 200 && c.green() < 50  && c.blue() < 50) sawRed  = true;
            if (c.blue()  > 200 && c.red()   < 50  && c.green() < 50) sawBlue = true;
            if (sawRed && sawBlue) break;
        }
        if (sawRed && sawBlue) break;
    }
    QVERIFY(sawRed);
    QVERIFY(sawBlue);
}
```

- [ ] **Step 3: Run, confirm fail**

```powershell
.\tests\run-tests.ps1
```

- [ ] **Step 4: Implement legend drawing in `renderBarChart`**

In `PlotEngine::renderBarChart`, just before returning, add:

```cpp
    if (!config.legendEntries.isEmpty()) {
        const int swatchW = 24;
        const int swatchH = 14;
        QFontMetrics fm(config.labelFont);
        const int rowH    = qMax(swatchH + 6, fm.height() + 4);
        const int padding = 12;

        int maxLabelW = 0;
        for (const auto& kv : config.legendEntries)
            maxLabelW = qMax(maxLabelW, fm.horizontalAdvance(kv.first));
        const int boxW = swatchW + 8 + maxLabelW + padding * 2;
        const int boxH = config.legendEntries.size() * rowH + padding;

        const int boxX = pxRight - boxW - 8;
        const int boxY = pxTop + 8;

        p.setPen(QPen(config.axisColor, 1));
        p.setBrush(QBrush(QColor(255, 255, 255, 230)));
        p.drawRect(boxX, boxY, boxW, boxH);

        p.setFont(config.labelFont);
        for (int i = 0; i < config.legendEntries.size(); ++i) {
            const auto& kv = config.legendEntries[i];
            const int rowY = boxY + padding / 2 + i * rowH;
            p.setPen(Qt::NoPen);
            p.setBrush(kv.second);
            p.drawRect(boxX + padding, rowY + (rowH - swatchH) / 2, swatchW, swatchH);
            p.setPen(config.axisColor);
            p.drawText(boxX + padding + swatchW + 8,
                       rowY + (rowH + fm.ascent()) / 2 - 2,
                       kv.first);
        }
    }
```

(`pxRight` and `pxTop` are local plot-area boundaries already in scope. Verify by rereading the function.)

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/plotting/PlotEngine.h src/plotting/PlotEngine.cpp tests/tst_plotengine/tst_plotengine.cpp
git commit -m "feat(plotting): PlotConfig::legendEntries — manual legend for bar charts"
```

---

## Phase 5 — New PptxWriter slide builders

### Task 16: `addConclusionsSlide` (simplest first)

**Files:**
- Modify: `src/reporting/PptxWriter.h`
- Modify: `src/reporting/PptxWriter.cpp`
- Modify: `tests/tst_pptxwriter/tst_pptxwriter.cpp`

- [ ] **Step 1: Failing test**

```cpp
private slots:
    void addConclusionsSlide_producesEmptyTextbox();
```

Body:

```cpp
void TestPptxWriter::addConclusionsSlide_producesEmptyTextbox()
{
    DVE::PptxWriter w;
    w.addConclusionsSlide();
    const QString tmp = QDir::tempPath() + "/dve_conclusions.pptx";
    QVERIFY(w.save(tmp));
    QVERIFY(QFileInfo(tmp).size() > 1000);
    QFile::remove(tmp);
}
```

- [ ] **Step 2: Add header declaration**

In `PptxWriter.h` public section:

```cpp
    void addConclusionsSlide();
```

- [ ] **Step 3: Run, confirm fail**

- [ ] **Step 4: Implement**

```cpp
void PptxWriter::addConclusionsSlide()
{
    Slide s;
    QString bgRid   = "rId2";
    QString logoRid = "rId3";
    s.media.append({QStringLiteral("media/background.png"), loadResourceImage("ccell_background.png")});
    s.media.append({QStringLiteral("media/logo.png"),       loadResourceImage("ccell_logo_full_white.png")});

    QString shapes;
    shapes += makeTextBox(100, 0.46, 0.30, 12.4, 0.80,
                          QStringLiteral("Conclusions"),
                          QStringLiteral("Segoe UI"), 3200, true,
                          QStringLiteral("FFFFFF"), QStringLiteral("l"));
    shapes += makeTextBox(101, 0.80, 1.40, 11.7, 5.50,
                          QStringLiteral(""),
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    s.xml = buildContentSlideXml(QStringLiteral("Conclusions"),
                                 SlideTable{}, QVector<SlideImage>{},
                                 bgRid, logoRid, {}, shapes);
    m_slides.append(std::move(s));
}
```

If `buildContentSlideXml` doesn't handle empty `SlideTable` cleanly, add a guard inside it: skip the table block if `table.headers.isEmpty()`. Apply that guard in this same task.

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/PptxWriter.h src/reporting/PptxWriter.cpp tests/tst_pptxwriter/tst_pptxwriter.cpp
git commit -m "feat(pptx): addConclusionsSlide — title + empty body textbox"
```

---

### Task 17: `addTestProtocolSlide`

- [ ] **Step 1: Failing test**

```cpp
private slots:
    void addTestProtocolSlide_producesTable();
```

Body:

```cpp
void TestPptxWriter::addTestProtocolSlide_producesTable()
{
    DVE::PptxWriter w;
    DVE::SlideTable t;
    t.headers = {"Test", "Objective", "Pass Criteria", "Equipment", "Quantity", "Est Duration"};
    t.rows.append({"Lifetime Test", "Run to depletion", "TPM > X", "Vape rig", "n=3", "1mL: 4h / 2mL: 8h"});
    w.addTestProtocolSlide(t);
    const QString tmp = QDir::tempPath() + "/dve_protocol.pptx";
    QVERIFY(w.save(tmp));
    QVERIFY(QFileInfo(tmp).size() > 1000);
    QFile::remove(tmp);
}
```

- [ ] **Step 2: Add header**

```cpp
    void addTestProtocolSlide(const SlideTable& sopTable);
```

- [ ] **Step 3: Run, confirm fail**

- [ ] **Step 4: Implement**

```cpp
void PptxWriter::addTestProtocolSlide(const SlideTable& sopTable)
{
    Slide s;
    QString bgRid   = "rId2";
    QString logoRid = "rId3";
    s.media.append({QStringLiteral("media/background.png"), loadResourceImage("ccell_background.png")});
    s.media.append({QStringLiteral("media/logo.png"),       loadResourceImage("ccell_logo_full_white.png")});

    s.xml = buildContentSlideXml(QStringLiteral("Test Protocol"),
                                 sopTable, QVector<SlideImage>{},
                                 bgRid, logoRid, {}, QString());
    m_slides.append(std::move(s));
}
```

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/PptxWriter.h src/reporting/PptxWriter.cpp tests/tst_pptxwriter/tst_pptxwriter.cpp
git commit -m "feat(pptx): addTestProtocolSlide — title + filtered SOP table"
```

---

### Task 18: `addTestOverviewSlide`

- [ ] **Step 1: Failing test**

```cpp
private slots:
    void addTestOverviewSlide_producesBullets();
```

Body:

```cpp
void TestPptxWriter::addTestOverviewSlide_producesBullets()
{
    DVE::PptxWriter w;
    w.addTestOverviewSlide("Standard performance evaluation of DeviceX across 3 tests.",
                           QStringList{"Lifetime Test", "Heavy Metals Test", "Big Headspace Test"});
    const QString tmp = QDir::tempPath() + "/dve_overview.pptx";
    QVERIFY(w.save(tmp));
    QVERIFY(QFileInfo(tmp).size() > 1000);
    QFile::remove(tmp);
}
```

- [ ] **Step 2: Add header**

```cpp
    void addTestOverviewSlide(const QString& description,
                              const QStringList& testNames);
```

- [ ] **Step 3: Run, confirm fail**

- [ ] **Step 4: Implement**

```cpp
void PptxWriter::addTestOverviewSlide(const QString& description,
                                      const QStringList& testNames)
{
    Slide s;
    QString bgRid   = "rId2";
    QString logoRid = "rId3";
    s.media.append({QStringLiteral("media/background.png"), loadResourceImage("ccell_background.png")});
    s.media.append({QStringLiteral("media/logo.png"),       loadResourceImage("ccell_logo_full_white.png")});

    QString shapes;
    shapes += makeTextBox(100, 0.46, 0.30, 12.4, 0.80,
                          QStringLiteral("Test Overview"),
                          QStringLiteral("Segoe UI"), 3200, true,
                          QStringLiteral("FFFFFF"), QStringLiteral("l"));
    shapes += makeTextBox(101, 0.80, 1.30, 11.7, 0.55,
                          description,
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    QStringList bulletLines;
    for (const QString& name : testNames)
        bulletLines << QStringLiteral("• ") + name;
    shapes += makeTextBox(102, 1.00, 2.10, 11.5, 4.30,
                          bulletLines.join("\n"),
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    shapes += makeTextBox(103, 0.80, 6.50, 11.7, 0.55,
                          QStringLiteral(""),
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    s.xml = buildContentSlideXml(QStringLiteral("Test Overview"),
                                 SlideTable{}, QVector<SlideImage>{},
                                 bgRid, logoRid, {}, shapes);
    m_slides.append(std::move(s));
}
```

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/PptxWriter.h src/reporting/PptxWriter.cpp tests/tst_pptxwriter/tst_pptxwriter.cpp
git commit -m "feat(pptx): addTestOverviewSlide — title, description, bullets, empty summary"
```

---

### Task 19: `addSectionDividerSlide`

- [ ] **Step 1: Failing test**

```cpp
private slots:
    void addSectionDividerSlide_centersFilename();
```

Body:

```cpp
void TestPptxWriter::addSectionDividerSlide_centersFilename()
{
    DVE::PptxWriter w;
    w.addSectionDividerSlide("DeviceX_Run_2026-04-22");
    const QString tmp = QDir::tempPath() + "/dve_divider.pptx";
    QVERIFY(w.save(tmp));
    QVERIFY(QFileInfo(tmp).size() > 1000);
    QFile::remove(tmp);
}
```

- [ ] **Step 2: Add header**

```cpp
    void addSectionDividerSlide(const QString& filename);
```

- [ ] **Step 3: Run, confirm fail**

- [ ] **Step 4: Implement**

```cpp
void PptxWriter::addSectionDividerSlide(const QString& filename)
{
    Slide s;
    QString bgRid   = "rId2";
    QString logoRid = "rId3";
    s.media.append({QStringLiteral("media/background.png"), loadResourceImage("ccell_background.png")});
    s.media.append({QStringLiteral("media/logo.png"),       loadResourceImage("ccell_logo_full_white.png")});

    s.xml = buildCoverSlideXml(filename, /*date=*/QString(), bgRid, logoRid);
    m_slides.append(std::move(s));
}
```

If `buildCoverSlideXml` doesn't skip the date shape on empty input, modify it to skip the date textbox when the date string is empty. Commit that guard with this task.

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/PptxWriter.h src/reporting/PptxWriter.cpp tests/tst_pptxwriter/tst_pptxwriter.cpp
git commit -m "feat(pptx): addSectionDividerSlide — cover-style with centered filename"
```

---

## Phase 6 — ReportGenerator integration

### Task 20: `loadSopRows` + filtering helper

- [ ] **Step 1: Failing test**

```cpp
private slots:
    void loadSopRows_filtersToRequestedTests();
```

Body:

```cpp
void TestReportGenerator::loadSopRows_filtersToRequestedTests()
{
    DVE::ReportGenerator gen;
    gen.setResourcePath(QFINDTESTDATA("../../resources"));
    const QStringList request = {"Lifetime Test"};
    QVector<DVE::SopEntry> filtered = gen.loadSopRowsForTesting(request);
    QCOMPARE(filtered.size(), 1);
    QCOMPARE(filtered[0].test, QString("Lifetime Test"));
}
```

- [ ] **Step 2: Add header**

Public test wrapper:

```cpp
    QVector<SopEntry> loadSopRowsForTesting(const QStringList& reportTests) const
        { return loadSopRows(reportTests); }
```

Private:

```cpp
    QVector<SopEntry> loadSopRows(const QStringList& reportTestNames) const;
```

- [ ] **Step 3: Run, confirm fail**

- [ ] **Step 4: Implement**

```cpp
#include "../utils/SopLoader.h"
#include <QSet>

QVector<SopEntry> ReportGenerator::loadSopRows(const QStringList& reportTestNames) const
{
    const QString xlsx = m_resourcePath +
        QStringLiteral("/templates/Standardized Test Template - December 2025.xlsx");
    QVector<SopEntry> all = SopLoader::load(xlsx);
    if (reportTestNames.isEmpty()) return all;

    QSet<QString> wantLower;
    for (const QString& n : reportTestNames) wantLower.insert(n.toLower());

    QVector<SopEntry> filtered;
    for (const SopEntry& e : all) {
        if (wantLower.contains(e.test.toLower()))
            filtered.append(e);
    }
    return filtered;
}
```

- [ ] **Step 5: Run, confirm pass**

- [ ] **Step 6: Commit**

```powershell
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp tests/tst_reportgenerator/tst_reportgenerator.cpp
git commit -m "feat(reporting): loadSopRows filters template SOPs to tests in the report"
```

---

### Task 21: Wire new slides into `generateFullReport`

**Files:**
- Modify: `src/reporting/ReportGenerator.cpp`

- [ ] **Step 1: Inject new slides into `generateFullReport`**

Right after the existing cover slide and before the per-sheet loop, insert:

```cpp
    // Test Protocol slide
    {
        QStringList testNames;
        for (const SheetResult& sh : data.sheets)
            if (sh.hasSamples())
                testNames << sh.sheetName;

        const QVector<SopEntry> sopRows = loadSopRows(testNames);
        if (!sopRows.isEmpty()) {
            SlideTable t;
            t.headers = {"Test", "Objective", "Pass Criteria", "Equipment", "Quantity", "Est Duration"};
            for (const SopEntry& e : sopRows) {
                QString dur = QStringLiteral("1mL: %1 / 2mL: %2")
                                  .arg(e.estDuration1mL.isEmpty() ? "-" : e.estDuration1mL,
                                       e.estDuration2mL.isEmpty() ? "-" : e.estDuration2mL);
                t.rows.append({e.test, e.objective, e.passCriteria, e.equipment, e.quantity, dur});
            }
            QSet<QString> covered;
            for (const SopEntry& e : sopRows) covered.insert(e.test.toLower());
            for (const QString& n : testNames)
                if (!covered.contains(n.toLower()))
                    t.rows.append({n, "—", "—", "—", "—", "—"});
            writer.addTestProtocolSlide(t);
        }
    }

    // Test Overview slide
    {
        QStringList testNames;
        for (const SheetResult& sh : data.sheets)
            if (sh.hasSamples())
                testNames << sh.sheetName;
        const QString desc = QStringLiteral("Standard performance evaluation of %1 across %2 tests.")
                                 .arg(displayFileName).arg(testNames.size());
        writer.addTestOverviewSlide(desc, testNames);
    }
```

At the end of `generateFullReport`, just before `reportProgress(progress, 95, ...)`, insert:

```cpp
    writer.addConclusionsSlide();
```

- [ ] **Step 2: Build**

```powershell
pushd build; mingw32-make -j8 release; popd
```

- [ ] **Step 3: Manual smoke**

Generate a Full Report on a real file. Open the resulting `.pptx` and verify slide order: Title → Test Protocol → Test Overview → existing data slides → existing image slides → Conclusions.

- [ ] **Step 4: Commit**

```powershell
git add src/reporting/ReportGenerator.cpp
git commit -m "feat(reporting): generateFullReport adds Test Protocol, Test Overview, Conclusions"
```

---

### Task 22: `generateCombinedFullReport`

**Files:**
- Modify: `src/reporting/ReportGenerator.h`
- Modify: `src/reporting/ReportGenerator.cpp`

- [ ] **Step 1: Add public declaration**

```cpp
    bool generateCombinedFullReport(const QVector<FileResult>& files,
                                    const ReportConfig& config,
                                    const QString& outputPath,
                                    ProgressFn progress = nullptr);
```

- [ ] **Step 2: Implement**

```cpp
bool ReportGenerator::generateCombinedFullReport(const QVector<FileResult>& files,
                                                  const ReportConfig& /*config*/,
                                                  const QString& outputPath,
                                                  ProgressFn progress)
{
    if (files.isEmpty()) {
        m_lastError = "No files to combine";
        return false;
    }

    PptxWriter writer;
    writer.setResourcePath(m_resourcePath);

    // 1. Cover
    QDate today = QDate::currentDate();
    QString dateStr = today.toString("MMMM d, yyyy");
    writer.addCoverSlide(QStringLiteral("Combined Standard Test Report"), dateStr);

    // 2. Test Protocol — union across files
    QStringList allTests;
    QSet<QString> seen;
    for (const FileResult& f : files) {
        for (const SheetResult& sh : f.sheets) {
            if (sh.hasSamples() && !seen.contains(sh.sheetName.toLower())) {
                seen.insert(sh.sheetName.toLower());
                allTests << sh.sheetName;
            }
        }
    }
    {
        const QVector<SopEntry> sopRows = loadSopRows(allTests);
        SlideTable t;
        t.headers = {"Test", "Objective", "Pass Criteria", "Equipment", "Quantity", "Est Duration"};
        QSet<QString> covered;
        for (const SopEntry& e : sopRows) {
            covered.insert(e.test.toLower());
            QString dur = QStringLiteral("1mL: %1 / 2mL: %2")
                              .arg(e.estDuration1mL.isEmpty() ? "-" : e.estDuration1mL,
                                   e.estDuration2mL.isEmpty() ? "-" : e.estDuration2mL);
            t.rows.append({e.test, e.objective, e.passCriteria, e.equipment, e.quantity, dur});
        }
        for (const QString& n : allTests)
            if (!covered.contains(n.toLower()))
                t.rows.append({n, "—", "—", "—", "—", "—"});
        if (!t.rows.isEmpty()) writer.addTestProtocolSlide(t);
    }

    // 3. Test Overview (combined)
    {
        const QString desc = QStringLiteral("Combined performance evaluation across %1 files and %2 unique tests.")
                                 .arg(files.size()).arg(allTests.size());
        writer.addTestOverviewSlide(desc, allTests);
    }

    // 4. Lifetime TPM Comparison
    {
        QVector<QString> labels;
        QVector<double>  values;
        QVector<QColor>  colors;
        QVector<QPair<QString, QColor>> legend;

        for (int fi = 0; fi < files.size(); ++fi) {
            const FileResult& f = files[fi];
            const SheetResult* lifetime = nullptr;
            for (const SheetResult& sh : f.sheets) {
                if (sh.sheetName.compare("Lifetime Test", Qt::CaseInsensitive) == 0) {
                    lifetime = &sh; break;
                }
            }
            if (!lifetime || !lifetime->hasSamples()) continue;

            int totalSamples = 0;
            for (const SampleResult& s : lifetime->samples)
                if (!s.rows.isEmpty()) ++totalSamples;

            QString fileLabel = QFileInfo(f.fileName).completeBaseName();
            legend.append({fileLabel, lifetimeBarColor(fi, 0, qMax(2, totalSamples))});

            int sIdx = 0;
            for (const SampleResult& s : lifetime->samples) {
                if (s.rows.isEmpty()) continue;
                QString name = s.sampleName.isEmpty() ? s.sampleID : s.sampleName;
                labels.append(name);
                values.append(s.averageTPM);
                colors.append(lifetimeBarColor(fi, sIdx, totalSamples));
                ++sIdx;
            }
        }

        if (!labels.isEmpty()) {
            SheetResult merged;
            merged.sheetName = "Lifetime Test";
            for (int i = 0; i < values.size(); ++i) {
                SampleResult s;
                s.averageTPM = values[i];
                DataRow r; r.tpm = values[i]; s.rows.append(r);
                merged.samples.append(s);
            }

            PlotConfig cfg = reportPlotConfig();
            cfg.title         = "Lifetime TPM Comparison";
            cfg.yLabel        = "Avg TPM (mg)";
            cfg.width         = 1200;
            cfg.height        = 600;
            cfg.autoScale     = false;
            cfg.yMin          = 0.0;
            cfg.yMax          = computeTpmYMax(merged);
            cfg.legendEntries = legend;
            cfg.labelFont     = QFont("Segoe UI", 18);

            QPixmap pm = PlotEngine::renderBarChart(labels, values, cfg, colors, /*stdDev=*/{});
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty()) {
                SlideImage img;
                img.pngData = png;
                img.x = 0.30; img.y = 1.10; img.w = 12.7; img.h = 6.20;
                writer.addImageSlide(QStringLiteral("Lifetime TPM Comparison"),
                                     QVector<SlideImage>{img});
            }
        }
    }

    // 5. Per-file sections
    const int totalFiles = files.size();
    for (int fi = 0; fi < totalFiles; ++fi) {
        const FileResult& f = files[fi];
        QString displayName = QFileInfo(f.fileName).completeBaseName();
        reportProgress(progress, 20 + (60 * fi / totalFiles),
                       "Adding section: " + displayName);

        writer.addSectionDividerSlide(displayName);

        QStringList fileTests;
        for (const SheetResult& sh : f.sheets)
            if (sh.hasSamples()) fileTests << sh.sheetName;
        const QString perFileDesc =
            QStringLiteral("Standard performance evaluation of %1 across %2 tests.")
                .arg(displayName).arg(fileTests.size());
        writer.addTestOverviewSlide(perFileDesc, fileTests);

        for (const SheetResult& origSheet : f.sheets) {
            if (!origSheet.hasSamples()) continue;
            SheetResult sheet = origSheet;
            sheet.samples.erase(
                std::remove_if(sheet.samples.begin(), sheet.samples.end(),
                               [](const SampleResult& s){ return s.rows.isEmpty(); }),
                sheet.samples.end());
            if (!sheet.hasSamples()) continue;

            SlideTable tbl = buildTable(sheet, ReportConfig{});
            QVector<QByteArray> plotPngs = buildPlots(sheet, true);

            static const struct { double x, y, w, h; } kPlotLayout[] = {
                { 0.10, 3.25, 4.32, 3.20 },
                { 4.51, 3.25, 4.32, 3.20 },
                { 8.91, 3.25, 4.32, 3.20 },
            };
            QVector<SlideImage> plotImages;
            for (int pi = 0; pi < plotPngs.size() && pi < 3; ++pi) {
                SlideImage img;
                img.pngData = std::move(plotPngs[pi]);
                img.x = kPlotLayout[pi].x; img.y = kPlotLayout[pi].y;
                img.w = kPlotLayout[pi].w; img.h = kPlotLayout[pi].h;
                plotImages.append(std::move(img));
            }
            writer.addContentSlide(sheet.sheetName, tbl, plotImages);

            // Image slides — same logic as generateFullReport
            for (const SampleResult& sr : sheet.samples) {
                if (sr.imagePaths.isEmpty()) continue;
                QString sampleDisplay = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
                const bool hasLayouts = (sr.imageLayouts.size() == sr.imagePaths.size());
                if (hasLayouts) {
                    QVector<SlideImage> slideImages;
                    for (int i = 0; i < sr.imagePaths.size(); ++i) {
                        QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                        if (data.isEmpty()) continue;
                        SlideImage si;
                        si.pngData = std::move(data);
                        const QRectF& r = sr.imageLayouts[i];
                        si.x = r.x(); si.y = r.y();
                        si.w = r.width(); si.h = r.height();
                        slideImages.append(std::move(si));
                    }
                    if (!slideImages.isEmpty())
                        writer.addImageSlide(sampleDisplay + " Images", slideImages);
                } else {
                    QVector<QByteArray> imgBytes;
                    for (int i = 0; i < sr.imagePaths.size(); ++i) {
                        QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                        if (!data.isEmpty()) imgBytes.append(data);
                    }
                    if (!imgBytes.isEmpty())
                        writer.addImageSlide(sampleDisplay + " Images", imgBytes);
                }
            }
            QVector<QByteArray> sheetImgs = collectImages(sheet);
            if (!sheetImgs.isEmpty())
                writer.addImageSlide(sheet.sheetName + " – Photos", sheetImgs);
        }
    }

    // 6. Conclusions
    writer.addConclusionsSlide();

    reportProgress(progress, 95, "Saving combined report...");
    bool ok = writer.save(outputPath);
    if (!ok) { m_lastError = writer.lastError(); return false; }
    reportProgress(progress, 100, "Saved: " + outputPath);
    return true;
}
```

- [ ] **Step 3: Build**

```powershell
pushd build; mingw32-make -j8 release; popd
```

Expected: clean build.

- [ ] **Step 4: Commit**

```powershell
git add src/reporting/ReportGenerator.h src/reporting/ReportGenerator.cpp
git commit -m "feat(reporting): generateCombinedFullReport assembles multi-file output"
```

---

## Phase 7 — MainWindow integration

### Task 23: Multi-file picker + batch loop in `onGenerateFullReport`

**Files:**
- Modify: `src/MainWindow.cpp`
- Modify: `src/MainWindow.h`

- [ ] **Step 1: Add `uniqueFilename` static helper**

In `MainWindow.h` private section:

```cpp
    static QString uniqueFilename(const QString& desiredPath);
```

In `MainWindow.cpp`:

```cpp
QString MainWindow::uniqueFilename(const QString& desiredPath)
{
    if (!QFile::exists(desiredPath)) return desiredPath;
    QFileInfo fi(desiredPath);
    QString stem = fi.completeBaseName();
    QString ext  = fi.suffix();
    QString dir  = fi.absolutePath();
    for (int i = 2; i < 1000; ++i) {
        QString candidate = QString("%1/%2 (%3).%4").arg(dir, stem).arg(i).arg(ext);
        if (!QFile::exists(candidate)) return candidate;
    }
    return desiredPath;
}
```

- [ ] **Step 2: Replace the body of `onGenerateFullReport`**

```cpp
void MainWindow::onGenerateFullReport()
{
    QStringList paths = QFileDialog::getOpenFileNames(
        this, "Select Files for Full Report",
        lastBrowseDir(),
        "Excel files (*.xlsx)"
    );
    if (paths.isEmpty()) return;
    setLastBrowseDir(paths.first());

    QVector<FileResult> files;
    QStringList loadFailures;
    for (int i = 0; i < paths.size(); ++i) {
        const QString& p = paths[i];
        bool reused = false;
        for (const FileResult& fr : m_loadedFiles) {
            if (fr.filePath == p) {
                files.append(buildCleanedFile(fr));
                reused = true;
                break;
            }
        }
        if (reused) continue;

        DataProcessor dp;
        FileResult fr = dp.processFile(p,
            [this, i, total = paths.size()](int pct, const QString& msg) {
                setProgress(pct, QString("Loading file %1 of %2: %3").arg(i+1).arg(total).arg(msg));
            });
        if (fr.filePath.isEmpty()) {
            loadFailures << QFileInfo(p).fileName();
            continue;
        }
        files.append(buildCleanedFile(fr));
    }
    setProgress(0, QString());

    if (files.isEmpty()) {
        showError("Full Report", "No files could be loaded.\n\n" + loadFailures.join("\n"));
        return;
    }

    if (files.size() == 1) {
        const QString def = lastBrowseDir() + "/" + files.first().fileName.chopped(5) + "_Report.pptx";
        const QString path = QFileDialog::getSaveFileName(this, "Save Full Report", def, "PowerPoint (*.pptx)");
        if (path.isEmpty()) return;
        setLastBrowseDir(path);

        ReportConfig cfg;
        cfg.outputPath = path;
        m_reportGen->setResourcePath(resourcePath());
        m_reportGen->generateFullReport(files.first(), cfg);
        return;
    }

    const QString outDir = QFileDialog::getExistingDirectory(
        this, "Select Output Folder", lastBrowseDir());
    if (outDir.isEmpty()) return;

    m_reportGen->setResourcePath(resourcePath());

    QStringList reportFailures = loadFailures;
    int succeeded = 0;
    const int total = files.size() + 1;

    for (int i = 0; i < files.size(); ++i) {
        const FileResult& f = files[i];
        QString outPath = outDir + "/" + f.fileName.chopped(5) + "_Report.pptx";
        outPath = uniqueFilename(outPath);
        ReportConfig cfg;
        cfg.outputPath = outPath;
        setProgress((100 * i) / total, QString("Generating %1 of %2").arg(i+1).arg(total));
        bool ok = m_reportGen->generateFullReport(f, cfg);
        if (!ok) reportFailures << f.fileName;
        else     ++succeeded;
    }

    {
        QString combinedPath = outDir + "/Combined_Report_" +
            QDate::currentDate().toString("yyyy-MM-dd") + ".pptx";
        combinedPath = uniqueFilename(combinedPath);
        setProgress((100 * files.size()) / total, "Generating combined report");
        ReportConfig cfg;
        cfg.outputPath = combinedPath;
        bool ok = m_reportGen->generateCombinedFullReport(files, cfg, combinedPath);
        if (!ok) reportFailures << "Combined_Report.pptx";
        else     ++succeeded;
    }

    setProgress(100, QString());

    QMessageBox box(this);
    box.setWindowTitle("Reports Generated");
    QString text = QString("Generated %1 of %2 reports.").arg(succeeded).arg(total);
    if (!reportFailures.isEmpty())
        text += "\n\nFailed:\n  " + reportFailures.join("\n  ");
    box.setText(text);
    auto* openBtn = box.addButton("Open Folder", QMessageBox::ActionRole);
    box.addButton(QMessageBox::Ok);
    box.exec();
    if (box.clickedButton() == openBtn)
        QDesktopServices::openUrl(QUrl::fromLocalFile(outDir));
}
```

- [ ] **Step 3: Build**

```powershell
pushd build; mingw32-make -j8 release; popd
```

Expected: clean build (add includes for `QFileDialog`, `QDesktopServices`, `QUrl` if not already).

- [ ] **Step 4: Manual smoke — single-file**

`Reports → Full Report` → pick one file → confirm previous behavior + new slides.

- [ ] **Step 5: Manual smoke — multi-file**

`Reports → Full Report` → Ctrl+click 2 files → pick output folder → 3 `.pptx` files appear.

- [ ] **Step 6: Commit**

```powershell
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(ui): Full Report supports multi-file selection + combined output"
```

---

## Phase 8 — Verification + ship

### Task 24: Manual checklist appended to deployment README

**Files:**
- Modify: `tests/deployment/README.md`

- [ ] **Step 1: Append the checklist**

Append at the end of `tests/deployment/README.md`:

```markdown
## Full Report manual checklist

Run after every install on the work machine that touches reporting code.

### Single-file Full Report
- [ ] Generates without error on a normal file (3+ sheets)
- [ ] Title slide unchanged
- [ ] Test Protocol slide shows 6 cols, only tests present in the file
- [ ] Test Overview shows auto-templated line + bullets + empty trailing textbox
- [ ] Each data slide: TPM Trend has markers, Y-axis matches the rule
- [ ] Image slides unchanged
- [ ] Conclusions slide: blank textbox

### Multi-file Combined Report (3 files)
- [ ] File picker accepts multi-select
- [ ] Output folder dialog appears once
- [ ] 3 individual reports + 1 combined report appear in chosen folder
- [ ] Combined order: Title → Protocol → Overview → Lifetime Comparison →
      Section divider → File 1 overview → File 1 data → ... → Conclusions
- [ ] Section divider slides: cover style, no date, filename centered
- [ ] Lifetime Comparison: bars colored by file, shaded by sample within file
- [ ] Files without "Lifetime Test" sheet are silently skipped on the comparison slide

### Edge cases
- [ ] Long Puff Lifetime Test: Y-axis 0–25 default, fallback to maxTPM+1 out of range
- [ ] Sheet with avgTPM > 7: Y-axis bumps to maxTPM+1
- [ ] Puffing regime "200mL/10s/60s": detected as Long Puff even on a non-Long-Puff sheet name
- [ ] File picked that's already loaded: reused (no re-parse)
- [ ] One report failing in batch: other reports still generate, summary dialog shows failures
- [ ] Filename collision in output dir: appends (2), no overwrite
```

- [ ] **Step 2: Commit**

```powershell
git add tests/deployment/README.md
git commit -m "docs(deployment): manual checklist for Full Report changes"
```

---

### Task 25: Push branch and merge to main

- [ ] **Step 1: Push the branch**

```powershell
git push -u origin feat/tpm-report-overhaul
```

- [ ] **Step 2: Final test-suite run**

```powershell
.\tests\run-tests.ps1
```

Expected: green.

- [ ] **Step 3: Fast-forward main**

```powershell
git checkout main
git merge --ff-only feat/tpm-report-overhaul
git push origin main
git checkout feat/tpm-report-overhaul
```

- [ ] **Step 4: Pull on the work machine** (user, not the agent)

```powershell
git pull origin main
qmake CONFIG+=release && mingw32-make -j8
tools\prepare_python_embed.bat
build_installer.bat
```

Then install `dist\DataViewer-setup.exe` and run through the Task 24 checklist.

---

## Out of scope (explicit, do not add)

- Refining auto-templated overview text wording.
- Graphical CI on GitHub Actions.
- Changing GUI plot fonts or markers — only report plots change.
- Sensory-mode reports.
- PowerPoint headless validation.
- Long Puff Lifetime Test or Rapid Puff Lifetime Test cross-file comparison slides.
