# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Environment

This is a **Windows-only** project. Always use Windows-compatible commands and paths.

- Default shell is PowerShell on the user's machines; Bash via Git Bash is also available. Pick the right one for each task — `$env:VAR`, not `$VAR`; backtick continuation, not backslash.
- Be aware of Windows encoding issues — write text files as UTF-8 explicitly when other tools will read them. PowerShell `Set-Content`/`Out-File` defaults to UTF-16 LE with BOM in Windows PowerShell 5.1; pass `-Encoding utf8` when needed.
- Watch for PowerShell variable interpolation gotchas (`$env:PATH` vs `$PATH`).
- Python multiprocessing on Windows uses spawn semantics — anything with `ProcessPoolExecutor` needs an `if __name__ == '__main__':` guard.
- Avoid Unix-only patterns (`/dev/null`, forward-slash paths in shell args, `&&` chaining in PowerShell 5.1) without checking platform first.

### MIP file encryption (this machine specifically)

The user's Windows account applies Microsoft Information Protection (MIP / AIP) sensitivity labels to source files at rest. Random-looking encryption hits `.cpp`/`.h`/`.py` files in the working tree — `g++`, `head`, `cat`, and the Edit/Write tools see ciphertext starting with `%TSD-Header-###%`. The user's Python interpreter is on the MIP allowlist and reads plaintext.

**File creation convention:** create new source files via Python's delete-and-rewrite pattern so they don't inherit MIP labels:

```python
import os
path = "src/foo/Bar.cpp"
content = "..."          # the file's full contents
if os.path.exists(path):
    os.remove(path)
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
```

The Edit tool's create-file path also works on this machine but the freshly-written file may pick up a MIP label as soon as it's closed. Subagents that produce code should write via Python through the Bash tool to be safe.

**Reactive decryption:** before any C++ build attempt, run `python tools/decrypt_via_copy.py --apply` from the repo root. The script scans `src/`, `tests/`, `tools/`, and the QXlsx vendor tree, identifies any file whose raw bytes start with `%TSD-Header-###%`, and rewrites it via Python to strip the MIP labels. Idempotent and fast.

The `decrypt-mip-files` user-level skill documents this in more detail and triggers when ciphertext markers appear in tool output.

## Project

DataViewer Enterprise — a C++17 / Qt 6 Windows desktop app for analyzing vape device test data. Reads `.xlsx` files produced from a standardized test template, displays measurement data in editable tables and plots, generates branded PowerPoint reports, and stores results in a shared PostgreSQL 16 database hosted on the office Synology NAS. Multi-user concurrent access with NOTIFY-driven live updates, presence indicators, optimistic-concurrency conflict resolution dialogs, and a read-only offline mode backed by a local SQLite snapshot all shipped in v2.0. Namespace is `DVE`. Target binary is `DataViewer.exe`.

The repo is now self-contained: a fresh `git clone` builds without submodule init or any manual step. QXlsx is vendored in-tree at the pinned upstream commit `9f54593` (see `external/QXlsx/LICENSE`).

Run the unit-test suite with `tests\run-tests.ps1`. 27 test classes (~6,800 lines) cover pipeline, plotting, reporting, database (Postgres + offline-snapshot integration), live updates (NOTIFY, presence), zip/xml utilities, and the Plan C offline-failover end-to-end checkpoint. The runner auto-detects Qt + MinGW under `C:\Qt\6.10.*` and builds incrementally via `tests/tests.pro` (a `SUBDIRS` template). Database-dependent suites use `tests\start-test-postgres.ps1` for an ephemeral `postgres:16` container; suites skip cleanly when `DVE_TEST_PG_CONN` is unset. Deployment is verified via `tests/deployment/Test-Deployment.ps1` after install.

`test_rcc_output.cpp` at the repo root is a generated Qt resource file, not a test.

## Toolchain

The project is set up for **qmake + MinGW** only (no CMake, no MSVC). Required versions per the .pro and the bundled installs at `C:\Qt`:

- Qt 6.10.x, Widgets/SQL/Network/Concurrent modules — `C:\Qt\6.10.2\mingw_64\bin\qmake.exe`
- MinGW 13.1.0 — `C:\Qt\Tools\mingw1310_64\bin\` (must be on PATH for qmake/make)
- Inno Setup 6 — `C:\Program Files (x86)\Inno Setup 6\ISCC.exe` (only for installer)

## Build / Run / Package

Out-of-tree debug build:

```bat
mkdir build && cd build
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
C:\Qt\6.10.2\mingw_64\bin\qmake.exe -spec win32-g++ ..\DataViewerEnterprise.pro
mingw32-make -j8
```

Pass `CONFIG+=release` to qmake for a release build (the installer expects `release\DataViewer.exe`).

Installer (writes `dist\DataViewer-setup.exe`):

```bat
tools\prepare_python_embed.bat   :: one-time per release: bundles Python 3.11 + openpyxl into release\python_bundle.zip
build_installer.bat              :: requires release\DataViewer.exe to already exist
```

CLI flags on the binary itself:
- `DataViewer.exe <file>` — opens the file in the running instance via `SingleInstance` IPC and brings the existing window to the front. `__RAISE__` is the sentinel for "no file, just raise the window."
- `DataViewer.exe --self-test [--self-test-out PATH]` — runs deployment diagnostics (no GUI), writes a JSON report, and exits. See `tests/deployment/README.md`.

## Release Workflow

The product ships through a fixed loop. Respect each stage:

1. **Develop here (home machine).** Feature work happens on a topic branch, then fast-forwards into `main`. Push tags for shipped releases (`v1.0.0`, `v1.1.0`, …). Code-only verification is what's possible from here — Synology Drive is *not* mounted on this machine, and we cannot test the deployed binary's network path locally.

2. **Pull on the work machine (`S1134987`).** Rebuild the installer:
   ```bat
   qmake CONFIG+=release && mingw32-make -j8
   tools\prepare_python_embed.bat
   build_installer.bat
   ```

3. **Run the deployment self-test.** After installing the new build:
   ```powershell
   .\tests\deployment\Test-Deployment.ps1
   ```
   Every phase must pass before continuing. See "Deployment Self-Test" below.

4. **Drop on Synology.** Move `dist\DataViewer-setup.exe` into:
   ```
   %USERPROFILE%\SynologyDrive\SDR\Device Group\Software Release\Current\v<X.Y.Z>\DataViewer-setup.exe
   ```
   The folder name format is enforced — `UpdateChecker::latestAvailable()` walks subdirs and parses `vX.Y.Z` (or `X.Y.Z`). The auto-updater hourly check on every live install offers the new build to users.

What this means for development sessions: **anything touching network paths, Synology layout, registry, file permissions, or installer behavior cannot be verified here.** It can only be verified via the work-machine self-test. Plan accordingly — stage UI/architectural changes for normal review here, batch deployment-touching changes for a single round-trip to the work machine.

## Deployment Self-Test

`tests/deployment/Test-Deployment.ps1` is the verification harness for stage 3 above. It runs three phases:

1. **Install tree** — every required Qt DLL, plugin, bundled Python file, and resource template is present.
2. **In-process diagnostics** — `DataViewer.exe --self-test` exercises the registry round-trip, Postgres `db.conf` load + connection ping, bundled Python + openpyxl, openpyxl Excel write+read, AppData write permission, the Synology version-scan via `UpdateChecker::probe()`, and the offline-snapshot path/readability checks. Output is JSON at `%TEMP%\dataviewer_selftest.json`.
3. **Independent Synology probe** — walks the synced update folder directly with PowerShell, listing every version subdirectory and flagging ones missing `DataViewer-setup.exe`.

Phases 2 and 3 verify the same thing two ways on purpose: phase 2 runs the production code path inside `DataViewer.exe`, phase 3 walks the directory with PowerShell. If a regression in `UpdateChecker` ever masks a working folder layout (or vice versa), the two will disagree and tell us which side broke.

When adding a new deployment-sensitive code path, add a corresponding `TestResult test*()` in `src/utils/SelfTest.cpp` and document it in `tests/deployment/README.md`. Each test must finish in seconds, leave no on-disk side effects, and be independent of other tests.

What is intentionally **not** covered:
- The auto-update **download + replace** flow. Stops at "version was detected." Validate by performing an actual install when needed.
- Drag-and-drop file loading.
- `SingleInstance` IPC (would need to launch twice; left out for the v1.0.0 baseline).

## Architecture

### Mode-switching MainWindow

`MainWindow` (`src/MainWindow.h/.cpp`) is one window that runs three workflows by swapping the central widget and ribbon contents in place:

- **TPM mode** — analyse Total Particulate Matter measurement data from Excel sheets.
- **Sensory mode** — five-metric panel evaluation sessions with a radar chart.
- **Detailed Sensory mode** — extended 14-question sensory form with dual radar charts.

Two `QStackedWidget`s do the swapping. `m_centralStack` holds the TPM splitter, `SensoryPanel`, and `DetailedSensoryPanel` as separate pages. `m_navStack` holds the file/sheet `QTreeWidget` and the sensory session `QListWidget`. Mode switches go through `toggleSensoryMode()` / `toggleDetailedSensoryMode()` and `updateRibbonForMode()`, which hides/shows per-mode `QToolButton` sets on the ribbon. Sensory panels are constructed lazily on first activation.

### The Excel pipeline

Data flow is: `ExcelReader` → `DataProcessor::processFile` → `SheetProcessors` → `TpmCalculator` → `FileResult`.

- **`ExcelReader`** does *not* parse `.xlsx` directly. It writes a temp Python script, runs the bundled `python\python.exe` with openpyxl, and returns each sheet as a 2-D `QVector<QVector<QVariant>>`. The reasoning (per the header): this transparently handles Microsoft Information Protection / IRM / AIP encrypted workbooks because openpyxl runs inside the authenticated user session, and `data_only=True` returns cached values instead of formula strings.
- **`DataProcessor`** iterates sheets and dispatches to `SheetProcessors`, which detect "new" vs "old" template versions. Sheets named like "Test SOP's" are flagged `isRawTable` and rendered as plain tables with no plot.
- **`TpmCalculator`** computes per-row TPM, power density, variation, and per-sample aggregates.
- The canonical in-memory shape is `FileResult` → `QVector<SheetResult>` → `QVector<SampleResult>` → `QVector<DataRow>`, in `src/pipeline/ReportData.h`. Editing in the UI mutates these structs and calls `recalculateSampleMetrics(sheet)` before pushing changed cells back into the Excel file.

### Write-back to Excel

Cell edits don't write through immediately. `MainWindow` debounces them with `m_excelWriteTimer` and `m_pendingWrites`, then calls `writeCellsToExcel` / `flushExcelWrites`, which spawn another Python subprocess via openpyxl to mutate the source workbook. `m_modifiedFilePaths` tracks dirty files; the same path drives `markFileModified()` and the database sync indicator.

### Database

DataViewer Enterprise stores its data in **PostgreSQL 16** hosted in a
Docker container on the office Synology NAS. The schema is deployed by
`deploy/postgres/init.sql`; the container is defined by
`deploy/postgres/docker-compose.yml` + `deploy/postgres/Dockerfile`
(the Dockerfile adds `postgresql-16-cron` on top of the official base
image). See `deploy/postgres/README.md` for NAS admin setup.

The Qt app connects via the **QPSQL** driver, with `libpq.dll` + 4
transitive DLLs bundled by `build_installer.bat` from
`vendor/libpq-16/`. Connection settings live at
`%PROGRAMDATA%\DataViewer\db.conf` (set by the installer; password
field is encrypted with a machine-bound key — copying the file to
another workstation invalidates the password).

The cross-machine `<dbPath>.lock` sidecar mechanism is gone; concurrency
is handled via Postgres row-level optimistic locking (per-row `updated_at`
timestamps checked on UPDATE). When the NAS is unreachable, the app falls
back to a read-only local SQLite snapshot at
`%LOCALAPPDATA%/DataViewer/snapshot.sqlite`, regenerated on every clean
online close. NOTIFY-driven live updates, presence indicators, and the
three conflict-resolution dialogs (stale-row, unique-violation, row-deleted)
all shipped in v2.0. See the
[plan index](docs/superpowers/plans/2026-05-11-postgres-multiuser-INDEX.md)
for the migration history.

### Local test database

Run `tests\start-test-postgres.ps1` to spin up a throwaway `postgres:16`
container on port 5433 with the data schema applied. The script sets
`$env:DVE_TEST_PG_CONN` and prepends `vendor\libpq-16` to `PATH` so test
binaries auto-detect both. Tear down with `docker rm -f dve-test-pg`.

### Reporting

`ReportGenerator` builds PowerPoint reports without a PPTX library. It constructs OOXML manually using `XmlBuilder` and `ZipWriter` (a thin zlib-backed zip writer). `PptxWriter` emits per-slide XML; `PlotEngine` renders chart images embedded as PNG.

Reports come in two flavours per mode (Test/Single + Full); the active report buttons are re-labelled by `updateRibbonForMode()` when the user toggles modes.

### UI structure

- **`RibbonWidget`** (`src/widgets/`) is a hand-built ribbon (not Qt's built-in). Tabs are Home / Reports / View / Tools. Buttons use word-wrap with fixed 80×76 px sizing — preserve this when adding new buttons; without it labels like "Single Sensory Report" truncate.
- **`SensoryPanel`** uses a horizontal split (cards left, chart right) — intentionally different from TPM mode's vertical split. See `tasks/lessons.md`.
- **`DetailedSensoryPanel`** uses a unified 2-column grid (questions 1–14 numbered for reading order) above a 4-quadrant chart area. Combo-box widths capped (280 px combo / 220 px line edit / 70 px spin). Radar charts inverted: best score (1) maps to the outer ring (9).
- **`DataCleanupDialog`** lets the user mark individual rows as excluded per (file, sheet, sample); exclusions live in `MainWindow::m_excludedRows` keyed by `"fileIdx:sheetIdx:sampleIdx"` and feed `buildCleanedSample/Sheet/File` for reports.

### Auto-update

`UpdateChecker` (`src/utils/`) walks the Synology folder hourly, looks for newer installer folders, and prompts the user. Version-scan accepts both `1.0.0` and `v1.0.0` folder names. `latestAvailable()` runs on a background thread to avoid stalling the UI. The public `probe()` API exists for the deployment self-test — production code uses `check()` directly.

### Bundled Python

The installer ships an embedded Python 3.11.9 + openpyxl in `release\python_bundle.zip`, extracted to `{app}\python\` at install time. `MainWindow::findPython()` resolves the interpreter and caches the result. Bundled site-packages are stripped of pip/setuptools/wheel because Inno Setup's solid compression corrupts those binaries.

### Document Translator

`dataviewer_translator/` is an optional sub-app (Python + PyInstaller, packaged via its own `installer.iss`) that calls the Anthropic API to translate `.xlsx` / `.pptx` files. The main app's `onLaunchTranslator()` writes its API key into a config file the translator reads at startup. The main installer ships the prebuilt `DocumentTranslator.exe` from `dataviewer_translator\dist\`.

## Conventions

- All app code lives in the `DVE` namespace.
- Hard-coded Qt 6.10.x — the `.pro` does not currently fall back gracefully to other Qt 6 minor versions.
- Qt 6.10's `rcc` emits binary TSD output rather than C++ source, so `RESOURCES += resources/resources.qrc` is **commented out** in the `.pro`. Icons load from disk via `MainWindow::resourcePath()`. Don't re-enable the `RESOURCES` line without a workaround.
- `tasks/lessons.md` is a project-level lessons log — append a new entry every time a UI/architectural assumption is corrected.
- The build is `-Werror -Wall -Wextra`. Don't downgrade the warning level to silence a single warning; fix the underlying code.

---

## General CLAUDE Instructions

### Workflow Orchestration

#### 1. Plan Mode Default
- Enter plan mode for ANY non-trivial task (3+ steps or architectural decisions).
- If something goes sideways, STOP and re-plan immediately — don't keep pushing.
- Use plan mode for verification steps, not just building.
- Write detailed specs upfront to reduce ambiguity.

#### 2. Subagent Strategy
- Use subagents liberally to keep the main context window clean.
- Offload research, exploration, and parallel analysis to subagents.
- For complex problems, throw more compute at it via subagents.
- One task per subagent for focused execution.

#### 3. Self-Improvement Loop
- After ANY correction from the user: update `tasks/lessons.md` with the pattern.
- Write rules for yourself that prevent the same mistake.
- Ruthlessly iterate on these lessons until mistake rate drops.
- Review lessons at session start for the relevant project.

#### 4. Verification before done
- Never mark a task complete without proving it works.
- Diff behaviour between main and your changes when relevant.
- Ask yourself: "Would a staff engineer approve this?"
- Run tests, check logs, demonstrate correctness.

#### 5. Demand elegance (balanced)
- For non-trivial changes, pause and ask "is there a more elegant way?"
- If a fix feels hacky: "Knowing everything I know now, implement the elegant solution."
- Skip this for simple, obvious fixes — don't over-engineer.
- Challenge your own work before presenting it.

#### 6. Autonomous bug fixing
- When given a bug report, just fix it. Don't ask for hand-holding.
- Point at logs, errors, failing tests — then resolve them.
- Zero context switching required from the user.
- Go fix failing CI tests without being told how.

### Task Management

1. **Plan first** — write the plan to `tasks/todo.md` with checkable items.
2. **Verify the plan** — find root causes. No temporary fixes. Senior-developer standards.
3. **Minimal impact** — changes should only touch what's necessary. Avoid introducing bugs.

### Approach Guidelines

- Always prefer the simplest, most efficient approach first. Before implementing, state the approach and estimated time/complexity. Do NOT over-engineer or add encryption / abstraction / complexity the user didn't ask for.
- For bulk data operations, prefer batch / pagination approaches over one-by-one API calls.
- Do not modify files or features that weren't explicitly requested.

### Data Analysis

- Before analyzing any data, confirm which dataset and which metrics/columns to use. State the data source and row count before proceeding.
- Never assume which data file to use — ask, or verify against the user's previous instructions.

### Memory & Performance

- For large datasets (100K+ rows), always use streaming / chunked writes to avoid `MemoryError` / OOM.
- When saving checkpoints, use `fsync` and atomic writes (write to temp, rename) to prevent truncation.
- Test with small batches before scaling up.

### Codebase Exploration

- Do not spend an entire session exploring code without producing deliverables.
- After 5 minutes of exploration, summarize findings and either begin implementation or present a concrete plan.
