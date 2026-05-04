# DataViewer Enterprise — Comprehensive Optimization Design

**Date:** 2026-04-10
**Status:** Approved
**Approach:** Inside-Out (Low-Risk → High-Risk), 5 phases with testing gates

## Constraints

- **No input/output changes** — all external behavior stays identical
- **Python subprocess stays** — company encryption requires Python's unobstructed permissions for Excel/PPTX
- **Excel load/write pipeline untouched** — same openpyxl + COM fallback, same QXlsx write path
- **PPTX generation untouched** — same XML templates, same ZIP structure, same slide output
- **Database schema unchanged** — no migrations, no new tables
- **Sensory metric ordering preserved** — Burnt Taste first in data entry, Overall Liking at 12 o'clock in plots
- **Executable naming preserved** — DataViewer.exe and DataViewer-setup.exe names must never change

## Codebase Baseline

- **~22K LOC** across 27 C++ modules + headers
- **6 architectural layers:** Presentation, Business Logic, Data, Rendering, Utility, Infrastructure
- **Qt 6.10.1** / MinGW GCC 13.1.0 / qmake / Inno Setup 6
- **10 test suites** covering core business logic and utilities
- **MainWindow.cpp: 3,668 lines, 82 slots** (God Object — primary refactoring target)

---

## Phase 1: Build & Infrastructure

**Goal:** Tighten compiler settings to catch more bugs at compile time, improve binary quality. Zero behavioral change.

### Changes

1. **DataViewerEnterprise.pro — optimized release flags:**
   - `-O3` (aggressive optimization — inlining, vectorization, loop unrolling)
   - `-flto` (link-time optimization — cross-file inlining, dead code elimination)
   - `-Werror` (treat warnings as errors)
   - `-Wpedantic` (catch non-standard C++ usage)
   - Remove duplicate `-Wextra` flag
   - Add `-Wno-deprecated-declarations` if needed for Qt compatibility

2. **Makefile regeneration** — `qmake` regenerates with new flags

3. **Installer (installer.iss)** — switch to `lzma2/ultra64` compression for smaller installer

### Testing Gate

- Full rebuild with new flags — fix any warnings surfaced by `-Werror`
- Run all 10 test suites
- Manual smoke test: load Excel, edit cells, generate PPTX, switch modes
- Compare output PPTX structure against known-good report

### Risk: Near-zero

---

## Phase 2: Security Fixes

**Goal:** Fix critical and high-severity vulnerabilities without changing external behavior. All fixes are internal — same inputs, same outputs.

### Critical

1. **SQL injection in `database_explorer.py`** (lines 158, 270)
   - Replace `.format()` and f-string SQL with parameterized queries (`?` placeholders)
   - Same query results, safe construction

2. **Hardcoded network paths in `database_manager.py`** (line 22)
   - Move Synology IP and share path to a JSON config file (`config.json` alongside the app)
   - App reads from config instead of source code; ships with default config

3. **API key storage in `MainWindow.cpp`** (lines 3603, 3655)
   - Replace Base64 encoding with Windows DPAPI encryption (`CryptProtectData`)
   - Existing config.dat migrated on first run (read old Base64, re-save encrypted)

### High

4. **Temp file permissions in `ExcelReader.cpp`** (line 1151)
   - Set `QFile::ReadOwner | QFile::WriteOwner` (0600) on Python script temp file

5. **XML entity escaping in `PptxWriter.cpp`** (line 773)
   - Add single-quote escaping (`'` → `&#39;`)

6. **BLOB size validation in `DatabaseManager.cpp`** (line 467)
   - Add 100MB cap on image file reads before `readAll()`

7. ~~**DB transaction wrapping**~~ — ALREADY IMPLEMENTED (line 285, with rollback on every failure)

8. ~~**Python subprocess timeout**~~ — ALREADY IMPLEMENTED (60s at line 1169)

### Medium

9. **Error message sanitization in `MainWindow.cpp`** (lines 1296-1410)
   - Log full error via `qDebug()`, show generic message to user

10. **Hardcoded dev path in `main.cpp`** (line 45)
    - Remove `C:/Users/S1134987/...` fallback from icon candidates

### Testing Gate

- Full rebuild + all 10 test suites
- Test Excel load with encrypted company files (verify Python path still works)
- Test PPTX generation — compare output structure against known-good report
- Test API key save/load round-trip (old config migration + new encryption)
- Test DB save with simulated mid-insert failure (verify rollback)
- Manual smoke test across all three modes (TPM, Sensory, Detailed Sensory)

### Risk: Low (isolated, targeted changes)

---

## Phase 3: Data Layer Hardening

**Goal:** Make the data layer more robust and defensive. No external behavior changes.

### Database

1. **Audit all `QSqlQuery` calls** — ensure all use `prepare()` + `addBindValue()`
2. ~~**Add `PRAGMA journal_mode=WAL`**~~ — ALREADY IMPLEMENTED (line 53, with network-share detection)
3. ~~**Make `PRAGMA busy_timeout` configurable**~~ — ALREADY IMPLEMENTED (line 51)

### Input Validation

4. **Add length/format validation at system boundaries:**
   - Sample names: max 255 chars, alphanumeric + spaces + hyphens
   - Tester/assessor names: max 255 chars
   - File paths: validate canonical path before processing
   - Applied in `SensoryPanel`, `DetailedSensoryPanel`, `HeaderEditDialog`, `DataCleanupDialog`

5. **Image file validation before DB storage:**
   - Verify magic bytes (PNG: `89 50 4E 47`, JPEG: `FF D8 FF`)
   - Reject non-image files with user-friendly message

### Python Subprocess

6. **Cache Python script for session lifetime**
   - Write once on first use, reuse same temp file path
   - Same script, same behavior, fewer disk writes

7. **Validate Python executable path** before `QProcess::start()`
   - Check `QFileInfo(pythonExe).isExecutable()`
   - Fail with clear error message

### Testing Gate

- Full rebuild + all 10 test suites
- Test DB operations: save, load, delete, re-save — verify data integrity
- Test WAL mode: concurrent read scenarios
- Test input validation edge cases (empty, 300 chars, `O'Brien`, unicode)
- Test Python subprocess: valid path, invalid path, missing Python, timeout
- Test image validation: valid PNG, valid JPEG, corrupt file, oversized file
- Full end-to-end smoke test (load Excel → edit → save → generate PPTX → verify)

### Risk: Low-medium

---

## Phase 4: MainWindow Decomposition

**Goal:** Break the 3,668-line God Object into 6 focused controller classes. Every signal, slot, and side effect stays identical.

### Controllers

| Controller | Responsibility | Est. Lines |
|-----------|----------------|-----------|
| `ReportController` | Report generation orchestration, progress, config dialogs | ~300 |
| `ImageController` | Image inbox, attach, view dialogs | ~300 |
| `ModeController` | TPM ↔ Sensory ↔ Detailed transitions, stack/dock switching | ~350 |
| `FileController` | File load, file tree, drag-drop, recent files, `m_loadedFiles` | ~500 |
| `ExcelController` | Cell edits, debounced writes, QXlsx management | ~400 |
| `DbSyncController` | Auto-save timer, DB sync, status indicator | ~250 |

**MainWindow after extraction: ~1,200 lines** — window setup, widget ownership, signal routing, settings.

### Extraction Order (safest first)

1. **ReportController** — most isolated, calls ReportGenerator/PptxWriter
2. **ImageController** — self-contained, minimal coupling
3. **ModeController** — touches UI state but not data
4. **FileController** — core data loading, after mode switching stable
5. **ExcelController** — depends on FileController's data
6. **DbSyncController** — depends on ExcelController, extracted last

### Decomposition Strategy

- Controllers are `QObject` subclasses (for signal/slot)
- Receive widget/data pointers via constructor injection
- MainWindow remains parent (Qt memory model — auto-deleted)
- One controller at a time: extract → compile → test → commit → next
- Each extraction is its own git commit (independently revertible)

### Review Protocol (Triple-Review Per Controller)

**Step 1: Extract & Compile**
- Move slot code to new controller, update wiring, compile

**Step 2: Automated Review**
- Run all 10 test suites
- Code-reviewer agent diffs extraction against original — verify no logic altered

**Step 3: Manual Regression**
- Exercise every feature the controller owns (specific checklist)

**Step 4: Cross-Review**
- Re-test all previously extracted controllers (catches interaction bugs)

**Step 5: Snapshot Comparison**
- Before Phase 4: generate "golden" PPTX report
- After each extraction: generate same report, diff XML structure
- Same test for Excel writes: edit cell, compare .xlsx before/after

### Rollback Plan

- Each extraction is its own git commit
- If regression found: `git revert` that single commit, all others intact

### Testing Gate (Per Controller + Final)

Per controller:
- Full rebuild + all 10 test suites
- Controller-specific regression test
- Cross-review of all previously extracted controllers
- PPTX snapshot comparison

Final (all 6 extracted):
- Full end-to-end smoke test of every feature
- Load Excel, edit cells, generate PPTX, compare against golden
- All three modes (TPM, Sensory, Detailed Sensory)
- Image attach/view, database browse, data cleanup

### Risk: Medium (mitigated by triple-review + per-commit rollback)

---

## Phase 5: Frontend Polish

**Goal:** Safe rendering efficiency improvements and signal safety. No visual changes.

### Plot Caching

1. **PlotWidget: cache rendered pixmap**
   - Add `m_plotDirty` flag, only re-render on data change
   - Same visual output, fewer redundant renders

2. **RadarChartWidget: cache polygon geometry**
   - Compute vertices on `setSessions()` only
   - `paintEvent()` draws from cache instead of recomputing

### Signal Safety

3. **Comprehensive `blockSignals()` guards**
   - Audit all programmatic combo/table updates
   - Ensure consistent `blockSignals(true)` wrapping

4. **Null-check signal senders in lambda captures**
   - Guard checks for widget pointers in lambdas

### Minor Rendering

5. **PlotEngine: pre-compute axis tick positions**
   - Compute `niceStep()` / `autoRange()` once, pass to grid + axes + labels
   - Same output, ~20% fewer FP operations per render

### Testing Gate

- Full rebuild + all 10 test suites
- Visual comparison: screenshot plots before/after (pixel-match expected)
- Stress test: rapidly toggle checkboxes, switch plot types, resize window
- Radar chart: verify overlap rendering with cached polygons
- Signal safety: rapid mode/sheet/cell switching — no freezes or loops
- Full end-to-end smoke test

### Risk: Low

---

## Summary

| Phase | Focus | Risk | Key Constraint |
|-------|-------|------|---------------|
| 1 | Build flags | Near-zero | Same binary behavior |
| 2 | Security | Low | No pipeline changes |
| 3 | Data layer | Low-medium | No schema changes |
| 4 | MainWindow decomposition | Medium | Identical behavior, triple-review |
| 5 | Frontend polish | Low | Same pixels on screen |

**Total estimated scope:** ~2,100 lines moved/refactored, ~500 lines new (controllers, validation, caching), ~50 lines removed (hardcoded paths, duplicates).

**Testing philosophy:** Every phase ends with a full test suite run + manual smoke test. Phase 4 adds triple-review per controller extraction with PPTX/Excel snapshot comparison.
