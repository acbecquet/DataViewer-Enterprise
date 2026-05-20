# DataViewer-Enterprise Codebase Cleanup Audit (v2.0.7)

**Date:** 2026-05-20
**Baseline:** `main` at v2.0.6 (commit `1991ab5`) + Pass 1 commits on `chore/v2.0.7-cleanup`
**Scope:** All five categories from the cleanup design — dead code, simplification, doc rot, tests, deploy/installer.

## Summary

| Category | High | Medium | Low | Total |
|---|---|---|---|---|
| Dead code (F) | 2 | 0 | 0 | **2** |
| Doc rot (D) | 25 | 0 | 0 | **25** |
| Tests (T) | 0 | 0 | 0 | **0** |
| Installer / Deploy (I) | 0 | 1 | 0 | **1** |
| Simplification (S) | 8 | 9 | 0 | **17** |
| **Total** | **35** | **10** | **0** | **45** |

**Mid-initiative checkpoint trigger:** Spec threshold was >40 simplification findings OR >5 high-risk simplifications. Actual is 17 simplification findings (8 High, 9 Medium). **No mid-initiative checkpoint needed** — proceed straight through Pass 3.

**Pass 3 execution plan (chosen disposition, see "Triage" below):**
- 25 doc-rot archive moves (low risk, mostly mechanical)
- 2 dead-code removals (F001 = SensoryDialog 1830 LOC; F002 = single-key findByKey variants ~50 LOC)
- 1 tools archive (db_integrity_check.py)
- ~6 low-risk simplifications (helper extraction, one-liner cleanups, XML-escape dedup)
- **Deferred to v2.0.8+:** all big file splits (MainWindow, DatabaseManager, OfflineSnapshot, SensoryPanel/DetailedSensoryPanel) — these are large focused refactors that don't belong in a cleanup release.

---

## Dead code (F)

| ID | Location | Description | Confidence | Disposition |
|---|---|---|---|---|
| F001 | `src/ui/SensoryDialog.cpp` + `src/ui/SensoryDialog.h` | Orphaned dialog class (1662+168 LOC). Not in `DataViewerEnterprise.pro`; zero `#include`s elsewhere; replaced by `SensoryPanel`. Also contains its own copy of `FlowLayout` + `SampleCard` (now duplicated with `SensoryPanel.h`). | **High** | **Remove** (archive copy first) |
| F002 | `DatabaseManager::findSensorySessionByKey` (`.h:180`, `.cpp:2554`) + `findDetailedSensorySessionByKey` (`.h:183`, `.cpp:2583`) | Single-key lookup variants replaced by the v2.0.6 bulk methods. Only reference in active code is a comment in `SensoryPanel.cpp:885`; no call sites. | **High** | **Remove** |

## Doc rot (D)

**25 plan/spec/handoff docs to archive — all describe work fully shipped in v1.3.0 through v2.0.6:**

| ID range | Files | Disposition |
|---|---|---|
| D001-D004 | `2026-03-26-sensory-fixes.md`, `2026-03-31-detailed-sensory-mode-design.md`, `2026-03-31-detailed-sensory-mode.md`, `2026-04-02-legacy-tpm-backwards-compat.md` | Archive (v1.3.0 work) |
| D005-D010 | `2026-04-08-auto-updater-design.md`, `2026-04-08-auto-updater.md`, `2026-04-08-translator-launcher-design.md`, `2026-04-08-translator-launcher.md`, `2026-04-10-optimization-plan-design.md`, `2026-04-10-optimization-plan.md` | Archive (v2.0.0 work) |
| D011-D014 | `2026-05-04-tpm-report-overhaul-design.md`, `2026-05-04-tpm-report-overhaul.md`, `2026-05-06-sensory-report-preview-design.md`, `2026-05-06-sensory-report-preview.md` | Archive (v2.0.4 work) |
| D015 | `2026-05-10-v1.3.0-undo-redo-snap-grid.md` | Archive (v1.3.0 work) |
| D016-D020 | All 5 postgres-multiuser docs (INDEX, plan-A, plan-B, plan-C, design) | Archive (v2.0.0 work) |
| D021-D022 | `2026-05-15-postwave-bugfix-batch-plan.md`, `2026-05-15-postwave-bugfix-batch-design.md` | Archive (v2.0.5 work) |
| D023-D024 | `2026-05-16-live-collab-plan.md`, `2026-05-16-live-collab-design.md` | Archive (v2.0.1 live-collab work; already shipped) |
| D025-D026 | `2026-05-20-codebase-cleanup-design.md` + `2026-05-20-codebase-cleanup.md` | **KEEP** — this is the active cleanup |
| D027 | `docs/handoff-2026-05-19.md` | Archive (v2.0.2 audit handoff; work shipped) |

**CLAUDE.md spot checks:** test count "27" if present should become "34"; spot-fix any references to deleted worktrees/branches if they're written as already-done.

**tasks/lessons.md:** all 5 lessons (sensory layout, Postgres-on-Windows, QTemporaryFile, MIP, pg_cron) remain accurate — **no changes**.

**README.md, CHANGELOG.md, deploy/postgres/README.md:** all accurate — **no changes**. CHANGELOG v2.0.7 entry added at the end of Pass 3.6.

## Tests (T)

**Zero findings.** 34/34 tests in `tests.pro`; all target classes exist in `src/`; no unconditional `QSKIP`; Python-availability and Postgres-availability probes follow a consistent guarded pattern. The MIP-encrypted xlsx fixture QSKIP added in Pass 1.7 is properly guarded.

## Installer / Deploy (I)

| ID | Location | Description | Confidence | Disposition |
|---|---|---|---|---|
| I001 | `tools/db_integrity_check.py` | SQLite diagnostics script. Useful but no callers (not in `build_installer.bat`, `tests/run-tests.ps1`, `installer.iss`, or any doc). | **Medium** | **Archive-only** |

`installer.iss`, `build_installer.bat`, `deploy/postgres/{Dockerfile,docker-compose.yml,README.md,migrations/}` are all clean and current.

## Simplification (S) — triaged

The audit surfaced 7 oversize-file analyses (S001-S007), each proposing 3-5 file splits + an "over-engineered patterns" list, plus 5 cross-file duplications. Total ~40 specific actions. After triage against "this is a cleanup release, not a re-architecture release":

### Will do in Pass 3.4

| ID | Action | LOC impact | Risk |
|---|---|---|---|
| S001a | Remove `MainWindow::writeCellToExcel` one-liner wrapper; callers use `writeCellsToExcel` with brace-init list | -10 | Low |
| S001b | Remove empty zoom/view stub methods (`onZoomIn`, `onZoomOut`, `onFitToWindow`, `onViewDataTable`, `onViewPlots`, `onViewBoth`) if confirmed empty | -10 | Low |
| S005 | **Delete SensoryDialog.cpp + .h entirely** (overlaps F001) | -1830 | Low (zero callers) |
| D001 | Extract `lastBrowseDir`/`setLastBrowseDir` to `src/utils/FileDialogHelpers.h/.cpp`; replace in 3 sites (MainWindow, SensoryPanel, DetailedSensoryPanel) | ~-30 net | Low |
| D005 | Replace hand-rolled XML escape in `PptxWriter::makeTextBox` and `makeTableCell` with `XmlBuilder::escapeXml` | -6 | Low |
| S006c | Move detailed-sensory JSON deserialization to a single function in `src/pipeline/DetailedSensoryData.cpp` (D002); collapse the two duplicate sites in `DatabaseManager.cpp` and `OfflineSnapshot.cpp` | ~-40 net | Medium |

**Estimated total LOC reduction for Pass 3.4: ~1925 lines (mostly from F001/S005 SensoryDialog deletion).**

### Deferred — out of scope for v2.0.7

These are each substantial focused refactors that should ship as standalone PRs after v2.0.7, not bundled into a "cleanup release":

| ID | Deferred action |
|---|---|
| S001 (rest) | MainWindow.cpp file splits (Ribbon, Presence, LiveSync, Excel, Images extraction). 5400-line file is bad, but splitting touches every ribbon button, LiveSync handler, and image path. High-blast-radius work. |
| S002 (rest) | DatabaseManager.cpp file splits (File/Sensory/DetailedSensory/Layout repositories). Touches every DB write path. |
| S003 (rest) | SensoryPanel.cpp splits (IO, Reports, LiveSync subfiles). |
| S004 (rest) | DetailedSensoryPanel.cpp splits (IO, Reports, LiveSync subfiles). |
| S006 (rest) | OfflineSnapshot.cpp splits (Schema, Regenerate, PendingEditQueue subfiles). |
| S007 (rest) | PptxWriter.cpp splits (Package, Shapes, Slides subfiles). |
| S004 partial | `buildQuestionForm` 220-line widget construction → loop-driven `addMetricRow` helper. Useful but UI-touching. |
| S002 partial | Three-site optimistic-update pattern dedup via `executeOptimisticUpdate` template. Useful but touches every DB write. |

These deferrals should be tracked in a `docs/superpowers/specs/2026-05-21-post-v2.0.7-refactors.md` follow-up doc — created at the end of Pass 3.6.

## Preserved symbols (do NOT remove)

Per the v2.0.6 hotfix doc, these are load-path workhorses:

- `DatabaseManager::NaturalKey` (struct)
- `DatabaseManager::SessionKeyMatch` (struct)
- `DatabaseManager::findSensorySessionsByKeys()` (bulk lookup)
- `DatabaseManager::findDetailedSensorySessionsByKeys()` (bulk lookup)
- `SensoryPanel::inheritExistingIdsAndVersions()` + its call from `SensoryPanel::loadSessions()`
- `DetailedSensoryPanel::inheritExistingIdsAndVersions()` + its call from `DetailedSensoryPanel::loadSessions()`

The audit confirmed the single-key variants (`findSensorySessionByKey`, `findDetailedSensorySessionByKey`) ARE unused and ARE candidates for removal (F002) — only the bulk variants are preserved.

## Pass 3 execution order

1. **Phase 3.1 (doc rot):** Archive D001-D024 + D027 to `DataViewer-Archive/2026-05-20-cleanup-pass-3-doc-rot/`. Edit CLAUDE.md spot fixes if needed. ~25 file moves.
2. **Phase 3.2 (test cleanup):** SKIP — audit found nothing to do.
3. **Phase 3.3 (dead-code):** Archive + delete SensoryDialog (F001/S005). Delete single-key findByKey (F002). Archive db_integrity_check.py (I001). Compile + test + self-test.
4. **Phase 3.4 (simplification, low-risk only):** S001a, S001b, D001, D005, S006c. Compile + test + self-test after each.
5. **Phase 3.5 (installer):** No further work — installer is clean.
6. **Phase 3.6 (polish):** VERSION bump 2.0.6 → 2.0.7, CHANGELOG entry, write `2026-05-21-post-v2.0.7-refactors.md` capturing deferrals, clean rebuild, installer build, surface to user for frontend smoke.

## Risk profile

- **Doc rot:** Pure markdown moves. Cannot break the build.
- **Dead-code (SensoryDialog):** File is not in `.pro` — already not in the binary. Removing source files is purely cosmetic. Zero risk of behavior change.
- **Dead-code (single-key findByKey):** Only mentioned in a comment; comment gets updated. Zero call-site risk.
- **`lastBrowseDir` extraction:** 3 sites, very narrow surface; smoke gate catches regressions.
- **XML-escape dedup:** Replaces hand-rolled `replace()` chain with the existing helper that handles a superset. If anything, more correct.
- **JSON deserialize consolidation:** Two sites collapse to one. Tests cover the load/save path; smoke gates catch regressions.

Confidence that Pass 3 completes without behavior regression: **high**. The biggest single change (deleting SensoryDialog) is also the lowest-risk because the file isn't in the build.
