# v2.0.6 Hotfix Context — Ctrl+U Freeze Fix

**Status:** Shipped 2026-05-20. Tag `v2.0.6` on `origin/main` at commit `1991ab5`. Installer staged at `dist\DataViewer-setup.exe` (ProductVersion 2.0.6, 74 MB); user transfers to Synology manually per the no-auto-Synology-drop standing instruction.

---

## TL;DR for the cleanup overhaul session

1. **Bump the cleanup target version from v2.0.6 → v2.0.7** everywhere in `docs/superpowers/specs/2026-05-20-codebase-cleanup-design.md`. v2.0.6 has been consumed by this hotfix.
2. **Re-base the cleanup branch from `main` at `v2.0.6` (commit `1991ab5`)**, not v2.0.5 (`fac258f`).
3. **Do not flag** the new `NaturalKey`, `SessionKeyMatch`, `findSensorySessionsByKeys`, or `findDetailedSensorySessionsByKeys` symbols as dead code in Pass 2 — they are the load-path workhorses introduced by this hotfix.
4. **Add a Pass-1 step** to fix the MIP auto-relabel of `.xlsx` test fixtures (2 pre-existing test failures, see "Test suite state" below).
5. The `next` branch has already been rebased so the cleanup spec sits cleanly on top of v2.0.6 — no manual git surgery needed before starting the cleanup work.

---

## Why the hotfix was needed

External users on slower LAN segments saw Windows "Not Responding" freezes on every Ctrl+U and every 5-second auto-save tick. v2.0.5 (commit `fac258f`) wired `inheritExistingIdsAndVersions()` into `MainWindow::onUpdateDatabase()`, which fired **N synchronous Postgres `SELECT`s on the UI thread** (one per loaded session) before each save loop. On slower LANs the cumulative round-trip time exceeded Windows' ~5 s "Not Responding" threshold; concurrent multi-user Ctrl+U amplified the freeze via NAS Postgres I/O saturation.

The user's initial guess was "concurrency problem". That turned out to be a red herring: `LiveSyncWorker` opens its own `QSqlDatabase` connection (`src/database/LiveSyncWorker.cpp:51`), so there is no thread-shared `QSqlDatabase` between the UI thread and the worker. The actual problem was plain blocking I/O on the UI thread.

The autosave tick is wired in `MainWindow.cpp:333-336`:

```cpp
m_dbSaveTimer = new QTimer(this);
m_dbSaveTimer->setInterval(5000);
connect(m_dbSaveTimer, &QTimer::timeout, this, [this]() { onUpdateDatabase(); });
```

…so every user with a live Postgres connection paid the cost every 5 s, not just on explicit Ctrl+U.

---

## What changed in v2.0.6 (commit `1991ab5`)

**Two-layer fix:**

1. **Load-path placement.** `inheritExistingIdsAndVersions()` is now called from `SensoryPanel::loadSessions` and `DetailedSensoryPanel::loadSessions` — the only points where new `id == -1` sessions actually enter panel state. After first save, ids stay set and the lookup would be a no-op anyway. Removed both calls from `MainWindow::onUpdateDatabase()` (lines that were at `MainWindow.cpp:4031-4032`).

2. **Bulk SELECT.** New `DatabaseManager::findSensorySessionsByKeys` and `findDetailedSensorySessionsByKeys` take a vector of natural keys and do one round-trip per chunk of 200 via `JOIN (VALUES …) AS k(session_name, tester_name, date)`. Replaces the per-session `findSensorySessionByKey` loops in both panel inherit methods.

The v2.0.5 UNIQUE-violation fix on re-import is **preserved** — inherit still runs on file import, just not on every save tick. Single-key `findSensorySessionByKey` / `findDetailedSensorySessionByKey` variants were **kept** in the header for any future direct caller and to keep the blast radius small.

---

## Files touched

| File | Change |
|---|---|
| `DataViewerEnterprise.pro` | `VERSION = 2.0.5` → `VERSION = 2.0.6` |
| `src/database/DatabaseManager.h` | Added `NaturalKey` + `SessionKeyMatch` structs and 2 bulk method decls |
| `src/database/DatabaseManager.cpp` | Implemented file-static `findSessionsByKeysImpl` + 2 public bulk wrappers |
| `src/ui/SensoryPanel.cpp` | Rewrote `inheritExistingIdsAndVersions()` to bulk; added `inheritExistingIdsAndVersions()` call at end of `loadSessions` |
| `src/ui/DetailedSensoryPanel.cpp` | Same as above for the detailed panel |
| `src/MainWindow.cpp` | Removed the two inherit calls from `onUpdateDatabase()`; left an explanatory comment |

**208 insertions, 45 deletions across 6 files.** Diff: `git diff fac258f..1991ab5`.

---

## New public API introduced (for cleanup audit awareness)

In `src/database/DatabaseManager.h`:

```cpp
// v2.0.6 — DO NOT REMOVE in Pass-2 audit. Load-path workhorses.
struct NaturalKey {
    QString sessionName;
    QString testerName;
    QString date;
};
struct SessionKeyMatch {
    QString sessionName;
    QString testerName;
    QString date;
    qint64  id      = -1;
    int     version = 0;
};
QVector<SessionKeyMatch>
    findSensorySessionsByKeys(const QVector<NaturalKey>& keys) const;
QVector<SessionKeyMatch>
    findDetailedSensorySessionsByKeys(const QVector<NaturalKey>& keys) const;
```

Single-key variants `findSensorySessionByKey` and `findDetailedSensorySessionByKey` also remain. They are no longer called from `SensoryPanel.cpp` or `DetailedSensoryPanel.cpp` — if the Pass-2 audit confirms zero call sites and the cleanup wants to delete them, that is acceptable.

---

## Git topology (after hotfix)

```
* b84d1e1 (HEAD -> next, origin/next)
|         docs(spec): codebase cleanup design for v2.0.6   ← rename to v2.0.7
* 1991ab5 (tag: v2.0.6, origin/main, origin/hotfix/v2.0.6-ctrlu-freeze)
|         fix(mw): eliminate Ctrl+U freeze via bulk lookup + load-path inherit (v2.0.6)
* fac258f (tag: v2.0.5)
|         fix(db): inherit existing IDs on file re-import (UNIQUE-violation fix)
...
```

**Operational notes for the cleanup session:**

- `next` was rebased onto `1991ab5`, so the cleanup spec sits cleanly on top of v2.0.6. The spec's commit SHA changed (`f03c81e` → `b84d1e1`) but its content is intact.
- The `hotfix/v2.0.6-ctrlu-freeze` branch is preserved on origin for traceability.
- **Local `main` is stale** at `fac258f` because the worktree at `.claude\worktrees\v2.0.2-fixes` has it checked out. Pass 1.1 of the cleanup (kill stale worktrees) will resolve this. Until then, use `origin/main` (= `1991ab5`) as the source of truth.
- The original cleanup spec (now at `b84d1e1`) assumed it would branch from v2.0.5 at `fac258f`. **Update Pass 1.1 to branch from `v2.0.6` at `1991ab5`** so the cleanup includes (rather than re-implements) the freeze fix.

---

## Test suite state

- **32 passed, 2 failed, 0 skipped** as of 2026-05-20 14:21 MST.
- Failures:
  - `tst_sopLoader::loadsKnownTemplate` — `'rows.size() >= 3' returned FALSE. (expected >= 3 SOP rows, got 0)`
  - `tst_reportgenerator::loadSopRows_filtersToRequestedTests` — `Compared values are not the same`
- Both fail because their `.xlsx` test fixtures (e.g. `resources/sops.xlsx`) are MIP-encrypted on this machine, and the MIP auto-labeler **re-encrypts the file faster than Python can strip the label** (verified: Python read-then-write rewrites the file to plaintext, but a `head -c 16` immediately afterward shows `%TSD-Header-###%` again). C++ tests via QXlsx / direct file I/O see ciphertext → zero rows loaded.
- **NOT caused by the v2.0.6 hotfix** — neither test touches any code modified in `1991ab5`.
- The S222 memory note ("34 passed, 0 failed" at v2.0.5 release time) must have been recorded during a brief window when the MIP auto-labeler had not yet processed the fixtures — likely right after a fresh `git checkout` reset their mtime.

**Candidate fix for the cleanup overhaul (v2.0.7):** extend `tools/decrypt_via_copy.py` (or add a sibling) to handle `.xlsx` fixtures by copying them out to a non-MIP-tracked path (e.g., `%TEMP%\dve_test_fixtures\`) before each test run, then pointing tests at the temp copy via an env var like `DVE_TEST_FIXTURES_DIR`. This is now a v2.0.7 cleanup item — add to the spec.

---

## Required edits to `2026-05-20-codebase-cleanup-design.md`

The cleanup overhaul session should make these edits to the spec **before** starting Pass 1:

1. **Bump the cleanup version label from v2.0.6 → v2.0.7** everywhere in the spec body (spec title, "deliverables" header, "Why this matters now" section, the `.pro` VERSION bump step, and any cross-references).
2. **Rename the cleanup branch** from `chore/v2.0.6-cleanup` to `chore/v2.0.7-cleanup`.
3. **Update Pass 1.1 baseline:** branch from `main` at tag `v2.0.6` (commit `1991ab5`), not v2.0.5 (`fac258f`).
4. **Add an explicit "do not touch" line** in Pass 2 audit guidance covering the new v2.0.6 symbols:
   - `DatabaseManager::NaturalKey`
   - `DatabaseManager::SessionKeyMatch`
   - `DatabaseManager::findSensorySessionsByKeys`
   - `DatabaseManager::findDetailedSensorySessionsByKeys`
   - The two `inheritExistingIdsAndVersions()` calls inside `SensoryPanel::loadSessions` and `DetailedSensoryPanel::loadSessions`
5. **Add a Pass 1 step (or extend Pass 1.8 / add Pass 1.9)** to address the MIP-relabel xlsx-fixture issue described in "Test suite state" above. This will get the test suite back to 34/34 green before the deep-audit pass begins.
6. **Optional cleanup target:** the inherit-helper bodies in `SensoryPanel.cpp` (lines 882-933) and `DetailedSensoryPanel.cpp` (lines 796-841) are now near-duplicates. If Pass 2 wants to factor them into a free function template `inheritFromDb<SessionT>(QVector<SessionT>& sessions, auto bulkFn)`, that's a clean refactor — but optional, not required.

---

## How to start the cleanup session that picks this up

Open a fresh Claude Code session in this repo and paste:

```
Read these two files in order, then proceed:

1. C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise\docs\superpowers\specs\v2_0_6_hotfix.md
2. C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise\docs\superpowers\specs\2026-05-20-codebase-cleanup-design.md

Update the cleanup spec for v2.0.7 per the hotfix doc's "Required edits" section,
then begin Pass 1 of the cleanup workflow.
```
