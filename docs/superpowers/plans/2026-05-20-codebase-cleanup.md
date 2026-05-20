# Codebase Cleanup Implementation Plan (v2.0.7)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Take a stable v2.0.6 codebase and remove every form of accumulated cruft (stale git refs, generated files committed by accident, dead source code, over-engineered abstractions, doc rot) without losing any user-visible functionality. Ship as v2.0.7.

**Architecture:** Two-pass approach — mechanical wins first (no audit), then a deep audit producing one document, then phased source execution gated by tests + comprehensive `--self-test` smoke. Source removals are archived to `C:\Users\S1134987\Documents\Python\DataViewer-Archive\`.

**Tech Stack:** C++17 / Qt 6.10.1 / MinGW 13.1.0 / qmake / PostgreSQL 16 / Inno Setup 6. Windows-only. PowerShell or Git Bash. MIP source-file encryption requires `python tools/decrypt_via_copy.py --apply` before every build.

---

## File Structure

**Created:**
- `docs/audits/2026-05-20-cleanup-audit.md` — full audit findings (Pass 2 output)
- `tools/stage_test_fixtures.py` — MIP-relabel xlsx workaround for test fixtures
- New methods in `src/utils/SelfTest.cpp` — comprehensive smoke harness
- `DataViewer-Archive/2026-05-20-cleanup-pass-1/`, `…pass-3-doc-rot/`, `…pass-3-deadcode/`, `…pass-3-simplification/` — archived removals with per-folder `notes.md` and `README.md` at archive root

**Modified:**
- `.gitignore` — add patterns to keep build artifacts and generated files out
- `tests/run-tests.ps1` — call staging script before launching suite
- `tests/tst_sopLoader/tst_sopLoader.cpp` — honor `DVE_TEST_FIXTURES_DIR`
- `tests/tst_reportgenerator/tst_reportgenerator.cpp` — honor `DVE_TEST_FIXTURES_DIR`
- `DataViewerEnterprise.pro` — `VERSION = 2.0.6` → `2.0.7` (Phase 3.6)
- `CHANGELOG.md` — v2.0.7 entry
- Files identified in Pass 2 audit (TBD at execution time)

**Deleted:**
- 3 worktrees and ~7 stale local branches and ~6 stale remote branches
- Repo-root build artifacts: `Makefile*`, `test_rcc_output.cpp`, `.qmake.stash`, `build*/`, `debug/`, `release/`
- `DATAVIEWER_UPDATES.txt` (archived first)
- `anthropic_api_key.txt` (sanitized archive if real key; rotate via user first)

---

## Pass 1 — Mechanical Cleanup

### Task 1.1: Set up the cleanup branch

**Files:**
- No file changes; pure git topology setup

- [ ] **Step 1: Verify hotfix doc and current branch state**

Run: `git branch --show-current && git log --oneline -3`
Expected: Current branch is `next`, top commit is the spec retarget commit (9b9e307 or similar).

- [ ] **Step 2: Remove worktrees**

Run:
```bash
git worktree remove --force ".claude/worktrees/cranky-hofstadter-afa8a2"
git worktree remove --force ".claude/worktrees/v2.0.2-fixes"
git worktree list
```
Expected: Only the main repo path remains in the worktree list.

- [ ] **Step 3: Fast-forward local main to origin/main**

Run:
```bash
git fetch origin
git branch -f main origin/main
git log --oneline main -3
```
Expected: Local `main` now points at `1991ab5` (tag `v2.0.6`).

- [ ] **Step 4: Create the cleanup branch and switch to it**

Run:
```bash
git checkout -b chore/v2.0.7-cleanup main
git log --oneline -3
```
Expected: HEAD is on `chore/v2.0.7-cleanup` at `1991ab5`.

- [ ] **Step 5: Cherry-pick the two spec commits from `next`**

Run:
```bash
git log next --oneline main..next
git cherry-pick <SHA-of-spec-creation> <SHA-of-spec-retarget>
git log --oneline -5
```
Expected: Both spec commits now sit on top of `1991ab5` on `chore/v2.0.7-cleanup`.

- [ ] **Step 6: Commit (none needed — cherry-picks already committed)**

Verification: `git status` shows clean working tree.

### Task 1.2: Delete stale local branches

**Files:** None (git ref deletion only)

- [ ] **Step 1: Verify each branch is merged to main**

Run:
```bash
for b in release/v2.0.3 release/v2.0.4 hotfix/v2.0.5-logging hotfix/v2.0.6-ctrlu-freeze feat/sensory-header-presets fix/ctrl-s-modes worktree-v2.0.2-fixes claude/cranky-hofstadter-afa8a2; do
  git merge-base --is-ancestor "$b" main && echo "$b: MERGED" || echo "$b: NOT MERGED"
done
```
Expected: All report MERGED. If any reports NOT MERGED, stop and surface to user.

- [ ] **Step 2: Delete merged branches**

Run:
```bash
git branch -d release/v2.0.3 release/v2.0.4 hotfix/v2.0.5-logging hotfix/v2.0.6-ctrlu-freeze feat/sensory-header-presets fix/ctrl-s-modes worktree-v2.0.2-fixes claude/cranky-hofstadter-afa8a2
git branch
```
Expected: Only `main`, `chore/v2.0.7-cleanup`, `next` remain locally. `-d` (not `-D`) refuses to delete unmerged branches.

- [ ] **Step 3: Delete `next` after confirming spec commits are on cleanup branch**

Run:
```bash
git log chore/v2.0.7-cleanup --oneline | grep -E "(cleanup design|retarget cleanup)"
git branch -d next
git branch
```
Expected: Both spec commits appear in `chore/v2.0.7-cleanup`. After deletion, only `main` and `chore/v2.0.7-cleanup` remain locally.

### Task 1.3: Prune stale remote branches

**Files:** None

- [ ] **Step 1: Prune origin refs**

Run:
```bash
git remote prune origin
git branch -r
```
Expected: Stale tracking refs gone.

- [ ] **Step 2: Delete remote feature branches**

Run:
```bash
git push origin --delete feat/sensory-report-preview fix/sensory-cumulative-per-test feat/tpm-report-overhaul feat/deployment-self-test optimize/comprehensive-review worktree-v2.0.2-fixes hotfix/v2.0.6-ctrlu-freeze next
git branch -r
```
Expected: Only `origin/HEAD`, `origin/main`, and any branches we explicitly kept remain.

- [ ] **Step 3: Push cleanup branch to origin**

Run:
```bash
git push -u origin chore/v2.0.7-cleanup
```
Expected: New upstream tracking on origin.

### Task 1.4: Remove repo-root generated artifacts

**Files:**
- Delete: `Makefile`, `Makefile.Debug`, `Makefile.Release`, `test_rcc_output.cpp`, `.qmake.stash`
- Delete: `build/`, `build-main-review/`, `build-release/`, `build-tests-review/`, `debug/`, `release/`
- Modify: `.gitignore`

- [ ] **Step 1: Confirm what's tracked**

Run:
```bash
git ls-files Makefile Makefile.Debug Makefile.Release test_rcc_output.cpp .qmake.stash
git ls-files build/ build-main-review/ build-release/ build-tests-review/ debug/ release/ | head
```
Expected: Lists the tracked files. (.qmake.stash may or may not be tracked.)

- [ ] **Step 2: Remove from git index, keep on disk for safety**

Run:
```bash
git rm --cached Makefile Makefile.Debug Makefile.Release test_rcc_output.cpp
git rm -r --cached build build-main-review build-release build-tests-review debug release 2>/dev/null
git rm --cached .qmake.stash 2>/dev/null
```
Expected: Files removed from index. Working tree retains them.

- [ ] **Step 3: Update `.gitignore`**

Add to `.gitignore` (preserve existing entries):
```
# Generated build artifacts (qmake + MinGW)
Makefile
Makefile.Debug
Makefile.Release
.qmake.stash
test_rcc_output.cpp

# Build output directories
/build/
/build-*/
/debug/
/release/
```

- [ ] **Step 4: Delete from working tree**

Run:
```bash
rm -f Makefile Makefile.Debug Makefile.Release test_rcc_output.cpp .qmake.stash
rm -rf build build-main-review build-release build-tests-review debug release
ls -la | grep -E "Makefile|test_rcc|build|debug|release" || echo "All cleared"
```
Expected: `All cleared`. (The `release/` directory recreated by future builds will be re-ignored.)

- [ ] **Step 5: Commit**

Run:
```bash
git add -A .gitignore
git status
git commit -m "$(cat <<'EOF'
chore(repo): remove generated build artifacts from tree

Makefiles, test_rcc_output.cpp, build dirs, debug/, release/
were committed by accident. .gitignore updated so they don't
return.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```
Expected: Single commit removing several files and updating `.gitignore`.

### Task 1.5: Handle anthropic_api_key.txt

**Files:**
- Inspect: `anthropic_api_key.txt`
- Possibly delete or archive
- Possibly trigger git history rewrite

- [ ] **Step 1: Inspect the file**

Run:
```bash
file anthropic_api_key.txt
wc -c anthropic_api_key.txt
head -c 64 anthropic_api_key.txt
git log --all --oneline -- anthropic_api_key.txt | head -5
```
Expected: Determine if file is text or MIP-encrypted, get content preview, see if it's in git history.

- [ ] **Step 2: Decision branch**

Three cases:
- **Case A — Placeholder / dummy / empty:** Archive a copy as-is. Delete from working tree and git index. Single commit.
- **Case B — Real key in working tree, NOT in git history:** Surface to user, ask them to rotate. After confirmation, archive a sanitized copy (replace the actual key with `<REDACTED-was-rotated-on-YYYY-MM-DD>`). Delete from working tree. Single commit.
- **Case C — Real key in working tree AND in git history:** Surface to user, ask them to rotate. After confirmation, archive sanitized copy. Use `git filter-repo --invert-paths --path anthropic_api_key.txt` to scrub history. User performs the force-push (it's destructive to public history).

- [ ] **Step 3: Archive (if real key or non-trivial content)**

Create `C:\Users\S1134987\Documents\Python\DataViewer-Archive\2026-05-20-cleanup-pass-1\` and place `anthropic_api_key.txt` (sanitized if needed) inside. Add a line to that pass's `notes.md`.

- [ ] **Step 4: Remove from tree**

Run:
```bash
git rm anthropic_api_key.txt
git commit -m "chore(secret): remove anthropic_api_key.txt from tree"
```
Expected: File gone. For Case C, follow up with filter-repo + user-driven force-push.

### Task 1.6: Archive and remove DATAVIEWER_UPDATES.txt

**Files:**
- Archive: `DATAVIEWER_UPDATES.txt`
- Delete from tree

- [ ] **Step 1: Archive the file**

Use Python (per CLAUDE.md MIP workaround):
```python
import shutil, os
src = r"C:\Users\S1134987\Documents\Python\DataViewer Dev\DataViewer-Enterprise\DATAVIEWER_UPDATES.txt"
dst_dir = r"C:\Users\S1134987\Documents\Python\DataViewer-Archive\2026-05-20-cleanup-pass-1"
os.makedirs(dst_dir, exist_ok=True)
shutil.copy2(src, os.path.join(dst_dir, "DATAVIEWER_UPDATES.txt"))
```

- [ ] **Step 2: Append a `notes.md` line**

Append to `C:\Users\S1134987\Documents\Python\DataViewer-Archive\2026-05-20-cleanup-pass-1\notes.md`:
```
- DATAVIEWER_UPDATES.txt — pre-CHANGELOG.md changelog (53KB). Superseded by CHANGELOG.md ~2026-03.
```

- [ ] **Step 3: Remove from tree**

Run:
```bash
git rm DATAVIEWER_UPDATES.txt
git commit -m "chore(repo): archive and remove DATAVIEWER_UPDATES.txt (pre-CHANGELOG)"
```
Expected: File removed from tree, archive populated.

### Task 1.7: Address MIP-relabeled xlsx test fixtures

**Files:**
- Create: `tools/stage_test_fixtures.py`
- Modify: `tests/run-tests.ps1` (add staging call before suite launch)
- Modify: `tests/tst_sopLoader/tst_sopLoader.cpp` (honor `DVE_TEST_FIXTURES_DIR`)
- Modify: `tests/tst_reportgenerator/tst_reportgenerator.cpp` (honor `DVE_TEST_FIXTURES_DIR`)

- [ ] **Step 1: Identify affected fixtures**

Run:
```bash
find tests resources -name "*.xlsx" | xargs -I{} sh -c 'echo "=== {} ==="; head -c 32 "{}" | xxd | head -3'
```
Expected: A short table showing which `.xlsx` files start with the `%TSD-Header-###%` MIP marker vs. which are plaintext.

- [ ] **Step 2: Write `tools/stage_test_fixtures.py`**

Create via Python delete-and-rewrite (CLAUDE.md MIP workaround) at `tools/stage_test_fixtures.py`:
```python
"""Stage MIP-encrypted .xlsx test fixtures to a non-MIP path.

Copies every .xlsx under tests/ and resources/ to %TEMP%\\dve_test_fixtures\\,
preserving relative paths. Sets DVE_TEST_FIXTURES_DIR for the test run.

The MIP auto-labeler re-encrypts files in tracked source paths faster than
decrypt_via_copy.py can strip them. By staging fixtures to %TEMP%, which
the labeler does not watch, C++ tests see plaintext xlsx bytes.

Idempotent. Safe to re-run.
"""
import os, shutil, sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
SOURCES = [REPO / "tests", REPO / "resources"]
DEST = Path(os.environ.get("TEMP", os.environ.get("TMP", "."))) / "dve_test_fixtures"

def is_mip_encrypted(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            head = f.read(32)
        return head.startswith(b"%TSD-Header")
    except OSError:
        return False

def main():
    DEST.mkdir(parents=True, exist_ok=True)
    staged = 0
    for root in SOURCES:
        if not root.exists():
            continue
        for src in root.rglob("*.xlsx"):
            rel = src.relative_to(REPO)
            dst = DEST / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            # Python read+write bypasses MIP at user-session level.
            with open(src, "rb") as fin:
                data = fin.read()
            if data.startswith(b"%TSD-Header"):
                print(f"WARNING: source {src} is itself MIP-encrypted; staged copy will be too.")
            with open(dst, "wb") as fout:
                fout.write(data)
            staged += 1
            print(f"  staged: {rel}")
    print(f"\n{staged} fixture(s) staged at: {DEST}")
    print(f"\nSet DVE_TEST_FIXTURES_DIR=\"{DEST}\" for the test run.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 3: Modify `tests/run-tests.ps1` to stage fixtures**

Edit `tests/run-tests.ps1`. Locate the section right before tests are invoked. Add:
```powershell
# v2.0.7 — stage MIP-encrypted .xlsx fixtures to %TEMP% so the MIP auto-labeler
# can't re-encrypt them between staging and read. Tests honor DVE_TEST_FIXTURES_DIR.
Write-Host "Staging test fixtures..."
python "$RepoRoot\tools\stage_test_fixtures.py"
if ($LASTEXITCODE -ne 0) { throw "Fixture staging failed" }
$env:DVE_TEST_FIXTURES_DIR = Join-Path $env:TEMP "dve_test_fixtures"
```
Place this before any `mingw32-make check` or test-binary invocation.

- [ ] **Step 4: Modify `tst_sopLoader.cpp` to honor env var**

In `tests/tst_sopLoader/tst_sopLoader.cpp`, find the hardcoded path to `resources/sops.xlsx` (or wherever the fixture lives). Wrap it with:
```cpp
QString fixturePath() {
    const QString envOverride = qEnvironmentVariable("DVE_TEST_FIXTURES_DIR");
    if (!envOverride.isEmpty()) {
        return envOverride + "/resources/sops.xlsx";
    }
    return QStringLiteral(REPO_ROOT) + "/resources/sops.xlsx";
}
```
Replace existing fixture-path references with `fixturePath()`.

- [ ] **Step 5: Modify `tst_reportgenerator.cpp` similarly**

Repeat the fixturePath() pattern for `tst_reportgenerator`. Identify the fixture path used by `loadSopRows_filtersToRequestedTests` and route it through the env-var override.

- [ ] **Step 6: Run the staging script manually to verify**

Run:
```bash
python tools/stage_test_fixtures.py
ls "$TEMP/dve_test_fixtures" 2>/dev/null || ls /c/Users/S1134987/AppData/Local/Temp/dve_test_fixtures
```
Expected: Staged fixtures appear in temp.

- [ ] **Step 7: Run the full test suite**

Run: `./tests/run-tests.ps1` (or `pwsh ./tests/run-tests.ps1`)
Expected: 34 passed, 0 failed, 0 skipped. (Previously was 32/2/0.)

- [ ] **Step 8: Commit**

Run:
```bash
git add tools/stage_test_fixtures.py tests/run-tests.ps1 tests/tst_sopLoader/tst_sopLoader.cpp tests/tst_reportgenerator/tst_reportgenerator.cpp
git commit -m "$(cat <<'EOF'
fix(tests): stage MIP-relabeled .xlsx fixtures to %TEMP%

The MIP auto-labeler re-encrypts .xlsx files in tracked paths
faster than decrypt_via_copy.py can strip them, breaking
tst_sopLoader and tst_reportgenerator. stage_test_fixtures.py
copies fixtures into %TEMP% before the suite runs and exposes
the path via DVE_TEST_FIXTURES_DIR. Restores 34/34 green.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 1.8: Build out comprehensive `--self-test` smoke

**Files:**
- Modify: `src/utils/SelfTest.cpp` (add new `TestResult test*()` methods)
- Modify: `src/utils/SelfTest.h` (declarations if needed)
- Modify: `tests/deployment/README.md` (document new tests)

- [ ] **Step 1: Read current SelfTest.cpp**

Read: `src/utils/SelfTest.cpp` and `src/utils/SelfTest.h` in full. Understand the existing pattern.

- [ ] **Step 2: Add `testAppStartupHeadless()` method**

Add a new test that exits early after constructing `QApplication` under `QT_QPA_PLATFORM=offscreen` to verify no GUI-thread crash on startup. The `--self-test` mode already runs without a display; this verifies the explicit offscreen path is operational.

- [ ] **Step 3: Add `testFileLoaderRoundtrip()` method**

Add a test that loads a representative `.xlsx` (path resolved via the same `DVE_TEST_FIXTURES_DIR` mechanism if set) through the `ExcelReader` → `DataProcessor` chain, asserts non-zero rows + non-empty `FileResult`, and exits cleanly.

- [ ] **Step 4: Add `testModeSwitch()` method**

Construct `MainWindow`, invoke `toggleSensoryMode()` then `toggleDetailedSensoryMode()` then back to TPM. Assert no leaks (`QPointer` checks on the central widget after each switch) and no crash.

- [ ] **Step 5: Add `testDbConnectAndSelect()` method**

If `DVE_TEST_PG_CONN` is set, open a connection via `DatabaseManager`, run a trivial `SELECT 1`, close. Skip cleanly if env var is unset.

- [ ] **Step 6: Add `testLiveSyncConnectCycle()` method**

If test postgres available: instantiate `LiveSync`, connect, wait briefly for worker-thread connection establishment, disconnect, verify clean teardown (no dangling worker thread, no orphaned `QSqlDatabase`).

- [ ] **Step 7: Add `testReportGenerationMinimal()` method**

Build a minimal `FileResult` in memory (one sheet, two rows, one sample). Call `ReportGenerator::buildTestPptx()` to a temp path. Verify the file is a valid ZIP (open with `QuaZip` or basic ZIP header check) containing at least `[Content_Types].xml` and one slide XML part.

- [ ] **Step 8: Add `testSensoryPresetRoundtrip()` method**

If test postgres available: call `DatabaseManager::saveSensoryHeaderPresets()` with a test set, then `loadSensoryHeaderPresets()`, verify equality, then delete the test rows. Skip if no DB.

- [ ] **Step 9: Wire all new tests into the suite**

Locate the test runner block in `SelfTest::run()`. Append calls to each new method in the order above. Each returns a `TestResult` aggregated into the JSON output.

- [ ] **Step 10: Build incrementally**

Run:
```bash
python tools/decrypt_via_copy.py --apply
cd build
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
mingw32-make -j8
```
Expected: Clean build, no `-Werror` violations.

- [ ] **Step 11: Run `--self-test` against the test postgres**

Run:
```bash
./tests/start-test-postgres.ps1   # starts dve-test-pg on port 5433
$env:DVE_TEST_PG_CONN = "host=localhost port=5433 dbname=dataviewer user=dataviewer password=test"
./release/DataViewer.exe --self-test --self-test-out "$env:TEMP/dve_selftest.json"
cat "$env:TEMP/dve_selftest.json"
```
Expected: All new tests appear in the JSON output. All pass (or skip cleanly when DB unavailable).

- [ ] **Step 12: Update `tests/deployment/README.md`**

Add one paragraph per new test method documenting what it verifies and any required preconditions (e.g., test postgres for DB tests).

- [ ] **Step 13: Commit**

Run:
```bash
git add src/utils/SelfTest.cpp src/utils/SelfTest.h tests/deployment/README.md
git commit -m "$(cat <<'EOF'
feat(selftest): comprehensive smoke harness for cleanup phases

Adds offscreen-startup, file-loader, mode-switch, DB,
LiveSync, report-gen, and sensory-preset roundtrip checks
to --self-test. These become the per-phase smoke gate for
Pass 3 of the v2.0.7 cleanup.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 1.9: Pass 1 commit checkpoint

- [ ] **Step 1: Run the full test suite**

Run: `./tests/run-tests.ps1`
Expected: 34/34 green.

- [ ] **Step 2: Run `--self-test` against running test postgres**

Run:
```bash
./tests/start-test-postgres.ps1
./release/DataViewer.exe --self-test --self-test-out "$env:TEMP/dve_selftest.json"
```
Expected: All new and existing self-test methods pass.

- [ ] **Step 3: Push the cleanup branch**

Run: `git push origin chore/v2.0.7-cleanup`
Expected: All Pass 1 commits land on origin.

- [ ] **Step 4: Move on to Pass 2.**

---

## Pass 2 — Deep Audit

### Task 2.1: Source dead-code scan

**Output:** Findings rows in `docs/audits/2026-05-20-cleanup-audit.md` under the "Dead code" section.

- [ ] **Step 1: Enumerate all source files**

Run:
```bash
find src tests -name "*.cpp" -o -name "*.h" | sort > /tmp/source_files.txt
wc -l /tmp/source_files.txt
```
Expected: A line count for the full set of source files (likely 150-300).

- [ ] **Step 2: Find files not referenced in .pro or by `#include`**

Dispatch an Explore subagent with this prompt:
```
List every .cpp/.h file in src/ and tests/ that meets ALL of:
  1. Not listed in DataViewerEnterprise.pro SOURCES, HEADERS, or its tests/*.pro equivalents
  2. Not #include'd by any other tracked .cpp or .h file (grep across the whole tree)
  3. Not built by qmake's MOC step (Q_OBJECT files in HEADERS still get auto-included)
Report a table: file path, evidence (which scan caught it), confidence (High if both 1 and 2, Medium if only one).
```

- [ ] **Step 3: Find unused public symbols**

For each `.h` file in `src/`, extract top-level public class/struct names and free-function declarations. For each symbol, grep the rest of the tree for usage. Flag symbols with zero hits (excluding the declaration site).

Skip symbols listed in the spec's "Preserved symbols" section.

- [ ] **Step 4: Find `#if 0` and commented-out implementations**

Run:
```bash
grep -rn "#if 0" src tests
grep -rn "^[[:space:]]*//.*[a-z]\+[[:space:]]*([a-z]" src tests | head -200
```
Expected: List of `#if 0` blocks and lines that look like commented-out function calls. Filter to obvious dead blocks of >5 lines.

- [ ] **Step 5: Write findings into the audit doc**

Use the table format from the spec (ID, Category, Location, Description, Confidence, Disposition, Archive?). Add F001..F0xx for each dead-code finding.

### Task 2.2: Source simplification scan

- [ ] **Step 1: Find files >800 lines**

Run:
```bash
find src -name "*.cpp" -o -name "*.h" | xargs wc -l | sort -nr | head -20
```
Expected: Top 20 largest files. Note any >800 lines for inspection.

- [ ] **Step 2: For each oversize file, identify split candidates**

Dispatch a feature-dev:code-explorer subagent per top-3 oversize file with prompt:
```
Read <file>. List its top-level responsibilities. For each responsibility, list which functions/classes/methods belong to it. Identify natural split lines (e.g., "lines 1-450 deal with X, lines 451-900 deal with Y"). Return one paragraph + a proposed split.
```

- [ ] **Step 3: Find single-use abstractions**

Look for files with `class FooFactory`, `class IFooImpl`, `class FooManager` style names. For each, count unique caller sites for the abstraction. Flag those with exactly 1 caller as candidates for inlining.

- [ ] **Step 4: Find redundant helpers across files**

Grep for function bodies that look near-identical across files. Start with helpers that have generic names (`makeXmlElement`, `buildHeader`, `escapeString`). Compare bodies; flag duplicates with `diff` of >70% similarity.

- [ ] **Step 5: Find wrapper functions**

For each function in `src/` whose body is 1-3 lines and just forwards arguments unchanged, flag as inlining candidate.

- [ ] **Step 6: Write findings into the audit doc**

Same table format. Confidence is Medium by default (these are judgment calls; user signs off per item).

### Task 2.3: Doc-rot scan

- [ ] **Step 1: Inventory all markdown files**

Run:
```bash
find . -name "*.md" -not -path "./.git/*" -not -path "./node_modules/*" -not -path "./.claude/*" -not -path "./external/*" | sort
```
Expected: ~30-50 markdown files.

- [ ] **Step 2: Identify shipped-work plans/specs**

In `docs/superpowers/plans/` and `docs/superpowers/specs/`, the convention is that a plan describes work that is now shipped. Read each filename's prefix date — anything from before 2026-05-15 is likely fully shipped.

For each, scan the file: if every task is checked-off or marked done, flag for archival (move to `DataViewer-Archive/2026-05-20-cleanup-pass-3-doc-rot/docs-superpowers/`).

- [ ] **Step 3: Check `CLAUDE.md` accuracy**

Read `CLAUDE.md` section by section. Cross-check each claim against current state:
- File paths still exist
- Counts (e.g., "27 test classes" — verify with `find tests -name "tst_*" | wc -l`)
- Tool versions (Qt 6.10.1, MinGW 13.1.0 — verify via `qmake -v`)
- Workflow steps still accurate (`build_installer.bat`, deploy steps)

Flag inaccurate sections for update.

- [ ] **Step 4: Check `tasks/lessons.md`**

Read `tasks/lessons.md`. For each lesson entry, check whether the pattern still applies to the code. Flag entries where the referenced code no longer exists or the pattern is now incorrect.

- [ ] **Step 5: Old handoff docs**

`docs/handoff-*.md` files older than 2 weeks should be archived (after a quick read to ensure no in-flight context).

- [ ] **Step 6: Write findings into the audit doc**

### Task 2.4: Test gaps & dead tests scan

- [ ] **Step 1: Cross-reference tests against source**

For each `tests/tst_*/` directory, identify which source class/module it tests. Verify that class still exists in `src/`.

- [ ] **Step 2: Find duplicated test setup**

Look for boilerplate `init()`, `cleanup()`, `initTestCase()` blocks that are near-identical across test files. Flag as shared-helper candidates.

- [ ] **Step 3: Find unconditionally-skipped tests**

```bash
grep -rn "QSKIP" tests | grep -v "if\|unless\|conditional"
```
Expected: List of `QSKIP` calls. Flag any that aren't guarded by a runtime check.

- [ ] **Step 4: Write findings into the audit doc**

### Task 2.5: Deploy / installer scan

- [ ] **Step 1: Read `installer.iss` end-to-end**

Read the full file. Flag any `[Code]` Pascal sections that look obsolete (e.g., references to v1.x migration logic). Flag any `[Files]` entries that point to non-existent files.

- [ ] **Step 2: Check migrations**

Run:
```bash
ls deploy/postgres/migrations/
```
Each migration should have been applied to prod. The user can confirm. Note any with errors that need separate handling.

- [ ] **Step 3: Check `deploy/postgres/README.md`**

Read the file. Verify it matches current container setup and connection details. Flag inaccurate sections.

- [ ] **Step 4: Write findings into the audit doc**

### Task 2.6: Finalize audit doc

- [ ] **Step 1: Add summary section to audit doc**

At the top of `docs/audits/2026-05-20-cleanup-audit.md`, add a summary:
```
| Category | High | Medium | Low | Total |
|---|---|---|---|---|
| Dead code | N | N | N | N |
| Simplification | N | N | N | N |
| Doc rot | N | N | N | N |
| Test gaps | N | N | N | N |
| Deploy | N | N | N | N |
```

- [ ] **Step 2: Re-read the preserved-symbols list**

Verify no flagged finding overlaps with the v2.0.6 hotfix preserved symbols.

- [ ] **Step 3: Decide on mid-initiative checkpoint**

If simplification findings > 40 OR high-risk simplifications > 5: flag in the audit doc summary that a user-smoke checkpoint will run after Phase 3.3 (before 3.4).

- [ ] **Step 4: Commit the audit doc**

Run:
```bash
mkdir -p docs/audits
git add docs/audits/2026-05-20-cleanup-audit.md
git commit -m "$(cat <<'EOF'
docs(audit): v2.0.7 codebase cleanup audit findings

Output of Pass 2. N findings across dead code, simplification,
doc rot, test gaps, and deploy categories. Drives Pass 3
execution order.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Pass 3 — Phased Execution

Each phase ends with: clean build + `tests/run-tests.ps1` 34/34 green + `--self-test` green + one commit. Audit findings drive the work; tasks below are the scaffolding.

### Task 3.1: Phase — Doc rot

- [ ] **Step 1: Archive shipped plans/specs**

For each plan/spec flagged in audit task 2.3 step 2:
```python
import shutil, pathlib
src = "<repo path>"
dst = r"C:\Users\S1134987\Documents\Python\DataViewer-Archive\2026-05-20-cleanup-pass-3-doc-rot\<original-path>"
pathlib.Path(dst).parent.mkdir(parents=True, exist_ok=True)
shutil.move(src, dst)
```
Append one-line summaries to the pass folder's `notes.md`.

- [ ] **Step 2: Update `CLAUDE.md`**

Per findings, fix inaccurate sections inline. Do not archive — git history is the record for in-tree edits.

- [ ] **Step 3: Update `tasks/lessons.md`**

Per findings, remove or update stale entries. Same rule: in-tree edits, no archive.

- [ ] **Step 4: Archive old handoff docs**

Move every `docs/handoff-*.md` older than 2 weeks to the archive folder.

- [ ] **Step 5: Build + test**

Run: `python tools/decrypt_via_copy.py --apply && cd build && mingw32-make -j8 && cd .. && ./tests/run-tests.ps1`
Expected: Build clean, 34/34 tests green. (Doc edits can't break the build, but verify anyway.)

- [ ] **Step 6: Commit**

```bash
git add -A
git commit -m "$(cat <<'EOF'
docs(cleanup): archive shipped plans + fix doc rot (Pass 3.1)

N plans/specs/handoffs archived to DataViewer-Archive.
CLAUDE.md and tasks/lessons.md updated where stale.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

### Task 3.2: Phase — Test cleanup

- [ ] **Step 1: Remove dead tests**

For each test flagged in audit task 2.4 step 1 (tests whose target class no longer exists): archive the test source file to `DataViewer-Archive\2026-05-20-cleanup-pass-3-deadcode\tests\` then `git rm` it.

- [ ] **Step 2: Consolidate duplicated test setup**

For each duplicated-setup finding from audit task 2.4 step 2: extract the shared helper into `tests/common/TestHelpers.h` (creating it if needed). Update each test to use the shared helper.

- [ ] **Step 3: Remove unconditional QSKIPs**

For each QSKIP from audit task 2.4 step 3: either implement the test fully or remove it entirely. No dead skipped tests.

- [ ] **Step 4: Build + test**

Run: full build + suite. Expected: 34/34 (count may shrink slightly if dead tests removed; document the new count).

- [ ] **Step 5: Run `--self-test`**

Run: `./release/DataViewer.exe --self-test`
Expected: All green.

- [ ] **Step 6: Commit**

### Task 3.3: Phase — Dead-code removal

- [ ] **Step 1: Per-finding execution loop**

For each High-confidence dead-code finding in the audit doc, in order:
1. Archive the file/section to `DataViewer-Archive\2026-05-20-cleanup-pass-3-deadcode\<original-path>\`
2. Append one line to that pass folder's `notes.md`
3. Remove from source tree (`git rm` for files, `Edit` for partial removals)
4. Update `DataViewerEnterprise.pro` if SOURCES/HEADERS changed
5. Run full build + suite + self-test — must stay green
6. Commit one finding at a time. Commit message: `chore(cleanup): remove <symbol> (audit F<NNN>)`

For Medium/Low confidence findings: surface to user, await confirmation, then proceed.

- [ ] **Step 2: Run final gate after all High items**

Full build + suite + self-test. Confirm green.

### Task 3.4: Phase — Simplification refactors

This phase is highest risk. Each finding gets its own commit. If audit has >40 simplification findings, pause and surface to user for triage.

- [ ] **Step 1: Sort findings by risk**

Order by: pure inlining (lowest risk) → file splits (medium) → abstraction collapses (highest).

- [ ] **Step 2: Per-finding execution loop**

For each finding:
1. Make the refactor
2. Run full build + suite + self-test
3. Commit one finding at a time

If the optional inherit-helper template refactor (spec line ~145) makes the cut, do it after all atomic refactors are done.

- [ ] **Step 3: Mid-initiative checkpoint (conditional)**

If the audit flagged this as needing user smoke: pause here, surface to user with a clean build and a list of what changed in 3.3 + 3.4 so far. Wait for confirmation before continuing to 3.5.

### Task 3.5: Phase — Installer / deploy cleanup

- [ ] **Step 1: Apply installer.iss findings**

Remove dead Pascal sections, fix incorrect `[Files]` entries per audit task 2.5 step 1.

- [ ] **Step 2: Update `deploy/postgres/README.md`**

Apply audit findings from task 2.5 step 3.

- [ ] **Step 3: Verify installer builds**

Run:
```bash
tools\prepare_python_embed.bat
build_installer.bat
```
Expected: `dist\DataViewer-setup.exe` produced. (Installer behavior validated by user at v2.0.7 build time.)

- [ ] **Step 4: Commit**

### Task 3.6: Phase — Final polish + v2.0.7 release

- [ ] **Step 1: Bump `VERSION` in `DataViewerEnterprise.pro`**

Change line `VERSION = 2.0.6` to `VERSION = 2.0.7`. Per CLAUDE.md, this requires a clean rebuild:
```bash
cd build
mingw32-make clean
mingw32-make -j8
```

- [ ] **Step 2: Add `CHANGELOG.md` entry**

Prepend a section to `CHANGELOG.md`:
```markdown
## v2.0.7 — 2026-05-21 — Codebase cleanup

No user-visible changes. Internal cleanup release:
- Removed N dead-code findings (files/classes/functions unreferenced)
- Simplified N source files (splits + abstraction collapses)
- Archived N shipped plan docs / handoffs to DataViewer-Archive
- Restored 34/34 green test suite (MIP-relabel xlsx fixture workaround)
- Added comprehensive `--self-test` smoke harness
- Removed N stale git branches and 3 abandoned worktrees
- Removed N generated build artifacts from version control
```

Replace N values with actual counts from the audit and execution.

- [ ] **Step 3: Verify the version embedded in the binary**

Run:
```bash
./release/DataViewer.exe --self-test --self-test-out "$env:TEMP/v207.json"
grep -i version "$env:TEMP/v207.json"
```
Expected: `2.0.7` everywhere.

- [ ] **Step 4: Build the installer**

Run:
```bash
tools\prepare_python_embed.bat
build_installer.bat
```
Expected: `dist\DataViewer-setup.exe` rebuilt at v2.0.7.

- [ ] **Step 5: Run the final test checklist** (see CHECKLIST below; user does this manually)

- [ ] **Step 6: Tag and merge**

After user confirms the test checklist passes:
```bash
git checkout main
git merge --ff-only chore/v2.0.7-cleanup
git tag -a v2.0.7 -m "v2.0.7 — codebase cleanup"
git push origin main v2.0.7
```

- [ ] **Step 7: User transfers installer to Synology** (per user-level rule — never automated).

---

## Final Test Checklist (for user, post-Pass 3.6)

This is the validation gate before v2.0.7 ships. The user runs this end-to-end on a fresh install.

### A. Install
- [ ] Run `dist\DataViewer-setup.exe`
- [ ] Setup wizard pre-fills DB password (no manual entry needed)
- [ ] About dialog shows version `2.0.7`
- [ ] First launch creates `%LOCALAPPDATA%\DataViewer\dataviewer.log`

### B. TPM mode
- [ ] Drag-and-drop a `.xlsx` file: loads without error, plot renders
- [ ] Edit a cell, observe excel write-back queues + flushes
- [ ] Generate a Single TPM report → opens valid `.pptx`
- [ ] Generate a Full TPM report → opens valid `.pptx`
- [ ] Switch sheets via tree → plot updates
- [ ] Data cleanup dialog → exclude a row → report excludes it

### C. Sensory mode
- [ ] Toggle to Sensory mode (ribbon updates)
- [ ] Load a sensory `.xlsx` from disk
- [ ] Add a sample card → fill in metrics → radar updates
- [ ] Save Test Headers → confirmation dialog appears with sample names
- [ ] Test Title and Media dropdowns populate from DB
- [ ] Sample name dropdown shows preset names
- [ ] Ctrl+S saves the file
- [ ] Ctrl+U commits to DB without freeze (the v2.0.6 hotfix path)
- [ ] Load same file again → Ctrl+U → no UNIQUE violation (v2.0.5 inherit path)
- [ ] Generate Single Sensory Report → opens valid `.pptx`

### D. Detailed Sensory mode
- [ ] Toggle to Detailed mode (ribbon updates again)
- [ ] Create new session → 14 questions render, dual radar charts on right
- [ ] Add multiple samples → navigate prev/next → values persist
- [ ] Save Test Headers → confirmation dialog appears
- [ ] Combo widths are correct (sample names visible, not truncated)
- [ ] Generate Detailed Sensory Report → opens valid `.pptx`

### E. Database & multi-user
- [ ] Load file → make edit → another user sees update (NOTIFY-driven)
- [ ] Two clients editing same cell → conflict dialog appears
- [ ] Take NAS offline → app falls back to read-only snapshot (no crash)
- [ ] Bring NAS back online → reconnects without restart

### F. Translator
- [ ] Tools menu → Launch Translator → opens with API key pre-filled
- [ ] Translate a small `.pptx` → output is valid

### G. Auto-update
- [ ] Check for updates → finds installed version, no false-positive update prompt
- [ ] (User does Synology transfer at end and verifies on a second install)

### H. Logs
- [ ] `%LOCALAPPDATA%\DataViewer\dataviewer.log` has session banner and recent entries
- [ ] No `qWarning`/`qCritical` lines from unexpected paths

### I. Regression spot-checks (things the cleanup specifically touched)
- [ ] App startup time isn't slower than v2.0.6 (one quick gut-check)
- [ ] No new console output (qDebug spam) during normal use
- [ ] No new dialogs appearing where they didn't before

**Pass criteria:** every box checked. Any unchecked or red → block v2.0.7 tag and surface back to me.

---

## Self-Review

**Spec coverage:** every section of the design spec maps to one or more tasks above:
- Pass 1.1-1.9 → Tasks 1.1-1.9 ✓
- Pass 2 categories (dead-code, simplification, doc rot, test gaps, deploy) → Tasks 2.1-2.5 ✓
- Pass 2 finalization → Task 2.6 ✓
- Pass 3 phases (doc rot, test cleanup, dead-code, simplification, installer, polish) → Tasks 3.1-3.6 ✓
- Archive policy → embedded in each task that touches files
- Preserved symbols → referenced in Task 2.1 step 3 ✓
- MIP fixture fix → Task 1.7 ✓
- Self-test fattening → Task 1.8 ✓
- Final test checklist → bottom of this plan ✓

**Placeholder scan:** Task 2 audit tasks contain TBD placeholder counts ("N findings") because the actual counts only exist after audit runs. Pass 3 task content depends on those findings. This is unavoidable for an audit-driven plan; the scaffolding is in place and the per-finding loops are explicit.

**Type/symbol consistency:** Names used (`DVE_TEST_FIXTURES_DIR`, `stage_test_fixtures.py`, `chore/v2.0.7-cleanup`, archive folder names) are consistent across all tasks.

**One risk worth flagging:** Pass 3.4 (simplification) is the highest-risk phase. If the audit surfaces a large refactor with cross-cutting impact, the per-finding-commit rule may produce a sequence of commits that don't each leave the tree in a working state. Mitigation: each commit must pass build + tests; if a refactor can't be split into commit-shaped chunks, it stays whole and isn't bundled with other changes.
