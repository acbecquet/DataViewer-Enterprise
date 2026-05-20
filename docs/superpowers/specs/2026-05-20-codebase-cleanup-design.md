# DataViewer-Enterprise Codebase Cleanup — Design

**Status:** Draft — pending user review
**Date:** 2026-05-20
**Author:** Claude (brainstorming session with @becquetcharlie)
**Target release:** v2.0.6 (no release until cleanup is fully done)

## Goal

Take a stable, fully-shipped v2.0.5 codebase and remove every form of accumulated cruft (stale git refs, generated files committed by accident, dead source code, over-engineered abstractions, doc rot) **without losing a single piece of user-visible functionality**. Backend optimizations are in-scope as long as they don't change behavior at the UI level. The product of this initiative is a v2.0.6 build that does exactly what v2.0.5 does, with a smaller, simpler codebase behind it.

## Non-goals

- Adding new features. None. This is pure cleanup.
- Changing UI behavior, layout, hotkeys, dialogs, or copy.
- Changing the database schema (migrations are sacrosanct — they ran against prod).
- Touching the Synology Drive (release transfer remains manual per the user-level rule in `C:\Users\S1134987\.claude\CLAUDE.md`).
- Refactoring the build system itself (still qmake + MinGW; same Qt 6.10.x).
- Touching the auto-update flow, the bundled Python stack, or the bundled libpq DLLs.

## Constraints

- Windows-only project. Always use PowerShell or Git Bash, never Unix-only tools.
- MIP encryption on source files. All builds preceded by `python tools/decrypt_via_copy.py --apply`.
- Test suite (`tests\run-tests.ps1`) must stay 34/34 green throughout. A red test halts the phase.
- Compiler is `-Werror -Wall -Wextra`. Don't downgrade to silence warnings; fix root cause.
- User-level rule: nothing under `%USERPROFILE%\SynologyDrive\` is touched at any point.

## Approach: two-pass with phased execution

```
Pass 1: Mechanical cleanup    (no audit needed; obvious "just delete" items)
   │
   ▼
Pass 2: Deep audit             (output: one audit doc; no code touched)
   │
   ▼
Pass 3: Phased source work     (executes from audit; test+smoke gates between phases)
   │
   ▼
v2.0.6 release                 (after user's frontend smoke passes)
```

The two-pass structure exists to keep the audit document focused on items that need judgment (source dead-code, simplification, doc rot). Removing mechanical noise first (build dirs, generated Makefiles, stale branches) shrinks the audit surface area and produces an audit that's worth reading.

## Pass 1 — Mechanical cleanup

Each step is its own commit so anything can be reverted in isolation. The whole pass should complete in one short working session.

### 1.1 — Worktrees
- Remove `.claude/worktrees/cranky-hofstadter-afa8a2` (associated with branch `claude/cranky-hofstadter-afa8a2` at 7dccc7e).
- Remove `.claude/worktrees/v2.0.2-fixes` (associated with `main` at fac258f — same commit as main repo).
- Confirm main repo is on the expected branch. Current state shows main repo checked out on `next`; switch back to `main` (which is at the same commit fac258f).

### 1.2 — Stale local branches
Verify each is merged to `main` before deletion. Delete:
- `release/v2.0.3`, `release/v2.0.4`
- `hotfix/v2.0.5-logging`
- `feat/sensory-header-presets`
- `fix/ctrl-s-modes`
- `worktree-v2.0.2-fixes`
- `claude/cranky-hofstadter-afa8a2`

Keep `main`, `next` for now (the `next` situation is sorted out in 1.1).

### 1.3 — Stale remote branches
`git remote prune origin`, then delete matching remote branches for the merged set in 1.2. Also delete pre-v2.0 remote feature branches: `feat/sensory-report-preview`, `fix/sensory-cumulative-per-test`, `feat/tpm-report-overhaul`, `feat/deployment-self-test`, `optimize/comprehensive-review`.

### 1.4 — Repo-root generated artifacts (delete + gitignore)
Remove from version control and add `.gitignore` entries for:
- `Makefile`, `Makefile.Debug`, `Makefile.Release` (qmake-generated, ~2.5MB)
- `test_rcc_output.cpp` (~2.2MB, Qt rcc output)
- `.qmake.stash`
- `build/`, `build-main-review/`, `build-release/`, `build-tests-review/`, `debug/`, `release/` directories at repo root

### 1.5 — Secret check
Inspect `anthropic_api_key.txt`. If it contains a real key:
- Notify user immediately; they rotate the key.
- Archive a sanitized copy (placeholder for the key value, plus a note explaining where it was used).
- Remove from working tree and from git history (`git filter-repo` or BFG).

If it's a placeholder/dummy: just remove and archive.

### 1.6 — Old top-level files
- `DATAVIEWER_UPDATES.txt` (53KB, pre-`CHANGELOG.md` changelog) → archive then delete.
- Untracked `docs/handoff-2026-05-17.md` → read it; archive if historical; just delete if superseded.

### 1.7 — Build out comprehensive `--self-test` smoke
Extend `src/utils/SelfTest.cpp` with new `TestResult test*()` methods covering:
- App startup under `QT_QPA_PLATFORM=offscreen` (no display required)
- File loader roundtrip on a representative `.xlsx` from `test data/`
- Mode switching: TPM → Sensory → Detailed Sensory → back, no leaks
- DB connection + a trivial SELECT (with test postgres running)
- LiveSync connect/disconnect cycle (test postgres)
- Minimal PPTX report generation; assert ZIP is valid and contains required parts
- Sensory header preset save → load → delete roundtrip

This becomes the **per-phase smoke gate for Pass 3**. Each new test stays in seconds and leaves no on-disk side effects.

### 1.8 — Commit checkpoint
After Pass 1, commit each step separately, run the full test suite + the newly-fattened `--self-test`, then ship the spec self-review for Pass 2.

## Pass 2 — Deep audit (output only)

No source files modified in this pass. Producing one comprehensive audit doc.

**Output:** `docs/audits/2026-05-20-cleanup-audit.md`

### Audit categories

**Source dead-code (highest-confidence findings):**
- Files that compile but are not `#include`d anywhere (header scan + `qmake` SOURCES check).
- Classes/structs with zero instantiation sites.
- Public methods with zero call sites (private/static excluded — too noisy).
- `#if 0` / `#ifdef NEVER` blocks.
- Commented-out implementations of >5 lines.

**Source simplification (judgment calls — needs user signoff per item):**
- Files >800 lines that mix multiple responsibilities. Candidates for split.
- Abstractions used in exactly one place ("interface" with one impl, factory with one product).
- Redundant helpers across files (e.g., two functions doing the same thing in different modules).
- Wrapper functions that just forward arguments unchanged.
- Stale comments referencing code that no longer exists.

**Doc rot:**
- `CLAUDE.md` sections that describe code or workflows that have changed.
- `tasks/lessons.md` entries whose underlying pattern no longer applies.
- Old handoff docs in `docs/` for shipped work.
- Plan docs under `docs/superpowers/plans/` for fully-shipped initiatives.
- `tasks/todo.md` if stale.

**Test gaps & dead tests:**
- Test files for removed code.
- Duplicated test setup that could be a shared helper.
- Tests skipping unconditionally.

**Deploy/installer:**
- Dead Pascal code in `installer.iss`.
- Old migrations under `deploy/postgres/migrations/` that have run against prod (keep them for history; flag if any have errors).
- Stale entries in `deploy/postgres/README.md`.

### Audit doc format

Per-finding table:

| ID | Category | Location | Description | Confidence | Disposition | Archive? |
|----|----------|----------|-------------|------------|-------------|----------|
| F001 | Dead | src/foo/Bar.cpp | Class Bar unused since X removed | High | Remove | Yes |
| F002 | Simplify | src/MainWindow.cpp:1234 | Wrapper around saveX with no extra logic | Medium | Inline | No |

`Disposition` is one of: Remove, Inline, Split, Update, Archive-only, Keep-with-note.
`Confidence` is High/Medium/Low. Low confidence items get flagged for user signoff.

A header section in the audit doc summarizes counts by category and confidence.

## Pass 3 — Phased execution

Each phase ends with **all** of:
1. `python tools/decrypt_via_copy.py --apply`
2. Clean incremental build (or full clean rebuild if `VERSION` changed)
3. `tests\run-tests.ps1` → 34/34 green
4. `release\DataViewer.exe --self-test` → all tests green (with test postgres running)
5. A single squash commit on `chore/v2.0.6-cleanup` describing the phase

### Phase 3.1 — Doc rot
Markdown-only edits. Cannot break the build. Archive copies of significantly-cut docs first.

### Phase 3.2 — Test cleanup
Remove dead tests, consolidate duplicated setup. After this, every remaining test exercises something real.

### Phase 3.3 — Dead-code removal
Archive then delete files/classes/functions flagged High confidence in the audit. Medium/Low items get user signoff before action.

### Phase 3.4 — Simplification refactors
Per-finding execution. Highest-risk phase. May want a mid-initiative user smoke checkpoint here depending on audit volume.

### Phase 3.5 — Installer/deploy cleanup
Last because installer regressions only show at install time. Frontend installer smoke happens at v2.0.6 build.

### Phase 3.6 — Final polish & v2.0.6
- Bump `VERSION` in `.pro` (forces full clean rebuild per CLAUDE.md note).
- New `CHANGELOG.md` entry summarizing the cleanup.
- Build installer (worktree → main repo dist per rebuild-dataviewer skill).
- User performs full frontend smoke (this is the only manual gate from user).
- Tag `v2.0.6`, fast-forward `next`.
- User transfers installer to Synology (per user-level rule).

## Archive policy

**Archive root:** `C:\Users\S1134987\Documents\Python\DataViewer-Archive\`

```
DataViewer-Archive/
├── README.md                                <- index keyed by date + pass
├── 2026-05-20-cleanup-pass-1/
│   ├── DATAVIEWER_UPDATES.txt
│   ├── docs-handoff-2026-05-17.md           (if archived)
│   ├── anthropic_api_key.txt.sanitized      (placeholder, never real key)
│   └── notes.md                             <- one line per file
├── 2026-05-20-cleanup-pass-3-doc-rot/
├── 2026-05-20-cleanup-pass-3-deadcode/
│   ├── src/                                 <- mirrors original tree
│   └── notes.md
└── 2026-05-20-cleanup-pass-3-simplification/
```

**Rules:**
- Source files (`.cpp`/`.h`/`.py`/`.iss`/`.sql`) archived before deletion. Original path preserved under each pass folder.
- Whole docs we remove (handoffs, old plans, etc.): archive a copy.
- Edited-out *sections* of docs that remain in tree: don't archive; git history is the record.
- Generated files, build dirs, qmake stash, branches: just delete. No archive.
- Real secrets are never archived. Sanitized placeholder + note explaining what it was.
- Each pass folder gets a `notes.md` with one line per file: original path, what it was, why removed.

## Verification strategy

### Per-phase gates (mandatory)
1. Compile clean — fail = stop the phase.
2. `tests\run-tests.ps1` 34/34 green — fail = stop the phase.
3. `release\DataViewer.exe --self-test` green (fattened in Pass 1.7) — fail = stop the phase.

### Mid-initiative checkpoint (conditional)
If Pass 2 surfaces >40 simplification findings or >5 high-risk simplifications, we pause after Phase 3.3 and have the user do a frontend smoke before continuing to 3.4. Audit volume decides this.

### Final gate (mandatory)
User performs full frontend smoke at v2.0.6 build. They drive the app for ~20 minutes across all three modes, exercise common workflows (file load, edit, save, report, DB update, sensory entry, live sync). Any regression = back to Pass 3.

## Sample test files

User will identify 1-2 representative `.xlsx` files from `test data/` for the `--self-test` smoke harness. Files should cover:
- A standard TPM template
- A sensory session (either standard or detailed)

If the user wants to provide additional smoke-test fixtures, we can extend.

## Risks & mitigations

| Risk | Mitigation |
|------|------------|
| Dead-code is actually called via reflection / qmake MOC magic | High-confidence findings only after manual grep + Q_OBJECT check. Run smoke after each deletion. |
| Simplification refactor changes behavior subtly | Per-finding execution with green tests + smoke between. Refactors land alone, not bundled. |
| MIP re-encryption mid-cleanup | `decrypt_via_copy.py --apply` before every build. |
| Branch deletion loses unmerged work | `git branch -d` only (not `-D`); Git refuses if unmerged. Verify each is in `main`'s ancestry first. |
| `git filter-repo` rewrites public history | Only if `anthropic_api_key.txt` actually held a real key, and only after user has rotated. Force-push only `main` and `next`; user does it. |
| Audit overlooks dead code that the linker silently keeps | Linker won't reveal unused public symbols. Cross-check by grepping for symbol name in the whole tree, not just headers. |

## Deliverables

- `docs/audits/2026-05-20-cleanup-audit.md` — full audit (output of Pass 2).
- `docs/superpowers/specs/2026-05-20-codebase-cleanup-design.md` — this design.
- `docs/superpowers/plans/2026-05-20-codebase-cleanup.md` — implementation plan from `writing-plans` skill.
- `DataViewer-Archive/2026-05-20-cleanup-pass-*/` populated as cleanup runs.
- New `--self-test` smoke methods in `src/utils/SelfTest.cpp`.
- Branch `chore/v2.0.6-cleanup` off `main`, merged at the end.
- Tag `v2.0.6` once frontend smoke passes.
- `CHANGELOG.md` entry under v2.0.6 describing the cleanup at high level.

## Open questions

- Whether `test data/` already contains acceptable smoke fixtures, or we need to commit a small minimal-but-representative `.xlsx` to `tests/fixtures/` for the smoke harness.
- Whether the mid-initiative user-smoke checkpoint is needed; decided in Pass 2 based on simplification count.
- Whether the `anthropic_api_key.txt` situation triggers a history-rewrite (only if it holds a real key).
