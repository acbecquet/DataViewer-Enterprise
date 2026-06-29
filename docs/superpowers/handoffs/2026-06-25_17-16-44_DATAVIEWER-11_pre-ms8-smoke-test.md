---
date: 2026-06-25T17:16:45-0700
researcher: Alexander Becquet
git_commit: 168f34a09aa95fc5a6ba35ff7ee89beabf57e15a
branch: feature/v2.6.0-backlog-batch
repository: DataViewer-Enterprise
topic: "v2.6.0 sprint — MS-4 + MS-5 shipped (awaiting smoke-test), MS-8 next"
tags: [implementation, strategy, sprint, DATAVIEWER-6, DATAVIEWER-13, DATAVIEWER-11, DATAVIEWER-15, sensory, notes-panel, stopwatch, phone-web-form]
status: complete
last_updated: 2026-06-25
last_updated_by: Alexander Becquet
type: implementation_strategy
---

# Handoff: DATAVIEWER-11 — smoke-test v2.5.11, then start MS-8 (phone sensory web form)

## Task(s)

Sprint **v2.5.x internal → v2.6.0 deployable** on branch `feature/v2.6.0-backlog-batch` (cut off `main` @ `v2.5.0`). Master spec: `docs/superpowers/specs/2026-06-23-backlog-master-spec.md`. Live dashboard: `docs/sprint-tracker.html`.

- **MS-4 / DATAVIEWER-6 — Editable TPM Notes "story" panel — DONE, awaiting owner smoke-test.** All 13 plan tasks committed; the legacy `m_dataTable` grid + its delegates + the dead "View" tab are gone and the NotesStoryPanel is the sole TPM edit surface; plot note-rings + raw/SOP empty-state added. Built as internal **v2.5.10**, version-verified, full suite green 45/45. (Iterated through v2.5.3–v2.5.9 on owner look-and-feel feedback before close-out.)
- **MS-5 / DATAVIEWER-13 — Sensory stopwatch + assignable hotkey — DONE, awaiting owner smoke-test.** Per-`SampleCard` Start/Stop button fills the existing "Puff length" field on stop; assignable **Space** hotkey toggles the focused card (suppressed while a text box / modal dialog has focus), rebindable in the Settings ribbon; focus outline on the active card. Built as internal **v2.5.11**, version-verified, full suite green 48/48.
- **MS-8 / DATAVIEWER-11 — Remote phone sensory web form → Postgres — PLANNED, NOT STARTED. This is the resume point.** Full implementation plan written (`docs/superpowers/plans/2026-06-25-DATAVIEWER-11-MS8-phone-sensory-plan.md`); spike + owner sign-off complete. The desktop merge-safety fix + DV-15 ride the v2.6.0 train; the DB migration + Flask service ship on an independent NAS track.
- **Already done earlier in the sprint:** MS-1/2/3 (DV-5/7/12 plot/ribbon/label polish, `4954369`); MS-7 (DV-9) resolved-minimal; MS-6 (DV-10) already-shipped/verify-only. **Still open small item:** DV-17 "View Raw Data" Tools button (separate follow-on, not started).

**The current installer `dist\DataViewer-setup.exe` is v2.5.11 and is CUMULATIVE** — it contains both the MS-4 panel close-out and the MS-5 stopwatch. One install smoke-tests both.

## Critical References

- `docs/superpowers/plans/2026-06-25-DATAVIEWER-11-MS8-phone-sensory-plan.md` — the MS-8 plan to execute next (3 phases, code-grounded).
- `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md` — AUTHORITATIVE MS-8 design (owner §10 sign-off); the plan operationalizes it.
- `docs/superpowers/specs/2026-06-23-backlog-master-spec.md` — sprint single-source-of-truth (cross-cutting build/test/data-safety standards bind every item).

## Recent changes

MS-4 close-out: `56716b1` removed `m_dataTable` + `CellFocusDelegate`/`RegimeComboDelegate` + dead View tab from `src/MainWindow.cpp/.h`; `1cccce5`+`c621d08` added per-point rings + `PlotWidget::selectPuff` in `src/plotting/PlotEngine.cpp` + `PlotWidget.cpp`; `3aec65f` added `NotesStoryPanel::showHint` in `src/widgets/NotesStoryPanel.cpp`.

MS-5: `3127c68` added the Start/Stop button + `clampPuffSeconds` (`src/ui/SensoryPanel.cpp/.h`, `src/pipeline/SensoryData.cpp/.h`); `19fdd1a` added the focus outline (`SampleCard` objectName/property + `SensoryPanel::onAppFocusChanged`/`cardOwning`); `371c61a` added the hotkey — `SensoryPanel::eventFilter` (Space, focus-guarded, `!activeModalWidget()`), `reloadStopwatchHotkey`, `m_stopwatchKey`, plus `MainWindow::buildSettingsTab` rebind control + `captureStopwatchKey`/`StopwatchKeyCatcher` dialog.

Plans/docs: `fdb69a6` (MS-5+MS-8 plans), `ca31b99` (MS-8 office-TZ + DB-port fixes), tracker commits `4648784`/`4f52ff4`/`7020aa2`/`168f34a`, release `c13e381` (v2.5.10) / `4954f4c` (v2.5.11). Lesson `cde6902` (the TPM grid was 5 coupled surfaces).

## Learnings

- **MIP / AIP encryption (this work machine).** Source `.cpp/.h/.py` in the working tree become ciphertext at rest after MIP re-labels them — the Read/Edit/Grep tools then see a short binary blob, NOT the source (e.g. a 315-line header reads as "102 lines"). Edit source RELIABLY by patching via the **allowlisted Python interpreter** (it reads plaintext); the build's `python tools/decrypt_via_copy.py --apply` plaintexts everything for g++. For committing non-decrypt-scope docs (`.md`/`.txt` under `tasks/`, `release_overview/`), write via Python then `git add` immediately and verify with `git show HEAD:<path> | head -1`.
- **Build-from-Bash invocation.** Use `MSYS_NO_PATHCONV=1 cmd.exe /c "..."` (single `/c`). `MSYS_NO_PATHCONV=1 cmd.exe //c` passes a literal `//c`, which cmd ignores → it prints its banner and **exits 0 without building** (a silent no-op that leaves the old exe). VERSION bumps need `mingw32-make clean` (main.o embeds the version).
- **Background agents are fragile across process restarts.** A Claude Code process restart this session killed two background subagents (their in-process state was lost) but their **git commits survived** — recovery = inspect `git log`/`git worktree list`, re-verify. Prefer running critical builds/tests as **background Bash commands** (re-runnable from one line) over agents. A worktree-isolated agent branched off `main` (not the feature branch) — cherry-pick its commits rather than merging the wrong base.
- **MS-8 design facts already resolved (don't re-derive):** the DB natural key is `(session_name, tester_name, date)` where `session_name` = **Test Title only**, round lives in `tester_name` as `" R1"/" R2"`, and `date` is a local-TZ TEXT string; write is **create-or-append by natural key**, not INSERT-only. Office TZ for the server-side date pin = **`America/Phoenix`** (Arizona has NO DST — do not use `America/Denver`). The DVE DB host/prod port is **5433**, but the co-located Flask service connects on the internal Docker network at **`dataviewer-db:5432`**. A **required desktop C++ fix** rides v2.6.0: `mergeSensoryPreservingDbScores` (`src/pipeline/SensoryData.cpp:103-126`) currently truncates `samples[]` to the in-memory count and would silently DROP a phone-appended sample on any whole-session save — the plan's Phase 1 fixes it. Spike line numbers had drifted; the plan carries re-verified anchors (live save `SensoryPanel.cpp:907`, Excel import `:2192`, panelist import `:2061`).

## Artifacts

- Plans: `docs/superpowers/plans/2026-06-25-DATAVIEWER-11-MS8-phone-sensory-plan.md`, `docs/superpowers/plans/2026-06-25-DATAVIEWER-13-MS5-sensory-stopwatch-plan.md`, `docs/superpowers/plans/2026-06-24-DATAVIEWER-6-editable-notes-panel-plan.md`.
- Specs: `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md`, `docs/superpowers/specs/2026-06-23-backlog-master-spec.md`.
- Installer: `dist\DataViewer-setup.exe` (ProductVersion 2.5.11, cumulative MS-4 + MS-5). Release notes: `release_overview/release_overview_v_2_5_10.txt`, `release_overview/release_overview_v_2_5_11.txt`.
- Dashboard: `docs/sprint-tracker.html` (per-item task panels for MS-4/MS-5/MS-8). Lessons: `tasks/lessons.md`.
- New user-level skill (NOT in this repo): `~/.claude/skills/sprint-tracker/` + a Stop hook in `~/.claude/settings.json` that nudges to keep the tracker current.

## Action Items & Next Steps

1. **OWNER smoke-tests `dist\DataViewer-setup.exe` (v2.5.11).** Covers BOTH features: (MS-4) TPM Notes panel — edits/freeze/Smell-Clog scroll/avatar circle/Puff column; (MS-5) Sensory stopwatch — Start/Stop fills Puff length and survives Ctrl+U + appears in report; **Space** hotkey toggles the *focused* card and does NOT fire while typing in a text box; focus outline on the active card; rebind works under Settings → "Stopwatch Hotkey". Get the verdict.
2. **On approval, start MS-8** per the plan. **Phase 1 (desktop C++, rides v2.6.0, under MS-7 heavy-verification discipline, TDD):** 1.1 add `sample_uid` to `SensorySample` + round-trip; 1.2 fix `mergeSensoryPreservingDbScores` to append DB-only tail samples + unit test; 1.3 grow cards in `SensoryPanel::applyMergedScoresToCurrentSession` + e2e regression (mirror `tst_saveintegrity_e2e` scenario 9); 1.4 DV-15 unify the 3 `session_name` derivations on title-only (`canonicalSensorySessionName`). **Phases 2–3 (independent NAS track, NOT in the v2.6.0 wrap):** DB migration `dve_append_sensory_sample` + least-privilege `sensory_web` role; Flask service at `deploy/sensory-collect/`. MS-8 deadline ~early July (stakeholder Isabel).
3. (Backlog) DV-17 "View Raw Data" Tools button — small separate follow-on.
4. **At sprint end:** wrap deployable **v2.6.0** (fast-forward feature branch → `main`, push branch+main+tags), then the **OWNER manually** drops `dist\DataViewer-setup.exe` on Synology. The agent must NEVER touch Synology and must NOT push to `main` / deploy without explicit owner approval.

## Other Notes

- **Standing constraints (absolute):** never read/write/list any path under `%USERPROFILE%\SynologyDrive\` (prod release channel; owner's manual gate); build installer in-repo only and surface `dist\DataViewer-setup.exe`; the repo is **PUBLIC** — never commit/bundle secrets (Anthropic key, `sensory_web` DB password); never `taskkill excel.exe` (owner runs Excel interactively); `-Werror -Wall -Wextra` — fix the code, never downgrade.
- **Build/test:** `qmake CONFIG+=release && mingw32-make clean && mingw32-make -j8` then `build_installer.bat` (see the `rebuild-dataviewer` skill); tests via `tests\run-tests.ps1` (DB-dependent suites skip without a test Postgres; the known-flaky `tst_responsivelayout` is not a gate). MIP decrypt before every build.
- **Branch state:** local only — feature branch NOT pushed to origin, NOT merged to main, NOT on Synology. Latest commit `168f34a`.
- **Memory updated this session:** `v2-6-0-backlog-batch`, `dataviewer-11-phone-sensory-collection`, `dataviewer-13-sensory-stopwatch`, `build-installer-from-bash-dotslash` (the `//c` trap).
- For MS-8 there remain two non-blocking NAS-side confirmations: the reverse-proxy/TLS hostname for the phone form (not in-repo), and the §10c "N/A" round merge semantics (currently defaulted to append-to-same-session).
