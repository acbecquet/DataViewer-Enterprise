# v2.5.x → v2.6.0 Backlog Batch — Multi-Phase Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL — use `superpowers:subagent-driven-development` to implement each phase task-by-task (implementer → spec review → quality review). UI tasks additionally go through the **visual feedback loop** (launch the sandbox app against the test container, screenshot, hand to the owner). **Every task references a Master-Spec section ID (MS-N), or is tagged `(setup)` / `(wrap)` for cross-cutting steps — keep that thread; never lose track of which spec section a task serves.**

**Anchor spec:** `docs/superpowers/specs/2026-06-23-backlog-master-spec.md` (THE single source of truth). When this plan and the spec disagree, the spec wins — fix it first, then re-derive here.

**Branch:** cut `feature/v2.6.0-backlog-batch` off `main` (current tag `v2.5.0`). The branch is named for the **deployable target (v2.6.0)**, not the internal-patch series — `feature/v2.5.0-*` would collide with the already-shipped minor. Do NOT push/merge until the owner install-tests the stacked builds. **Never touch Synology.** Build is `-Werror -Wall -Wextra -Wpedantic`.

**Versioning:** each DataViewer.exe phase produces one or more internal **v2.5.x** patches; the whole batch wraps into the deployable **v2.6.0** on the owner's install-test approval (owner does the Synology drop manually). **MS-8 (the NAS web service) ships on its own independent track and is NOT part of the v2.6.0 wrap.**

**Standing per-task rules (apply to every task):**
- **MIP:** `python tools/decrypt_via_copy.py --apply` before any build/edit. Create NEW files via the Python delete-and-rewrite convention + immediate `git add`; verify the blob (`git show HEAD:<path> | head`).
- **VERSION bump → clean rebuild** (`mingw32-make clean`) or `main.o` ships stale.
- **Tests:** `tests\run-tests.ps1`; DB suites via `tests\start-test-postgres.ps1` (auto-applies migrations). If Docker is off, **ask the owner to start it** — never autonomously. `tst_responsivelayout` is pre-existing flaky — not a gate.
- **`/ponytail-review`** on each non-trivial diff before commit; cut what it flags except validation/security/accessibility/data-safety.
- Commit trailer: `Co-Authored-By: Claude Opus 4.8 <noreply@anthropic.com>`.

---

## Phase 0 — Branch + sprint kickoff

**Goal:** stand up the sprint branch and resolve the blocking product forks *before* any gated work starts.

- [ ] **Task 0.1 (setup)** Cut `feature/v2.6.0-backlog-batch` off `main`; confirm `.pro` VERSION baseline = 2.5.0; MIP decrypt; baseline `tests\run-tests.ps1` to record the starting green/red count.
- [ ] **Task 0.2 → MS-4, MS-7, MS-8** Surface the three product forks to the owner and record decisions in the spec:
  - MS-4: additive Notes panel (option A) vs drop in-app TPM editing (option B — spins out into its own plan).
  - MS-7: minimal (close asymmetry + retire redundant Ctrl+U) vs literal "sole writer" (spins out into its own plan).
  - MS-8: **RESOLVED 2026-06-24** (see the spike sub-spec `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md`). 5-metric Sensory; internet-reachable via reverse proxy + TLS + per-tester tokens (office-WiFi rejected); online-only; **create-or-append by natural key** (not INSERT-only); dedicated least-privilege `sensory_web` role; **+ one required desktop C++ merge fix in v2.6.0** (Phase 5.0). Remaining sign-off items in sub-spec §10.
- [ ] **Task 0.3 → MS-6** Request the owner-approved read-only live query to confirm the backfill marker + zero residual string scores (per the live-query guardrail — ask, don't run unprompted). **Non-blocking** — if the owner is unavailable, proceed and leave the live check as a follow-up chore (Phase 3 closes the ticket on code-evidence).

**VERIFICATION GATE:** branch exists off `main` with the v2.6.0 name; baseline suite count recorded; the MS-4/MS-7 fork decisions captured in the spec (or their phases explicitly parked); MS-8 spike scope noted. No code yet.
**Internal version:** none (setup). **Wraps into:** v2.6.0.

---

## Phase 1 — Quick wins: plot y-axis + clipping + labels (S, no forks)

**Goal:** ship the three low-risk, high-visibility UI-polish fixes. **Strong candidate for a parallel multi-agent run** (`superpowers:dispatching-parallel-agents`) — MS-1, MS-2, MS-3 are independent, render/label-only, and touch disjoint files (PlotEngine/ReportGenerator vs MainWindow ribbon/image-bar vs PlotWidget/SensoryPanel save buttons).

- [ ] **Task 1.1 → MS-1** Anchor yMin=0 in `PlotEngine::autoRange` (`PlotEngine.cpp:53-82`) before the nice-tick snap; apply the same clamp to the dual-axis right/secondary axis (`:774-786`). Add a test accessor on autoRange in `PlotEngine.h`. (Confirm negative-value policy from MS-1 open questions — default hard 0.)
- [ ] **Task 1.2 → MS-1** Fix `computeTpmYMax` (`ReportGenerator.cpp:684-702`) to return `max(7.0, maxTPM + headroom)` in the non-long-puff branch; preserve the long-puff + >7 branches.
- [ ] **Task 1.3 → MS-1** Tests: new `tst_plotengine` assertion that auto-scaled line plots produce yMin==0; update the 5 `computeTpmYMax` assertions in `tst_reportgenerator` (lines 242-270) + add a max>avg-over-7 anti-clip case. Append `tasks/lessons.md` (0-anchor + average-vs-peak yMax).
- [ ] **Task 1.4 → MS-2** Defect A: shorten the path-button label at `MainWindow.cpp:830` per the owner's chosen wording; **visually verify** the longest post-wrap run fits 74px (shorten further if it still clips). Optional: fix the stale "56×70" comment in `RibbonWidget.h:29-32`.
- [ ] **Task 1.5 → MS-2** Defect B: `imgBar->setFixedHeight(32) → 40` (`MainWindow.cpp:1348`); eyeball the left-splitter layout in both TPM + sensory modes. Append `tasks/lessons.md` (ribbon label cap; `RibbonWidget.cpp` is the size source of truth).
- [ ] **Task 1.6 → MS-3** TPM: add a "Save plot/chart" caption to `m_saveBtn` (`PlotWidget.cpp:64`/`:88`, prefer adjacent QLabel if it clips). Sensory: add a matching `QLabel` before `saveChartBtn` (`SensoryPanel.cpp:914`). One `tr()`-wrapped wording across both. Document the Detailed no-button gap in `tasks/lessons.md` + flag to owner.

**VERIFICATION GATE:** `tst_plotengine` + `tst_reportgenerator` green; full `tests\run-tests.ps1` no new failures; build clean under `-Werror …`; `/ponytail-review` clean (no axis-policy abstraction / config knob); **visual loop signed off by owner** for the y-anchor look (real `.xlsx` in UI + a generated PPTX), Defect A/B clipping, and the save-button captions in both modes.
**Internal version:** v2.5.1. **Wraps into:** v2.6.0.

---

## Phase 2 — Medium UI: stopwatch, then Notes panel

**Goal:** deliver the unblocked stopwatch, then the Notes panel once its fork is resolved.

### Phase 2a — Sensory stopwatch (S, unblocked)

- [ ] **Task 2.1 → MS-5** Add a Start/Stop `QPushButton` + `QElapsedTimer` value member to `SampleCard` inside `puffRow` (`SensoryPanel.cpp:~500-516`; header members in `SensoryPanel.h`). Toggle: start → relabel "Stop" (optional ~200ms tick `QTimer`); stop → `qBound`-clamp elapsed to the spin range + `m_puffLengthSpin->setValue(elapsed)` → reuses the existing `cellCommitted` chain.
- [ ] **Task 2.2 → MS-5** Lifecycle: stop/clear the timer on Stop, on card destruction/removal, and reset in `fromSample()`; no dangling timer, no spurious commit. Decide + document the >60s clamp/warn behavior.
- [ ] **Task 2.3 → MS-5** Verify the existing `tst_sensorydataplaceholder` puff round-trip still passes; optionally unit-test any extracted rounding/clamp helper.

**VERIFICATION GATE (2a):** Start→Stop sets the spin to elapsed (±0.2s, 1-decimal); value survives Ctrl+U + appears in the sensory report; on a persisted (id>0) LiveSync session, Stop streams a single `puff_length_sec` per-cell commit and reload shows a JSON number; build clean; suite green; **visual loop** confirms card layout/wrapping unchanged.
**Internal version:** v2.5.2.

### Phase 2b — TPM Notes panel (L; **GATED on the MS-4 fork from Phase 0**)

> Proceed only after Task 0.2 resolves MS-4 to **option A (additive, grid stays)**. Option B (destructive) is its own larger plan covering the live-sync/offline/add-row/raw-sheet teardown — do NOT implement it inline here.

- [ ] **Task 2.4 → MS-4** Add a right-side read-only `QListWidget` (word-wrap) Notes panel echoing the Navigator section-header style (reuse `propHeader` + AppTheme). New member `m_notesPanel` in `MainWindow.h`.
- [ ] **Task 2.5 → MS-4** Restructure `setupCentralWidget` (`MainWindow.cpp:857-1017`) so `PlotWidget` gets center stretch and the Notes panel sits right (reuse the existing `QSplitter`). Keep the grid as the editor (collapsible/secondary tab).
- [ ] **Task 2.6 → MS-4** Populate in `displayCurrentSample` (`:3730-3980`) using the **same** visible-row filter; format `"Note N: Puff <puffs>, Current TPM <tpm>, Average TPM <averageTPM>, <notes>"`; verify values match the grid + `ReportGenerator` output.
- [ ] **Task 2.7 → MS-4** Retire/repurpose the Layout ribbon group (`buildViewTab :790-803`, handlers `:4270-4272`); ensure raw/SOP sheet rendering + mode switches still work (Notes TPM-only). Add a MainWindow-level Notes-population test. Append `tasks/lessons.md` (the grid is the editing + live-collaboration surface).

**VERIFICATION GATE (2b):** build clean (no `-Wunused`); Notes panel lists one wrapped entry per visible row with values matching grid+report (incl. zero-weight filter); per option A, per-row editing + add/remove + live-sync still function; raw/SOP display intact; new Notes test + full suite green; **visual loop** signed off (panel echoes Navigator, no truncation, plot center-fills).
**Internal version:** v2.5.3.
**Phase 2 wraps into:** v2.6.0.

---

## Phase 3 — Schema-touching: JSONB backfill verify-and-close

**Goal:** confirm the already-shipped string-score backfill ran on the live DB and close the ticket — **no code.**

- [ ] **Task 3.1 → MS-6** Confirm in-repo that `dve_normalize_legacy_json()` is present + byte-identical in `init.sql:500`, the migration, and `DatabaseManager.cpp:512-566`; the one-shot is `schema_meta`-gated under `pg_advisory_xact_lock(4242002)` (`:568-614`); the nightly cron (`init.sql:583`) + manual Repair button (`DatabaseBrowserDialog.cpp:1235`) are wired; `tst_databasemanager.cpp:3055/:3224/:3171` cover it.
- [ ] **Task 3.2 → MS-6** With the owner-approved query from Task 0.3: verify the marker holds a timestamp AND residual string-score count = 0 in both session tables. If >0, owner clicks Repair once (or wait for the 03:17 cron). **If the owner is unavailable, do not block** — close on code-evidence and leave the live check as a follow-up chore.
- [ ] **Task 3.3 → MS-6** Transition DATAVIEWER-10 → Done/Duplicate (link commit `1a31c11` + the four code paths); append `tasks/lessons.md` retiring the "pending backfill" assumption.

**VERIFICATION GATE:** code-evidence confirmed (Task 3.1); marker + zero residual confirmed where owner-available; **no new SQL function / migration / schema_meta key / C++ / UI added** (reject any such diff at `/ponytail-review`); ticket closed with the cross-link.
**Internal version:** none (verify/close only). **Wraps into:** v2.6.0.

---

## Phase 4 — Risky architecture: LiveSync persistence (heavy gate)

**Goal:** close the `oil_smell_liking`/session-level asymmetry and (if chosen) retire the redundant Ctrl+U action — **without** removing the load-bearing whole-session path. **GATED on the MS-7 fork from Phase 0.** Proceed only on the minimal-scope decision; the literal "sole writer" decision spins out into its own multi-phase plan (NOT this batch).

- [ ] **Task 4.1 → MS-7** TDD: add a session-level-field two-writer scenario to `tst_saveintegrity_e2e` (mirror scenario 9) — a stale-and-untouched in-memory `oil_smell_liking` must NOT clobber a newer DB value; a value edited this run must persist. Confirm it FAILS on current code. Add a `tst_databasemanager` unit test for session-field arbitration. (`superpowers:test-driven-development`.)
- [ ] **Task 4.2 → MS-7** Extend `mergeDetailedSensoryPreservingDbScores` (`DetailedSensoryData.cpp:153-179`) — and the sensory twin (`SensoryData.cpp:103-126`) if needed — to arbitrate session-level fields via the existing dirty-aware pattern (touched-this-run stays in-memory-authoritative; untouched takes DB). Mark oil_smell_liking/clog/mouthpiece dirty in `commitSessionField` (`DetailedSensoryPanel.cpp:2064`, dirtyCells `:2107-2135`).
- [ ] **Task 4.3 → MS-7** (If the fork chose it) retire the user-facing Ctrl+U action (`MainWindow.cpp:1577-1589`) while keeping `onUpdateDatabase` callable internally from the 5s auto-save timer + close paths; update comments + the DB-sync-indicator tooltip that reference "press Ctrl+U." No dangling QAction/connect.
- [ ] **Task 4.4 → MS-7** Confirm **no** over-removal: `tryWrite*Core`, the merge helpers, dirtyCells, `flushNowAndWait` remain (INSERT path + gapped-LiveSync safety net + reconnect catch-up). Mark `docs/superpowers/specs/2026-06-05-bugfix-batch-design.md` §8 resolved/re-scoped; append `tasks/lessons.md` (whole-session path is load-bearing).

**VERIFICATION GATE:** the new e2e FAILS pre-change, PASSES post-change; new-session (id≤0) INSERT for both sensory + detailed still works; RC1/RC2/Task-3 gapped-LiveSync scenarios still pass; if Ctrl+U removed, close-without-Ctrl+U manual test persists edits + no tooltip/comment still says "press Ctrl+U"; **tst_twoclient_e2e + tst_saveintegrity_e2e green against `dve-test-pg`** + a **manual two-client session-field edit** on the work machine behaves per the chosen conflict model; build clean; full suite green except `tst_responsivelayout`.
**Internal version:** v2.5.4. **Wraps into:** v2.6.0. **NOT** a parallel-agent candidate — single-threaded, high-risk, needs careful sequencing + manual concurrency verification.

---

## Phase 5 — New feature: remote phone sensory web form (spike → build)

**Goal:** stand up the standalone NAS web service. **Spike DONE (2026-06-24) — see `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md`; awaiting owner sign-off on §10.** The web service runs in **parallel-friendly isolation** (separate `deploy/sensory-collect/` Docker stack) on an **independent versioning track** — NOT part of the v2.6.0 wrap. **Exception:** the spike found ONE **required desktop-side C++ merge-safety fix** (so a desktop save can't silently drop a phone-appended sample, R-M) — that fix **DOES ship in v2.6.0** (Phase 5.0 below, gated behind the MS-7 heavy-verification discipline since it touches the load-bearing whole-session path).

### Phase 5.0 — Desktop merge-safety fix (C++; ships in v2.6.0, gated like MS-7)

- [ ] **Task 5.0 → MS-8 (sub-spec §4.1)** Fix `mergeSensoryPreservingDbScores` (`src/pipeline/SensoryData.cpp:107-124`) to APPEND DB-only trailing samples (when `dbSamples.size() > memSamples.size()`) instead of truncating to the in-memory count; grow the card list in `SensoryPanel::applyMergedScoresToCurrentSession`. Add a regression test (in-memory N, DB N+1 → save → reload still N+1) in `tst_sensorydataplaceholder`. **TDD + the MS-7 heavy-verification discipline.** Without this, any desktop save of an open session deletes a concurrently phone-appended sample (R-M).

**VERIFICATION GATE (5.0):** the append regression FAILS pre-change, PASSES post-change; existing sensory round-trip + save-integrity e2e still green; `-Werror` clean. **Folds into the v2.6.0 train.**

### Phase 5a — Design spike (produces its own sub-spec) — DONE 2026-06-24, awaiting §10 sign-off

- [x] **Task 5.1 → MS-8 — DONE 2026-06-24.** Spike written: `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md` — 5-metric Sensory; internet + TLS + per-tester tokens; online-only; **create-or-append by natural key** under FOR UPDATE via a new `dve_append_sensory_sample`; dedicated least-privilege `sensory_web` role; the exact JSON contract (numeric scores, the five literal keys, top-level columns); per-sample `sample_uid`; and the required Phase-5.0 desktop merge fix. Owner decisions captured; remaining sign-off items in §10.

**VERIFICATION GATE (5a):** sub-spec written ✔ (2026-06-24); network + auth + create-or-append + least-privilege-role decisions recorded ✔; JSON contract pinned to `sensorySessionToJson` shape ✔. **Owner sign-off on sub-spec §10 still required before service code (5b–5d).** Phase 5.0 (the desktop merge fix) proceeds within the v2.6.0 train under the MS-7 discipline.

### Phase 5b — Build the service (Flask + Docker, create-or-append)

- [ ] **Task 5.2 → MS-8** `deploy/sensory-collect/app.py`: GET serves the mobile-first form (5 sliders 1–9, per-sample name+comments, header test title/tester/assessor/media/date, "add another sample"); POST validates + enforces DATAVIEWER-8 (else 400), builds `json_data` with **numeric** scores, runs one parameterized INSERT (id=−1 semantics, `updated_by="web/<tester>"`), but per the sub-spec the write is **create-or-append by natural key** (resolve-then-append under FOR UPDATE via `dve_append_sensory_sample`, with `sample_uid` idempotency) as the least-privilege `sensory_web` role — NOT a fresh INSERT / auto-suffix-fork. See sub-spec §3 + §5.
- [ ] **Task 5.3 → MS-8** `Dockerfile` + `docker-compose.yml` mirroring `deploy/postgres` + bug-form-app; new stack on the same Docker network (reaches DB by service name, not the firewalled host port); password from the NAS env file (never committed); phone reachability per the chosen fork.
- [ ] **Task 5.4 → MS-8** `tests/test_submit.py`: POST a form, read the row back, assert numeric scores 1–9 + top-level columns. Manual e2e: a desktop on the test container shows the session, renders the radar, re-saves cleanly. Append `tasks/lessons.md`.

**VERIFICATION GATE (5b):** a non-developer submits a 5-metric session from a phone on the target network; stored scores are JSON numbers (pytest); a running desktop gets it live via NOTIFY, opens + renders + re-saves with no version/merge error; empty title/tester rejected; existing (title,tester,date) does not overwrite (original provably unchanged); **repo grep finds no secret**; zero `DataViewer.exe` changes (C++ build + suite still green); stack survives a restart + reaches the DB by internal name.
**Version:** independent NAS service version (not part of the v2.6.0 wrap). Ship when the network decision + spike land.

---

## Phase 6 — Wrap to deployable v2.6.0

**Goal:** fold the internal v2.5.x DataViewer.exe patches into one deployable minor. (Excludes MS-8.)

- [ ] **Task 6.1 (wrap)** Bump `.pro` VERSION → 2.6.0; **clean rebuild** (`mingw32-make clean && mingw32-make -j8`).
- [ ] **Task 6.2 (wrap)** Fresh full `tests\run-tests.ps1` (green except `tst_responsivelayout`); `/ponytail-review` final pass on the cumulative diff.
- [ ] **Task 6.3 (wrap)** Build the installer (`rebuild-dataviewer` flow); verify `release\DataViewer.exe` + `dist\DataViewer-setup.exe` FileVersion = 2.6.0. **No Synology.**
- [ ] **Task 6.4 (wrap)** Run `tests\deployment\Test-Deployment.ps1` on the work machine — all phases pass.
- [ ] **Task 6.5 (wrap)** Final visual pass over every shipped change (y-anchor, clipping, save labels, stopwatch, Notes panel if Phase 2b shipped) + a regression glance at TPM/sensory views; screenshots to the owner.
- [ ] **Task 6.6 (wrap)** Write + commit the v2.6.0 release overview; update the master spec with the final fork decisions; on owner approval, fast-forward into `main`, tag `v2.6.0`, push, cut the next branch (per the branch-to-main workflow).

**VERIFICATION GATE:** clean rebuild + FileVersion match; full suite + deployment self-test green; owner visual sign-off; installer left in-repo for the owner's manual Synology drop.
**Deployable version:** **v2.6.0** (the single minor dropped on Synology by the owner).

---

## Parallel / multi-agent notes

- **Phase 1 (MS-1, MS-2, MS-3)** — best parallel candidate: three independent render/label-only items on disjoint files, no shared state, no forks. Dispatch one agent per item via `superpowers:dispatching-parallel-agents`.
- **Phase 2a (MS-5)** vs **Phase 5b service (MS-8)** can also run in parallel (different file trees: `src/ui/SensoryPanel` vs `deploy/sensory-collect/`), once both gates clear.
- **Phase 4 (MS-7)** — **NOT** a parallel candidate: single-threaded, high data-loss risk, needs sequential e2e + manual concurrency verification.
- **Phase 2b (MS-4)** — keep on one agent (central-widget restructure touches many coupled `MainWindow` paths).

## Untriaged items note

**DATAVIEWER-11/12/13 were triaged in Plane on 2026-06-24** (enriched, labeled `type:feature`, owner decisions captured) and map to **MS-8 / MS-3 / MS-5** respectively. Owner outcomes folded into the spec: **DV-12** = TPM + Sensory only (Detailed fully out, not a separate item); **DV-13** = approved as specced; **DV-11** = the spike sub-spec above (create-or-append, internet exposure, least-privilege role, + the Phase-5.0 desktop fix). All three moved to **Todo** in Plane. No additional untriaged 11/12/13 exist in the repo.
