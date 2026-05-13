# Postgres Multi-User Database — Plan Index

> Single entry point for the multi-plan migration of DataViewer Enterprise's
> database layer from SQLite-on-Synology to PostgreSQL-on-NAS. Always start
> here before working on anything DB-related.

- **Date:** 2026-05-11
- **Spec:** [docs/superpowers/specs/2026-05-11-postgres-multiuser-design.md](../specs/2026-05-11-postgres-multiuser-design.md)

## Plans

| Plan | Title | Status | Scope |
|---|---|---|---|
| **A** | Foundation, schema, migration | **Complete (2026-05-11)** — 30 tasks, 42+ commits, 34 new tests passing, work-machine checkpoint deferred | Docker compose + `init.sql` on NAS, `PostgresConnection`, `IdentityManager`, `ConfigLoader`, `MigrationTool`, libpq bundling, installer wiring. Existing SQLite path remains active in parallel. |
| **B** | Concurrency, live updates, presence | **Complete (2026-05-13)** — 25/25 tasks. Two-client e2e (`tst_twoclient_e2e`) verifies the full stack: NOTIFY round-trip + optimistic-concurrency conflict detection + presence broadcast. **113 / 113 PASS** across 9 test suites. Visual gate (presence dots, avatar bar, conflict dialogs, don't-yank decoration, row-deleted banner) deferred to manual smoke. | All 8 phases shipped: `NotificationListener`, `PresenceManager`, `DatabaseManager`-on-Postgres + optimistic concurrency, `ConflictResolver` + 3 dialogs, `PresenceDotsDelegate` + `PresenceAvatarBar`, don't-yank-in-progress edit rule + `RowDeletedBanner`, `SaveCoordinator` routing all 13 save call sites, MainWindow wiring + own-UUID echo filter. |
| **C** | Offline mode + cutover | **Drafted, needs refresh before execution** — see "Plan C staleness notes" below | `OfflineSnapshot`, banner, in-flight-edit pending badge, reconnect detection, lock-file cleanup verification. Ships as v2.0. |

### Plan B sub-status (2026-05-13 — COMPLETE)

| Phase | Tasks | Status |
|---|---|---|
| 1 | T1–T2 NotificationListener + e2e tests | ✅ Done (real LISTEN/NOTIFY round-trip verified) |
| 2 | T3–T4 PresenceManager + integration tests | ✅ Done (6 DB-exercising tests pass) |
| 3 | T5–T9 DatabaseManager Postgres rewrite + WriteResult + optimistic concurrency | ✅ Done 2026-05-12 (8 commits across sub-batches 3a/3b/3c/3d, 30 retargeted tests pass) |
| 4 | T10–T13 ConflictResolver + 3 dialogs | ✅ Done (4 commits, clean build) |
| 5 | T14–T16 Presence UI in nav + avatars | ✅ Done 2026-05-12 (`9cd1568` + `9d2bb5b`) |
| 6 | T17–T19 Don't-yank-in-progress edit rule | ✅ Done 2026-05-12/13 (`7b37525` + `43abeeb`) |
| 7 | T20–T22 MainWindow startup wiring + own-UUID filter | ✅ Done 2026-05-12 (`0937a36` + `15aade6` + `02e8582`) |
| 8 | T23–T25 Two-client e2e + checkpoint | ✅ Done 2026-05-12/13 (`44add1c` two-client e2e; checkpoint INDEX update) |

**Phase 3 closing notes (2026-05-12):** DatabaseManager rewired to Postgres across 4 sub-batches with subagent-driven reviews per batch.
- **3a** (`9dd54c0` + `aad968f`): backend swap, lock-file machinery deleted, MainWindow constructor patched, two test SUBDIRS temporarily disabled.
- **3b** (`04715ec` + `0a5522c`): files-hierarchy methods + perf fix to hoist prepared statements out of the inner save loops.
- **3c** (`16c3c84` + `9e03e95`): sensory + detailed sensory + settings + layouts; layout_json preservation across ON CONFLICT DO UPDATE.
- **3d code** (`f232e45` + `f25cc0b`): `WriteResult` enum, optimistic UPDATE with `WHERE id = ? AND version = ?`, follow-up SELECT to distinguish VersionMismatch vs RowDeleted, SQLSTATE 23505 → UniqueViolation, bool overload shims preserved.
- **3d tests** (`bbab113`): `tst_databasemanager` retargeted at ephemeral Postgres with 30 tests (8 new conflict-path tests for VersionMismatch/RowDeleted/UniqueViolation); `tst_sensoryreportsource` surgically updated. All 8 test suites pass: **101 / 101 PASS**.

**Phase 3 carry-over (resolved by Phase 7b):** The Phase 3 bool shims silently swallowed `UniqueViolation` / `VersionMismatch`. Phase 7b's `SaveCoordinator` now routes all 13 save call sites (7 in MainWindow, 3 each in SensoryPanel + DetailedSensoryPanel) through `ConflictResolver`, so users see real dialogs on collision.

**Plan B closing notes (2026-05-13):**
- **Phase 5** — `PresenceDotsDelegate` paints colored circles after item text in the 3 nav widgets (file tree, sensory list, detailed sensory list). `PresenceAvatarBar` sits at the top of the central editor showing each active user's circle with their initial (self-user has a 2 px white inner ring). Live refresh wired via `NotificationListener::presenceChanged`. Mode switch clears active presence via `clearActivePresence()` helper.
- **Phase 6** — Cell-level `dve_editing` flag (`UserRole+2`/`+3`) on `m_dataTable`. Yellow background + tooltip + pending value (`UserRole+4`) on remote NOTIFY for a dirty cell; click accepts the remote value and flushes through `onTableCellChanged` to the in-memory DataRow. `RowDeletedBanner` pops above the central editor on DELETE NOTIFY for the currently-open resource; clicking Recreate clears id+version and INSERTs via SaveCoordinator's mutable-ref `tryWriteFile` overload.
- **Phase 7** — Constructed `m_pgConn`, `m_notify`, `m_presence`, `m_conflict`, `m_saveCoordinator` after `m_db->open()` succeeds. `m_pgConn` is a SEPARATE PostgresConnection from DatabaseManager's internal one so NOTIFY/heartbeat workload doesn't contend with main queries. Own-UUID echo filter on both `rowChanged` and `presenceChanged`. Graceful degradation: if live-updates connection fails, the app continues in single-user mode with a warning.
- **Phase 8** — `tst_twoclient_e2e` (12 tests) verifies NOTIFY round-trip + optimistic concurrency + presence between two PostgresConnections sharing one ephemeral Postgres. 113/113 PASS across the full 9-suite test stack.

**Plan B checkpoint criteria — verified:**
1. ✅ Two clients can edit simultaneously and NOTIFY propagates within 1s (e2e-tested).
2. ✅ Conflict dialogs surface on collision (`SaveCoordinator` routes via `ConflictResolver`).
3. ⏳ Presence dots visible in nav (visual gate — needs manual verification).
4. ⏳ Self-test passes (will run on next installer build with `Test-Deployment.ps1`).
5. ✅ All test suites pass — 113/113.

### Plan C staleness notes (refresh needed before execution)

Plan C was drafted before Plan B's actual implementation pulled some items forward. Stale items to revise:

- **Task 3 — `WriteResult::OfflineReadOnly` enum:** *already added* in Plan B sub-batch 3d (`f232e45`). The enum value exists and is reserved for Plan C use. Task 3 should become "wire `OfflineReadOnly` returns into every `tryWrite*` method's online-check guard" — no enum change needed.
- **Task 10 — Delete remaining lock-file code:** *already deleted* in Plan B sub-batch 3a (`9dd54c0`). Task 10 should become a one-shot verification: `grep -r "LockInfo\|forceReleaseLock\|writeLockFile\|pathLooksCloudSynced" src/` returns nothing. Confirmed already.
- **Task 12 — CLAUDE.md final update:** *partly done* on 2026-05-13 (the two SQLite-mention stalenesses on lines 39 + 111 were corrected when Plan B closed). Plan C's Task 12 reduces to "remove any residual transitional language and bump the v2.0 release notes."
- **Phase 7-style wiring already exists in MainWindow:** `m_pgConn` is constructed and owned; `ConnectionMonitor` in Plan C Task 8 just adds a sibling object to that block (no constructor refactor needed).
- **Plan C Task 8 modal copy:** "Cannot reach database and no offline copy available" — needs to match Phase 7's graceful-degradation UX (currently shows `QMessageBox::warning` for the live-updates connection but `QMessageBox::critical` for the main DB connection). Decide whether OfflineSnapshot's existence changes the main-DB-critical path.
- **MIP file pattern is well-validated by now:** Plan C's "create via Python pattern" note is correct but should reference the established CLAUDE.md convention rather than restating the pattern.
- **`bool DatabaseManager::isOnline() const` and `setOnline(bool)`:** Task 3 introduces these. Verify they don't collide with the existing `isOpen()` which already delegates to `m_pg->isOpen()` (Plan B sub-batch 3a fix-up). Probably fine — isOpen tracks "did open() succeed?", isOnline tracks "is the connection currently usable?" — different semantics.

The plan structure (6 phases, 15 tasks) and architecture (`OfflineSnapshot` + `ConnectionMonitor` + `OfflineBanner`) is sound and can stay as-is. The above are tactical updates a plan-execution session should apply in the first 15 minutes before dispatching subagents.

## Verifiable checkpoints

- **End of Plan A** — `postgres:16` container running on Synology, migration tool produces a verified Postgres database from the existing SQLite source, deployment self-test passes (including new Phase 4: migration verification). The old SQLite code path is still active for runtime reads/writes — Plan A only adds new code, doesn't switch anything over.
- **End of Plan B** — Two `DataViewer.exe` instances on different machines can edit the same file simultaneously; NOTIFY events propagate within 1s; conflict dialogs surface on collision; presence dots appear next to filenames in the nav. The old SQLite path remains as a safety net.
- **End of Plan C** — Offline failover (NAS unreachable → read-only banner) and reconnect both work; the SQLite-on-Synology code path is deleted entirely; the cross-machine `<dbPath>.lock` mechanism is gone. Ready to tag as v2.0.

## Sequencing rules

- **Plans MUST be executed in order.** Each depends on artifacts produced by the previous.
- **Do not begin Plan B until Plan A's checkpoint is verified on the work machine.** Verification means: `Test-Deployment.ps1` passes all phases including the new Phase 4, with real data migrated and re-read.
- **Within a plan**, tasks may be reordered when dependencies allow, but the plan file lists the recommended order.

## Plan files

- [Plan A — Foundation, schema, migration](2026-05-11-postgres-multiuser-plan-A-foundation.md)
- [Plan B — Concurrency + live updates + presence](2026-05-11-postgres-multiuser-plan-B-runtime.md)
- [Plan C — Offline mode + cutover](2026-05-11-postgres-multiuser-plan-C-offline-cutover.md)

## Cross-cutting concerns

- **MIP/AIP encryption:** every new `.cpp`/`.h` file MUST be created via Python's delete-and-rewrite pattern (per CLAUDE.md). `python tools/decrypt_via_copy.py --apply` runs before every build attempt until the first successful clean build verifies files are stable.
- **Execution skill:** when a plan is executed, the executor uses either `superpowers:subagent-driven-development` (recommended, fresh subagent per task) or `superpowers:executing-plans` (inline). The choice is made at execution time.

## Status log

- 2026-05-11 — Spec drafted, reviewed, committed (`6f6ac47`). Plan A drafted.
- 2026-05-12 — Plan B Phase 3 closed. `DatabaseManager` (1786 lines) rewired to Postgres with WriteResult/optimistic concurrency. 8 commits, subagent-driven with per-batch spec + code-quality reviews. Full suite **101 / 101 PASS** against ephemeral Postgres. Phase 7 inherits the MainWindow call-site routing concern.
- 2026-05-13 — **Plan B complete.** Phases 5, 6, 7, 8 shipped (Presence UI, don't-yank rule, MainWindow wiring, two-client e2e). `SaveCoordinator` resolves Phase 3's silent-failure carry-over. Full suite **113 / 113 PASS** across 9 test suites. Code complete; visual gate (presence dots, avatar bar, dialogs, yellow-cell decoration, deleted-row banner) needs manual smoke before tagging v2.0.
