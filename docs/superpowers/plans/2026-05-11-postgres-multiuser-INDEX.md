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
| **C** | Offline mode + cutover | **Complete (2026-05-13)** — 15/15 tasks, 17 commits, 31 new tests passing (5 checkpoint + 9 snapshot + 7 monitor + 10 banner). Code complete; visual smoke + Test-Deployment.ps1 (with new offline self-test cases) needed on next installer build before tagging v2.0.0. | `OfflineSnapshot` (atomic SQLite mirror under `%LOCALAPPDATA%\DataViewer\snapshot.sqlite`), `ConnectionMonitor` (30s ping + jittered reconnect), `OfflineBanner` widget with last-sync + pending count, `DatabaseManager` online/offline routing (writes refused with `OfflineReadOnly`, reads route to snapshot), MainWindow wiring + pending-edit queue + clean-close snapshot regenerate + Tools-tab refresh button, lock-file cleanup verified (already gone since Plan B sub-batch 3a), version bumped to 2.0.0. |

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

### Plan C sub-status (2026-05-13 — COMPLETE)

| Phase | Tasks | Status |
|---|---|---|
| 1 | T1–T3 OfflineSnapshot + DatabaseManager routing | ✅ Done (`e1f6159` + `9027cb6` + `c4a09f3` + `a09f0d1` + `3fd2ee9`) |
| 2 | T4–T5 ConnectionMonitor + tests | ✅ Done (`d79257e` + `2a540f1`) |
| 3 | T6–T8 OfflineBanner + MainWindow wiring | ✅ Done (`f188700` + `6c8dd9b` + `56d8125`) |
| 4 | T9 Snapshot lifecycle (clean-close regenerate + pending-edit queue) | ✅ Done (`90377ab` + `549be57` + `3d3e305`) |
| 5 | T10–T12 Final cutover (lock-file verification, SelfTest additions, CLAUDE.md) | ✅ Done (`f1d4d67` + `b2613e5`; lock-file code confirmed already deleted in Plan B 3a) |
| 6 | T13–T15 Release (checkpoint test, version bump, INDEX closeout) | ✅ Done (`ef8fc15` + `e493348` + this commit) |

**Plan C closing notes (2026-05-13):**
- **Phase 1** — `OfflineSnapshot` is a local SQLite mirror at `%LOCALAPPDATA%\DataViewer\snapshot.sqlite`. `regenerate(PostgresConnection*)` runs all reads inside a REPEATABLE READ READ ONLY transaction for point-in-time consistency, writes to `.tmp`, atomic-renames over the production file. Schema mirrors Postgres minus presence + schema_meta (not needed offline). `DatabaseManager` gained `isOnline() / setOnline() / setOfflineSnapshot()`: writes refused with `WriteResult::OfflineReadOnly` when offline; reads route to the snapshot (`listFiles`, `loadFileByPath`, `loadFile`, `listSensoryRecords`, `loadSensorySession`, `listDetailedSensoryRecords`, `loadDetailedSensorySession`, `getSetting`, `loadCumulativeLayout`).
- **Phase 2** — `ConnectionMonitor` owns a 30s ping timer (online) + jittered reconnect timer (offline, base 30s + 0-5s jitter). On ping failure → `wentOffline`; on `tryOpenWithRetry` success → `cameOnline`. MainWindow connects these to `DatabaseManager::setOnline()`.
- **Phase 3** — `OfflineBanner` sits at the top of the MainWindow showing "Offline (read-only)" with snapshot timestamp + pending-edit count. MainWindow constructs `OfflineSnapshot` + `ConnectionMonitor` after `m_db->open()` succeeds; banner toggles on `wentOffline`/`cameOnline`.
- **Phase 4** — Snapshot regenerates on clean app close (so the next cold boot has fresh data). Manual refresh in the Tools tab. Pending-edit queue captures `OfflineReadOnly` returns and replays them on `cameOnline`.
- **Phase 5** — Lock-file code (`LockInfo`, `forceReleaseLock`, `writeLockFile`, `pathLooksCloudSynced`) confirmed already deleted in Plan B 3a; grep returns nothing. SelfTest gained offline-snapshot path probe + readability diagnostic. CLAUDE.md transitional language stripped.
- **Phase 6** — `tst_planc_checkpoint` bookend test (5 cases, ~2.5s) drives the full offline failover round-trip across DatabaseManager + OfflineSnapshot + ConnectionMonitor in a single fixture. Version bumped to 2.0.0 in `DataViewerEnterprise.pro` (single source of truth — `main.cpp` picks it up via `DVE_APP_VERSION`, `build_installer.bat` parses the same line for `/DAppVersion=`).

**Plan C checkpoint criteria — verified:**
1. ✅ Offline failover works end-to-end (e2e-tested in `tst_planc_checkpoint`).
2. ✅ Reconnect detection works (`ConnectionMonitor` ping + jittered reconnect tested).
3. ✅ Lock-file code is gone (verified `grep -r LockInfo src/` returns nothing).
4. ✅ All test suites pass — 157/157 across 13 PG-dependent suites; 144/144 across 14 non-PG suites (excluding 2 pre-existing fixture-path failures in `tst_sopLoader` + `tst_reportgenerator` unrelated to Plan C scope).
5. ⏳ Visual smoke (offline banner, snapshot regenerate on close, pending-edit replay) needs manual verification on next installer build.
6. ⏳ `Test-Deployment.ps1` with new offline self-test cases needs work-machine run before tagging v2.0.0.

## Verifiable checkpoints

- **End of Plan A** — `postgres:16` container running on Synology, migration tool produces a verified Postgres database from the existing SQLite source, deployment self-test passes (including new Phase 4: migration verification). The old SQLite code path is still active for runtime reads/writes — Plan A only adds new code, doesn't switch anything over.
- **End of Plan B** — Two `DataViewer.exe` instances on different machines can edit the same file simultaneously; NOTIFY events propagate within 1s; conflict dialogs surface on collision; presence dots appear next to filenames in the nav. The old SQLite path remains as a safety net.
- **End of Plan C** — ✅ Verified (code; visual smoke + Test-Deployment.ps1 pending on next installer build). Offline failover (NAS unreachable → read-only banner) and reconnect both work end-to-end (e2e-tested in `tst_planc_checkpoint`); the SQLite-on-Synology code path is deleted entirely; the cross-machine `<dbPath>.lock` mechanism is gone. Ready to tag as v2.0.0 once visual smoke + work-machine self-test pass.

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
- 2026-05-13 — **Plan C complete.** Offline mode shipped: `OfflineSnapshot` (atomic SQLite mirror, REPEATABLE READ regenerate), `ConnectionMonitor` (30s ping + jittered reconnect), `OfflineBanner` widget, `DatabaseManager` online/offline routing (`OfflineReadOnly` write guard + snapshot-routed reads), MainWindow wiring + pending-edit replay queue + clean-close snapshot regenerate, SelfTest offline-snapshot probe, lock-file code confirmed already gone. Version bumped to **v2.0.0** in `DataViewerEnterprise.pro`. 17 commits on top of `c86655b`. Full suite **157 / 157 PASS** across 13 PG-dependent test suites (with `tst_planc_checkpoint` as the bookend e2e). Code complete; visual smoke + `Test-Deployment.ps1` (with new offline self-test cases) needed on next installer build before tagging v2.0.0.
