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
| **B** | Concurrency, live updates, presence | **Partial (2026-05-12)** — Phases 1, 2, 3, 4 done (15/25 tasks). Remaining: Phase 5 (Presence UI), Phase 6 (don't-yank rule), Phase 7 (MainWindow wiring), Phase 8 (e2e tests). | Audit columns + version triggers used by app, `NotificationListener` ✅, `PresenceManager` ✅, `DatabaseManager` rewired to Postgres with optimistic concurrency ✅, `ConflictResolver` + three conflict dialogs ✅, presence dots in nav + avatars top-right. |
| **C** | Offline mode + cutover | **Drafted, not started** | `OfflineSnapshot`, banner, in-flight-edit pending badge, reconnect detection, deletion of SQLite-on-Synology code paths. Ships as v2.0. |

### Plan B sub-status (2026-05-12)

| Phase | Tasks | Status |
|---|---|---|
| 1 | T1–T2 NotificationListener + e2e tests | ✅ Done (real LISTEN/NOTIFY round-trip verified) |
| 2 | T3–T4 PresenceManager + integration tests | ✅ Done (6 DB-exercising tests pass) |
| 3 | T5–T9 DatabaseManager Postgres rewrite + WriteResult + optimistic concurrency | ✅ Done 2026-05-12 (8 commits across sub-batches 3a/3b/3c/3d, 30 retargeted tests pass) |
| 4 | T10–T13 ConflictResolver + 3 dialogs | ✅ Done (4 commits, clean build) |
| 5 | T14–T16 Presence UI in nav + avatars | ⏳ Not started |
| 6 | T17–T19 Don't-yank-in-progress edit rule | ⏳ Not started |
| 7 | T20–T22 MainWindow startup wiring + own-UUID filter | ⏳ Not started — **blocks the Phase 3 follow-up: see below** |
| 8 | T23–T25 Two-client e2e + checkpoint | ⏳ Not started |

**Phase 3 closing notes (2026-05-12):** DatabaseManager rewired to Postgres across 4 sub-batches with subagent-driven reviews per batch.
- **3a** (`9dd54c0` + `aad968f`): backend swap, lock-file machinery deleted, MainWindow constructor patched, two test SUBDIRS temporarily disabled.
- **3b** (`04715ec` + `0a5522c`): files-hierarchy methods + perf fix to hoist prepared statements out of the inner save loops.
- **3c** (`16c3c84` + `9e03e95`): sensory + detailed sensory + settings + layouts; layout_json preservation across ON CONFLICT DO UPDATE.
- **3d code** (`f232e45` + `f25cc0b`): `WriteResult` enum, optimistic UPDATE with `WHERE id = ? AND version = ?`, follow-up SELECT to distinguish VersionMismatch vs RowDeleted, SQLSTATE 23505 → UniqueViolation, bool overload shims preserved.
- **3d tests** (`bbab113`): `tst_databasemanager` retargeted at ephemeral Postgres with 30 tests (8 new conflict-path tests for VersionMismatch/RowDeleted/UniqueViolation); `tst_sensoryreportsource` surgically updated. All 8 test suites pass: **101 / 101 PASS**.

**Phase 3 carry-over for Phase 7:** Several MainWindow save call sites depend on the old natural-key UPSERT semantics. After 3d-code, fresh structs sharing natural keys with existing DB rows now return `WriteResult::UniqueViolation`, which the bool shims map to `false` (silent failure for now). Phase 7 (Task 21) must route every save through `ConflictResolver` so these surface to the user. Specifically `MainWindow.cpp:2908` (re-import-from-disk flow) needs a fresh-struct → existing-row id+version copy before save.

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
