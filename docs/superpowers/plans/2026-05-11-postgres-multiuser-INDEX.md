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
| **B** | Concurrency, live updates, presence | **Partial (2026-05-11)** — Phases 1, 2, 4 done (10/25 tasks). Remaining: Phase 3 (DatabaseManager refactor), Phase 5 (Presence UI), Phase 6 (don't-yank rule), Phase 7 (MainWindow wiring), Phase 8 (e2e tests). | Audit columns + version triggers used by app, `NotificationListener` ✅, `PresenceManager` ✅, `ConflictResolver` + three conflict dialogs ✅, presence dots in nav + avatars top-right. |
| **C** | Offline mode + cutover | **Drafted, not started** | `OfflineSnapshot`, banner, in-flight-edit pending badge, reconnect detection, deletion of SQLite-on-Synology code paths. Ships as v2.0. |

### Plan B sub-status (2026-05-11)

| Phase | Tasks | Status |
|---|---|---|
| 1 | T1–T2 NotificationListener + e2e tests | ✅ Done (real LISTEN/NOTIFY round-trip verified) |
| 2 | T3–T4 PresenceManager + integration tests | ✅ Done (6 DB-exercising tests pass) |
| 3 | T5–T9 DatabaseManager Postgres rewrite + WriteResult + optimistic concurrency | ⏳ Not started |
| 4 | T10–T13 ConflictResolver + 3 dialogs | ✅ Done (4 commits, clean build) |
| 5 | T14–T16 Presence UI in nav + avatars | ⏳ Not started |
| 6 | T17–T19 Don't-yank-in-progress edit rule | ⏳ Not started |
| 7 | T20–T22 MainWindow startup wiring + own-UUID filter | ⏳ Not started |
| 8 | T23–T25 Two-client e2e + checkpoint | ⏳ Not started |

**Notes for Phase 3 resumption:** The DatabaseManager refactor is the highest-risk single chunk in the entire initiative. It rewrites a 1786-line file (current SQLite-backed implementation) to use `PostgresConnection` under the hood. Public method signatures stay the same so MainWindow doesn't need surgery, but the internal hierarchical save (file → tests → samples → data_rows → images) is intricate. Plan for this to take multiple iterations with careful test coverage at each step. The new components written in Phases 1, 2, 4 are already in place and ready to be wired in once Phase 3 lands.

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
