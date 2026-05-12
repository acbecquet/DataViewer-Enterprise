# Postgres Multi-User Database — Plan Index

> Single entry point for the multi-plan migration of DataViewer Enterprise's
> database layer from SQLite-on-Synology to PostgreSQL-on-NAS. Always start
> here before working on anything DB-related.

- **Date:** 2026-05-11
- **Spec:** [docs/superpowers/specs/2026-05-11-postgres-multiuser-design.md](../specs/2026-05-11-postgres-multiuser-design.md)

## Plans

| Plan | Title | Status | Scope |
|---|---|---|---|
| **A** | Foundation, schema, migration | **Complete (2026-05-11)** — 30 tasks, 42+ commits, 34 new tests passing, deferred work-machine checkpoint only | Docker compose + `init.sql` on NAS, `PostgresConnection`, `IdentityManager`, `ConfigLoader`, `MigrationTool`, libpq bundling, installer wiring. Existing SQLite path remains active in parallel. |
| **B** | Concurrency, live updates, presence | **Drafted, in flight** | Audit columns + version triggers used by app, `NotificationListener`, `PresenceManager`, `ConflictResolver`, three conflict dialogs, presence dots in nav + avatars top-right. |
| **C** | Offline mode + cutover | **Drafted, not started** | `OfflineSnapshot`, banner, in-flight-edit pending badge, reconnect detection, deletion of SQLite-on-Synology code paths. Ships as v2.0. |

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
