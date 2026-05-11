# Postgres Multi-User Database — Design

- **Date:** 2026-05-11
- **Status:** Draft (pending user review)
- **Scope (v1):** Replace the single-writer SQLite-on-Synology setup with a
  PostgreSQL instance on the NAS. Apply to all four data types: TPM files,
  sensory sessions, detailed sensory sessions, and embedded images. Add
  optimistic-concurrency conflict resolution, live UI updates via NOTIFY,
  presence indicators, and a hybrid online/offline mode (read-only when the
  NAS is unreachable).

## Goal

Remove the cross-machine file-lock (`<dbPath>.lock`) and the
SQLite-on-Synology-Drive architecture entirely. Replace with a real
PostgreSQL server on the Synology so that multiple users on the LAN can
read and write the database simultaneously without corruption, with the
following user-visible behavior:

1. Anyone can open and edit any file/session at any time — no lock errors.
2. Edits from other users appear in the UI within ~1 second (no manual
   refresh).
3. Each file/session in the list shows a colored dot per active user, with
   hover tooltip showing names and intent (viewing vs editing).
4. When two users edit the same row, the second saver gets a clear
   "merge / overwrite / keep" dialog — never silent data loss.
5. When the NAS is unreachable, the app opens against a local snapshot in
   read-only mode with a polite banner. Writes are blocked until the
   connection returns.
6. Existing data migrates losslessly. The pre-migration SQLite file is
   preserved on Synology for one release cycle as a rollback escape hatch.

## Non-goals (v1)

- **DB-level user accounts.** Single shared `dataviewer_app` role. Per-user
  Postgres accounts can come in v2 if audit compliance demands.
- **Real-time cursor/selection broadcasting.** Presence is at the
  resource level only ("Sarah is in this file"), not "Sarah's cursor is
  at row 12 col 3".
- **Server-side replication / HA.** Single Postgres container on one NAS.
  Disaster recovery is Hyper Backup, not hot standby.
- **Separate audit log table.** Audit columns (`updated_by`, `updated_at`,
  `version`) cover the immediate need. A dedicated full-history audit log
  is a clean v2 add.
- **TLS between client and Postgres.** LAN-only behind corporate firewall.
  v1.1 add.
- **Excel file concurrency on Synology Drive.** Different problem with
  separate solutions. (Long-term, the answer may be native data entry
  replacing Excel ingestion — tracked in user memory, not in v1 scope.)
- **Cell-level auto-merge.** Conflicts always surface a dialog; the dialog
  supports per-field selection but does not silently auto-merge two
  edits to different fields of the same row.
- **Postgres major-version upgrades.** Pinned to `postgres:16`. Major
  upgrades are deliberate events with planned data-dir migration.
- **Offline writes.** Offline is strictly read-only; no write queue.
- **Validating auto-update download+replace** and `SingleInstance` IPC —
  same exclusions as today's deployment test.

## Architecture

### Topology

```
┌──────────────────────┐    TCP/5432 (LAN only)     ┌────────────────────────────┐
│  DataViewer.exe      │ ─────────────────────────▶ │  Synology NAS              │
│  (each workstation)  │     LISTEN/NOTIFY rides    │  └─ Container Manager      │
│                      │     the same connection    │     └─ postgres:16         │
│  ┌────────────────┐  │                            │        ├─ data dir         │
│  │ Local SQLite   │  │    optimistic upserts      │        │   (bind-mount)    │
│  │ snapshot       │  │    (version-checked)       │        └─ port 5432        │
│  │ (read-only     │  │                            │                            │
│  │  when offline) │  │                            │  └─ Hyper Backup job       │
│  └────────────────┘  │                            │                            │
└──────────────────────┘                            └────────────────────────────┘
```

### New components in `src/database/`

- **`PostgresConnection`** — owns the live TCP connection, retry logic,
  online/offline detection. Holds two `QSqlDatabase` instances: one for
  the notification thread (LISTEN blocks a connection), one for the UI
  thread's queries.
- **`NotificationListener`** — wraps `QSqlDriver::subscribeToNotification()`;
  parses NOTIFY payloads and emits typed Qt signals
  (`rowChanged(table, op, id, updatedByUuid)`,
  `presenceChanged(uuid, resourceType, resourceId, intent)`).
- **`PresenceManager`** — 10-second heartbeat upsert, single-resource-per-
  user rule, stale cleanup; publishes presence to UI.
- **`OfflineSnapshot`** — local SQLite mirror at
  `%LOCALAPPDATA%\DataViewer\snapshot.sqlite`. Refreshed on clean app
  close + manual menu item. Read-only when used.
- **`ConflictResolver`** — three dialog types
  (version-mismatch / row-deleted / unique-violation) and the
  multi-row transaction rollback logic.
- **`IdentityManager`** — first-launch **modal dialog** (dismissed once,
  never shown again) prompts for display name + color; generates a
  stable UUID per install; persisted to `QSettings`. Subsequent
  identity changes go through Tools menu, not a modal.

### Components that change

- **`DatabaseManager`** — becomes a thin facade over `PostgresConnection`
  + `OfflineSnapshot`. Existing public API
  (`saveFile`, `loadFileByPath`, `saveSensorySession`, ...) keeps its
  signatures; internal implementation is fully rewired. The lock-file
  code path (`writeLockFile`, `pathLooksCloudSynced`, `LockInfo`,
  `forceReleaseLock`) is **deleted**.

### Components untouched

- All in-memory data structures (`FileResult`, `SensorySession`,
  `DetailedSensorySession`, image BLOBs).
- Excel pipeline, plotting, reporting, write-back, bundled Python +
  openpyxl.
- Mode-switching UI, ribbon, sensory panels.
- Auto-updater, translator launcher, self-test framework (beyond the
  new DB-related cases added to `SelfTest.cpp`).

### Deletions

- `<dbPath>.lock` sidecar mechanism and "Force Release Lock" UI — gone.
- `pathLooksCloudSynced` / DELETE-journal fallback — SQLite-on-Synology
  is no longer a deployment target.

## Schema design

### Type translation (SQLite → PostgreSQL)

| SQLite | PostgreSQL | Notes |
|--------|-----------|-------|
| `INTEGER PRIMARY KEY AUTOINCREMENT` | `BIGSERIAL PRIMARY KEY` | future-proof |
| `INTEGER` | `INTEGER` | unchanged |
| `REAL` | `DOUBLE PRECISION` | unchanged precision |
| `TEXT` | `TEXT` | unchanged |
| `BLOB` (images) | `BYTEA` | unchanged semantics |
| `TEXT` containing JSON (`json_data`, `layout_json`) | `JSONB` | indexable; migration validates each row |
| `TEXT` ISO date strings (`loaded_at`, `date`, `timestamp`) | **kept as `TEXT`** | preserves exact existing values, zero parse risk during migration |

### Audit columns

Added to every editable table (`files`, `tests`, `samples`, `data_rows`,
`images`, `sensory_sessions`, `sensory_images`,
`detailed_sensory_sessions`, `detailed_sensory_images`, `settings`):

```sql
updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
updated_by TEXT        NOT NULL DEFAULT 'migration',
version    INTEGER     NOT NULL DEFAULT 1
```

A `BEFORE UPDATE` trigger increments `version` and stamps `updated_at`
automatically. The app supplies `updated_by` (its UUID) on every write.

### New tables

```sql
CREATE TABLE presence (
    user_uuid      UUID        NOT NULL,
    user_name      TEXT        NOT NULL,
    user_color     TEXT        NOT NULL,
    resource_type  TEXT        NOT NULL,   -- 'file' | 'sensory_session' | 'detailed_sensory_session'
    resource_id    BIGINT      NOT NULL,
    intent         TEXT        NOT NULL DEFAULT 'viewing',  -- 'viewing' | 'editing'
    last_heartbeat TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (user_uuid, resource_type, resource_id)
);
CREATE INDEX idx_presence_resource ON presence(resource_type, resource_id);

CREATE TABLE schema_meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
-- seeded: ('schema_version', '2'), ('migrated_from', '<path>'),
--         ('migrated_at', '<iso>'), ('source_sha256', '<hex>')
```

### Constraints preserved

- `UNIQUE(files.file_path)`
- `UNIQUE(sensory_sessions(session_name, tester_name, date))`
- `UNIQUE(detailed_sensory_sessions(session_name, tester_name, date))`
- All existing `ON DELETE CASCADE` FK relationships.

### NOTIFY plumbing

Two channels:

- `dataviewer_changes` — fired by `AFTER INSERT/UPDATE/DELETE` triggers on
  every editable table. Payload: `{table, op, id, updated_by}`.
- `dataviewer_presence` — fired by trigger on `presence`. Same shape.

A single trigger function (`notify_row_change`) is reused across all
data tables; a parallel function (`notify_presence_change`) covers
presence.

## Online / offline flow

### Startup decision tree

```
Try TCP connect to Postgres (3s timeout)
    │
    ├── success ──▶ ONLINE mode
    │              ├─ subscribe to NOTIFY channels
    │              ├─ start presence heartbeat
    │              └─ regenerate offline snapshot in background
    │
    └── failure ──▶ check for local snapshot
                   │
                   ├── snapshot exists ──▶ OFFLINE mode (read-only)
                   │                       └─ banner: "Working offline — read-only.
                   │                                    Last synced [timestamp].
                   │                                    [Retry connection]"
                   │
                   └── no snapshot ──────▶ Modal: "Cannot reach database and no
                                                   offline copy is available.
                                                   Connect to the network and retry."
```

### Mid-session disconnect detection

Three signals, whichever fires first:

1. The NOTIFY socket dies (`QSqlDriver::notification` stops + ping fails).
2. A query fails with a connection-level SQLSTATE (`08006`, `57P01`, ...).
3. The 30-second `SELECT 1` ping fails.

On detected disconnect → switch to offline mode mid-session, banner
appears, NOTIFY subscription dropped, presence heartbeat stops.

### In-flight edits during disconnect

- Edits the user is currently typing stay in memory. Nothing is silently
  discarded.
- The cell shows a small "pending" badge. The banner gets a second line:
  `"1 unsaved change. Will retry when reconnected."`
- On reconnect, the app re-attempts the save under optimistic
  concurrency. If the row was changed remotely while offline → standard
  conflict dialog. If not → save proceeds, badge clears.

### Reconnect detection

Background `SELECT 1` ping every 30 s while offline (with 0–5 s jitter
to avoid thundering-herd on NAS reboot). On success: re-subscribe to
NOTIFY, restart presence, attempt to flush any in-flight edits, hide
the banner with a brief "Reconnected" toast.

### Snapshot refresh policy

- **On clean app close (online):** full dump to
  `%LOCALAPPDATA%\DataViewer\snapshot.sqlite`. Atomic write (write
  `.tmp`, then `MoveFileEx` with `REPLACE_EXISTING`).
- **Manual refresh:** File → "Refresh Offline Snapshot" menu item.
- **No automatic mid-session refresh in v1** — adds complexity (delta
  sync) without much gain.

The snapshot is the only place SQLite is still used. Existing
`DatabaseManager` SQLite code paths are retargeted to read/write this
snapshot file.

## Concurrency model

### Read pattern

```sql
SELECT id, sample_name, viscosity, ..., version
  FROM samples WHERE test_id = $1;
```

App keeps `version` alongside data in memory. UI never displays it.

### Single-row write

```sql
UPDATE samples
   SET sample_name = $1, viscosity = $2, ..., updated_by = $user_uuid
 WHERE id = $row_id AND version = $expected_version;
```

App inspects `numRowsAffected()`:

- `1` → success, local version = `expected_version + 1`.
- `0` → conflict. A diagnostic follow-up query disambiguates:

```sql
SELECT version FROM samples WHERE id = $row_id;
-- no row returned  → DELETED conflict
-- different version → VERSION_MISMATCH conflict
```

### Multi-row write

```sql
BEGIN;
  UPDATE files SET ... WHERE id=$1 AND version=$fv;   -- rowcount must be 1
  UPDATE tests SET ... WHERE id=$2 AND version=$tv;   -- rowcount must be 1
  ... dozens more ...
COMMIT;
```

If any UPDATE returns 0 → `ROLLBACK`, present **one** conflict dialog for
the offending entity. No partial state ever lands.

### The three conflict dialogs

**1. Version mismatch (most common)** — full-row side-by-side picker
with per-field radio buttons. Identical fields shown as `●same` and not
selectable. Differences default to "yours" but flip with one click.
Bulk buttons: `Keep all mine` / `Take all theirs` / `Apply selection`.
On submit → re-fetch current version, apply merged values, retry save.

**2. Row deleted** — message + list of user's unsaved changes;
`Recreate as new` (INSERT, new id, version=1, user's edits preserved)
vs `Discard my changes`.

**3. UNIQUE violation** — only relevant for `files(file_path)` and
sensory `(session_name, tester_name, date)`. Options:
`Open existing` / `Rename mine to "X (2)"` / `Cancel`.

### Live-update rule (incoming NOTIFY)

| Condition | Behavior |
|---|---|
| `updated_by_uuid == my UUID` | Ignore (own write echo). |
| Row not visible in UI | Refresh affected list/table on next render. |
| Row visible, no field focused | Re-query and update in place. Brief flash highlight. |
| Row visible, user currently typing in a field of that row | **Don't yank.** Yellow border on the conflicting cell + tooltip `"changed by Sarah just now — click to take her value"`. Other fields update normally. |
| Row deleted, user has it open | Non-modal toast: `"This row was deleted by Sarah. Click here to recreate."` |

`"Currently being edited"` detection: focused widget sets
`setProperty("dve_editing", true)` while it has focus and value differs
from loaded value; cleared on `editingFinished` or focus loss.

## Live updates & presence

### Subscription (Qt)

```cpp
QSqlDriver* drv = m_notifyDb.driver();
drv->subscribeToNotification("dataviewer_changes");
drv->subscribeToNotification("dataviewer_presence");
connect(drv, &QSqlDriver::notification, this, &NotificationListener::onNotify);
```

### Heartbeat lifecycle

- **Activate a resource** (click a file/session in nav): UPSERT to
  `presence`, `intent='viewing'`.
- **First dirty keystroke**: UPDATE `intent='editing'`.
- **Save successfully**: UPDATE `intent='viewing'`.
- **Every 10 s**: same UPSERT refreshes `last_heartbeat`.
- **Switch to a different resource**: DELETE old presence row + INSERT
  new one in a single transaction.
- **App exits cleanly**: DELETE all rows for this `user_uuid`.
- **App crashes**: rows go stale. Cleanup query runs on each client's
  presence read, plus a `pg_cron` job every minute:
  ```sql
  DELETE FROM presence WHERE last_heartbeat < now() - INTERVAL '30 seconds';
  ```

### Single-resource-per-user rule

Each user has at most one active presence row per `resource_type`.
Switching files = delete + insert. Keeps the table at
~`N_users × 3 resource_types` rows max.

### UI placement of dots

- **TPM file tree** (left nav): colored dot to the right of each filename.
  Hover tooltip: `"Bob Miller (editing), Sarah Chen (viewing)"`.
- **Sensory session list** (left nav, sensory mode): same.
- **Detailed sensory session list**: same.
- **Top-right of central editor** when a resource is loaded: avatar row
  showing all active users; clickable for a mini-modal with names +
  intents. Your own avatar is rendered with a thicker ring.

Colors from `presence.user_color`, picked by user at first launch.
Default palette: **12 hand-picked hex codes** (not algorithmic HSL),
chosen for distinguishability on both light and dark UI themes.
Exact values picked during implementation.

### Identity

A UUID generated at first launch (per install) alongside the display
name and color. UUID is the canonical key for audit and presence;
display name is for humans. Renaming yourself is allowed — UUID
stays. Stored in `QSettings`.

## Edge cases & common pitfalls (must-handle)

1. **Multi-row save atomicity** — entire save wrapped in transaction;
   ROLLBACK on any conflict, single dialog presented.
2. **Two connections per client** — `LISTEN` blocks a connection;
   separate connection for queries.
3. **Identity needs a UUID, not just a name** — two "Charlie"s would
   otherwise collide in audit/presence.
4. **Three conflict types, three dialogs** — version mismatch, row
   deleted, UNIQUE violation. App distinguishes via SQLSTATE +
   diagnostic query.
5. **Migration safety net** — single Postgres transaction, per-table
   row-count verification, original SQLite renamed (not deleted) to
   `<name>.pre-migration.sqlite` and kept for one release cycle.
6. **BLOB size ceiling** — 50 MB cap on image inserts; clear error
   message past that.
7. **Connection-storm jitter** — 0–5 s random delay on reconnect to
   prevent thundering-herd when NAS reboots.
8. **Pin Postgres version** — `postgres:16`, never `postgres:latest`.

## Migration & deployment

### Migration tool

New CLI mode on `DataViewer.exe`:

```
DataViewer.exe --migrate-from-sqlite=<path> --to-postgres="<conn_string>"
```

Headless, writes JSON report to `%TEMP%\dataviewer_migration.json`.

1. Open source SQLite read-only.
2. Refuse if `schema_meta.migrated_at` already set (unless `--force`,
   which drops all Postgres data and re-runs migration from scratch — use
   only after rolling back via the pre-migration SQLite escape hatch).
3. Per-table, in FK-dependency order: stream rows, transform types
   (TEXT→JSONB validated, BLOB→BYTEA), batched INSERT with
   `OVERRIDING SYSTEM VALUE` to **preserve original `id` values**
   (critical: layout JSON references session ids).
4. Per-table `COUNT(*)` match check against SQLite.
5. Bump each table's sequence to `MAX(id)+1`.
6. Entire migration in **one** Postgres transaction. Any failure →
   `ROLLBACK`, Postgres empty, non-zero exit.
7. On success: write `schema_meta`, `COMMIT`, rename Synology SQLite
   `<name>.sqlite` → `<name>.pre-migration.sqlite`.

### Malformed JSON handling

Any row with invalid `json_data` / `layout_json` aborts migration with
a precise error message pointing at the offending row id. No silent
data drop.

### Migration verification

New phase in `tests\deployment\Test-Deployment.ps1`: reads
`<name>.pre-migration.sqlite` read-only and live Postgres; compares
per-table row counts and field-by-field on a random sample of N=20
rows per table. Fails deployment if anything diverges.

### Rollback path

If Postgres proves unstable in the first weeks:

1. Reinstall pre-migration build from `Software Release\Archive\v1.3.x`.
2. Rename `<name>.pre-migration.sqlite` → `<name>.sqlite`.
3. Resume. No data lost.

### Docker compose

`deploy/postgres/docker-compose.yml`:

```yaml
services:
  dataviewer-db:
    image: postgres:16
    container_name: dataviewer-db
    restart: unless-stopped
    environment:
      POSTGRES_DB: dataviewer
      POSTGRES_USER: dataviewer_app
      POSTGRES_PASSWORD: ${DVE_DB_PASSWORD}   # from .env, not committed
    ports:
      - "5432:5432"
    volumes:
      - /volume1/docker/dataviewer-db/data:/var/lib/postgresql/data
      - ./init.sql:/docker-entrypoint-initdb.d/01-init.sql:ro
    healthcheck:
      test: ["CMD-SHELL", "pg_isready -U dataviewer_app -d dataviewer"]
      interval: 10s
```

`init.sql` runs on first container start: schema, indexes, trigger
functions/triggers, `CREATE EXTENSION pg_cron`, and the embedded cron
job for presence stale-cleanup. Idempotent (`IF NOT EXISTS`). `pg_cron`
requires a `shared_preload_libraries = 'pg_cron'` line in
`postgresql.conf` — added via a small `command:` override in the
compose file so the bind-mounted data dir doesn't need manual edits.

### Client-side libpq bundling

QPSQL needs `libpq.dll` + 4 transitive dependencies
(`libcrypto-3-x64.dll`, `libssl-3-x64.dll`, `libintl-9.dll`,
`libiconv-2.dll`). ~5 MB. Sourced from the official Postgres 16
Windows zip distribution. Added to `release\` and the Inno Setup
`[Files]` block in `installer.iss`.

### Connection configuration

`%PROGRAMDATA%\DataViewer\db.conf` (set by installer, editable by
sysadmin):

```ini
[postgres]
host                = dve-db.smoorecig.internal
port                = 5432
database            = dataviewer
user                = dataviewer_app
password_encrypted  = <AES-encrypted, machine+user-salted>
```

Password encrypted with `QAESEncryption` (same pattern translator uses
for the Anthropic API key). Installer prompts for the password once
during first install on each machine.

### First-time setup workflow

| Who | Where | One-time | Action |
|---|---|---|---|
| NAS admin | DSM Container Manager | ✔ | Drop `docker-compose.yml` + `init.sql` + `.env`; `docker compose up -d`; verify `pg_isready`. |
| NAS admin | DSM Firewall | ✔ | Allow TCP/5432 from user VLAN. |
| NAS admin | Hyper Backup | ✔ | Add `/volume1/docker/dataviewer-db/data` to backup job. |
| You | Work machine | ✔ | Install v2 build; run `DataViewer.exe --migrate-from-sqlite=... --to-postgres=...`; check exit 0. |
| You | Work machine | ✔ | Run `Test-Deployment.ps1` (with new migration phase). All green = go. |
| Each user | Workstation | per machine | Auto-updater installs v2; installer prompts for DB password once. |

### IT permissions required

- DSM admin on the NAS.
- Static IP or internal DNS name for the NAS (e.g.,
  `dve-db.smoorecig.internal`).
- Synology firewall rule allowing inbound TCP/5432 from user VLAN.
- Workstation VLAN ACL allowing outbound TCP/5432 to the NAS.
- Storage allocation on the NAS Postgres volume (~1–5 GB realistic).
- Backup policy includes the Postgres data volume.

**Not required:** external port forwarding, VPN access, AD/SSO
integration, additional ports beyond 5432.

## Testing strategy

### Infrastructure

A new test subdirectory `tests/tst_postgresconcurrency/` joins the
existing 11 in `tests.pro`. Integration tests look for an env var:

```
DVE_TEST_PG_CONN="host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"
```

- If set: tests run against that Postgres. Each test owns its own
  schema (`CREATE SCHEMA tst_<n>` / `DROP SCHEMA tst_<n> CASCADE`).
- If unset: tests `QSKIP` with a clear message. Pure-logic unit tests
  still run.

A helper `tests/start-test-postgres.ps1` spins up a throwaway
`postgres:16` container on port 5433, sets the env var, prints
teardown command. `Test-Deployment.ps1` does this automatically.

### Three tiers

| Tier | What | Where | When |
|---|---|---|---|
| **Unit** (pure logic) | conflict-type detection, schema-version comparison, NOTIFY payload parsing, color-palette assignment, AES password encryption roundtrip | Qt Test, no DB | every run |
| **Integration** (one client + DB) | optimistic concurrency, NOTIFY delivery, presence heartbeat, migration roundtrip, offline snapshot read | Qt Test + ephemeral Postgres | every run when env var set |
| **Multi-client / E2E** | two app instances editing same row, NOTIFY propagation, presence dots, conflict dialogs end-to-end | Qt Test spawns 2 processes | nightly / pre-release |

### Full test class inventory (~28 new test classes)

**Section: offline/online flow (5)**
- `tst_offlineFailover` — stop container; banner appears within 30 s;
  writes blocked; reads from snapshot succeed.
- `tst_offlineReconnect` — restart container; "Reconnected" toast;
  writes resume within 30 s.
- `tst_inflightEditDuringDisconnect` — edit retained in memory;
  reconnect; save attempt fires.
- `tst_noSnapshotGracefulFail` — modal appears; app doesn't crash.
- `tst_snapshotAtomicWrite` — crash mid-write; previous snapshot
  intact.

**Section: concurrency & conflicts (7)**
- `tst_optimisticConcurrencySingleRow`
- `tst_multiRowTransactionRollback`
- `tst_conflictDialogChoices` (all three dialog types, all paths)
- `tst_notifyIgnoreSelf`
- `tst_notifyDontYankInProgressEdit`
- `tst_uniqueViolationOnSensoryCreate`
- `tst_deletedWhileEditing`

**Section: live updates + presence (8)**
- `tst_notifyDelivery`
- `tst_notifyTwoChannelsIndependent`
- `tst_heartbeatRefresh`
- `tst_presenceSwitchResource`
- `tst_presenceCrashCleanup`
- `tst_intentTransition`
- `tst_presenceDotInTree`
- `tst_colorCollisionHandled`

**Section: migration & deployment (7)**
- `tst_migrationRoundTrip`
- `tst_migrationPreservesIds`
- `tst_migrationRejectsExistingDb`
- `tst_migrationMalformedJsonAborts`
- `tst_migrationRollbackOnError`
- `tst_migrationSequenceBump`
- `tst_libpqDllsPresent`

**E2E smoke (3)**
- `tst_e2eOpenEditSaveCycle`
- `tst_e2eTwoClientsCollab`
- `tst_e2eOfflineToOnlineRoundTrip`

### Existing tests that change

- `tst_databasemanager` — lock-specific tests retired (~20%);
  save/load tests retargeted at Postgres facade (~30% rewritten);
  the rest retained.
- `tests\deployment\Test-Deployment.ps1` — Phase 2 (`--self-test`)
  gains DB roundtrip checks; new Phase 4 = migration verification.
- `src/utils/SelfTest.cpp` — new `testPostgresConnection()`,
  `testNotifyRoundTrip()`, `testPresenceUpsert()` cases.

### Manual test checklist

Added to `tests\deployment\README.md`:

- Presence dots render with correct colors and tooltip copy.
- Conflict dialog displays both versions with differences highlighted;
  per-field selection works.
- Offline banner appears within 30 s of network drop and reads cleanly
  to non-technical users.
- "Reconnected" toast appears on reconnect; banner clears.
- Yellow border + tooltip appears on cell when remote change arrives
  mid-edit; user's text preserved.
- First-launch identity prompt is clear; color picker works.

## File-level change inventory

### New files

```
src/database/PostgresConnection.h/.cpp
src/database/NotificationListener.h/.cpp
src/database/PresenceManager.h/.cpp
src/database/OfflineSnapshot.h/.cpp
src/database/ConflictResolver.h/.cpp
src/database/IdentityManager.h/.cpp
src/database/MigrationTool.h/.cpp          (CLI mode in main)
src/database/ConfigLoader.h/.cpp           (db.conf parsing, AES decrypt)

deploy/postgres/docker-compose.yml
deploy/postgres/init.sql                   (schema + triggers + pg_cron)
deploy/postgres/.env.example
deploy/postgres/README.md                  (setup workflow)

tests/tst_postgresconcurrency/             (28 test classes)
tests/start-test-postgres.ps1
```

### Modified files

```
src/database/DatabaseManager.h/.cpp        (facade rewire; delete lock code)
src/main.cpp                                (--migrate-from-sqlite flag)
src/MainWindow.h/.cpp                       (offline banner, presence dots,
                                             conflict dialog wiring,
                                             identity-prompt first launch)
src/utils/SelfTest.cpp                      (3 new cases)
src/widgets/RibbonWidget.h/.cpp             (presence avatars top-right)
DataViewerEnterprise.pro                    (libpq DLL bundling notes)
build_installer.bat                         (copy libpq DLLs)
installer.iss                               ([Files] block adds 5 DLLs,
                                             db.conf prompt)
tests/tests.pro                             (add tst_postgresconcurrency
                                             subdir)
tests/deployment/Test-Deployment.ps1        (Phase 2 DB cases, Phase 4
                                             migration verification)
tests/deployment/README.md                  (manual test checklist)
CLAUDE.md                                   (replace SQLite-on-Synology
                                             section with Postgres setup)
```

### Deleted files

```
(none — lock-file code lives inside DatabaseManager.cpp and is removed
 in-place rather than via file deletion)
```

## Risks & known limitations

1. **NAS reboot impact**: every workstation reconnects simultaneously.
   Mitigated by 0–5 s jitter, but a 5–10 s "Reconnecting…" period is
   expected for the user pool.
2. **Long-running save transactions**: huge sensory sessions with
   embedded images may hold a transaction open for several seconds.
   Other clients' small saves wait. Mitigation: batch image inserts
   separately from metadata; cap individual image at 50 MB.
3. **Stale snapshot at home**: offline reads can be hours old. The
   banner explicitly shows the timestamp so the user knows.
4. **Conflict dialog under heavy collaboration**: if many users hammer
   the same row repeatedly, the dialog can show up frequently. This is
   the intended fail-loud behavior; the alternative (silent overwrite)
   was explicitly rejected.
5. **`pg_cron` extension dependency**: required for the presence
   stale-cleanup safety net. Bundled in `postgres:16` Docker image but
   needs `CREATE EXTENSION pg_cron` in `init.sql` and a `shared_
   preload_libraries` config line. Documented in
   `deploy/postgres/README.md`.
6. **Password encryption is machine-bound**: `db.conf`'s
   `password_encrypted` cannot be copied between machines (machine
   salt). Sysadmin reinstalls per machine. Acceptable for a small
   team; documented.
7. **Microsoft Information Protection on new source files**: this
   project's Windows account applies MIP/AIP sensitivity labels to
   newly-written `.cpp`/`.h` files, returning ciphertext starting with
   `%TSD-Header-###%` to compilers. CLAUDE.md documents the workaround
   (write source via Python's delete-and-rewrite pattern;
   `tools/decrypt_via_copy.py --apply` as reactive cleanup). This
   change-set introduces ~10 new `.h`/`.cpp` pairs, materially raising
   the chance of a MIP-encrypted file blocking the build. Mitigation
   baked into the implementation plan: every new source file is
   created via Python, and `decrypt_via_copy.py --apply` runs as a
   build pre-step until the first successful clean build verifies the
   files are stable.

## Open questions resolved at spec time

All four open questions raised in the design discussion are decided:

- **NOTIFY payload shape**: `{table, op, id, updated_by}`. Field naming
  may still evolve in implementation; the rest of the architecture
  does not depend on field names being final.
- **`pg_cron` deployment**: embedded inside `init.sql` (not a
  sidecar), with the required `shared_preload_libraries` set via
  compose `command:` override.
- **Color palette**: 12 hand-picked hex codes, not HSL-algorithmic.
- **Identity prompt**: modal dialog at first launch, dismissed once,
  with subsequent identity changes through Tools menu.
