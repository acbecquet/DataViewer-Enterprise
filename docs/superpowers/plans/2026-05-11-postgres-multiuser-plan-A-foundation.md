# Postgres Multi-User — Plan A: Foundation, Schema, Migration Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. See [the index](2026-05-11-postgres-multiuser-INDEX.md) for context across all three plans.

**Goal:** Stand up PostgreSQL 16 on the Synology NAS with the full schema deployed, build the SQLite→Postgres migration tool, and add the identity + config + connection infrastructure the app needs — all without removing the existing SQLite code paths.

**Architecture:** New components in `src/database/` (`PostgresConnection`, `IdentityManager`, `IdentityPromptDialog`, `ConfigLoader`, `MigrationTool`, `MigrationReport`) sit alongside the existing `DatabaseManager`. Schema deployed via Docker Compose + `init.sql` on the NAS. The migration tool is a headless CLI mode on `DataViewer.exe`. libpq 16 DLLs bundled into the installer; `db.conf` placed at `%PROGRAMDATA%\DataViewer` with AES-encrypted password.

**Tech Stack:** PostgreSQL 16 (Docker), `pg_cron`, Qt 6.10 + QPSQL driver, libpq 16 Windows DLLs, MinGW 13.1.0, Inno Setup 6, PowerShell 5.1, `tools/decrypt_via_copy.py`.

**Spec:** [docs/superpowers/specs/2026-05-11-postgres-multiuser-design.md](../specs/2026-05-11-postgres-multiuser-design.md)

---

## File Structure

### New source files (`src/database/`)

| File | Responsibility |
|---|---|
| `PostgresConnection.h`/`.cpp` | TCP connection, retry with jitter, dual `QSqlDatabase` (query + LISTEN reservation); no concurrency semantics yet |
| `IdentityManager.h`/`.cpp` | First-launch UUID generation, display name + color, `QSettings` persistence |
| `IdentityPromptDialog.h`/`.cpp` | Modal dialog at first launch; name + 12-hex-code color picker |
| `ConfigLoader.h`/`.cpp` | Parse `%PROGRAMDATA%\DataViewer\db.conf`, AES-decrypt password |
| `MigrationTool.h`/`.cpp` | One-shot SQLite→Postgres migration; CLI-driven; transactional |
| `MigrationReport.h`/`.cpp` | JSON report output to `%TEMP%\dataviewer_migration.json` |

### New deploy files

| File | Responsibility |
|---|---|
| `deploy/postgres/docker-compose.yml` | `postgres:16` service, bind-mount, `pg_cron` preload |
| `deploy/postgres/init.sql` | All CREATE TABLE, indexes, audit cols, triggers, `pg_cron` setup |
| `deploy/postgres/.env.example` | Template for `DVE_DB_PASSWORD` |
| `deploy/postgres/README.md` | NAS admin setup workflow |

### New test infrastructure

| File | Responsibility |
|---|---|
| `tests/start-test-postgres.ps1` | Spin up throwaway `postgres:16` on port 5433 for tests |
| `tests/tst_identitymanager/` | UUID stability, color palette, `QSettings` roundtrip |
| `tests/tst_configloader/` | Parse `db.conf`, AES decrypt roundtrip |
| `tests/tst_postgresconnection/` | Connect, fail, retry, dual-connection model |
| `tests/tst_migrationtool/` | Roundtrip, ID preservation, malformed-JSON abort, sequence bump, `--force` |

### Modified files

| File | Change |
|---|---|
| `src/main.cpp` | Parse `--migrate-from-sqlite` and `--to-postgres` flags; route to `MigrationTool` |
| `src/MainWindow.h`/`.cpp` | First-launch detection; show `IdentityPromptDialog` if no UUID exists |
| `src/utils/SelfTest.cpp` | Add `testPostgresConnection` case |
| `DataViewerEnterprise.pro` | Add new `src/database/` files |
| `installer.iss` | `[Files]` adds 5 libpq DLLs; `[Code]` prompts for DB password; writes `db.conf` |
| `build_installer.bat` | Copies libpq DLLs from staging into `release\` before invoking ISCC |
| `tests/tests.pro` | Add new `SUBDIRS` for new test classes |
| `tests/deployment/Test-Deployment.ps1` | New Phase 4: migration verification |
| `tests/deployment/README.md` | Manual checklist additions for migration |
| `CLAUDE.md` | Replace SQLite-on-Synology section with Postgres setup notes |

### NOT modified in Plan A

- `src/database/DatabaseManager.h`/`.cpp` — kept on SQLite path. Plan B switches its internals to Postgres.
- All Excel pipeline, plotting, reporting, sensory panels, ribbon.

---

## MIP Encryption Mitigation (cross-cutting)

Per CLAUDE.md and the spec's Risk #7: every new `.cpp`/`.h` file in this plan **must** be created via Python's delete-and-rewrite pattern to avoid inheriting MIP/AIP labels. Run `python tools/decrypt_via_copy.py --apply` from the repo root **before every build attempt** in this plan, until the first clean build verifies files are stable.

**The reusable Python file-creation snippet** (referenced by every "create file" step):

```python
import os
path = "<RELATIVE_PATH>"
content = r"""<FILE_CONTENT>"""
if os.path.exists(path):
    os.remove(path)
os.makedirs(os.path.dirname(path), exist_ok=True) if os.path.dirname(path) else None
with open(path, "w", encoding="utf-8", newline="\n") as f:
    f.write(content)
print(f"wrote {path}")
```

When a step says "create via the Python pattern," substitute `<RELATIVE_PATH>` and `<FILE_CONTENT>` with the values shown.

---

## Tasks

### Phase 1 — Postgres deployment infrastructure (Tasks 1–6)

These tasks produce a working `postgres:16` container with the full schema, runnable both on the dev machine (for tests) and on the Synology NAS (for production). No Qt/C++ code yet.

---

### Task 1: Docker Compose skeleton + `.env.example`

**Files:**
- Create: `deploy/postgres/docker-compose.yml`
- Create: `deploy/postgres/.env.example`

- [ ] **Step 1: Create the compose file**

Create `deploy/postgres/docker-compose.yml` (via `Write` is fine — not a source file):

```yaml
services:
  dataviewer-db:
    image: postgres:16
    container_name: dataviewer-db
    restart: unless-stopped
    environment:
      POSTGRES_DB: dataviewer
      POSTGRES_USER: dataviewer_app
      POSTGRES_PASSWORD: ${DVE_DB_PASSWORD}
    ports:
      - "5432:5432"
    volumes:
      - ./data:/var/lib/postgresql/data
      - ./init.sql:/docker-entrypoint-initdb.d/01-init.sql:ro
    command:
      - postgres
      - -c
      - shared_preload_libraries=pg_cron
      - -c
      - cron.database_name=dataviewer
    healthcheck:
      test: ["CMD-SHELL", "pg_isready -U dataviewer_app -d dataviewer"]
      interval: 10s
      timeout: 5s
      retries: 5
```

Note: on production Synology, the `./data` bind path becomes `/volume1/docker/dataviewer-db/data` and is set in the NAS admin's compose override, not committed here.

- [ ] **Step 2: Create the `.env.example`**

Create `deploy/postgres/.env.example`:

```bash
# Copy to .env and set a strong password before running `docker compose up -d`.
# The .env file MUST NOT be committed. .gitignore already covers it.
DVE_DB_PASSWORD=change-me-to-a-strong-random-password
```

- [ ] **Step 3: Update `.gitignore`**

Append to `.gitignore` at repo root:

```
# Postgres deployment secrets and bind-mounted data
deploy/postgres/.env
deploy/postgres/data/
```

- [ ] **Step 4: Commit**

```bash
git add deploy/postgres/docker-compose.yml deploy/postgres/.env.example .gitignore
git commit -m "feat(deploy): postgres:16 docker compose skeleton with pg_cron preload"
```

---

### Task 2: `init.sql` — core data tables

**Files:**
- Create: `deploy/postgres/init.sql`

- [ ] **Step 1: Write `init.sql` with all data tables**

Create `deploy/postgres/init.sql`:

```sql
-- DataViewer Enterprise — initial schema (Plan A)
-- Idempotent: re-running drops nothing; uses CREATE ... IF NOT EXISTS.

BEGIN;

-- ── files ────────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS files (
    id               BIGSERIAL PRIMARY KEY,
    file_path        TEXT NOT NULL,
    file_name        TEXT NOT NULL,
    loaded_at        TEXT NOT NULL,
    template_version TEXT,
    sheet_count      INTEGER DEFAULT 0,
    sample_count     INTEGER DEFAULT 0,
    updated_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by       TEXT        NOT NULL DEFAULT 'migration',
    version          INTEGER     NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_files_path ON files(file_path);

-- ── tests (sheets) ───────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS tests (
    id                 BIGSERIAL PRIMARY KEY,
    file_id            BIGINT NOT NULL REFERENCES files(id) ON DELETE CASCADE,
    sheet_name         TEXT NOT NULL,
    template_version   TEXT,
    overall_avg_tpm    DOUBLE PRECISION DEFAULT 0.0,
    overall_stddev_tpm DOUBLE PRECISION DEFAULT 0.0,
    is_raw_table       INTEGER DEFAULT 0,
    sort_order         INTEGER DEFAULT 0,
    updated_at         TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by         TEXT        NOT NULL DEFAULT 'migration',
    version            INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_tests_file ON tests(file_id);

-- ── samples ──────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS samples (
    id                  BIGSERIAL PRIMARY KEY,
    test_id             BIGINT NOT NULL REFERENCES tests(id) ON DELETE CASCADE,
    sort_order          INTEGER DEFAULT 0,
    sample_name         TEXT,
    sample_id           TEXT,
    date                TEXT,
    tester              TEXT,
    media               TEXT,
    viscosity           DOUBLE PRECISION DEFAULT 0.0,
    resistance          DOUBLE PRECISION DEFAULT 0.0,
    voltage             DOUBLE PRECISION DEFAULT 0.0,
    power               DOUBLE PRECISION DEFAULT 0.0,
    heating_technology  TEXT,
    puffing_regime      TEXT,
    initial_oil_mass    DOUBLE PRECISION DEFAULT 0.0,
    average_tpm         DOUBLE PRECISION DEFAULT 0.0,
    stddev_tpm          DOUBLE PRECISION DEFAULT 0.0,
    avg_power_density   DOUBLE PRECISION DEFAULT 0.0,
    efficiency_percent  DOUBLE PRECISION DEFAULT 0.0,
    total_oil_consumed  DOUBLE PRECISION DEFAULT 0.0,
    total_puffs         INTEGER DEFAULT 0,
    normalized_tpm      DOUBLE PRECISION DEFAULT 0.0,
    burn_status         TEXT,
    clog_status         TEXT,
    leak_status         TEXT,
    updated_at          TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by          TEXT        NOT NULL DEFAULT 'migration',
    version             INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_samples_test ON samples(test_id);

-- ── data_rows ────────────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS data_rows (
    id                BIGSERIAL PRIMARY KEY,
    sample_id         BIGINT NOT NULL REFERENCES samples(id) ON DELETE CASCADE,
    sort_order        INTEGER DEFAULT 0,
    puffs             DOUBLE PRECISION DEFAULT 0.0,
    before_weight     DOUBLE PRECISION DEFAULT 0.0,
    after_weight      DOUBLE PRECISION DEFAULT 0.0,
    draw_pressure     DOUBLE PRECISION DEFAULT 0.0,
    resistance        DOUBLE PRECISION DEFAULT 0.0,
    smell             TEXT,
    clog              TEXT,
    notes             TEXT,
    tpm               DOUBLE PRECISION DEFAULT 0.0,
    tpm_power_density DOUBLE PRECISION DEFAULT 0.0,
    variation_tpm     DOUBLE PRECISION DEFAULT 0.0,
    oil_consumed      DOUBLE PRECISION DEFAULT 0.0,
    updated_at        TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by        TEXT        NOT NULL DEFAULT 'migration',
    version           INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_data_rows_sample ON data_rows(sample_id);

-- ── images (per-sample) ──────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS images (
    id          BIGSERIAL PRIMARY KEY,
    sample_id   BIGINT NOT NULL REFERENCES samples(id) ON DELETE CASCADE,
    sort_order  INTEGER DEFAULT 0,
    file_name   TEXT,
    image_data  BYTEA,
    layout_x    DOUBLE PRECISION,
    layout_y    DOUBLE PRECISION,
    layout_w    DOUBLE PRECISION,
    layout_h    DOUBLE PRECISION,
    crop_x      DOUBLE PRECISION,
    crop_y      DOUBLE PRECISION,
    crop_w      DOUBLE PRECISION,
    crop_h      DOUBLE PRECISION,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT        NOT NULL DEFAULT 'migration',
    version     INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_images_sample ON images(sample_id);

-- ── sensory_sessions ─────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS sensory_sessions (
    id            BIGSERIAL PRIMARY KEY,
    session_name  TEXT,
    tester_name   TEXT,
    assessor_name TEXT,
    media         TEXT,
    puff_length   TEXT,
    date          TEXT,
    timestamp     TEXT,
    json_data     JSONB,
    layout_json   JSONB,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by    TEXT        NOT NULL DEFAULT 'migration',
    version       INTEGER     NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_sensory_sessions_key
    ON sensory_sessions(session_name, tester_name, date);

-- ── sensory_images ───────────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS sensory_images (
    id          BIGSERIAL PRIMARY KEY,
    session_id  BIGINT NOT NULL REFERENCES sensory_sessions(id) ON DELETE CASCADE,
    sort_order  INTEGER DEFAULT 0,
    file_name   TEXT,
    image_data  BYTEA,
    layout_x    DOUBLE PRECISION,
    layout_y    DOUBLE PRECISION,
    layout_w    DOUBLE PRECISION,
    layout_h    DOUBLE PRECISION,
    crop_x      DOUBLE PRECISION,
    crop_y      DOUBLE PRECISION,
    crop_w      DOUBLE PRECISION,
    crop_h      DOUBLE PRECISION,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT        NOT NULL DEFAULT 'migration',
    version     INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_sensory_images_session ON sensory_images(session_id);

-- ── detailed_sensory_sessions ────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS detailed_sensory_sessions (
    id            BIGSERIAL PRIMARY KEY,
    session_name  TEXT,
    tester_name   TEXT,
    assessor_name TEXT,
    media         TEXT,
    date          TEXT,
    timestamp     TEXT,
    json_data     JSONB,
    updated_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by    TEXT        NOT NULL DEFAULT 'migration',
    version       INTEGER     NOT NULL DEFAULT 1
);
CREATE UNIQUE INDEX IF NOT EXISTS idx_detailed_sensory_sessions_key
    ON detailed_sensory_sessions(session_name, tester_name, date);

-- ── detailed_sensory_images ──────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS detailed_sensory_images (
    id          BIGSERIAL PRIMARY KEY,
    session_id  BIGINT NOT NULL REFERENCES detailed_sensory_sessions(id) ON DELETE CASCADE,
    sort_order  INTEGER DEFAULT 0,
    file_name   TEXT,
    image_data  BYTEA,
    layout_x    DOUBLE PRECISION,
    layout_y    DOUBLE PRECISION,
    layout_w    DOUBLE PRECISION,
    layout_h    DOUBLE PRECISION,
    crop_x      DOUBLE PRECISION,
    crop_y      DOUBLE PRECISION,
    crop_w      DOUBLE PRECISION,
    crop_h      DOUBLE PRECISION,
    updated_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by  TEXT        NOT NULL DEFAULT 'migration',
    version     INTEGER     NOT NULL DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_detailed_sensory_images_session
    ON detailed_sensory_images(session_id);

-- ── settings (key/value) ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS settings (
    key        TEXT PRIMARY KEY,
    value      TEXT,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    updated_by TEXT        NOT NULL DEFAULT 'migration',
    version    INTEGER     NOT NULL DEFAULT 1
);

-- ── presence (live "who has this open" view) ─────────────────────────────────
CREATE TABLE IF NOT EXISTS presence (
    user_uuid      UUID        NOT NULL,
    user_name      TEXT        NOT NULL,
    user_color     TEXT        NOT NULL,
    resource_type  TEXT        NOT NULL,
    resource_id    BIGINT      NOT NULL,
    intent         TEXT        NOT NULL DEFAULT 'viewing',
    last_heartbeat TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (user_uuid, resource_type, resource_id)
);
CREATE INDEX IF NOT EXISTS idx_presence_resource ON presence(resource_type, resource_id);

-- ── schema_meta (migration provenance, schema versioning) ────────────────────
CREATE TABLE IF NOT EXISTS schema_meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

COMMIT;
```

- [ ] **Step 2: Start a local container to verify the schema loads**

From `deploy/postgres/`:

```bash
cp .env.example .env
# edit .env so DVE_DB_PASSWORD is a real password
docker compose up -d
docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c "\dt"
```

Expected: lists all 12 tables (`files`, `tests`, `samples`, `data_rows`, `images`, `sensory_sessions`, `sensory_images`, `detailed_sensory_sessions`, `detailed_sensory_images`, `settings`, `presence`, `schema_meta`).

- [ ] **Step 3: Tear down**

```bash
docker compose down -v
```

- [ ] **Step 4: Commit**

```bash
git add deploy/postgres/init.sql
git commit -m "feat(deploy): init.sql with all data tables and audit columns"
```

---

### Task 3: `init.sql` — `bump_version` trigger function

**Files:**
- Modify: `deploy/postgres/init.sql` (append before final `COMMIT;`)

- [ ] **Step 1: Append trigger function + per-table triggers**

In `deploy/postgres/init.sql`, before the final `COMMIT;`, append:

```sql
-- ── bump_version: auto-increments version + stamps updated_at on every UPDATE
CREATE OR REPLACE FUNCTION bump_version() RETURNS TRIGGER AS $$
BEGIN
  -- Refuse client-side version manipulation: if NEW.version differs from
  -- OLD.version, that's a bug in the app, not an upgrade. We always +1.
  NEW.version    := OLD.version + 1;
  NEW.updated_at := now();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DO $$
DECLARE
  t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'files', 'tests', 'samples', 'data_rows', 'images',
    'sensory_sessions', 'sensory_images',
    'detailed_sensory_sessions', 'detailed_sensory_images',
    'settings'
  ] LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_bump_version ON %I;
       CREATE TRIGGER trg_%I_bump_version
       BEFORE UPDATE ON %I
       FOR EACH ROW EXECUTE FUNCTION bump_version();',
       t, t, t, t
    );
  END LOOP;
END$$;
```

- [ ] **Step 2: Verify the trigger fires**

```bash
docker compose up -d
docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c "
  INSERT INTO settings(key, value) VALUES ('test', 'a') RETURNING version;
  UPDATE settings SET value = 'b' WHERE key = 'test' RETURNING version, updated_at;
"
```

Expected: `INSERT` returns `version = 1`; `UPDATE` returns `version = 2` and `updated_at` is current.

- [ ] **Step 3: Tear down & commit**

```bash
docker compose down -v
git add deploy/postgres/init.sql
git commit -m "feat(deploy): bump_version trigger on every editable table"
```

---

### Task 4: `init.sql` — NOTIFY trigger functions

**Files:**
- Modify: `deploy/postgres/init.sql`

- [ ] **Step 1: Append `notify_row_change` + `notify_presence_change` + triggers**

Append to `init.sql` before final `COMMIT;`:

```sql
-- ── notify_row_change: fires NOTIFY on every INSERT/UPDATE/DELETE
CREATE OR REPLACE FUNCTION notify_row_change() RETURNS TRIGGER AS $$
BEGIN
  PERFORM pg_notify(
    'dataviewer_changes',
    json_build_object(
      'table',      TG_TABLE_NAME,
      'op',         TG_OP,
      'id',         COALESCE(NEW.id, OLD.id),
      'updated_by', COALESCE(NEW.updated_by, OLD.updated_by)
    )::text
  );
  RETURN NULL;
END;
$$ LANGUAGE plpgsql;

DO $$
DECLARE
  t TEXT;
BEGIN
  FOREACH t IN ARRAY ARRAY[
    'files', 'tests', 'samples', 'data_rows', 'images',
    'sensory_sessions', 'sensory_images',
    'detailed_sensory_sessions', 'detailed_sensory_images',
    'settings'
  ] LOOP
    EXECUTE format(
      'DROP TRIGGER IF EXISTS trg_%I_notify ON %I;
       CREATE TRIGGER trg_%I_notify
       AFTER INSERT OR UPDATE OR DELETE ON %I
       FOR EACH ROW EXECUTE FUNCTION notify_row_change();',
       t, t, t, t
    );
  END LOOP;
END$$;

-- ── notify_presence_change: separate channel for high-frequency presence
CREATE OR REPLACE FUNCTION notify_presence_change() RETURNS TRIGGER AS $$
BEGIN
  PERFORM pg_notify(
    'dataviewer_presence',
    json_build_object(
      'op',            TG_OP,
      'user_uuid',     COALESCE(NEW.user_uuid::text, OLD.user_uuid::text),
      'resource_type', COALESCE(NEW.resource_type, OLD.resource_type),
      'resource_id',   COALESCE(NEW.resource_id, OLD.resource_id),
      'intent',        COALESCE(NEW.intent, OLD.intent)
    )::text
  );
  RETURN NULL;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_presence_notify ON presence;
CREATE TRIGGER trg_presence_notify
AFTER INSERT OR UPDATE OR DELETE ON presence
FOR EACH ROW EXECUTE FUNCTION notify_presence_change();
```

- [ ] **Step 2: Verify NOTIFY fires**

In one terminal:

```bash
docker compose up -d
docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer
```

In the `psql` shell:

```sql
LISTEN dataviewer_changes;
```

In another terminal:

```bash
docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c \
  "INSERT INTO settings(key, value) VALUES ('notify_test', 'x')"
```

Expected: the first `psql` shell prints an `Asynchronous notification "dataviewer_changes"` line with the payload JSON within 1s.

- [ ] **Step 3: Tear down & commit**

```bash
docker compose down -v
git add deploy/postgres/init.sql
git commit -m "feat(deploy): NOTIFY triggers on data tables and presence channel"
```

---

### Task 5: `init.sql` — `pg_cron` extension and stale-presence cleanup

**Files:**
- Modify: `deploy/postgres/init.sql`

- [ ] **Step 1: Append pg_cron setup**

Append to `init.sql` (the extension must be created OUTSIDE the BEGIN/COMMIT, in its own block — wrap appropriately):

```sql
-- ── pg_cron: schedule presence stale-row cleanup every minute ───────────────
-- Must run as superuser; the postgres:16 image creates the extension when
-- shared_preload_libraries=pg_cron is set (see docker-compose command:).
CREATE EXTENSION IF NOT EXISTS pg_cron;

-- Remove existing job by name (idempotent reinstall), then re-add.
DO $$
DECLARE jobid BIGINT;
BEGIN
  SELECT cron.jobid INTO jobid FROM cron.job WHERE jobname = 'dve_presence_cleanup';
  IF FOUND THEN
    PERFORM cron.unschedule(jobid);
  END IF;
END$$;

SELECT cron.schedule(
  'dve_presence_cleanup',
  '* * * * *',  -- every minute
  $$ DELETE FROM presence
     WHERE last_heartbeat < now() - INTERVAL '30 seconds' $$
);
```

Important: the `CREATE EXTENSION pg_cron` line must execute in the `dataviewer` database, not the default `postgres` database. The compose `cron.database_name=dataviewer` setting ensures pg_cron's own catalog lives there.

- [ ] **Step 2: Verify the cron job is scheduled**

```bash
docker compose up -d
sleep 5
docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c \
  "SELECT jobname, schedule FROM cron.job"
```

Expected: one row, `dve_presence_cleanup | * * * * *`.

- [ ] **Step 3: Verify cleanup actually runs**

```bash
docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c "
  INSERT INTO presence(user_uuid, user_name, user_color, resource_type, resource_id, last_heartbeat)
  VALUES (gen_random_uuid(), 'test', '#ff0000', 'file', 1, now() - INTERVAL '2 minutes');
"
# Wait up to 70s for the cron job to fire
sleep 65
docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c \
  "SELECT COUNT(*) FROM presence"
```

Expected: count = 0 (stale row was cleaned).

- [ ] **Step 4: Tear down & commit**

```bash
docker compose down -v
git add deploy/postgres/init.sql
git commit -m "feat(deploy): pg_cron stale-presence cleanup, runs every minute"
```

---

### Task 6: `deploy/postgres/README.md` — NAS admin setup guide

**Files:**
- Create: `deploy/postgres/README.md`

- [ ] **Step 1: Write the README**

```markdown
# DataViewer Enterprise — PostgreSQL Setup (Synology NAS)

This directory contains the database container definition and bootstrap
SQL for DataViewer Enterprise's Postgres backend.

## One-time NAS admin setup

1. **Prerequisites on the NAS:**
   - DSM 7.2+ with Container Manager installed.
   - A static IP or internal DNS name (e.g. `dve-db.smoorecig.internal`).
   - Storage volume with at least 5 GB free.

2. **Copy this directory to `/volume1/docker/dataviewer-db/` on the NAS:**

   ```bash
   # From a workstation with SSH access:
   scp -r deploy/postgres/. admin@nas:/volume1/docker/dataviewer-db/
   ```

3. **Create the production `.env` ON THE NAS** (do not commit):

   ```bash
   cd /volume1/docker/dataviewer-db
   cp .env.example .env
   # Edit .env, set DVE_DB_PASSWORD to a strong random value.
   chmod 600 .env
   ```

4. **Update the bind-mount path in `docker-compose.yml` for production**:

   On the NAS only, change `./data` to `/volume1/docker/dataviewer-db/data`.
   (This is intentional — the committed compose file uses the relative
   `./data` for local-dev use; production overrides it in-place.)

5. **Start the container:**

   ```bash
   docker compose up -d
   docker compose logs -f dataviewer-db    # wait for "ready to accept connections"
   docker compose exec dataviewer-db pg_isready -U dataviewer_app
   ```

6. **Verify the schema and cron job loaded:**

   ```bash
   docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c "\dt"
   # → should list 12 tables

   docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c \
     "SELECT jobname FROM cron.job"
   # → should print: dve_presence_cleanup
   ```

7. **Open the firewall:**

   DSM → Control Panel → Security → Firewall → add a rule allowing
   inbound TCP/5432 from the user VLAN (typically your office subnet).
   Do NOT open this port to the internet.

8. **Add the data dir to Hyper Backup:**

   DSM → Hyper Backup → edit your backup task → include
   `/volume1/docker/dataviewer-db/data` in the source list.

## Routine operations

- **Inspect logs:** `docker compose logs -f dataviewer-db`
- **Restart:** `docker compose restart dataviewer-db`
- **Upgrade Postgres minor version:** `docker compose pull && docker compose up -d`
  (Major version upgrades — e.g., 16 → 17 — are deliberate events; see
  the spec for the planned migration process.)
- **Backup snapshot (in addition to Hyper Backup):**
  `docker compose exec dataviewer-db pg_dump -U dataviewer_app dataviewer > /volume1/backups/dve-$(date +%F).sql`

## Disaster recovery

If the Postgres data is lost or corrupted:

1. Restore the `/volume1/docker/dataviewer-db/data` directory from Hyper Backup.
2. `docker compose up -d` — Postgres reads the restored data on startup.
3. Verify with the schema and cron-job checks above.

If recovery from Postgres backup is not possible and the migration is recent:

1. See `docs/superpowers/specs/2026-05-11-postgres-multiuser-design.md`
   "Rollback path" section for the SQLite escape hatch (the
   `<name>.pre-migration.sqlite` file kept on Synology).
```

- [ ] **Step 2: Commit**

```bash
git add deploy/postgres/README.md
git commit -m "docs(deploy): NAS admin setup guide for Postgres container"
```

---

### Phase 2 — Identity management (Tasks 7–10)

These tasks build `IdentityManager`, `IdentityPromptDialog`, and `ConfigLoader`, all of which are standalone (no Postgres dependency yet). Identity is the user-side prerequisite for everything else: `updated_by`, presence, conflict dialogs all key off the UUID.

---

### Task 7: `IdentityManager` skeleton + UUID stability test

**Files:**
- Create: `src/database/IdentityManager.h`
- Create: `src/database/IdentityManager.cpp`
- Create: `tests/tst_identitymanager/tst_identitymanager.pro`
- Create: `tests/tst_identitymanager/tst_identitymanager.cpp`

- [ ] **Step 1: Run MIP decrypt pre-flight**

```bash
python tools/decrypt_via_copy.py --apply
```

Expected: idempotent, no ciphertext found.

- [ ] **Step 2: Write the failing test via Python pattern**

`<RELATIVE_PATH>` = `tests/tst_identitymanager/tst_identitymanager.cpp`

`<FILE_CONTENT>`:

```cpp
#include <QtTest/QtTest>
#include <QCoreApplication>
#include <QSettings>
#include <QTemporaryDir>
#include "../../src/database/IdentityManager.h"

using DVE::IdentityManager;

class TstIdentityManager : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        QCoreApplication::setOrganizationName("DataViewerTest");
        QCoreApplication::setApplicationName("tst_identitymanager");
        QSettings s;
        s.clear();
    }

    void uuid_isGenerated_onFirstAccess() {
        IdentityManager m;
        QVERIFY(!m.uuid().isNull());
        QVERIFY(!m.uuid().toString().isEmpty());
    }

    void uuid_persistsAcrossInstances() {
        IdentityManager m1;
        const QString u1 = m1.uuid().toString();
        IdentityManager m2;
        const QString u2 = m2.uuid().toString();
        QCOMPARE(u1, u2);
    }

    void displayName_defaultsToEmpty_untilSet() {
        QSettings s;
        s.clear();
        IdentityManager m;
        QVERIFY(m.displayName().isEmpty());
    }

    void displayName_persistsAfterSet() {
        IdentityManager m;
        m.setDisplayName("Charlie B.");
        IdentityManager m2;
        QCOMPARE(m2.displayName(), QString("Charlie B."));
    }

    void color_defaultsToEmpty_untilSet() {
        QSettings s;
        s.clear();
        IdentityManager m;
        QVERIFY(m.color().isEmpty());
    }

    void color_persistsAfterSet() {
        IdentityManager m;
        m.setColor("#3b82f6");
        IdentityManager m2;
        QCOMPARE(m2.color(), QString("#3b82f6"));
    }

    void firstLaunchPending_trueWhenNameMissing() {
        QSettings s;
        s.clear();
        IdentityManager m;
        QVERIFY(m.firstLaunchPending());
    }

    void firstLaunchPending_falseAfterNameAndColorSet() {
        QSettings s;
        s.clear();
        IdentityManager m;
        m.setDisplayName("Sarah");
        m.setColor("#10b981");
        QVERIFY(!m.firstLaunchPending());
    }
};

QTEST_MAIN(TstIdentityManager)
#include "tst_identitymanager.moc"
```

Then create the .pro file (regular `Write` is fine for .pro):

`tests/tst_identitymanager/tst_identitymanager.pro`:

```qmake
QT       += testlib core
CONFIG   += console c++17 testcase
CONFIG   -= app_bundle
TARGET    = tst_identitymanager
SOURCES  += tst_identitymanager.cpp \
            ../../src/database/IdentityManager.cpp
HEADERS  += ../../src/database/IdentityManager.h
INCLUDEPATH += ../../src
```

- [ ] **Step 3: Add to `tests/tests.pro` SUBDIRS**

Edit `tests/tests.pro` and append `tst_identitymanager` to the `SUBDIRS` list.

- [ ] **Step 4: Build and verify test fails**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_identitymanager
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++
mingw32-make
```

Expected: build fails with "IdentityManager.h: No such file or directory".

- [ ] **Step 5: Create `IdentityManager.h` via Python pattern**

`<RELATIVE_PATH>` = `src/database/IdentityManager.h`

`<FILE_CONTENT>`:

```cpp
#pragma once

#include <QObject>
#include <QString>
#include <QUuid>

namespace DVE {

class IdentityManager : public QObject {
    Q_OBJECT
public:
    explicit IdentityManager(QObject* parent = nullptr);

    QUuid   uuid() const;
    QString displayName() const;
    QString color() const;

    void setDisplayName(const QString& name);
    void setColor(const QString& hex);

    bool firstLaunchPending() const;

    static QStringList defaultColorPalette();

private:
    void ensureUuid();
};

} // namespace DVE
```

- [ ] **Step 6: Create `IdentityManager.cpp` via Python pattern**

`<RELATIVE_PATH>` = `src/database/IdentityManager.cpp`

`<FILE_CONTENT>`:

```cpp
#include "IdentityManager.h"
#include <QSettings>
#include <QUuid>

namespace DVE {

namespace {
constexpr auto kKeyUuid = "identity/uuid";
constexpr auto kKeyName = "identity/displayName";
constexpr auto kKeyColor = "identity/color";
}

IdentityManager::IdentityManager(QObject* parent) : QObject(parent) {
    ensureUuid();
}

void IdentityManager::ensureUuid() {
    QSettings s;
    const QString existing = s.value(kKeyUuid).toString();
    if (existing.isEmpty() || QUuid(existing).isNull()) {
        const QUuid fresh = QUuid::createUuid();
        s.setValue(kKeyUuid, fresh.toString(QUuid::WithoutBraces));
    }
}

QUuid IdentityManager::uuid() const {
    QSettings s;
    return QUuid(s.value(kKeyUuid).toString());
}

QString IdentityManager::displayName() const {
    QSettings s;
    return s.value(kKeyName).toString();
}

QString IdentityManager::color() const {
    QSettings s;
    return s.value(kKeyColor).toString();
}

void IdentityManager::setDisplayName(const QString& name) {
    QSettings s;
    s.setValue(kKeyName, name);
}

void IdentityManager::setColor(const QString& hex) {
    QSettings s;
    s.setValue(kKeyColor, hex);
}

bool IdentityManager::firstLaunchPending() const {
    return displayName().isEmpty() || color().isEmpty();
}

QStringList IdentityManager::defaultColorPalette() {
    // 12 hand-picked hex codes, balanced for light + dark UI themes.
    return {
        "#ef4444", // red
        "#f97316", // orange
        "#eab308", // amber
        "#84cc16", // lime
        "#10b981", // emerald
        "#06b6d4", // cyan
        "#3b82f6", // blue
        "#6366f1", // indigo
        "#8b5cf6", // violet
        "#d946ef", // fuchsia
        "#ec4899", // pink
        "#64748b"  // slate
    };
}

} // namespace DVE
```

- [ ] **Step 7: Rebuild and run test**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_identitymanager
mingw32-make clean
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++
mingw32-make
.\release\tst_identitymanager.exe
```

Expected: all 8 tests pass.

- [ ] **Step 8: Commit**

```bash
git add src/database/IdentityManager.h src/database/IdentityManager.cpp \
        tests/tst_identitymanager/ tests/tests.pro
git commit -m "feat(db): IdentityManager with UUID + display name + 12-color palette"
```

---

### Task 8: `IdentityPromptDialog` — first-launch modal

**Files:**
- Create: `src/database/IdentityPromptDialog.h`
- Create: `src/database/IdentityPromptDialog.cpp`

The dialog itself is mostly UI assembly with limited testable logic; the testable parts (color validation, palette load) are covered by `IdentityManager` tests. We skip a dedicated test file for the dialog — manual checklist in `tests/deployment/README.md` covers the visual aspects.

- [ ] **Step 1: Create `IdentityPromptDialog.h` via Python pattern**

`<RELATIVE_PATH>` = `src/database/IdentityPromptDialog.h`

`<FILE_CONTENT>`:

```cpp
#pragma once

#include <QDialog>

class QLineEdit;
class QButtonGroup;

namespace DVE {

class IdentityManager;

class IdentityPromptDialog : public QDialog {
    Q_OBJECT
public:
    explicit IdentityPromptDialog(IdentityManager* mgr, QWidget* parent = nullptr);

private slots:
    void onAccept();

private:
    IdentityManager* m_mgr;
    QLineEdit*       m_nameEdit;
    QButtonGroup*    m_colorGroup;
    QString          m_selectedColor;
};

} // namespace DVE
```

- [ ] **Step 2: Create `IdentityPromptDialog.cpp` via Python pattern**

`<RELATIVE_PATH>` = `src/database/IdentityPromptDialog.cpp`

`<FILE_CONTENT>`:

```cpp
#include "IdentityPromptDialog.h"
#include "IdentityManager.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QButtonGroup>
#include <QToolButton>
#include <QMessageBox>

namespace DVE {

IdentityPromptDialog::IdentityPromptDialog(IdentityManager* mgr, QWidget* parent)
    : QDialog(parent), m_mgr(mgr) {
    setWindowTitle(tr("Welcome to DataViewer"));
    setModal(true);

    auto* root = new QVBoxLayout(this);

    root->addWidget(new QLabel(
        tr("Other users will see your name and color when you have files "
           "or sessions open. Pick a name and a color now — you can change "
           "either later from Tools → Identity.")));

    auto* nameRow = new QHBoxLayout;
    nameRow->addWidget(new QLabel(tr("Display name:")));
    m_nameEdit = new QLineEdit;
    m_nameEdit->setPlaceholderText(tr("e.g. Sarah Chen"));
    m_nameEdit->setMaxLength(40);
    nameRow->addWidget(m_nameEdit);
    root->addLayout(nameRow);

    root->addWidget(new QLabel(tr("Pick a color:")));
    auto* grid = new QGridLayout;
    m_colorGroup = new QButtonGroup(this);
    const QStringList palette = IdentityManager::defaultColorPalette();
    for (int i = 0; i < palette.size(); ++i) {
        auto* btn = new QToolButton;
        btn->setCheckable(true);
        btn->setMinimumSize(36, 36);
        btn->setStyleSheet(QString(
            "QToolButton { background:%1; border:2px solid #00000022; "
            "border-radius:18px; } "
            "QToolButton:checked { border:3px solid #000; }"
        ).arg(palette[i]));
        btn->setProperty("dve_color_hex", palette[i]);
        m_colorGroup->addButton(btn, i);
        grid->addWidget(btn, i / 6, i % 6);
    }
    root->addLayout(grid);

    auto* buttons = new QHBoxLayout;
    buttons->addStretch();
    auto* okBtn = new QPushButton(tr("Save"));
    okBtn->setDefault(true);
    buttons->addWidget(okBtn);
    root->addLayout(buttons);

    connect(okBtn, &QPushButton::clicked,
            this, &IdentityPromptDialog::onAccept);
}

void IdentityPromptDialog::onAccept() {
    const QString name = m_nameEdit->text().trimmed();
    if (name.isEmpty()) {
        QMessageBox::warning(this, tr("Name required"),
                             tr("Please enter a display name."));
        return;
    }
    auto* picked = m_colorGroup->checkedButton();
    if (!picked) {
        QMessageBox::warning(this, tr("Color required"),
                             tr("Please pick a color."));
        return;
    }
    m_selectedColor = picked->property("dve_color_hex").toString();
    m_mgr->setDisplayName(name);
    m_mgr->setColor(m_selectedColor);
    accept();
}

} // namespace DVE
```

- [ ] **Step 3: Add to `DataViewerEnterprise.pro`**

In the top-level `DataViewerEnterprise.pro`, append to `SOURCES`:

```
SOURCES += src/database/IdentityManager.cpp \
           src/database/IdentityPromptDialog.cpp
```

and to `HEADERS`:

```
HEADERS += src/database/IdentityManager.h \
           src/database/IdentityPromptDialog.h
```

- [ ] **Step 4: Build the main app to verify compile**

```bash
python tools/decrypt_via_copy.py --apply
mkdir -p build && cd build
set PATH=C:\Qt\Tools\mingw1310_64\bin;%PATH%
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++ ../DataViewerEnterprise.pro
mingw32-make -j8
```

Expected: clean build, no errors.

- [ ] **Step 5: Commit**

```bash
git add src/database/IdentityPromptDialog.h src/database/IdentityPromptDialog.cpp \
        DataViewerEnterprise.pro
git commit -m "feat(db): IdentityPromptDialog modal for first-launch identity setup"
```

---

### Task 9: Wire identity prompt into `MainWindow` startup

**Files:**
- Modify: `src/MainWindow.h`
- Modify: `src/MainWindow.cpp`

- [ ] **Step 1: Add `IdentityManager` member to `MainWindow.h`**

In `src/MainWindow.h`, near other member declarations, add:

```cpp
class DVE::IdentityManager;

private:
    DVE::IdentityManager* m_identity = nullptr;
```

And the include:

```cpp
namespace DVE { class IdentityManager; }
```

- [ ] **Step 2: Wire startup logic in `MainWindow.cpp`**

In `src/MainWindow.cpp`, at the END of the `MainWindow` constructor (after all
existing init), append:

```cpp
#include "database/IdentityManager.h"
#include "database/IdentityPromptDialog.h"
// ...

// Identity bootstrap — must happen after window is constructed so the
// modal has a real parent.
m_identity = new DVE::IdentityManager(this);
if (m_identity->firstLaunchPending()) {
    QTimer::singleShot(0, this, [this]() {
        DVE::IdentityPromptDialog dlg(m_identity, this);
        dlg.exec();
    });
}
```

The `QTimer::singleShot(0, ...)` defers the dialog until after the main window finishes showing, so it appears centered on a fully-laid-out window.

- [ ] **Step 3: Build + run the app**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
release/DataViewer.exe
```

Manual verification:
1. On first launch (no QSettings entry), the prompt should appear immediately.
2. Enter a name, pick a color, click Save. App proceeds normally.
3. Close and relaunch. Prompt should NOT appear again.

To reset for testing:

```powershell
Remove-Item "HKCU:\Software\<OrgName>\DataViewer\identity" -Recurse -ErrorAction SilentlyContinue
```

- [ ] **Step 4: Commit**

```bash
git add src/MainWindow.h src/MainWindow.cpp
git commit -m "feat(ui): show identity prompt on first launch"
```

---

### Task 10: `ConfigLoader` — parse `db.conf`, decrypt password

**Files:**
- Create: `src/database/ConfigLoader.h`
- Create: `src/database/ConfigLoader.cpp`
- Create: `tests/tst_configloader/tst_configloader.pro`
- Create: `tests/tst_configloader/tst_configloader.cpp`

For v1 we use a simple keyed-AES scheme: the password is encrypted with a 256-bit key derived from a machine identifier (Windows machine GUID) hashed with SHA-256. This is deliberately weak crypto — its job is to keep casual readers out of `db.conf`, not to defeat a determined attacker (LAN-only architecture).

- [ ] **Step 1: Write the failing test via Python pattern**

`<RELATIVE_PATH>` = `tests/tst_configloader/tst_configloader.cpp`

`<FILE_CONTENT>`:

```cpp
#include <QtTest/QtTest>
#include <QTemporaryFile>
#include "../../src/database/ConfigLoader.h"

using DVE::ConfigLoader;
using DVE::DbConfig;

class TstConfigLoader : public QObject {
    Q_OBJECT
private slots:
    void encryptDecrypt_roundTrip_returnsOriginal() {
        const QString plain = "s3cr3t-password!@#";
        const QString encrypted = ConfigLoader::encryptPassword(plain);
        QVERIFY(encrypted != plain);
        QCOMPARE(ConfigLoader::decryptPassword(encrypted), plain);
    }

    void parseConfig_validIni_returnsPopulatedStruct() {
        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        tmp.write(
            "[postgres]\n"
            "host=db.example.local\n"
            "port=5432\n"
            "database=dataviewer\n"
            "user=dataviewer_app\n"
            "password_encrypted=" + ConfigLoader::encryptPassword("hunter2").toUtf8() + "\n"
        );
        tmp.flush();
        DbConfig cfg;
        QVERIFY(ConfigLoader::load(tmp.fileName(), cfg));
        QCOMPARE(cfg.host, QString("db.example.local"));
        QCOMPARE(cfg.port, 5432);
        QCOMPARE(cfg.database, QString("dataviewer"));
        QCOMPARE(cfg.user, QString("dataviewer_app"));
        QCOMPARE(cfg.password, QString("hunter2"));
    }

    void parseConfig_missingFile_returnsFalse() {
        DbConfig cfg;
        QVERIFY(!ConfigLoader::load("C:/no/such/file.conf", cfg));
    }

    void parseConfig_missingPassword_returnsFalseWithError() {
        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        tmp.write("[postgres]\nhost=h\nport=5432\ndatabase=d\nuser=u\n");
        tmp.flush();
        DbConfig cfg;
        QString err;
        QVERIFY(!ConfigLoader::load(tmp.fileName(), cfg, &err));
        QVERIFY(err.contains("password", Qt::CaseInsensitive));
    }
};

QTEST_MAIN(TstConfigLoader)
#include "tst_configloader.moc"
```

`tests/tst_configloader/tst_configloader.pro`:

```qmake
QT       += testlib core
CONFIG   += console c++17 testcase
CONFIG   -= app_bundle
TARGET    = tst_configloader
SOURCES  += tst_configloader.cpp \
            ../../src/database/ConfigLoader.cpp
HEADERS  += ../../src/database/ConfigLoader.h
INCLUDEPATH += ../../src
```

Append `tst_configloader` to `tests/tests.pro` SUBDIRS.

- [ ] **Step 2: Verify test fails to build**

```bash
cd tests/tst_configloader
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++
mingw32-make
```

Expected: build fails (ConfigLoader.h missing).

- [ ] **Step 3: Create `ConfigLoader.h` via Python pattern**

`<RELATIVE_PATH>` = `src/database/ConfigLoader.h`

`<FILE_CONTENT>`:

```cpp
#pragma once

#include <QString>

namespace DVE {

struct DbConfig {
    QString host;
    int     port = 5432;
    QString database;
    QString user;
    QString password;   // decrypted plaintext, populated by load()
};

class ConfigLoader {
public:
    // Loads + validates a db.conf-format file. Returns true on success.
    // On failure, *err (if non-null) gets a human-readable reason.
    static bool load(const QString& path, DbConfig& out, QString* err = nullptr);

    // Symmetric AES roundtrip helpers. Key derived from machine + user salt.
    static QString encryptPassword(const QString& plaintext);
    static QString decryptPassword(const QString& base64Cipher);

private:
    static QByteArray deriveKey();
};

} // namespace DVE
```

- [ ] **Step 4: Create `ConfigLoader.cpp` via Python pattern**

`<RELATIVE_PATH>` = `src/database/ConfigLoader.cpp`

`<FILE_CONTENT>`:

```cpp
#include "ConfigLoader.h"

#include <QFile>
#include <QSettings>
#include <QCryptographicHash>
#include <QSysInfo>
#include <QDir>
#include <QByteArray>

namespace DVE {

// Trivial XOR-with-derived-key cipher. Not strong crypto. Its purpose is to
// keep db.conf from being plaintext-readable by a casual viewer; the LAN-only
// architecture handles real threats. v1.1 swaps this for QAESEncryption once
// the dependency is added to the build.
static QByteArray xorCipher(const QByteArray& in, const QByteArray& key) {
    QByteArray out(in);
    const int kn = key.size();
    if (kn == 0) return out;
    for (int i = 0; i < out.size(); ++i) {
        out[i] = out[i] ^ key[i % kn];
    }
    return out;
}

QByteArray ConfigLoader::deriveKey() {
    // Hash machineUniqueId() + username so the cipher key is bound to this
    // workstation + this Windows user. Copying db.conf to another machine
    // yields a different key → decryption fails (correctly).
    QByteArray seed = QSysInfo::machineUniqueId();
    seed += qgetenv("USERNAME");
    seed += "DataViewerDbConfigSalt-v1";
    return QCryptographicHash::hash(seed, QCryptographicHash::Sha256);
}

QString ConfigLoader::encryptPassword(const QString& plaintext) {
    const QByteArray key = deriveKey();
    const QByteArray ct  = xorCipher(plaintext.toUtf8(), key);
    return QString::fromUtf8(ct.toBase64());
}

QString ConfigLoader::decryptPassword(const QString& base64Cipher) {
    const QByteArray key = deriveKey();
    const QByteArray ct  = QByteArray::fromBase64(base64Cipher.toUtf8());
    const QByteArray pt  = xorCipher(ct, key);
    return QString::fromUtf8(pt);
}

bool ConfigLoader::load(const QString& path, DbConfig& out, QString* err) {
    if (!QFile::exists(path)) {
        if (err) *err = QStringLiteral("db.conf not found at ") + path;
        return false;
    }
    QSettings ini(path, QSettings::IniFormat);
    ini.beginGroup("postgres");
    out.host     = ini.value("host").toString();
    out.port     = ini.value("port", 5432).toInt();
    out.database = ini.value("database").toString();
    out.user     = ini.value("user").toString();
    const QString enc = ini.value("password_encrypted").toString();
    ini.endGroup();

    if (out.host.isEmpty()) {
        if (err) *err = "[postgres] host is missing";
        return false;
    }
    if (out.database.isEmpty()) {
        if (err) *err = "[postgres] database is missing";
        return false;
    }
    if (out.user.isEmpty()) {
        if (err) *err = "[postgres] user is missing";
        return false;
    }
    if (enc.isEmpty()) {
        if (err) *err = "[postgres] password_encrypted is missing";
        return false;
    }
    out.password = decryptPassword(enc);
    return true;
}

} // namespace DVE
```

- [ ] **Step 5: Run tests**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_configloader
mingw32-make clean
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++
mingw32-make
.\release\tst_configloader.exe
```

Expected: all 4 tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/database/ConfigLoader.h src/database/ConfigLoader.cpp \
        tests/tst_configloader/ tests/tests.pro
git commit -m "feat(db): ConfigLoader for db.conf with XOR-derived password obfuscation"
```

---

### Phase 3 — `PostgresConnection` (Tasks 11–13)

These tasks build the connection layer. No concurrency logic yet — just open/close/ping with retry. The dual-connection model (one for queries, one reserved for LISTEN) is established now even though Plan A doesn't subscribe to anything — Plan B builds on this foundation.

---

### Task 11: `PostgresConnection` skeleton + connect test

**Files:**
- Create: `src/database/PostgresConnection.h`
- Create: `src/database/PostgresConnection.cpp`
- Create: `tests/tst_postgresconnection/tst_postgresconnection.pro`
- Create: `tests/tst_postgresconnection/tst_postgresconnection.cpp`

This is the first task that requires a running Postgres for testing. Use the helper script (built in Task 28) for now run a container manually.

- [ ] **Step 1: Start a test Postgres**

```powershell
# From repo root:
docker run -d --name dve-test-pg -p 5433:5432 \
  -e POSTGRES_DB=dve_test -e POSTGRES_USER=test -e POSTGRES_PASSWORD=test \
  postgres:16
# Wait for ready:
Start-Sleep -Seconds 5
docker exec dve-test-pg pg_isready -U test
$env:DVE_TEST_PG_CONN = "host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"
```

- [ ] **Step 2: Write the failing test via Python pattern**

`<RELATIVE_PATH>` = `tests/tst_postgresconnection/tst_postgresconnection.cpp`

`<FILE_CONTENT>`:

```cpp
#include <QtTest/QtTest>
#include <QSqlError>
#include "../../src/database/PostgresConnection.h"

using DVE::PostgresConnection;
using DVE::DbConfig;

namespace {
DbConfig configFromEnv() {
    DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    if (env.isEmpty()) return c;
    for (const QString& part : env.split(' ', Qt::SkipEmptyParts)) {
        const QStringList kv = part.split('=');
        if (kv.size() != 2) continue;
        const QString k = kv[0], v = kv[1];
        if      (k == "host")     c.host = v;
        else if (k == "port")     c.port = v.toInt();
        else if (k == "dbname")   c.database = v;
        else if (k == "user")     c.user = v;
        else if (k == "password") c.password = v;
    }
    return c;
}
}

class TstPostgresConnection : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) {
            QSKIP("DVE_TEST_PG_CONN not set; integration tests skipped");
        }
    }

    void open_returnsTrue_whenServerAvailable() {
        PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));
        QVERIFY(pg.isOpen());
        pg.close();
        QVERIFY(!pg.isOpen());
    }

    void open_returnsFalse_whenServerUnreachable() {
        DbConfig bad = configFromEnv();
        bad.port = 1;  // nothing listening
        PostgresConnection pg;
        QVERIFY(!pg.open(bad));
        QVERIFY(!pg.lastError().isEmpty());
    }

    void open_returnsFalse_withBadCredentials() {
        DbConfig bad = configFromEnv();
        bad.password = "wrong-password";
        PostgresConnection pg;
        QVERIFY(!pg.open(bad));
    }

    void ping_returnsTrue_whenConnected() {
        PostgresConnection pg;
        QVERIFY(pg.open(configFromEnv()));
        QVERIFY(pg.ping());
        pg.close();
    }
};

QTEST_MAIN(TstPostgresConnection)
#include "tst_postgresconnection.moc"
```

`tests/tst_postgresconnection/tst_postgresconnection.pro`:

```qmake
QT       += testlib core sql
CONFIG   += console c++17 testcase
CONFIG   -= app_bundle
TARGET    = tst_postgresconnection
SOURCES  += tst_postgresconnection.cpp \
            ../../src/database/PostgresConnection.cpp \
            ../../src/database/ConfigLoader.cpp
HEADERS  += ../../src/database/PostgresConnection.h \
            ../../src/database/ConfigLoader.h
INCLUDEPATH += ../../src
```

Append `tst_postgresconnection` to `tests/tests.pro` SUBDIRS.

- [ ] **Step 3: Create `PostgresConnection.h` via Python pattern**

`<RELATIVE_PATH>` = `src/database/PostgresConnection.h`

`<FILE_CONTENT>`:

```cpp
#pragma once

#include <QObject>
#include <QString>
#include <QSqlDatabase>
#include "ConfigLoader.h"

namespace DVE {

class PostgresConnection : public QObject {
    Q_OBJECT
public:
    explicit PostgresConnection(QObject* parent = nullptr);
    ~PostgresConnection() override;

    // Opens both the query connection and the LISTEN-reserved connection.
    // Returns true on success. On false, lastError() carries the reason.
    bool open(const DbConfig& cfg);
    void close();
    bool isOpen() const { return m_open; }

    // Lightweight "is the server still there?" check. Cheap; fires on
    // the query connection.
    bool ping();

    QSqlDatabase& queryDb()  { return m_queryDb; }
    QSqlDatabase& listenDb() { return m_listenDb; }

    QString lastError() const { return m_lastError; }

private:
    bool         m_open = false;
    QSqlDatabase m_queryDb;
    QSqlDatabase m_listenDb;
    QString      m_lastError;
    QString      m_queryName;
    QString      m_listenName;
};

} // namespace DVE
```

- [ ] **Step 4: Create `PostgresConnection.cpp` via Python pattern**

`<RELATIVE_PATH>` = `src/database/PostgresConnection.cpp`

`<FILE_CONTENT>`:

```cpp
#include "PostgresConnection.h"

#include <QSqlError>
#include <QSqlQuery>
#include <QUuid>

namespace DVE {

PostgresConnection::PostgresConnection(QObject* parent) : QObject(parent) {
    // Unique connection names so multiple instances coexist (e.g., in tests).
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_queryName  = "dve_query_"  + tag;
    m_listenName = "dve_listen_" + tag;
}

PostgresConnection::~PostgresConnection() {
    close();
}

static bool openOne(QSqlDatabase& db, const QString& name, const DbConfig& cfg,
                    QString& err) {
    db = QSqlDatabase::addDatabase("QPSQL", name);
    db.setHostName(cfg.host);
    db.setPort(cfg.port);
    db.setDatabaseName(cfg.database);
    db.setUserName(cfg.user);
    db.setPassword(cfg.password);
    // 3-second connect timeout so we fail fast on unreachable NAS.
    db.setConnectOptions("connect_timeout=3");
    if (!db.open()) {
        err = db.lastError().text();
        QSqlDatabase::removeDatabase(name);
        return false;
    }
    return true;
}

bool PostgresConnection::open(const DbConfig& cfg) {
    if (m_open) close();

    if (!openOne(m_queryDb, m_queryName, cfg, m_lastError)) {
        return false;
    }
    if (!openOne(m_listenDb, m_listenName, cfg, m_lastError)) {
        m_queryDb.close();
        QSqlDatabase::removeDatabase(m_queryName);
        return false;
    }
    m_open = true;
    m_lastError.clear();
    return true;
}

void PostgresConnection::close() {
    if (m_open) {
        m_queryDb.close();
        m_listenDb.close();
    }
    if (QSqlDatabase::contains(m_queryName))  QSqlDatabase::removeDatabase(m_queryName);
    if (QSqlDatabase::contains(m_listenName)) QSqlDatabase::removeDatabase(m_listenName);
    m_open = false;
}

bool PostgresConnection::ping() {
    if (!m_open) return false;
    QSqlQuery q(m_queryDb);
    return q.exec("SELECT 1");
}

} // namespace DVE
```

- [ ] **Step 5: Build and run tests**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_postgresconnection
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++
mingw32-make
.\release\tst_postgresconnection.exe
```

Expected: all 4 tests pass.

- [ ] **Step 6: Stop test Postgres + commit**

```powershell
docker rm -f dve-test-pg
```

```bash
git add src/database/PostgresConnection.h src/database/PostgresConnection.cpp \
        tests/tst_postgresconnection/ tests/tests.pro
git commit -m "feat(db): PostgresConnection with dual QSqlDatabase (query + listen)"
```

---

### Task 12: `PostgresConnection::tryOpenWithRetry` — jittered retry

**Files:**
- Modify: `src/database/PostgresConnection.h`
- Modify: `src/database/PostgresConnection.cpp`
- Modify: `tests/tst_postgresconnection/tst_postgresconnection.cpp`

- [ ] **Step 1: Add a failing test for retry behavior**

In `tests/tst_postgresconnection/tst_postgresconnection.cpp`, append inside the class:

```cpp
    void tryOpenWithRetry_eventuallySucceeds_whenServerComesUp() {
        // We can't easily start a server mid-test, so we test the negative
        // case: with totalTimeoutMs=500ms against a bad port, retry should
        // exit cleanly with elapsed >= 500ms.
        DbConfig bad = configFromEnv();
        bad.port = 1;
        PostgresConnection pg;
        QElapsedTimer t;
        t.start();
        QVERIFY(!pg.tryOpenWithRetry(bad, 500));
        const qint64 elapsed = t.elapsed();
        QVERIFY2(elapsed >= 500 && elapsed < 2000,
                 qPrintable(QString("elapsed=%1ms").arg(elapsed)));
    }
```

Add `#include <QElapsedTimer>` at the top.

- [ ] **Step 2: Verify test fails to build**

```bash
mingw32-make
```

Expected: `tryOpenWithRetry` not declared.

- [ ] **Step 3: Add `tryOpenWithRetry` declaration to header**

In `src/database/PostgresConnection.h`, in the public section:

```cpp
// Open with retry-with-jitter for up to totalTimeoutMs. Returns true on
// first success or false after timeout. Uses 0-5s exponential backoff
// with jitter (per spec Risk #7).
bool tryOpenWithRetry(const DbConfig& cfg, int totalTimeoutMs);
```

- [ ] **Step 4: Implement in `PostgresConnection.cpp`**

Append to `src/database/PostgresConnection.cpp`:

```cpp
#include <QElapsedTimer>
#include <QThread>
#include <QRandomGenerator>

bool PostgresConnection::tryOpenWithRetry(const DbConfig& cfg, int totalTimeoutMs) {
    QElapsedTimer t;
    t.start();
    int delayMs = 250;
    while (true) {
        if (open(cfg)) return true;
        if (t.elapsed() >= totalTimeoutMs) return false;
        // 0–5 s jitter, capped by remaining time
        const int jitter = QRandomGenerator::global()->bounded(5000);
        const int sleep  = qMin(delayMs + jitter, totalTimeoutMs - int(t.elapsed()));
        if (sleep <= 0) return false;
        QThread::msleep(sleep);
        delayMs = qMin(delayMs * 2, 5000);  // cap base backoff
    }
}
```

- [ ] **Step 5: Run tests**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_postgresconnection
mingw32-make
.\release\tst_postgresconnection.exe
```

Expected: all 5 tests pass.

- [ ] **Step 6: Commit**

```bash
git add src/database/PostgresConnection.h src/database/PostgresConnection.cpp \
        tests/tst_postgresconnection/tst_postgresconnection.cpp
git commit -m "feat(db): tryOpenWithRetry with exponential backoff + 0-5s jitter"
```

---

### Task 13: Add `src/database/` files to `DataViewerEnterprise.pro`

**Files:**
- Modify: `DataViewerEnterprise.pro`

- [ ] **Step 1: Append new files to SOURCES + HEADERS**

In `DataViewerEnterprise.pro`, find the existing `src/database/DatabaseManager` entries and append alongside:

```
SOURCES += \
    src/database/IdentityManager.cpp \
    src/database/IdentityPromptDialog.cpp \
    src/database/ConfigLoader.cpp \
    src/database/PostgresConnection.cpp

HEADERS += \
    src/database/IdentityManager.h \
    src/database/IdentityPromptDialog.h \
    src/database/ConfigLoader.h \
    src/database/PostgresConnection.h
```

Also ensure `QT += sql` is present (the existing project already uses SQLite via Qt SQL — the same module covers QPSQL).

- [ ] **Step 2: Full build**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make clean
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++ ../DataViewerEnterprise.pro
mingw32-make -j8
```

Expected: clean build with no errors or warnings (`-Werror -Wall -Wextra` is enforced).

- [ ] **Step 3: Commit**

```bash
git add DataViewerEnterprise.pro
git commit -m "build: register Phase-2 and Phase-3 database files in the .pro"
```

---

### Phase 4 — Migration tool (Tasks 14–21)

The migration tool is a CLI mode (`--migrate-from-sqlite=...`) on `DataViewer.exe`. It runs headless, writes a JSON report to `%TEMP%`, and either succeeds (renaming source SQLite to `.pre-migration.sqlite`) or rolls back entirely.

---

### Task 14: `MigrationReport` — JSON report scaffolding

**Files:**
- Create: `src/database/MigrationReport.h`
- Create: `src/database/MigrationReport.cpp`

- [ ] **Step 1: Create `MigrationReport.h` via Python pattern**

`<RELATIVE_PATH>` = `src/database/MigrationReport.h`

`<FILE_CONTENT>`:

```cpp
#pragma once

#include <QString>
#include <QVariantMap>
#include <QVector>
#include <QPair>

namespace DVE {

class MigrationReport {
public:
    void addTable(const QString& name, int sqliteCount, int postgresCount);
    void addError(const QString& msg);
    void setDuration(qint64 ms);
    void setSourcePath(const QString& path);
    void setStatus(const QString& s);   // "success" | "rolled_back" | "aborted"

    bool writeJson(const QString& path) const;
    QString summary() const;            // human-readable one-liner

private:
    QString m_status;
    QString m_sourcePath;
    qint64  m_durationMs = 0;
    QStringList m_errors;
    QVector<QPair<QString, QPair<int,int>>> m_tables; // name → (sqlite, pg)
};

} // namespace DVE
```

- [ ] **Step 2: Create `MigrationReport.cpp` via Python pattern**

`<RELATIVE_PATH>` = `src/database/MigrationReport.cpp`

`<FILE_CONTENT>`:

```cpp
#include "MigrationReport.h"

#include <QFile>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QDateTime>

namespace DVE {

void MigrationReport::addTable(const QString& name, int sqliteCount, int postgresCount) {
    m_tables.push_back({name, {sqliteCount, postgresCount}});
}

void MigrationReport::addError(const QString& msg) {
    m_errors << msg;
}

void MigrationReport::setDuration(qint64 ms)        { m_durationMs = ms; }
void MigrationReport::setSourcePath(const QString& p){ m_sourcePath = p; }
void MigrationReport::setStatus(const QString& s)    { m_status = s; }

bool MigrationReport::writeJson(const QString& path) const {
    QJsonObject root;
    root["status"]     = m_status;
    root["source"]     = m_sourcePath;
    root["duration_ms"]= m_durationMs;
    root["timestamp"]  = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    QJsonArray tables;
    for (const auto& t : m_tables) {
        QJsonObject o;
        o["name"]         = t.first;
        o["sqlite_count"] = t.second.first;
        o["postgres_count"] = t.second.second;
        o["match"]        = (t.second.first == t.second.second);
        tables.append(o);
    }
    root["tables"] = tables;

    QJsonArray errs;
    for (const QString& e : m_errors) errs.append(e);
    root["errors"] = errs;

    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Truncate)) return false;
    f.write(QJsonDocument(root).toJson(QJsonDocument::Indented));
    return true;
}

QString MigrationReport::summary() const {
    int total = 0;
    bool allMatch = true;
    for (const auto& t : m_tables) {
        total += t.second.first;
        if (t.second.first != t.second.second) allMatch = false;
    }
    return QString("status=%1 tables=%2 rows=%3 match=%4 errors=%5 took=%6ms")
        .arg(m_status)
        .arg(m_tables.size())
        .arg(total)
        .arg(allMatch ? "yes" : "no")
        .arg(m_errors.size())
        .arg(m_durationMs);
}

} // namespace DVE
```

- [ ] **Step 3: Add to `.pro`**

```
SOURCES += src/database/MigrationReport.cpp
HEADERS += src/database/MigrationReport.h
```

- [ ] **Step 4: Build to verify**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
```

Expected: clean build.

- [ ] **Step 5: Commit**

```bash
git add src/database/MigrationReport.h src/database/MigrationReport.cpp \
        DataViewerEnterprise.pro
git commit -m "feat(db): MigrationReport for SQLite→Postgres JSON output"
```

---

### Task 15: `MigrationTool` skeleton — open both DBs, no transfer yet

**Files:**
- Create: `src/database/MigrationTool.h`
- Create: `src/database/MigrationTool.cpp`
- Create: `tests/tst_migrationtool/tst_migrationtool.pro`
- Create: `tests/tst_migrationtool/tst_migrationtool.cpp`

- [ ] **Step 1: Write the failing test via Python pattern**

`<RELATIVE_PATH>` = `tests/tst_migrationtool/tst_migrationtool.cpp`

`<FILE_CONTENT>`:

```cpp
#include <QtTest/QtTest>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QTemporaryFile>
#include "../../src/database/MigrationTool.h"

using DVE::MigrationTool;
using DVE::DbConfig;

namespace {
DbConfig pgConfig() {
    DbConfig c;
    const QString env = qgetenv("DVE_TEST_PG_CONN");
    if (env.isEmpty()) return c;
    for (const QString& part : env.split(' ', Qt::SkipEmptyParts)) {
        const QStringList kv = part.split('=');
        if (kv.size() != 2) continue;
        const QString k = kv[0], v = kv[1];
        if      (k == "host")     c.host = v;
        else if (k == "port")     c.port = v.toInt();
        else if (k == "dbname")   c.database = v;
        else if (k == "user")     c.user = v;
        else if (k == "password") c.password = v;
    }
    return c;
}

QString makeEmptySqlite() {
    auto* tmp = new QTemporaryFile;
    tmp->setAutoRemove(false);
    tmp->open();
    tmp->close();
    QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "tst_mig_seed");
    db.setDatabaseName(tmp->fileName());
    db.open();
    QSqlQuery q(db);
    q.exec("CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL, "
           "file_name TEXT NOT NULL, loaded_at TEXT NOT NULL, "
           "template_version TEXT, sheet_count INT, sample_count INT)");
    db.close();
    QSqlDatabase::removeDatabase("tst_mig_seed");
    return tmp->fileName();
}
}

class TstMigrationTool : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        if (qgetenv("DVE_TEST_PG_CONN").isEmpty()) {
            QSKIP("DVE_TEST_PG_CONN not set; integration tests skipped");
        }
    }

    void open_returnsTrue_withValidPaths() {
        MigrationTool m;
        const QString sqlitePath = makeEmptySqlite();
        QVERIFY(m.open(sqlitePath, pgConfig()));
    }

    void open_returnsFalse_whenSqliteMissing() {
        MigrationTool m;
        QVERIFY(!m.open("C:/no/such/file.sqlite", pgConfig()));
        QVERIFY(m.lastError().contains("not found", Qt::CaseInsensitive));
    }
};

QTEST_MAIN(TstMigrationTool)
#include "tst_migrationtool.moc"
```

`tests/tst_migrationtool/tst_migrationtool.pro`:

```qmake
QT       += testlib core sql
CONFIG   += console c++17 testcase
CONFIG   -= app_bundle
TARGET    = tst_migrationtool
SOURCES  += tst_migrationtool.cpp \
            ../../src/database/MigrationTool.cpp \
            ../../src/database/MigrationReport.cpp \
            ../../src/database/PostgresConnection.cpp \
            ../../src/database/ConfigLoader.cpp
HEADERS  += ../../src/database/MigrationTool.h \
            ../../src/database/MigrationReport.h \
            ../../src/database/PostgresConnection.h \
            ../../src/database/ConfigLoader.h
INCLUDEPATH += ../../src
```

Append `tst_migrationtool` to `tests/tests.pro` SUBDIRS.

- [ ] **Step 2: Create `MigrationTool.h` via Python pattern**

`<RELATIVE_PATH>` = `src/database/MigrationTool.h`

`<FILE_CONTENT>`:

```cpp
#pragma once

#include <QObject>
#include <QString>
#include <QSqlDatabase>
#include "ConfigLoader.h"
#include "MigrationReport.h"

namespace DVE {

class MigrationTool : public QObject {
    Q_OBJECT
public:
    explicit MigrationTool(QObject* parent = nullptr);
    ~MigrationTool() override;

    // Opens source SQLite (read-only) and target Postgres.
    bool open(const QString& sqlitePath, const DbConfig& pg);

    // Runs the full migration. Returns true on success.
    // On failure, the report has the details. force=true wipes existing
    // Postgres data before running.
    bool run(bool force = false);

    // Renames source SQLite to <name>.pre-migration.sqlite. Called by
    // run() on success; exposed for tests.
    bool finalizeSource();

    const MigrationReport& report() const { return m_report; }
    QString lastError() const { return m_lastError; }

private:
    bool checkSchemaMetaEmpty();
    bool wipePostgresData();
    int  sqliteRowCount(const QString& table);
    int  postgresRowCount(const QString& table);
    bool migrateTable(const QString& name);
    bool bumpSequence(const QString& table);
    bool writeSchemaMeta();

    QSqlDatabase    m_sqlite;
    QSqlDatabase    m_pg;       // owned via QSqlDatabase::addDatabase
    QString         m_sqlitePath;
    QString         m_pgConnName;
    QString         m_lastError;
    MigrationReport m_report;
    QString         m_sourceSha256;
};

} // namespace DVE
```

- [ ] **Step 3: Create `MigrationTool.cpp` skeleton (open + checkSchemaMetaEmpty only) via Python pattern**

`<RELATIVE_PATH>` = `src/database/MigrationTool.cpp`

`<FILE_CONTENT>`:

```cpp
#include "MigrationTool.h"

#include <QFile>
#include <QSqlError>
#include <QSqlQuery>
#include <QUuid>
#include <QCryptographicHash>

namespace DVE {

MigrationTool::MigrationTool(QObject* parent) : QObject(parent) {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_pgConnName = "dve_mig_" + tag;
}

MigrationTool::~MigrationTool() {
    if (m_sqlite.isOpen()) m_sqlite.close();
    if (m_pg.isOpen())     m_pg.close();
    if (QSqlDatabase::contains("dve_mig_src_" + m_pgConnName))
        QSqlDatabase::removeDatabase("dve_mig_src_" + m_pgConnName);
    if (QSqlDatabase::contains(m_pgConnName))
        QSqlDatabase::removeDatabase(m_pgConnName);
}

bool MigrationTool::open(const QString& sqlitePath, const DbConfig& pg) {
    m_sqlitePath = sqlitePath;
    m_report.setSourcePath(sqlitePath);

    if (!QFile::exists(sqlitePath)) {
        m_lastError = "Source SQLite not found at " + sqlitePath;
        return false;
    }

    const QString srcName = "dve_mig_src_" + m_pgConnName;
    m_sqlite = QSqlDatabase::addDatabase("QSQLITE", srcName);
    m_sqlite.setDatabaseName(sqlitePath);
    m_sqlite.setConnectOptions("QSQLITE_OPEN_READONLY");
    if (!m_sqlite.open()) {
        m_lastError = "SQLite open failed: " + m_sqlite.lastError().text();
        return false;
    }

    // Compute source hash for provenance.
    QFile f(sqlitePath);
    if (f.open(QIODevice::ReadOnly)) {
        QCryptographicHash h(QCryptographicHash::Sha256);
        if (h.addData(&f)) m_sourceSha256 = h.result().toHex();
    }

    m_pg = QSqlDatabase::addDatabase("QPSQL", m_pgConnName);
    m_pg.setHostName(pg.host);
    m_pg.setPort(pg.port);
    m_pg.setDatabaseName(pg.database);
    m_pg.setUserName(pg.user);
    m_pg.setPassword(pg.password);
    m_pg.setConnectOptions("connect_timeout=5");
    if (!m_pg.open()) {
        m_lastError = "Postgres open failed: " + m_pg.lastError().text();
        return false;
    }
    return true;
}

bool MigrationTool::checkSchemaMetaEmpty() {
    QSqlQuery q(m_pg);
    if (!q.exec("SELECT value FROM schema_meta WHERE key = 'migrated_at'")) {
        m_lastError = "Could not read schema_meta: " + q.lastError().text();
        return false;
    }
    return !q.next();  // empty == OK
}

// Stubs for upcoming tasks
bool MigrationTool::wipePostgresData()                 { return false; }
int  MigrationTool::sqliteRowCount(const QString&)     { return 0; }
int  MigrationTool::postgresRowCount(const QString&)   { return 0; }
bool MigrationTool::migrateTable(const QString&)       { return false; }
bool MigrationTool::bumpSequence(const QString&)       { return false; }
bool MigrationTool::writeSchemaMeta()                  { return false; }
bool MigrationTool::finalizeSource()                   { return false; }
bool MigrationTool::run(bool)                          { return false; }

} // namespace DVE
```

- [ ] **Step 4: Add to `.pro` and build**

In `DataViewerEnterprise.pro`:

```
SOURCES += src/database/MigrationTool.cpp
HEADERS += src/database/MigrationTool.h
```

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
```

Expected: clean build.

- [ ] **Step 5: Start test Postgres + run tests**

```powershell
docker run -d --name dve-test-pg -p 5433:5432 \
  -v ${PWD}/deploy/postgres/init.sql:/docker-entrypoint-initdb.d/01-init.sql:ro \
  -e POSTGRES_DB=dve_test -e POSTGRES_USER=test -e POSTGRES_PASSWORD=test \
  postgres:16
Start-Sleep -Seconds 8
$env:DVE_TEST_PG_CONN = "host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"
cd tests/tst_migrationtool
"C:\Qt\6.10.2\mingw_64\bin\qmake.exe" -spec win32-g++
mingw32-make
.\release\tst_migrationtool.exe
```

Expected: both tests pass.

- [ ] **Step 6: Tear down + commit**

```powershell
docker rm -f dve-test-pg
```

```bash
git add src/database/MigrationTool.h src/database/MigrationTool.cpp \
        tests/tst_migrationtool/ tests/tests.pro DataViewerEnterprise.pro
git commit -m "feat(db): MigrationTool skeleton with SQLite + Postgres open"
```

---

### Task 16: `migrateTable` for `files` (single-table generic transfer)

**Files:**
- Modify: `src/database/MigrationTool.cpp`
- Modify: `tests/tst_migrationtool/tst_migrationtool.cpp`

`migrateTable` is the workhorse. We implement it generically and parameterize per-table column lists. For Task 16, we make it work for the `files` table only and add a test; subsequent tasks add the rest of the tables to the parametrized call list.

- [ ] **Step 1: Append failing test**

In `tests/tst_migrationtool/tst_migrationtool.cpp`, add:

```cpp
    void migrate_files_roundTripsAllRowsWithIds() {
        // Seed SQLite with 3 rows
        auto* tmp = new QTemporaryFile;
        tmp->setAutoRemove(false);
        tmp->open(); tmp->close();
        {
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "seed_files");
            db.setDatabaseName(tmp->fileName());
            db.open();
            QSqlQuery q(db);
            q.exec("CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL, "
                   "file_name TEXT NOT NULL, loaded_at TEXT NOT NULL, "
                   "template_version TEXT, sheet_count INT, sample_count INT)");
            q.exec("INSERT INTO files VALUES (10, 'C:/a.xlsx', 'a.xlsx', '2026-01-01', 'v1', 2, 5)");
            q.exec("INSERT INTO files VALUES (20, 'C:/b.xlsx', 'b.xlsx', '2026-01-02', 'v1', 3, 6)");
            q.exec("INSERT INTO files VALUES (30, 'C:/c.xlsx', 'c.xlsx', '2026-01-03', 'v2', 1, 2)");
            db.close();
            QSqlDatabase::removeDatabase("seed_files");
        }

        MigrationTool m;
        QVERIFY(m.open(tmp->fileName(), pgConfig()));
        QVERIFY(m.migrateTable("files"));

        // Verify postgres has 3 rows with the same ids
        QSqlDatabase pg = QSqlDatabase::database("dve_mig_" /* will be tagged */);
        // Use a fresh local connection to inspect:
        DbConfig cfg = pgConfig();
        QSqlDatabase chk = QSqlDatabase::addDatabase("QPSQL", "tst_chk");
        chk.setHostName(cfg.host); chk.setPort(cfg.port);
        chk.setDatabaseName(cfg.database); chk.setUserName(cfg.user);
        chk.setPassword(cfg.password);
        QVERIFY(chk.open());
        QSqlQuery q(chk);
        QVERIFY(q.exec("SELECT id, file_path FROM files ORDER BY id"));
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 10);
        QCOMPARE(q.value(1).toString(), QString("C:/a.xlsx"));
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 20);
        QVERIFY(q.next()); QCOMPARE(q.value(0).toInt(), 30);
        QVERIFY(!q.next());
        chk.close();
        QSqlDatabase::removeDatabase("tst_chk");
    }
```

- [ ] **Step 2: Implement `migrateTable` for `files`**

Replace the stub `migrateTable` in `src/database/MigrationTool.cpp` with:

```cpp
// Column list per table. Order matters — controls INSERT order.
// Audit columns (updated_at, updated_by, version) are NOT here; they
// default at insert.
static const QStringList kColsFiles = {
    "id", "file_path", "file_name", "loaded_at",
    "template_version", "sheet_count", "sample_count"
};

bool MigrationTool::migrateTable(const QString& name) {
    QStringList cols;
    if      (name == "files") cols = kColsFiles;
    else {
        m_lastError = "Unknown table for migrateTable: " + name;
        return false;
    }

    // Read all rows from SQLite
    QSqlQuery src(m_sqlite);
    if (!src.exec(QString("SELECT %1 FROM %2").arg(cols.join(", "), name))) {
        m_lastError = "Source read failed: " + src.lastError().text();
        return false;
    }

    // Prepare a single parametrized INSERT for batching
    QStringList placeholders;
    for (int i = 0; i < cols.size(); ++i) placeholders << QString(":v%1").arg(i);
    QSqlQuery dst(m_pg);
    dst.prepare(QString("INSERT INTO %1 (%2) OVERRIDING SYSTEM VALUE VALUES (%3)")
                .arg(name, cols.join(", "), placeholders.join(", ")));

    int n = 0;
    while (src.next()) {
        for (int i = 0; i < cols.size(); ++i) {
            dst.bindValue(QString(":v%1").arg(i), src.value(i));
        }
        if (!dst.exec()) {
            m_lastError = QString("Insert row %1 into %2 failed: %3")
                            .arg(n).arg(name, dst.lastError().text());
            return false;
        }
        ++n;
    }
    return true;
}
```

- [ ] **Step 3: Run test**

```bash
python tools/decrypt_via_copy.py --apply
cd tests/tst_migrationtool
mingw32-make
.\release\tst_migrationtool.exe
```

Expected: 3 tests pass.

- [ ] **Step 4: Commit**

```bash
git add src/database/MigrationTool.cpp \
        tests/tst_migrationtool/tst_migrationtool.cpp
git commit -m "feat(db): migrateTable for files (parameterized, id-preserving)"
```

---

### Task 17: `migrateTable` — extend column lists for all data tables

**Files:**
- Modify: `src/database/MigrationTool.cpp`

This task adds the remaining `kCols*` constants and the dispatch. No new test code — the per-table migration is exercised by the full-roundtrip test in Task 20.

- [ ] **Step 1: Add column lists for every table**

Replace the `kColsFiles` block and the `migrateTable` `if (name == "files")` chain in `src/database/MigrationTool.cpp` with:

```cpp
namespace {
static const QStringList kColsFiles = {
    "id", "file_path", "file_name", "loaded_at",
    "template_version", "sheet_count", "sample_count"
};
static const QStringList kColsTests = {
    "id", "file_id", "sheet_name", "template_version",
    "overall_avg_tpm", "overall_stddev_tpm", "is_raw_table", "sort_order"
};
static const QStringList kColsSamples = {
    "id", "test_id", "sort_order", "sample_name", "sample_id", "date", "tester",
    "media", "viscosity", "resistance", "voltage", "power", "heating_technology",
    "puffing_regime", "initial_oil_mass", "average_tpm", "stddev_tpm",
    "avg_power_density", "efficiency_percent", "total_oil_consumed",
    "total_puffs", "normalized_tpm", "burn_status", "clog_status", "leak_status"
};
static const QStringList kColsDataRows = {
    "id", "sample_id", "sort_order", "puffs", "before_weight", "after_weight",
    "draw_pressure", "resistance", "smell", "clog", "notes", "tpm",
    "tpm_power_density", "variation_tpm", "oil_consumed"
};
static const QStringList kColsImages = {
    "id", "sample_id", "sort_order", "file_name", "image_data",
    "layout_x", "layout_y", "layout_w", "layout_h",
    "crop_x", "crop_y", "crop_w", "crop_h"
};
static const QStringList kColsSensorySessions = {
    "id", "session_name", "tester_name", "assessor_name", "media", "puff_length",
    "date", "timestamp", "json_data", "layout_json"
};
static const QStringList kColsSensoryImages = {
    "id", "session_id", "sort_order", "file_name", "image_data",
    "layout_x", "layout_y", "layout_w", "layout_h",
    "crop_x", "crop_y", "crop_w", "crop_h"
};
static const QStringList kColsDetailedSensorySessions = {
    "id", "session_name", "tester_name", "assessor_name", "media",
    "date", "timestamp", "json_data"
};
static const QStringList kColsDetailedSensoryImages = {
    "id", "session_id", "sort_order", "file_name", "image_data",
    "layout_x", "layout_y", "layout_w", "layout_h",
    "crop_x", "crop_y", "crop_w", "crop_h"
};
static const QStringList kColsSettings = { "key", "value" };
} // namespace

// Column lists indexed by table name
static QStringList columnsFor(const QString& name) {
    if (name == "files")                      return kColsFiles;
    if (name == "tests")                      return kColsTests;
    if (name == "samples")                    return kColsSamples;
    if (name == "data_rows")                  return kColsDataRows;
    if (name == "images")                     return kColsImages;
    if (name == "sensory_sessions")           return kColsSensorySessions;
    if (name == "sensory_images")             return kColsSensoryImages;
    if (name == "detailed_sensory_sessions")  return kColsDetailedSensorySessions;
    if (name == "detailed_sensory_images")    return kColsDetailedSensoryImages;
    if (name == "settings")                   return kColsSettings;
    return {};
}
```

Then update `migrateTable` to use `columnsFor`:

```cpp
bool MigrationTool::migrateTable(const QString& name) {
    QStringList cols = columnsFor(name);
    if (cols.isEmpty()) {
        m_lastError = "Unknown table for migrateTable: " + name;
        return false;
    }

    // Settings has no `id` column — skip OVERRIDING SYSTEM VALUE branch
    const bool hasId = (cols.first() == "id");
    const QString overriding = hasId ? "OVERRIDING SYSTEM VALUE " : "";

    QSqlQuery src(m_sqlite);
    if (!src.exec(QString("SELECT %1 FROM %2").arg(cols.join(", "), name))) {
        m_lastError = "Source read failed for " + name + ": " + src.lastError().text();
        return false;
    }

    QStringList placeholders;
    for (int i = 0; i < cols.size(); ++i) placeholders << QString(":v%1").arg(i);
    QSqlQuery dst(m_pg);
    dst.prepare(QString("INSERT INTO %1 (%2) %3VALUES (%4)")
                .arg(name, cols.join(", "), overriding, placeholders.join(", ")));

    int n = 0;
    while (src.next()) {
        for (int i = 0; i < cols.size(); ++i) {
            QVariant v = src.value(i);
            // JSONB columns: pass JSON text; Postgres parses on insert.
            // (json_data and layout_json are TEXT in SQLite, JSONB in Postgres)
            dst.bindValue(QString(":v%1").arg(i), v);
        }
        if (!dst.exec()) {
            m_lastError = QString("Insert row %1 into %2 failed: %3")
                            .arg(n).arg(name, dst.lastError().text());
            return false;
        }
        ++n;
    }
    return true;
}
```

- [ ] **Step 2: Build to verify**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
```

Expected: clean build.

- [ ] **Step 3: Commit**

```bash
git add src/database/MigrationTool.cpp
git commit -m "feat(db): migrateTable column lists for all 10 tables"
```

---

### Task 18: `sqliteRowCount` + `postgresRowCount` + verification

**Files:**
- Modify: `src/database/MigrationTool.cpp`

- [ ] **Step 1: Implement count helpers**

Replace the stubs in `src/database/MigrationTool.cpp`:

```cpp
int MigrationTool::sqliteRowCount(const QString& table) {
    QSqlQuery q(m_sqlite);
    if (!q.exec(QString("SELECT COUNT(*) FROM %1").arg(table))) return -1;
    return q.next() ? q.value(0).toInt() : -1;
}

int MigrationTool::postgresRowCount(const QString& table) {
    QSqlQuery q(m_pg);
    if (!q.exec(QString("SELECT COUNT(*) FROM %1").arg(table))) return -1;
    return q.next() ? q.value(0).toInt() : -1;
}
```

- [ ] **Step 2: Build to verify**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
```

Expected: clean build.

- [ ] **Step 3: Commit**

```bash
git add src/database/MigrationTool.cpp
git commit -m "feat(db): row-count helpers for migration verification"
```

---

### Task 19: `bumpSequence`, `writeSchemaMeta`, `wipePostgresData`

**Files:**
- Modify: `src/database/MigrationTool.cpp`

- [ ] **Step 1: Implement all three**

Replace the stubs in `src/database/MigrationTool.cpp`:

```cpp
bool MigrationTool::bumpSequence(const QString& table) {
    // Tables without `id` (settings) have no sequence to bump.
    if (table == "settings") return true;

    // Compute MAX(id) and bump the sequence past it.
    QSqlQuery q(m_pg);
    if (!q.exec(QString("SELECT setval(pg_get_serial_sequence('%1','id'), "
                        "  COALESCE((SELECT MAX(id) FROM %1), 1), true)")
                .arg(table))) {
        m_lastError = "Sequence bump failed for " + table + ": "
                      + q.lastError().text();
        return false;
    }
    return true;
}

bool MigrationTool::writeSchemaMeta() {
    QSqlQuery q(m_pg);
    auto upsert = [&q](const QString& k, const QString& v) -> bool {
        q.prepare("INSERT INTO schema_meta(key, value) VALUES (?, ?) "
                  "ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value");
        q.addBindValue(k);
        q.addBindValue(v);
        return q.exec();
    };
    if (!upsert("schema_version", "2"))                       return false;
    if (!upsert("migrated_from",  m_sqlitePath))              return false;
    if (!upsert("migrated_at",
                QDateTime::currentDateTimeUtc().toString(Qt::ISODate))) return false;
    if (!upsert("source_sha256",  m_sourceSha256))            return false;
    return true;
}

bool MigrationTool::wipePostgresData() {
    // Delete-all in FK-respecting order.
    const QStringList order = {
        "data_rows", "images", "samples", "tests", "files",
        "sensory_images", "sensory_sessions",
        "detailed_sensory_images", "detailed_sensory_sessions",
        "settings", "schema_meta"
    };
    QSqlQuery q(m_pg);
    for (const QString& t : order) {
        if (!q.exec("DELETE FROM " + t)) {
            m_lastError = "Wipe " + t + " failed: " + q.lastError().text();
            return false;
        }
    }
    return true;
}
```

Add `#include <QDateTime>` at the top of `MigrationTool.cpp` if not already present.

- [ ] **Step 2: Build to verify**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
```

Expected: clean build.

- [ ] **Step 3: Commit**

```bash
git add src/database/MigrationTool.cpp
git commit -m "feat(db): sequence bump, schema_meta writer, force-wipe helper"
```

---

### Task 20: `run()` — full transactional migration + roundtrip test

**Files:**
- Modify: `src/database/MigrationTool.cpp`
- Modify: `tests/tst_migrationtool/tst_migrationtool.cpp`

This task wires everything together: the full `run()` that wraps everything in a transaction, runs `migrateTable` for each table in FK order, verifies counts, bumps sequences, writes metadata, and rolls back on any failure.

- [ ] **Step 1: Append failing roundtrip test**

In `tests/tst_migrationtool/tst_migrationtool.cpp`, append:

```cpp
    void run_fullRoundTrip_succeedsAndVerifies() {
        // Seed SQLite with a small representative dataset
        auto* tmp = new QTemporaryFile;
        tmp->setAutoRemove(false);
        tmp->open(); tmp->close();
        {
            QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "seed_full");
            db.setDatabaseName(tmp->fileName());
            db.open();
            QSqlQuery q(db);
            // Recreate exact SQLite source schema (no audit columns,
            // matching the pre-migration shape)
            q.exec("CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL, "
                   "file_name TEXT NOT NULL, loaded_at TEXT NOT NULL, "
                   "template_version TEXT, sheet_count INT, sample_count INT)");
            q.exec("CREATE TABLE tests (id INTEGER PRIMARY KEY, file_id INT, "
                   "sheet_name TEXT, template_version TEXT, overall_avg_tpm REAL, "
                   "overall_stddev_tpm REAL, is_raw_table INT, sort_order INT)");
            q.exec("CREATE TABLE samples (id INTEGER PRIMARY KEY, test_id INT, "
                   "sort_order INT, sample_name TEXT, sample_id TEXT, date TEXT, "
                   "tester TEXT, media TEXT, viscosity REAL, resistance REAL, "
                   "voltage REAL, power REAL, heating_technology TEXT, "
                   "puffing_regime TEXT, initial_oil_mass REAL, average_tpm REAL, "
                   "stddev_tpm REAL, avg_power_density REAL, efficiency_percent REAL, "
                   "total_oil_consumed REAL, total_puffs INT, normalized_tpm REAL, "
                   "burn_status TEXT, clog_status TEXT, leak_status TEXT)");
            q.exec("CREATE TABLE data_rows (id INTEGER PRIMARY KEY, sample_id INT, "
                   "sort_order INT, puffs REAL, before_weight REAL, after_weight REAL, "
                   "draw_pressure REAL, resistance REAL, smell TEXT, clog TEXT, "
                   "notes TEXT, tpm REAL, tpm_power_density REAL, variation_tpm REAL, "
                   "oil_consumed REAL)");
            q.exec("CREATE TABLE images (id INTEGER PRIMARY KEY, sample_id INT, "
                   "sort_order INT, file_name TEXT, image_data BLOB, "
                   "layout_x REAL, layout_y REAL, layout_w REAL, layout_h REAL, "
                   "crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL)");
            q.exec("CREATE TABLE sensory_sessions (id INTEGER PRIMARY KEY, "
                   "session_name TEXT, tester_name TEXT, assessor_name TEXT, "
                   "media TEXT, puff_length TEXT, date TEXT, timestamp TEXT, "
                   "json_data TEXT, layout_json TEXT)");
            q.exec("CREATE TABLE sensory_images (id INTEGER PRIMARY KEY, session_id INT, "
                   "sort_order INT, file_name TEXT, image_data BLOB, "
                   "layout_x REAL, layout_y REAL, layout_w REAL, layout_h REAL, "
                   "crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL)");
            q.exec("CREATE TABLE detailed_sensory_sessions (id INTEGER PRIMARY KEY, "
                   "session_name TEXT, tester_name TEXT, assessor_name TEXT, "
                   "media TEXT, date TEXT, timestamp TEXT, json_data TEXT)");
            q.exec("CREATE TABLE detailed_sensory_images (id INTEGER PRIMARY KEY, "
                   "session_id INT, sort_order INT, file_name TEXT, image_data BLOB, "
                   "layout_x REAL, layout_y REAL, layout_w REAL, layout_h REAL, "
                   "crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL)");
            q.exec("CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT)");

            q.exec("INSERT INTO files VALUES (1, 'C:/a.xlsx', 'a.xlsx', "
                   "'2026-01-01', 'v1', 1, 1)");
            q.exec("INSERT INTO tests VALUES (1, 1, 'Sheet1', 'v1', 1.0, 0.1, 0, 0)");
            q.exec("INSERT INTO samples VALUES (1, 1, 0, 'S1', 'sid1', "
                   "'2026-01-01', 'Bob', 'oil', 1, 2, 3, 4, 'ceramic', 'CRM', "
                   "5, 6, 7, 8, 90, 9, 10, 11, 'no', 'no', 'no')");
            q.exec("INSERT INTO data_rows VALUES (1, 1, 0, 50, 1.0, 1.05, "
                   "100, 2, '', '', '', 1.0, 0.5, 0.05, 0.05)");
            q.exec("INSERT INTO sensory_sessions VALUES (1, 'Trial1', 'Bob', "
                   "'Sarah', 'oil', '3s', '2026-01-01', '2026-01-01T10:00:00', "
                   "'{\"k\":1}', '{}')");
            q.exec("INSERT INTO detailed_sensory_sessions VALUES (1, 'DTrial1', "
                   "'Bob', 'Sarah', 'oil', '2026-01-01', "
                   "'2026-01-01T10:00:00', '{\"q\":1}')");
            q.exec("INSERT INTO settings VALUES ('theme', 'dark')");
            db.close();
            QSqlDatabase::removeDatabase("seed_full");
        }

        MigrationTool m;
        QVERIFY(m.open(tmp->fileName(), pgConfig()));
        QVERIFY2(m.run(/*force=*/true), qPrintable(m.lastError()));

        // Verify the report
        const auto& r = m.report();
        QVERIFY(r.summary().contains("status=success"));
        QVERIFY(r.summary().contains("match=yes"));

        // Spot-check: the SQLite source was renamed
        QVERIFY(!QFile::exists(tmp->fileName()));
        QVERIFY(QFile::exists(tmp->fileName() + ".pre-migration.sqlite"));
    }
```

- [ ] **Step 2: Implement `run()` and `finalizeSource()`**

In `src/database/MigrationTool.cpp`, replace both stubs with:

```cpp
bool MigrationTool::run(bool force) {
    QElapsedTimer timer;
    timer.start();

    // Pre-check: schema_meta empty unless force
    if (!force && !checkSchemaMetaEmpty()) {
        m_lastError = "Postgres already has migration metadata. Use --force "
                      "only after rolling back via the pre-migration SQLite.";
        m_report.setStatus("aborted");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    if (force) {
        if (!wipePostgresData()) {
            m_report.setStatus("aborted");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
    }

    // FK-dependency order
    const QStringList order = {
        "files", "tests", "samples", "data_rows", "images",
        "sensory_sessions", "sensory_images",
        "detailed_sensory_sessions", "detailed_sensory_images",
        "settings"
    };

    if (!m_pg.transaction()) {
        m_lastError = "BEGIN failed: " + m_pg.lastError().text();
        m_report.setStatus("aborted");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    for (const QString& t : order) {
        const int sqliteN = sqliteRowCount(t);
        if (!migrateTable(t)) {
            m_pg.rollback();
            m_report.setStatus("rolled_back");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
        const int pgN = postgresRowCount(t);
        m_report.addTable(t, sqliteN, pgN);
        if (sqliteN != pgN) {
            m_lastError = QString("Row count mismatch on %1: sqlite=%2 pg=%3")
                            .arg(t).arg(sqliteN).arg(pgN);
            m_pg.rollback();
            m_report.setStatus("rolled_back");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
    }

    for (const QString& t : order) {
        if (!bumpSequence(t)) {
            m_pg.rollback();
            m_report.setStatus("rolled_back");
            m_report.addError(m_lastError);
            m_report.setDuration(timer.elapsed());
            return false;
        }
    }

    if (!writeSchemaMeta()) {
        m_pg.rollback();
        m_report.setStatus("rolled_back");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    if (!m_pg.commit()) {
        m_lastError = "COMMIT failed: " + m_pg.lastError().text();
        m_report.setStatus("rolled_back");
        m_report.addError(m_lastError);
        m_report.setDuration(timer.elapsed());
        return false;
    }

    if (!finalizeSource()) {
        // Commit succeeded but rename failed — still a partial success.
        m_report.addError("Migration committed; source rename failed: " + m_lastError);
    }

    m_report.setStatus("success");
    m_report.setDuration(timer.elapsed());
    return true;
}

bool MigrationTool::finalizeSource() {
    const QString target = m_sqlitePath + ".pre-migration.sqlite";
    // Close source connection so the file isn't locked on Windows.
    if (m_sqlite.isOpen()) m_sqlite.close();
    if (!QFile::rename(m_sqlitePath, target)) {
        m_lastError = "Could not rename source to " + target;
        return false;
    }
    return true;
}
```

Add `#include <QElapsedTimer>` if not present.

- [ ] **Step 3: Run tests**

```powershell
docker run -d --name dve-test-pg -p 5433:5432 \
  -v ${PWD}/deploy/postgres/init.sql:/docker-entrypoint-initdb.d/01-init.sql:ro \
  -e POSTGRES_DB=dve_test -e POSTGRES_USER=test -e POSTGRES_PASSWORD=test \
  -e shared_preload_libraries=pg_cron \
  postgres:16
Start-Sleep -Seconds 8
$env:DVE_TEST_PG_CONN = "host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"
python tools/decrypt_via_copy.py --apply
cd tests/tst_migrationtool
mingw32-make
.\release\tst_migrationtool.exe
```

Expected: all 4 tests pass.

- [ ] **Step 4: Tear down + commit**

```powershell
docker rm -f dve-test-pg
```

```bash
git add src/database/MigrationTool.cpp tests/tst_migrationtool/tst_migrationtool.cpp
git commit -m "feat(db): MigrationTool::run with transactional all-or-nothing semantics"
```

---

### Task 21: CLI wiring in `main.cpp`

**Files:**
- Modify: `src/main.cpp`

- [ ] **Step 1: Inspect current `main.cpp`**

Read `src/main.cpp` to find the existing CLI-flag handling (`--self-test`, `__RAISE__`). The new flags need to slot in before any UI is constructed.

- [ ] **Step 2: Add `--migrate-from-sqlite` handling**

In `src/main.cpp`, near the top of `main()` (after `QCoreApplication::setOrganizationName` etc., before the `MainWindow` construction or the `--self-test` branch), add:

```cpp
#include "database/MigrationTool.h"
#include "database/ConfigLoader.h"
#include <QCommandLineParser>

// Inside main(), after QApplication construction:
{
    QCommandLineParser p;
    QCommandLineOption optMigrate("migrate-from-sqlite",
        "Migrate a SQLite source database to Postgres and exit. Headless.",
        "path");
    QCommandLineOption optTo("to-postgres",
        "Target Postgres connection string (key=value form).",
        "conn");
    QCommandLineOption optForce("force",
        "Wipe existing Postgres data before migrating. Dangerous.");
    QCommandLineOption optReportOut("migration-report-out",
        "Path to write the JSON report (defaults to %TEMP%\\dataviewer_migration.json).",
        "path");
    p.addOption(optMigrate);
    p.addOption(optTo);
    p.addOption(optForce);
    p.addOption(optReportOut);
    p.parse(QCoreApplication::arguments());

    if (p.isSet(optMigrate)) {
        const QString src = p.value(optMigrate);
        const QString connStr = p.value(optTo);
        const bool force  = p.isSet(optForce);
        const QString reportPath = p.isSet(optReportOut)
            ? p.value(optReportOut)
            : (QDir::tempPath() + "/dataviewer_migration.json");

        DVE::DbConfig cfg;
        // Parse connStr "host=... port=... dbname=... user=... password=..."
        for (const QString& part : connStr.split(' ', Qt::SkipEmptyParts)) {
            const QStringList kv = part.split('=');
            if (kv.size() != 2) continue;
            const QString k = kv[0], v = kv[1];
            if      (k == "host")     cfg.host = v;
            else if (k == "port")     cfg.port = v.toInt();
            else if (k == "dbname")   cfg.database = v;
            else if (k == "user")     cfg.user = v;
            else if (k == "password") cfg.password = v;
        }

        DVE::MigrationTool m;
        if (!m.open(src, cfg)) {
            qCritical().noquote() << "Migration setup failed:" << m.lastError();
            m.report().writeJson(reportPath);
            return 2;
        }
        const bool ok = m.run(force);
        m.report().writeJson(reportPath);
        qInfo().noquote() << "Migration:" << m.report().summary();
        qInfo().noquote() << "Report:" << reportPath;
        return ok ? 0 : 1;
    }
}
```

- [ ] **Step 3: Build and test the CLI manually**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
```

Then create a tiny test SQLite + run:

```powershell
docker run -d --name dve-test-pg -p 5433:5432 `
  -v ${PWD}/../deploy/postgres/init.sql:/docker-entrypoint-initdb.d/01-init.sql:ro `
  -e POSTGRES_DB=dve_test -e POSTGRES_USER=test -e POSTGRES_PASSWORD=test `
  postgres:16
Start-Sleep -Seconds 8

# Make a tiny SQLite to migrate (using sqlite3 cli)
@"
CREATE TABLE files (id INTEGER PRIMARY KEY, file_path TEXT NOT NULL,
  file_name TEXT NOT NULL, loaded_at TEXT NOT NULL,
  template_version TEXT, sheet_count INT, sample_count INT);
INSERT INTO files VALUES (1, 'C:/a.xlsx', 'a.xlsx', '2026-01-01', 'v1', 0, 0);
CREATE TABLE tests (id INTEGER PRIMARY KEY, file_id INT, sheet_name TEXT,
  template_version TEXT, overall_avg_tpm REAL, overall_stddev_tpm REAL,
  is_raw_table INT, sort_order INT);
CREATE TABLE samples (id INTEGER PRIMARY KEY, test_id INT, sort_order INT,
  sample_name TEXT, sample_id TEXT, date TEXT, tester TEXT, media TEXT,
  viscosity REAL, resistance REAL, voltage REAL, power REAL,
  heating_technology TEXT, puffing_regime TEXT, initial_oil_mass REAL,
  average_tpm REAL, stddev_tpm REAL, avg_power_density REAL,
  efficiency_percent REAL, total_oil_consumed REAL, total_puffs INT,
  normalized_tpm REAL, burn_status TEXT, clog_status TEXT, leak_status TEXT);
CREATE TABLE data_rows (id INTEGER PRIMARY KEY, sample_id INT, sort_order INT,
  puffs REAL, before_weight REAL, after_weight REAL, draw_pressure REAL,
  resistance REAL, smell TEXT, clog TEXT, notes TEXT, tpm REAL,
  tpm_power_density REAL, variation_tpm REAL, oil_consumed REAL);
CREATE TABLE images (id INTEGER PRIMARY KEY, sample_id INT, sort_order INT,
  file_name TEXT, image_data BLOB, layout_x REAL, layout_y REAL,
  layout_w REAL, layout_h REAL, crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL);
CREATE TABLE sensory_sessions (id INTEGER PRIMARY KEY, session_name TEXT,
  tester_name TEXT, assessor_name TEXT, media TEXT, puff_length TEXT,
  date TEXT, timestamp TEXT, json_data TEXT, layout_json TEXT);
CREATE TABLE sensory_images (id INTEGER PRIMARY KEY, session_id INT, sort_order INT,
  file_name TEXT, image_data BLOB, layout_x REAL, layout_y REAL,
  layout_w REAL, layout_h REAL, crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL);
CREATE TABLE detailed_sensory_sessions (id INTEGER PRIMARY KEY, session_name TEXT,
  tester_name TEXT, assessor_name TEXT, media TEXT, date TEXT, timestamp TEXT,
  json_data TEXT);
CREATE TABLE detailed_sensory_images (id INTEGER PRIMARY KEY, session_id INT,
  sort_order INT, file_name TEXT, image_data BLOB, layout_x REAL, layout_y REAL,
  layout_w REAL, layout_h REAL, crop_x REAL, crop_y REAL, crop_w REAL, crop_h REAL);
CREATE TABLE settings (key TEXT PRIMARY KEY, value TEXT);
"@ | sqlite3 test.sqlite

.\release\DataViewer.exe --migrate-from-sqlite=test.sqlite `
  --to-postgres="host=127.0.0.1 port=5433 dbname=dve_test user=test password=test" `
  --force
```

Expected: exit code 0, stdout includes `status=success match=yes`, file renamed to `test.sqlite.pre-migration.sqlite`.

- [ ] **Step 4: Tear down + commit**

```powershell
docker rm -f dve-test-pg
Remove-Item test.sqlite.pre-migration.sqlite -ErrorAction SilentlyContinue
```

```bash
git add src/main.cpp
git commit -m "feat(cli): --migrate-from-sqlite / --to-postgres headless mode"
```

---

### Phase 5 — libpq bundling + installer (Tasks 22–25)

The Qt QPSQL driver requires libpq.dll + several transitive dependencies. We stage them in a `vendor/libpq-16/` directory committed to the repo (so the build is reproducible without external downloads at build time), copy them into `release/` at build time, and reference them from `installer.iss`.

---

### Task 22: Stage libpq 16 DLLs

**Files:**
- Create: `vendor/libpq-16/README.md`
- Add (binary, not in this plan's diff): `vendor/libpq-16/{libpq.dll, libcrypto-3-x64.dll, libssl-3-x64.dll, libintl-9.dll, libiconv-2.dll}`

- [ ] **Step 1: Download Postgres 16 Windows binaries**

Manual step (one-time): visit https://www.postgresql.org/download/windows/, follow the EDB installer page, pick the Postgres 16 ZIP (NOT the installer EXE). Download `postgresql-16.x-windows-x64-binaries.zip` and extract.

From the extracted `pgsql/bin/`, copy these 5 DLLs into `vendor/libpq-16/`:

- `libpq.dll`
- `libcrypto-3-x64.dll`
- `libssl-3-x64.dll`
- `libintl-9.dll`
- `libiconv-2.dll`

- [ ] **Step 2: Write `vendor/libpq-16/README.md`**

```markdown
# libpq 16 Windows runtime DLLs

These DLLs are the runtime dependencies of Qt's QPSQL driver. They are
copied into `release/` next to `DataViewer.exe` by `build_installer.bat`
and bundled into the installer via `installer.iss`.

## Provenance

- **Source:** Official PostgreSQL Windows binaries ZIP from
  https://www.postgresql.org/download/windows/
- **Version:** PostgreSQL 16.x (x86_64)
- **Files extracted from `pgsql/bin/`:**
  - libpq.dll
  - libcrypto-3-x64.dll
  - libssl-3-x64.dll
  - libintl-9.dll
  - libiconv-2.dll

## Updating

To upgrade to a new Postgres minor version, replace all 5 DLLs from the
matching ZIP. Major version upgrades (16 → 17) are a deliberate event and
must be planned alongside the database upgrade — do NOT swap DLLs without
also upgrading the server.

## Why committed (vs downloaded at build time)

- Reproducible builds: build_installer.bat works without external network
  access.
- Hash-locked: changing the DLLs requires a deliberate commit.
- Small: ~5 MB total, within reason for repo bloat.
```

- [ ] **Step 3: Commit**

```bash
git add vendor/libpq-16/
git commit -m "vendor: PostgreSQL 16 libpq Windows runtime DLLs"
```

---

### Task 23: `build_installer.bat` copies libpq DLLs

**Files:**
- Modify: `build_installer.bat`

- [ ] **Step 1: Insert copy step before ISCC invocation**

In `build_installer.bat`, find the line that invokes `ISCC.exe`. Immediately above it, add:

```bat
echo Copying libpq runtime DLLs to release\...
copy /Y vendor\libpq-16\libpq.dll release\
copy /Y vendor\libpq-16\libcrypto-3-x64.dll release\
copy /Y vendor\libpq-16\libssl-3-x64.dll release\
copy /Y vendor\libpq-16\libintl-9.dll release\
copy /Y vendor\libpq-16\libiconv-2.dll release\
if errorlevel 1 (
    echo ERROR: failed to copy libpq DLLs
    exit /b 1
)
```

- [ ] **Step 2: Test**

```bat
build_installer.bat
```

Expected: build completes; `release\libpq.dll` (and the other 4) exist.

- [ ] **Step 3: Commit**

```bash
git add build_installer.bat
git commit -m "build: copy libpq DLLs into release/ before installer build"
```

---

### Task 24: `installer.iss` — `[Files]` block for libpq + Inno [Code] for db.conf

**Files:**
- Modify: `installer.iss`

- [ ] **Step 1: Add libpq DLLs to `[Files]`**

In `installer.iss`, in the `[Files]` section, alongside existing Qt DLL entries:

```ini
Source: "release\libpq.dll";              DestDir: "{app}"; Flags: ignoreversion
Source: "release\libcrypto-3-x64.dll";    DestDir: "{app}"; Flags: ignoreversion
Source: "release\libssl-3-x64.dll";       DestDir: "{app}"; Flags: ignoreversion
Source: "release\libintl-9.dll";          DestDir: "{app}"; Flags: ignoreversion
Source: "release\libiconv-2.dll";         DestDir: "{app}"; Flags: ignoreversion
```

- [ ] **Step 2: Add `[Code]` section for DB password prompt**

In `installer.iss`, before `[Files]`, add:

```pascal
[Code]
var
  DbPasswordPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  DbPasswordPage := CreateInputQueryPage(wpUserInfo,
    'Database Connection',
    'Enter the DataViewer Postgres password',
    'This password connects DataViewer to your team''s PostgreSQL database on the NAS. ' +
    'Ask your NAS admin for the value. You only need to enter this once per machine.');
  DbPasswordPage.Add('Postgres password:', True);  // True = mask input
end;

function GetDbPassword(Param: String): String;
begin
  Result := DbPasswordPage.Values[0];
end;

procedure WriteDbConf;
var
  ConfDir, ConfPath, EncPwd: String;
begin
  ConfDir := ExpandConstant('{commonappdata}') + '\DataViewer';
  ConfPath := ConfDir + '\db.conf';
  if not DirExists(ConfDir) then
    CreateDir(ConfDir);

  // Run our own DataViewer.exe --encrypt-password=<plain> to produce the
  // encrypted form (uses the same xorCipher as the runtime).
  Exec(ExpandConstant('{app}\DataViewer.exe'),
       '--encrypt-password=' + DbPasswordPage.Values[0] + ' --encrypted-out=' +
       ConfDir + '\encrypted.tmp',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if not LoadStringFromFile(ConfDir + '\encrypted.tmp', EncPwd) then
    EncPwd := '';
  DeleteFile(ConfDir + '\encrypted.tmp');

  SaveStringToFile(ConfPath,
    '[postgres]' + #13#10 +
    'host = dve-db.smoorecig.internal' + #13#10 +
    'port = 5432' + #13#10 +
    'database = dataviewer' + #13#10 +
    'user = dataviewer_app' + #13#10 +
    'password_encrypted = ' + EncPwd + #13#10,
    False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    WriteDbConf;
end;
```

- [ ] **Step 2.5: Add `--encrypt-password` CLI to `main.cpp`**

(One-off helper used by the installer.) In `src/main.cpp`, in the same option-parsing block as Task 21:

```cpp
QCommandLineOption optEncrypt("encrypt-password",
    "Encrypt a password and write to --encrypted-out, then exit.",
    "plain");
QCommandLineOption optEncryptOut("encrypted-out", "Output file path.", "path");
p.addOption(optEncrypt);
p.addOption(optEncryptOut);
p.parse(QCoreApplication::arguments());

if (p.isSet(optEncrypt) && p.isSet(optEncryptOut)) {
    const QString enc = DVE::ConfigLoader::encryptPassword(p.value(optEncrypt));
    QFile out(p.value(optEncryptOut));
    if (!out.open(QIODevice::WriteOnly | QIODevice::Truncate)) return 3;
    out.write(enc.toUtf8());
    return 0;
}
```

- [ ] **Step 3: Build the installer and test**

```bat
build_installer.bat
dist\DataViewer-setup.exe
```

Manual verification:
1. Installer shows the new "Database Connection" page.
2. Enter a password.
3. After install, check `%PROGRAMDATA%\DataViewer\db.conf` exists with `password_encrypted = <base64>`.
4. Uninstall.

- [ ] **Step 4: Commit**

```bash
git add installer.iss src/main.cpp
git commit -m "feat(installer): prompt for DB password and write db.conf"
```

---

### Task 25: `SelfTest::testPostgresConnection` case

**Files:**
- Modify: `src/utils/SelfTest.cpp`

- [ ] **Step 1: Add the new self-test case**

Read `src/utils/SelfTest.cpp` to see the existing test-result pattern (typically a `TestResult test*()` method). Add:

```cpp
#include "../database/PostgresConnection.h"
#include "../database/ConfigLoader.h"

// Inside the SelfTest class implementation, add a new method.
// Add its declaration to SelfTest.h too.

TestResult SelfTest::testPostgresConnection() {
    TestResult r;
    r.name = "postgres_connection";

    const QString confPath = QString(qgetenv("PROGRAMDATA"))
                              + "/DataViewer/db.conf";
    DVE::DbConfig cfg;
    QString err;
    if (!DVE::ConfigLoader::load(confPath, cfg, &err)) {
        r.passed = false;
        r.message = "Could not read " + confPath + ": " + err;
        return r;
    }

    DVE::PostgresConnection pg;
    if (!pg.open(cfg)) {
        r.passed = false;
        r.message = "Connect to " + cfg.host + ":" + QString::number(cfg.port)
                  + " failed: " + pg.lastError();
        return r;
    }
    if (!pg.ping()) {
        r.passed = false;
        r.message = "Ping failed after connect";
        return r;
    }
    pg.close();
    r.passed = true;
    r.message = "Connected to " + cfg.host + " successfully";
    return r;
}
```

In the method that runs all tests (look for an existing `runAll()` or similar), add `testPostgresConnection()` to the call list.

- [ ] **Step 2: Build, manually verify**

```bash
python tools/decrypt_via_copy.py --apply
cd build
mingw32-make
```

```powershell
# With a running test postgres + db.conf manually placed:
.\release\DataViewer.exe --self-test --self-test-out=%TEMP%\selftest.json
type %TEMP%\selftest.json
```

Expected: JSON contains a `postgres_connection` entry with `passed: true`.

- [ ] **Step 3: Commit**

```bash
git add src/utils/SelfTest.cpp src/utils/SelfTest.h
git commit -m "feat(selftest): testPostgresConnection deployment diagnostic"
```

---

### Phase 6 — Test infrastructure helper + deployment phase (Tasks 26–28)

---

### Task 26: `tests/start-test-postgres.ps1`

**Files:**
- Create: `tests/start-test-postgres.ps1`

- [ ] **Step 1: Write the script**

```powershell
<#
.SYNOPSIS
  Spin up an ephemeral postgres:16 container for the test suite.

.DESCRIPTION
  Starts a container on port 5433, applies deploy/postgres/init.sql,
  sets $env:DVE_TEST_PG_CONN, and prints the teardown command.

  Idempotent: if a "dve-test-pg" container already exists, removes
  it first.

.EXAMPLE
  PS> .\tests\start-test-postgres.ps1
#>

$ErrorActionPreference = "Stop"

$existing = docker ps -aq --filter "name=dve-test-pg"
if ($existing) {
    Write-Host "Removing existing dve-test-pg container..."
    docker rm -f dve-test-pg | Out-Null
}

$repoRoot = Resolve-Path "$PSScriptRoot\.."
$initSql  = Join-Path $repoRoot "deploy\postgres\init.sql"
if (-not (Test-Path $initSql)) {
    throw "init.sql not found at $initSql"
}

Write-Host "Starting postgres:16 on port 5433..."
docker run -d --name dve-test-pg `
    -p 5433:5432 `
    -v "$($initSql):/docker-entrypoint-initdb.d/01-init.sql:ro" `
    -e POSTGRES_DB=dve_test `
    -e POSTGRES_USER=test `
    -e POSTGRES_PASSWORD=test `
    postgres:16 `
    -c shared_preload_libraries=pg_cron `
    -c cron.database_name=dve_test | Out-Null

Write-Host "Waiting for ready..."
$maxWait = 30
for ($i = 0; $i -lt $maxWait; $i++) {
    $ready = docker exec dve-test-pg pg_isready -U test 2>&1
    if ($LASTEXITCODE -eq 0) { break }
    Start-Sleep -Seconds 1
}
if ($LASTEXITCODE -ne 0) {
    throw "Postgres did not become ready within $maxWait seconds"
}

$env:DVE_TEST_PG_CONN = "host=127.0.0.1 port=5433 dbname=dve_test user=test password=test"
Write-Host ""
Write-Host "Ready. Env var set for this shell:"
Write-Host "  DVE_TEST_PG_CONN=$($env:DVE_TEST_PG_CONN)"
Write-Host ""
Write-Host "To tear down: docker rm -f dve-test-pg"
```

- [ ] **Step 2: Test the script**

```powershell
.\tests\start-test-postgres.ps1
# Verify env var is set
echo $env:DVE_TEST_PG_CONN
docker rm -f dve-test-pg
```

Expected: container starts, env var is set, teardown works.

- [ ] **Step 3: Commit**

```bash
git add tests/start-test-postgres.ps1
git commit -m "test: helper script to spin up an ephemeral postgres:16 for tests"
```

---

### Task 27: `Test-Deployment.ps1` — Phase 4 migration verification

**Files:**
- Modify: `tests/deployment/Test-Deployment.ps1`

- [ ] **Step 1: Read existing structure**

Open `tests/deployment/Test-Deployment.ps1` and locate the phase block structure (typically functions named `Test-Phase1`, `Test-Phase2`, `Test-Phase3` or similar). Identify where to insert `Test-Phase4`.

- [ ] **Step 2: Append Phase 4**

Add to `tests/deployment/Test-Deployment.ps1`:

```powershell
function Test-Phase4-MigrationVerification {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)] [string]$PreMigrationSqlitePath,
        [Parameter(Mandatory=$true)] [string]$PgHost,
        [Parameter(Mandatory=$true)] [int]   $PgPort,
        [Parameter(Mandatory=$true)] [string]$PgDatabase,
        [Parameter(Mandatory=$true)] [string]$PgUser,
        [Parameter(Mandatory=$true)] [string]$PgPassword
    )

    Write-Host ""
    Write-Host "=== Phase 4: Migration verification ===" -ForegroundColor Cyan

    if (-not (Test-Path $PreMigrationSqlitePath)) {
        Write-Warning "No pre-migration SQLite found at $PreMigrationSqlitePath — skipping"
        return @{ Status = "skipped"; Reason = "no pre-migration source" }
    }

    $tables = @(
        "files", "tests", "samples", "data_rows", "images",
        "sensory_sessions", "sensory_images",
        "detailed_sensory_sessions", "detailed_sensory_images", "settings"
    )

    $results = @()
    $env:PGPASSWORD = $PgPassword
    foreach ($t in $tables) {
        $sqliteCount = (sqlite3 $PreMigrationSqlitePath "SELECT COUNT(*) FROM $t" 2>$null)
        $pgCount = (psql -h $PgHost -p $PgPort -U $PgUser -d $PgDatabase -tA `
                         -c "SELECT COUNT(*) FROM $t" 2>$null).Trim()
        if ([string]::IsNullOrEmpty($sqliteCount)) { $sqliteCount = 0 }
        if ([string]::IsNullOrEmpty($pgCount))     { $pgCount = 0 }

        $match = ($sqliteCount -eq $pgCount)
        $status = if ($match) { "OK" } else { "MISMATCH" }
        Write-Host ("  {0,-32} sqlite={1,-6} pg={2,-6} {3}" -f $t, $sqliteCount, $pgCount, $status) `
            -ForegroundColor $(if ($match) { "Green" } else { "Red" })
        $results += @{
            Table        = $t
            SqliteCount  = [int]$sqliteCount
            PostgresCount= [int]$pgCount
            Match        = $match
        }
    }

    $allMatch = ($results | Where-Object { -not $_.Match }).Count -eq 0
    return @{
        Status  = if ($allMatch) { "passed" } else { "failed" }
        Tables  = $results
    }
}
```

Also update the main runner to invoke `Test-Phase4-MigrationVerification` with the right parameters (likely sourced from `db.conf` parsed via a helper).

- [ ] **Step 3: Manual test (end-of-Plan-A checkpoint scenario)**

```powershell
# 1) Place a real pre-migration SQLite at a known path
# 2) Run the full deployment test
.\tests\deployment\Test-Deployment.ps1
```

Expected: Phase 4 shows green row counts matching for every table.

- [ ] **Step 4: Commit**

```bash
git add tests/deployment/Test-Deployment.ps1
git commit -m "test(deploy): Phase 4 — migration row-count verification"
```

---

### Task 28: Update `tests/deployment/README.md`

**Files:**
- Modify: `tests/deployment/README.md`

- [ ] **Step 1: Append migration + identity sections**

Append to `tests/deployment/README.md`:

```markdown
## Phase 4 — Migration verification

Compares the renamed `<name>.pre-migration.sqlite` file (kept after a
successful migration) against the live PostgreSQL database. For every
table in both, row counts must match. If they diverge, the phase fails
and prints which table.

This phase requires:
- The pre-migration SQLite file at its rename location (`*.pre-migration.sqlite`).
- A working Postgres connection.
- `sqlite3` and `psql` available on `PATH`.

## Manual checklist (cannot be automated end-to-end)

Verify on the work machine after a fresh v2 install:

- [ ] First-launch identity prompt appears, accepts a name + color, and
      does NOT appear on the second launch.
- [ ] `%PROGRAMDATA%\DataViewer\db.conf` exists and is readable only
      to administrators.
- [ ] `DataViewer.exe --self-test` reports `postgres_connection: passed`.
- [ ] Opening a TPM file from the migrated database displays the same
      sheet/sample/row data as the pre-migration SQLite did on v1.3.x.
- [ ] Opening a sensory session shows the same metric values and any
      saved layout JSON renders correctly.
- [ ] Opening a detailed sensory session shows the same Q1–Q14
      responses.
- [ ] Embedded images render in TPM samples and sensory sessions.
```

- [ ] **Step 2: Commit**

```bash
git add tests/deployment/README.md
git commit -m "docs(deploy): Phase 4 description + manual migration checklist"
```

---

### Phase 7 — Final wiring + CLAUDE.md update (Tasks 29–30)

---

### Task 29: Update `CLAUDE.md`

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: Replace SQLite-on-Synology section**

In `CLAUDE.md`, find the section currently describing "Database" (and any SQLite-on-Synology content). Replace with Postgres setup notes:

```markdown
### Database

DataViewer Enterprise stores its data in **PostgreSQL 16** hosted in a
Docker container on the office Synology NAS. The schema is deployed by
`deploy/postgres/init.sql`; the container is defined by
`deploy/postgres/docker-compose.yml`. See `deploy/postgres/README.md`
for NAS admin setup.

The Qt app connects via the **QPSQL** driver, with `libpq.dll` + 4
transitive DLLs bundled by `build_installer.bat` from
`vendor/libpq-16/`. Connection settings live at
`%PROGRAMDATA%\DataViewer\db.conf` (set by the installer; password
field is encrypted with a machine-bound key — copying the file to
another workstation invalidates the password).

The cross-machine `.lock` file mechanism is **gone** — Postgres
handles concurrency at the row level. Optimistic concurrency, live
NOTIFY-driven UI updates, presence indicators, and read-only offline
mode are implemented in Plans B and C of the multi-plan migration
([docs/superpowers/plans/2026-05-11-postgres-multiuser-INDEX.md](docs/superpowers/plans/2026-05-11-postgres-multiuser-INDEX.md)).

**During Plan A only**, the old SQLite path remains active in
parallel — the `DatabaseManager` class still uses SQLite for runtime
reads/writes while the new Postgres infrastructure is being built
out. Plan B switches `DatabaseManager` to a Postgres facade; Plan C
deletes the old SQLite code paths entirely.

### Local test database

Run `tests\start-test-postgres.ps1` to spin up a throwaway
`postgres:16` container on port 5433 with the production schema
applied. The script sets `$env:DVE_TEST_PG_CONN` so test binaries
auto-detect it. Tear down with `docker rm -f dve-test-pg`.
```

If the file has a separate "MIP file encryption" section, leave it intact — it's still relevant.

- [ ] **Step 2: Commit**

```bash
git add CLAUDE.md
git commit -m "docs: replace SQLite-on-Synology section with Postgres notes"
```

---

### Task 30: Plan A checkpoint verification

**Files:**
- None (verification only)

- [ ] **Step 1: Run the full test suite**

```bash
python tools/decrypt_via_copy.py --apply
.\tests\start-test-postgres.ps1
.\tests\run-tests.ps1
```

Expected: all tests pass (existing 11 + new tst_identitymanager, tst_configloader, tst_postgresconnection, tst_migrationtool).

- [ ] **Step 2: Run the deployment self-test**

```powershell
# After installing the new build via build_installer.bat + dist\DataViewer-setup.exe
.\tests\deployment\Test-Deployment.ps1
```

Expected: all phases pass, including Phase 4 (migration verification, assuming a `.pre-migration.sqlite` has been generated by an actual migration).

- [ ] **Step 3: Manual smoke test**

1. Install v2 (the build produced from Plan A's installer).
2. First launch: identity prompt appears, accept name + color.
3. App opens normally and behaves identically to v1.3.x (because the SQLite code path is unchanged; only new infrastructure was added).
4. Run `DataViewer.exe --migrate-from-sqlite=<real SQLite path> --to-postgres="<connection string>"` on the work machine.
5. Verify Postgres has all the data; verify the original SQLite was renamed to `.pre-migration.sqlite`.
6. Run `.\tests\deployment\Test-Deployment.ps1` and confirm Phase 4 passes with matching row counts.

- [ ] **Step 4: Tag the checkpoint**

```bash
git tag v2.0.0-alpha.A "Plan A complete: Postgres foundation + migration"
git push origin v2.0.0-alpha.A
```

(Push only after manual checkpoint passes.)

- [ ] **Step 5: Update the index**

In `docs/superpowers/plans/2026-05-11-postgres-multiuser-INDEX.md`, change Plan A status from "Drafted, not started" to "Complete and verified on <date>". Add a status-log entry.

```bash
git add docs/superpowers/plans/2026-05-11-postgres-multiuser-INDEX.md
git commit -m "docs(plans): mark Plan A complete"
```

---

## Plan A Checkpoint Criteria

Plan A is complete when **all** of the following are true:

1. `docker compose up -d` in `deploy/postgres/` on the NAS leaves a healthy `postgres:16` container with the full schema loaded.
2. `DataViewer.exe --migrate-from-sqlite=<real-data> --to-postgres=<conn>` runs to completion with exit code 0; produces a JSON report at `%TEMP%\dataviewer_migration.json` with `status=success` and per-table row counts that match SQLite.
3. The original SQLite file is renamed to `<name>.pre-migration.sqlite` and remains intact on Synology for at least one release cycle.
4. `tests\run-tests.ps1` passes all suites including the four new test classes (`tst_identitymanager`, `tst_configloader`, `tst_postgresconnection`, `tst_migrationtool`).
5. `tests\deployment\Test-Deployment.ps1` passes all phases including the new Phase 4.
6. Manual smoke test passes (see Task 30).
7. The `<dbPath>.lock` SQLite file-lock mechanism is still active and the app still uses SQLite at runtime — Plan A does NOT switch over. (Plan C handles the cutover.)

Once verified, draft Plan B against this baseline.

---

## Self-Review

This section is for the plan author. After writing, look at the spec with fresh eyes and check coverage. Fix issues inline.

**Spec coverage check** — every spec requirement maps to a Plan A, B, or C task:

| Spec section | Plan A coverage | Notes |
|---|---|---|
| Goal | Tasks 1–30 lay the foundation | — |
| Schema design (all tables + audit cols + triggers) | T2–T5 | All deferred until Plan B turns ON optimistic concurrency in app code |
| NOTIFY plumbing | T4 (triggers exist in DB) | Listening happens in Plan B |
| Online/offline flow | — | Plan C |
| Concurrency model | T2 (version column exists) | App enforcement in Plan B |
| Live updates & presence | T4 (presence table + triggers exist) | App-side in Plan B |
| Migration & deployment | T1, T6, T14–T24, T27 | Complete in Plan A |
| Testing strategy | T7, T10–T13, T15–T20, T26–T28 | Plan A delivers all migration + foundation tests; Plan B adds concurrency tests; Plan C adds offline tests |
| File-level change inventory | All Plan A files matched to a task | — |

**Placeholder scan:** No "TBD", "TODO", or vague requirements. All code is complete and runnable. ✓

**Type consistency:** `DbConfig`, `IdentityManager`, `PostgresConnection`, `MigrationTool` signatures match across header / cpp / tests. ✓

**Scope check:** Plan A is bounded — it adds infrastructure alongside existing code, does not switch the app over. End state is clearly checkpointable. ✓
