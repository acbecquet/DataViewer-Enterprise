# DataViewer Enterprise — PostgreSQL Setup (Synology NAS)

This directory contains the database container definition and bootstrap
SQL for DataViewer Enterprise's Postgres backend.

## What's in this directory

- `docker-compose.yml` — service definition; builds the local Dockerfile and runs the resulting image.
- `Dockerfile` — derives from `postgres:16` and adds `postgresql-16-cron`. The base image does not bundle pg_cron, so this layer is required.
- `init.sql` — schema, audit columns, version-bump trigger, NOTIFY triggers, and the `pg_cron` stale-presence cleanup job. Runs on first container start via `/docker-entrypoint-initdb.d/`.
- `.env.example` — template for `DVE_DB_PASSWORD`. Copy to `.env` (gitignored) and set a real password.

## One-time NAS admin setup

1. **Prerequisites on the NAS:**
   - DSM 7.2+ with Container Manager installed.
   - A static IP or internal DNS name (e.g. `dve-db.smoorecig.internal`).
   - Storage volume with at least 5 GB free.

2. **Copy these four files to `/volume1/docker/dataviewer-db/` on the NAS:**

   ```bash
   # From a workstation with SSH access:
   scp deploy/postgres/docker-compose.yml \
       deploy/postgres/Dockerfile \
       deploy/postgres/init.sql \
       deploy/postgres/.env.example \
       admin@nas:/volume1/docker/dataviewer-db/
   ```

3. **Create the production `.env` ON THE NAS** (do not commit):

   ```bash
   cd /volume1/docker/dataviewer-db
   cp .env.example .env
   # Edit .env, set DVE_DB_PASSWORD to a strong random value.
   chmod 600 .env
   ```

4. **Update the bind-mount path in `docker-compose.yml` for production:**

   On the NAS only, change the data volume from `./data` to `/volume1/docker/dataviewer-db/data`:

   ```yaml
   volumes:
     - /volume1/docker/dataviewer-db/data:/var/lib/postgresql/data
     - ./init.sql:/docker-entrypoint-initdb.d/01-init.sql:ro
   ```

   The committed compose file uses the relative `./data` for local-dev convenience; production overrides it in-place.

5. **Build and start the container:**

   ```bash
   docker compose up -d --build       # builds dve-postgres:16 (postgres:16 + pg_cron), then starts
   docker compose logs -f dataviewer-db   # watch for "ready to accept connections"
   docker compose exec dataviewer-db pg_isready -U dataviewer_app -d dataviewer
   ```

   The `--build` flag forces a fresh build on first run. Subsequent `docker compose up -d` calls reuse the image unless the Dockerfile changes.

6. **Verify the schema and cron job loaded:**

   ```bash
   docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c "\dt"
   # → should list 12 tables (files, tests, samples, data_rows, images,
   #   sensory_sessions, sensory_images, detailed_sensory_sessions,
   #   detailed_sensory_images, settings, presence, schema_meta)

   docker compose exec dataviewer-db psql -U dataviewer_app -d dataviewer -c \
     "SELECT jobname, schedule FROM cron.job"
   # → should print: dve_presence_cleanup | * * * * *
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
- **Rebuild after Dockerfile change:** `docker compose up -d --build`
- **Upgrade Postgres minor version:** edit the `FROM postgres:16` line in the Dockerfile (e.g., to `postgres:16.5`), then `docker compose up -d --build`. Major version upgrades (16 → 17) are deliberate events; see the spec for the planned migration process.
- **Backup snapshot (in addition to Hyper Backup):**
  ```bash
  docker compose exec dataviewer-db pg_dump -U dataviewer_app dataviewer \
    > /volume1/backups/dve-$(date +%F).sql
  ```

## Important: clearing the data directory

`docker compose down -v` removes Compose-managed volumes but does **not** delete the host-path bind mount (`./data` or `/volume1/docker/dataviewer-db/data`). If you bring the stack down with `-v` and then `up -d` again, Postgres will see the existing data directory and **skip the init.sql** — your schema changes won't apply.

For a full clean re-init (e.g., recovering from a botched first run during dev):

```bash
docker compose down
sudo rm -rf ./data            # local dev path; on Synology: /volume1/docker/dataviewer-db/data
docker compose up -d --build
```

On Synology, you may need to remove the directory via DSM File Station rather than the shell, depending on permissions.

## Disaster recovery

If the Postgres data is lost or corrupted:

1. Restore the `/volume1/docker/dataviewer-db/data` directory from Hyper Backup.
2. `docker compose up -d` — Postgres reads the restored data on startup.
3. Verify with the schema and cron-job checks above.

If recovery from Postgres backup is not possible and the migration is recent:

1. See `docs/superpowers/specs/2026-05-11-postgres-multiuser-design.md`
   "Rollback path" section for the SQLite escape hatch (the
   `<name>.pre-migration.sqlite` file kept on Synology).
