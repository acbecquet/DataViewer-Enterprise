# DataViewer Enterprise — Phone Sensory Collection Service (`sensory-collect`)

A standalone, mobile-first web form (Flask + gunicorn, Docker) that lets an
offsite tester submit a single 5-metric **Sensory** sample from a phone. On
submit the sample is **created-or-appended by natural key** into the same
PostgreSQL `sensory_sessions` schema the desktop app uses, so every desktop sees
it live via `NOTIFY`.

This service is on an **independent versioning track** — it is NOT part of a
`DataViewer.exe` release and is NOT wrapped into the v2.6.0 desktop sprint.
Implements DATAVIEWER-11 (MS-8). Design source of truth:
`docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md`.

## What's in this directory

- `app.py` — one GET (the form) + one POST (`/submit`): validates input, then
  calls the `dve_append_sensory_sample` stored function as the least-privilege
  `sensory_web` role. It never calls the generic `dve_commit_cell*` mutators.
- `templates/form.html` — the mobile form (vanilla JS `fetch`, no framework).
- `requirements.txt` — Flask, gunicorn, psycopg2-binary, Flask-Limiter.
- `Dockerfile` — `python:3.11-slim`, gunicorn on port 8000.
- `docker-compose.yml` — the service joined to the **external** `dve-net`
  Docker network so it reaches the DB by service name `dataviewer-db:5432`.
- `.env.example` — `DVE_SENSORY_WEB_PASSWORD=change-me` placeholder. The real
  secret lives ONLY in the NAS `.env` (this repo is PUBLIC — never commit it).
- `tests/test_append.py` — `pytest` for the create-or-append path, idempotency,
  least privilege, and the HTTP validation gates. Runs against an ephemeral DB;
  skips cleanly when no test DB is configured.
- `tests/test_mfused_form.py` — DB-free `pytest` that renders the form and
  asserts the Mfused host relabels Round as Mode (A/B/C -> round 1/2/3) while the
  default host is unchanged.

## Mfused customer variant (host-gated)

One offsite customer is served a variant of this same form, selected purely by the **Host** header (`DVE_MFUSED_HOSTS`, default `mfused-sensory.ccell-sdr.com`).
The variant swaps the free-text header inputs for fixed dropdowns whose option lists are hard-coded in `app.py` (`MFUSED_OPTIONS` / `MFUSED_MODE_CHOICES`) and rendered straight into the page.
Those lists never query the database, so the anonymous endpoint still cannot enumerate the test catalog.
The POST path and the `dve_append_sensory_sample` call are identical to the default form.

The variant relabels **Round** as **Mode** with letter choices **A / B / C**.
Each `<option>` carries the round number it maps to, so the browser POSTs (and the DB stores) the same `round` values as always:

| Mode (shown) | `round` posted | Resulting session `tester_name` |
| ------------ | -------------- | ------------------------------- |
| A            | `1`            | `<tester> R1`                   |
| B            | `2`            | `<tester> R2`                   |
| C            | `3`            | `<tester>` (no suffix, matching existing round-3 behavior) |

This is a temporary per-customer relabel and is presentational only: it changes no backend code, no `/submit` handling, and no schema.

## Security model (no user-facing auth)

The endpoint is **anonymous and frictionless by design** (owner decision
2026-06-24): a one-time offsite customer submits with no password or token. All
control is backend and invisible:

- The service connects as **`sensory_web`**, a least-privilege role with only
  `SELECT, INSERT, UPDATE` on `sensory_sessions` and `EXECUTE` on
  `dve_append_sensory_sample`. It is denied every other table and the generic
  `dve_commit_cell*` mutators. A service bug cannot escalate beyond appending
  sensory rows.
- Strict server-side validation (required Test Title + Tester; scores must be
  JSON numbers in [1,9]; the SQL function re-validates as defense in depth).
- An invisible per-IP rate limit (Flask-Limiter) plus a 64 KiB request-size cap.

**Accepted residual risk:** an open anonymous endpoint can post junk sensory
rows. Bounded by the least-privilege role, deletable from any desktop. This is
the deliberate trade for a frictionless one-time-customer UX.

---

## One-time NAS admin setup

> Prereqs: the `dataviewer-db` Postgres stack is already running on the NAS
> (`deploy/postgres/`), and the Phase-2 DB migrations have been applied to it:
> `2026-06-25-dv11-append-sensory-sample.sql` and
> `2026-06-25-dv11-sensory-web-role.sql` (creates the `dve_append_sensory_sample`
> function and the `sensory_web` role — with NO password set).

### 1. Create the shared Docker network (once)

Both stacks must share a user-defined bridge network so the web service can
reach the DB by container name. The committed compose files declare `dve-net`
as **external**, so create it once before bringing either stack up:

```bash
docker network create dve-net
```

If `dataviewer-db` was already running on the default network, recreate it so it
joins `dve-net` (see step 4) — `docker compose up -d` re-reads the updated
compose and reconnects the container without data loss (the data lives in the
host bind-mount, not the container).

### 2. Set the `sensory_web` password secret

The migration creates `sensory_web` with **no password**. Set one out-of-band
from a NAS-only secret. First put the secret in the NAS `.env` (gitignored,
never committed):

```bash
cd /volume1/docker/dataviewer-sensory-collect    # wherever you place this stack
cp .env.example .env
# Edit .env, set DVE_SENSORY_WEB_PASSWORD to a strong random value.
chmod 600 .env
```

Then apply that same value to the role in Postgres (run once, after the role
migration). Read it from the `.env` so the literal never appears in shell
history:

```bash
set -a; . ./.env; set +a    # loads DVE_SENSORY_WEB_PASSWORD into the env

docker compose -f /volume1/docker/dataviewer-db/docker-compose.yml \
  exec -T dataviewer-db \
  psql -U dataviewer_app -d dataviewer \
  -c "ALTER ROLE sensory_web LOGIN PASSWORD '$DVE_SENSORY_WEB_PASSWORD';"
```

(The desktop app and `pytest` set the same password the same way; the value is
known only to the NAS `.env` and the role.)

### 3. Copy this stack to the NAS

```bash
# From a workstation with SSH access:
scp -r deploy/sensory-collect \
    admin@nas:/volume1/docker/dataviewer-sensory-collect/
# Then create the production .env on the NAS as in step 2 (do NOT scp a .env).
```

### 4. Bring up the stack

```bash
cd /volume1/docker/dataviewer-sensory-collect
docker compose up -d --build
docker compose logs -f sensory-collect      # watch for the gunicorn boot line
```

`docker-compose.yml` joins the service to the external `dve-net` and points it
at `dataviewer-db:5432` (the DB container's **internal** listening port). Note
the host/prod port `5433` is for host-side tools (the desktop app) only — the
co-located service never uses it.

If `dataviewer-db`'s own compose did not yet declare `dve-net`, it has now been
updated to (see `deploy/postgres/docker-compose.yml`); re-run its
`docker compose up -d` so it joins the network too.

### 5. Expose over the internet via the existing reverse proxy + TLS

Phones reach the service from anywhere through the **existing NAS reverse proxy +
TLS** — the same precedent as the `bug-form-app` that files Plane reports. This
proxy config is **NAS-only and not in this repo**.

> **Confirm with the owner** before wiring the route:
> - the public hostname / route to map to this service (e.g.
>   `https://sensory.<your-domain>/` → `sensory-collect:8000`);
> - that TLS terminates at the proxy (DSM Reverse Proxy or the existing
>   Nginx/Traefik front end), as it does for `bug-form-app`.

In DSM: **Control Panel → Login Portal → Advanced → Reverse Proxy → Create**,
source = the chosen HTTPS hostname, destination = `localhost:8000` (or the
container, if the proxy shares `dve-net`). Keep the office firewall rule on the
DB port (5433) unchanged — only the proxy is internet-facing, never Postgres.

### 6. Verify a phone OFF the office network can submit

1. On a phone using **cellular data** (not office Wi-Fi), open the public HTTPS
   URL. The form should load over TLS.
2. Fill Test Title + Tester + the five 1–9 scores, tap **Submit sample**.
3. Confirm the success response, then on a desktop confirm the session appears
   (live via `NOTIFY`, or after a refresh). A second submit with the same Test
   Title + Tester + Round appends a second sample to the SAME session row; a
   different header creates a new row.
4. Re-submitting the same sample (same `sample_uid`, e.g. a retried tap) does
   NOT duplicate.

---

## Running the tests

`tests/test_append.py` needs an ephemeral PostgreSQL 16 with `init.sql` + the
Phase-2 migrations applied, reachable as a **superuser** (it runs DDL +
`ALTER ROLE`). On the work machine the simplest path reuses the project's test
container:

```powershell
pip install pytest psycopg2-binary "Flask==3.0.*" "Flask-Limiter==3.*"
.\tests\start-test-postgres.ps1            # exports DVE_TEST_PG_CONN, applies migrations
pytest deploy/sensory-collect/tests/test_append.py -v
docker rm -f dve-test-pg                    # teardown
```

The suite reads `DVE_TEST_PG_CONN` (a libpq DSN) or the discrete
`DVE_DB_HOST` / `DVE_DB_PORT` / `DVE_DB_NAME` / `DVE_DB_USER` / `DVE_DB_PASSWORD`
vars, and **skips cleanly** when none is set. It pins a throwaway password on
`sensory_web` in the test DB to exercise the least-privilege login.

## Routine operations

- **Logs:** `docker compose logs -f sensory-collect`
- **Restart:** `docker compose restart sensory-collect`
- **Rebuild after a code change:** `docker compose up -d --build`
- **Rotate the `sensory_web` password:** update the NAS `.env`, re-run the
  `ALTER ROLE ... PASSWORD` from step 2, then `docker compose up -d` to restart
  the service with the new secret.
