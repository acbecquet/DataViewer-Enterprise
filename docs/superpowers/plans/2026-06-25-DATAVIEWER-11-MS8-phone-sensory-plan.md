# MS-8 — Remote Phone Sensory Web Form → Postgres — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.
>
> **AUTHORITATIVE DESIGN:** `docs/superpowers/specs/2026-06-24-DATAVIEWER-11-remote-phone-sensory-spike.md` (owner-signed §10, 2026-06-24). This plan operationalizes it. Where they disagree, the spike wins — update it first.

**Goal:** Let an offsite tester submit a 5-metric Sensory sample from a phone; on Submit it is **created-or-appended by natural key** into the same Postgres `sensory_sessions` schema the desktop uses, so every desktop sees it live via NOTIFY — with the desktop made data-safe against the phone's appends.

**Architecture:** Two independent versioning tracks. (1) **Desktop (rides v2.6.0):** a surgical merge-safety fix so a desktop save of an open session can't drop a phone-appended sample, a stable per-sample `sample_uid`, and the DV-15 `session_name` unification. (2) **NAS track (independent of `DataViewer.exe`):** a new DB migration (`dve_append_sensory_sample` stored function + a least-privilege `sensory_web` role) and a standalone Flask Docker stack at `deploy/sensory-collect/` that validates input and calls the stored function — never the generic `dve_commit_cell*`.

**Tech Stack:** Desktop — C++17 / Qt 6.10, qmake + MinGW, `-Werror`. NAS — PostgreSQL 16 (plpgsql migration), Python 3.11 Flask + gunicorn + psycopg2, Docker Compose. Deadline ~early July 2026 (Isabel).

---

## Owner decisions (locked 2026-06-24, sub-spec §10)

- Reachable **from anywhere** (internet + reverse proxy + TLS). Office-Wi-Fi-only **rejected**.
- **No user-facing auth — anonymous & frictionless.** All control is backend: least-privilege `sensory_web` role + strict server-side validation + invisible per-IP rate-limit.
- **Online-only** (no offline/PWA).
- **Create-or-append by natural key** — a sample whose headers match an existing session is appended, not forked/rejected.
- **Do the §4.1 desktop merge fix in the v2.6.0 train** (a) — DV-11 is **not** zero-C++.
- **DV-15** (unify the 3 `session_name` derivations) created as a **prerequisite** (d).
- Scores are **JSON numbers in [1,9]** — never strings (DATAVIEWER-4 data-loss class).

## CRITICAL: the natural key is NOT "title - tester - R#"

Grounded (sub-spec §2, re-verified 2026-06-25): the unique index is `idx_sensory_sessions_key ON sensory_sessions(session_name, tester_name, date)` — three raw TEXT columns. On the **live** path `session_name` = **Test Title alone**; round lives in `tester_name` as a trailing `" R1"`/`" R2"`; `date` is a local-`yyyy-MM-dd` TEXT string. The web service must therefore **resolve the row by the three key columns and append**, creating only on a true miss — and pin `date` to the **office TZ server-side**.

## Grounded anchors (verified 2026-06-25 via `git show d7ca362`; spike line numbers had drifted)

| Thing | Location |
|---|---|
| `mergeSensoryPreservingDbScores` (the drop-the-tail bug) | `src/pipeline/SensoryData.cpp:103-126`; called `DatabaseManager.cpp:1783` |
| `applyMergedScoresToCurrentSession` (overlay + `applySession` rebuild) | `src/ui/SensoryPanel.cpp:1434-1448` |
| `sensorySessionToJson`/`FromJson` (tolerant; ignores unknown keys) | `src/pipeline/SensoryData.cpp:10-101` |
| 5 metric keys (`kSensoryMetrics`) | `src/pipeline/SensoryData.h:~48-50` |
| `session_name` live-save = title-only | `src/ui/SensoryPanel.cpp:907` |
| `combineTesterRound` (" R1"/" R2"); split regex `^(.*\S)\s+R([12])$` | `src/ui/TesterRound.h:31-37` (+ split `:18`) |
| `session_name` Excel-import = `title + " - " + tester` | `src/ui/SensoryPanel.cpp:2192` |
| `session_name` panelist-import = `title (+ " - " + tester)` | `src/ui/SensoryPanel.cpp:2061` |
| Detailed live-save = `testTitle` (already title-only) | `src/ui/DetailedSensoryPanel.cpp:1245` |
| INSERT branch of `tryWriteSensoryCore` (column list, id=-1, updated_by) | `src/database/DatabaseManager.cpp:1859-1889` |
| Table + natural-key index | `deploy/postgres/init.sql:129-146` |
| `bump_version` / `notify_row_change` triggers (auto version + NOTIFY) | `deploy/postgres/init.sql:258-266`, `322-344`, attach `357-366` |
| `dve_commit_cell` / `dve_commit_cell_json` (generic mutators) | `deploy/postgres/init.sql:426-474` |
| **Zero** `CREATE ROLE`/`GRANT`/`REVOKE` anywhere | `init.sql` + all `migrations/*.sql` (confirmed) |
| Compose (no `networks:` block, host port 5432, `.env.example`) | `deploy/postgres/docker-compose.yml`, `.env.example` |
| Sensory round-trip/merge tests | `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp:37,52-63` |
| e2e two-writer/reconnect template | `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp:754` (`scenario9_…`), `:830` (`scenario9b_…`) |

> **MIP / build:** `python tools/decrypt_via_copy.py --apply` before every C++ build; new C++ via the Python delete-and-rewrite convention. `.sql/.py/.yml/.html` web-stack files: create via Python delete-and-rewrite too if MIP labels them (`.py` is MIP-sensitive on this machine). Desktop changes gated under the **MS-7 heavy-verification discipline** (whole-session save is load-bearing) — TDD, e2e proof, `/ponytail-review`.

---

# PHASE 1 — Desktop (C++, rides v2.6.0)

> Gate: the desktop is the load-bearing whole-session save path. Each task writes its failing test first and proves it green. `-Werror` throughout.

### Task 1.1: Add a stable per-sample `sample_uid`

**Files:** `src/pipeline/SensoryData.h`/`.cpp`, `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp`.

- [ ] **Step 1: Failing round-trip test** — add to `tst_sensorydataplaceholder.cpp`:

```cpp
void sampleUidSurvivesJsonRoundTrip();
// ...
void TstSensoryDataPlaceholder::sampleUidSurvivesJsonRoundTrip() {
    SensorySession s; SensorySample smp; smp.sampleUid = "uid-abc-123";
    s.samples.append(smp);
    const SensorySession back = sensorySessionFromJson(sensorySessionToJson(s));
    QCOMPARE(back.samples.size(), 1);
    QCOMPARE(back.samples[0].sampleUid, QString("uid-abc-123"));
}
```

Run → FAIL (`sampleUid` undefined).

- [ ] **Step 2: Add the field + serialize** — in `SensoryData.h` add `QString sampleUid;` to `SensorySample`. In `sensorySessionToJson` (per-sample object build, `SensoryData.cpp:~38`) add `obj["sample_uid"] = smp.sampleUid;` **only when non-empty** (keep the contract clean for desktop-authored samples). In `sensorySessionFromJson` (per-sample read, `:~60-90`) add `smp.sampleUid = sObj.value("sample_uid").toString();` (tolerant — absent ⇒ empty). Run → PASS.

- [ ] **Step 3: Commit** — `feat(sensory): add per-sample sample_uid for web-append idempotency (DATAVIEWER-11)` (+ Co-Authored-By).

### Task 1.2: Merge-safety fix — preserve DB-only tail samples

**Files:** `src/pipeline/SensoryData.cpp`, `tests/tst_sensorydataplaceholder/tst_sensorydataplaceholder.cpp`.

- [ ] **Step 1: Failing unit test**:

```cpp
void mergeSensory_appendsDbTailSamplesWhenMoreInDb();
// ...
void TstSensoryDataPlaceholder::mergeSensory_appendsDbTailSamplesWhenMoreInDb() {
    QJsonObject mem;  QJsonArray memArr;
    memArr.append(QJsonObject{{"name","S1"},{"Burnt Taste",4.0}});
    mem["samples"] = memArr;
    QJsonObject db;   QJsonArray dbArr;
    dbArr.append(QJsonObject{{"name","S1"},{"Burnt Taste",7.0}});       // untouched → DB wins
    dbArr.append(QJsonObject{{"name","S2-phone"},{"Burnt Taste",6.0}}); // DB-only tail
    db["samples"] = dbArr;
    const QJsonObject merged = mergeSensoryPreservingDbScores(mem, db, /*dirty=*/{});
    const QJsonArray out = merged["samples"].toArray();
    QCOMPARE(out.size(), 2);
    QCOMPARE(out[1].toObject().value("name").toString(), QString("S2-phone"));
    QCOMPARE(out[0].toObject().value("Burnt Taste").toDouble(), 7.0);
}
```

Run → FAIL (returns 1 sample today).

- [ ] **Step 2: Fix** — in `mergeSensoryPreservingDbScores` (`SensoryData.cpp:103-126`), between the overlap loop and `merged["samples"] = memSamples;` insert:

```cpp
    // DATAVIEWER-11: keep samples that exist only in the DB (e.g. appended by
    // the phone web form while this session was open). The desktop never had
    // them in memory, so there is no local edit to reconcile — copy verbatim.
    for (int i = memSamples.size(); i < dbSamples.size(); ++i)
        memSamples.append(dbSamples.at(i));
```

Run → PASS. Verify the existing `merge_*` tests still pass (the change is append-only past the overlap).

- [ ] **Step 3: Commit** — `fix(sensory): merge preserves DB-only tail samples (phone appends) (DATAVIEWER-11)`.

### Task 1.3: Apply side — grow the card list + e2e regression

**Files:** `src/ui/SensoryPanel.cpp`, `tests/tst_saveintegrity_e2e/tst_saveintegrity_e2e.cpp`.

- [ ] **Step 1: Grow on apply** — in `applyMergedScoresToCurrentSession` (`SensoryPanel.cpp:1434-1448`), after `overlayMergedScores(sess, mergedSession);` and before `applySession(sess);`:

```cpp
    // DATAVIEWER-11: overlayMergedScores copies scores, not samples. If the
    // merged blob carries DB-only samples (phone appends) the desktop never had,
    // grow the in-memory struct so applySession() rebuilds cards for them.
    const QJsonArray mergedSamples = mergedSession.value("samples").toArray();
    if (mergedSamples.size() > sess.samples.size()) {
        const SensorySession parsed = sensorySessionFromJson(mergedSession);
        for (int i = sess.samples.size(); i < parsed.samples.size(); ++i)
            sess.samples.append(parsed.samples.at(i));
    }
```

- [ ] **Step 2: e2e regression** — in `tst_saveintegrity_e2e.cpp`, mirroring `scenario9_…` (`:754`), add `scenario10_phoneAppendSurvivesDesktopSave()`: seed a sensory row (id>0) with N samples; open it in the panel (in-memory N); via a direct SQL `UPDATE … jsonb_set` (or `dve_append_sensory_sample` once Phase 2 lands) append an (N+1)th sample to the DB row; trigger a whole-session save (`onUpdateDatabase` flush=true); reload the row; **assert the reloaded session has N+1 samples and the tail sample survived**. This FAILS on pre-Task-1.2 code and PASSES after 1.2 + 1.3. Register it in the slots list. (DB-dependent: skips cleanly when `DVE_TEST_PG_CONN` is unset.)

- [ ] **Step 3:** Build `-Werror`; run `tst_sensorydataplaceholder` + `tst_saveintegrity_e2e` (with a test DB up). **Commit** `fix(sensory): grow cards for DB-only samples + e2e (DATAVIEWER-11)`.

### Task 1.4: DV-15 — unify the 3 `session_name` derivations (code)

**Files:** `src/ui/TesterRound.h` (or `SensoryData.h`), `src/ui/SensoryPanel.cpp`, `tests/tst_sensorydataplaceholder/` (helper test).

- [ ] **Step 1: Failing helper test** — assert the canonical derivation is title-only + trimmed:

```cpp
void canonicalSensorySessionNameIsTitleOnly();
// ...
void TstSensoryDataPlaceholder::canonicalSensorySessionNameIsTitleOnly() {
    QCOMPARE(canonicalSensorySessionName("  Mango v2  "), QString("Mango v2"));
}
```

- [ ] **Step 2: Add the helper** — in `TesterRound.h` (it already owns tester/round string logic): `inline QString canonicalSensorySessionName(const QString& testTitle) { return testTitle.trimmed(); }`. Run → PASS.

- [ ] **Step 3: Route all three sites through it** — replace:
  - live-save `SensoryPanel.cpp:907` `sess.sessionName = testTitle;` → `sess.sessionName = canonicalSensorySessionName(testTitle);`
  - Excel-import `:2192` `sess.sessionName = testTitle + " - " + testerName;` → `sess.sessionName = canonicalSensorySessionName(testTitle);`
  - panelist-import `:2061` `sess.sessionName = testTitle + (testerName.isEmpty() ? "" : " - " + testerName);` → `sess.sessionName = canonicalSensorySessionName(testTitle);`

- [ ] **Step 4: Audit Detailed** — confirm `DetailedSensoryPanel.cpp:1245` is already `sess.sessionName = sess.testTitle;` (title-only). If any Detailed import path diverges, route it through the helper too; otherwise add a one-line note that Detailed is already canonical.

- [ ] **Step 5:** Build `-Werror`; full suite green. `/ponytail-review` (one helper, three call sites — no new abstraction). **Commit** `fix(sensory): unify session_name on title-only across save+imports — DV-15 (DATAVIEWER-15)`. Append `tasks/lessons.md`: the 3 divergent `session_name` derivations forked imported sessions on re-save; canonical = title-only.

> **Note:** the *data* re-key of already-forked rows is **Task 2.3** (DB, owner-gated). Code unification (this task) is safe on its own; new saves are consistent immediately.

---

# PHASE 2 — DB migration (NAS track, independent of v2.6.0)

> New files under `deploy/postgres/migrations/`. Validate against an ephemeral DB via `tests\start-test-postgres.ps1` (auto-applies migrations). Repo is PUBLIC — never commit the `sensory_web` password.

### Task 2.1: `dve_append_sensory_sample` stored function

**Files:** create `deploy/postgres/migrations/2026-06-25-dv11-append-sensory-sample.sql`.

- [ ] **Step 1: Write the migration** (create via Python delete-and-rewrite if `.sql` shows MIP ciphertext):

```sql
-- DATAVIEWER-11: create-or-append a single sensory sample by natural key.
-- Called ONLY by the sensory_web role from the phone web form. Hard-codes the
-- sensory_sessions table (it is NOT the generic dve_commit_cell* primitive).
CREATE OR REPLACE FUNCTION dve_append_sensory_sample(
    p_session_name TEXT,   -- trimmed Test Title ONLY (canonical key)
    p_tester       TEXT,   -- trimmed tester, no round suffix
    p_round        TEXT,   -- '1' | '2' | anything-else => no suffix
    p_assessor     TEXT,
    p_media        TEXT,
    p_sample       JSONB,  -- one sample object: numeric scores [1,9] + sample_uid
    p_office_tz    TEXT DEFAULT 'America/New_York'
) RETURNS BIGINT
LANGUAGE plpgsql AS $$
DECLARE
    v_tester TEXT := btrim(coalesce(p_tester, ''));
    v_date   TEXT;
    v_id     BIGINT;
    v_uid    TEXT := p_sample->>'sample_uid';
    v_metric TEXT;
BEGIN
    -- Defense in depth: scores must be JSON numbers in [1,9] (string scores are
    -- the DATAVIEWER-4 reset-to-5 data-loss class). The service validates too.
    FOREACH v_metric IN ARRAY ARRAY['Burnt Taste','Vapor Volume','Overall Flavor',
                                    'Smoothness','Overall Liking'] LOOP
        IF jsonb_typeof(p_sample->v_metric) IS DISTINCT FROM 'number'
           OR (p_sample->>v_metric)::numeric < 1
           OR (p_sample->>v_metric)::numeric > 9 THEN
            RAISE EXCEPTION 'sensory score "%" must be a JSON number in [1,9]', v_metric;
        END IF;
    END LOOP;

    -- Key columns, computed exactly as the desktop live path does.
    IF p_round = '1' THEN v_tester := v_tester || ' R1';
    ELSIF p_round = '2' THEN v_tester := v_tester || ' R2';
    END IF;
    v_date := to_char((now() AT TIME ZONE p_office_tz)::date, 'YYYY-MM-DD');  -- office-TZ date

    -- Resolve-or-create, then lock the row (serializes concurrent appends).
    INSERT INTO sensory_sessions(session_name, tester_name, assessor_name, media,
                                 date, timestamp, json_data, updated_by)
    VALUES (btrim(p_session_name), v_tester, p_assessor, p_media, v_date,
            to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS'),
            '{"samples":[]}'::jsonb, 'web/' || v_tester)
    ON CONFLICT (session_name, tester_name, date) DO NOTHING;

    SELECT id INTO v_id FROM sensory_sessions
     WHERE session_name = btrim(p_session_name)
       AND tester_name  = v_tester
       AND date         = v_date
     FOR UPDATE;

    -- Idempotency: a retried POST with the same sample_uid is a no-op.
    IF v_uid IS NOT NULL THEN
        PERFORM 1 FROM sensory_sessions s,
                      jsonb_array_elements(coalesce(s.json_data->'samples','[]'::jsonb)) e
         WHERE s.id = v_id AND e->>'sample_uid' = v_uid;
        IF FOUND THEN RETURN v_id; END IF;
    END IF;

    -- Tail-append only (never reorder/remove existing elements).
    UPDATE sensory_sessions
       SET json_data = jsonb_set(json_data, '{samples}',
                                 coalesce(json_data->'samples','[]'::jsonb) || p_sample)
     WHERE id = v_id;   -- bump_version + notify_row_change fire automatically

    RETURN v_id;
END;
$$;
```

- [ ] **Step 2: Verify** — apply via `tests\start-test-postgres.ps1`; in `psql`: call it twice with the same key+different `sample_uid` → one row, two samples; call again with a duplicate `sample_uid` → still two; a string score → `RAISE EXCEPTION`; confirm `version` bumped and a `NOTIFY dataviewer_changes` fired (`LISTEN` in a second session). Add these as a `pytest` in Phase 3's suite (Task 3.3) since they exercise the same function the service calls.

- [ ] **Step 3: Commit** — `feat(db): dve_append_sensory_sample create-or-append by natural key (DATAVIEWER-11)`.

### Task 2.2: Least-privilege `sensory_web` role

**Files:** create `deploy/postgres/migrations/2026-06-25-dv11-sensory-web-role.sql`; update `deploy/postgres/README.md` + `deploy/postgres/.env.example`.

- [ ] **Step 1: Write the migration** (NO password in SQL — repo is public):

```sql
-- DATAVIEWER-11: dedicated least-privilege role for the phone web form.
-- The password is set out-of-band on the NAS from a .env secret (never committed).
DO $$ BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'sensory_web') THEN
        CREATE ROLE sensory_web LOGIN;
    END IF;
END $$;

-- Strip everything, then grant ONLY what the form needs.
REVOKE ALL ON ALL TABLES    IN SCHEMA public FROM sensory_web;
REVOKE ALL ON ALL FUNCTIONS IN SCHEMA public FROM sensory_web;
REVOKE ALL ON ALL SEQUENCES IN SCHEMA public FROM sensory_web;
GRANT  USAGE ON SCHEMA public TO sensory_web;
GRANT  SELECT, INSERT, UPDATE ON sensory_sessions TO sensory_web;
GRANT  USAGE, SELECT ON SEQUENCE sensory_sessions_id_seq TO sensory_web;
GRANT  EXECUTE ON FUNCTION dve_append_sensory_sample(TEXT,TEXT,TEXT,TEXT,TEXT,JSONB,TEXT) TO sensory_web;
-- Explicitly deny the generic mutators (their SQL-safety is delegated to the
-- C++ client allowlist, which a web caller bypasses).
REVOKE EXECUTE ON FUNCTION dve_commit_cell(TEXT,BIGINT,TEXT,TEXT,TEXT)               FROM sensory_web;
REVOKE EXECUTE ON FUNCTION dve_commit_cell_json(TEXT,BIGINT,TEXT,TEXT[],TEXT,TEXT)   FROM sensory_web;
```

- [ ] **Step 2: Password out-of-band** — document in `README.md`: the NAS admin runs `ALTER ROLE sensory_web PASSWORD '<from .env DVE_SENSORY_WEB_PASSWORD>';` once after the migration. Add `DVE_SENSORY_WEB_PASSWORD=change-me` to `.env.example` (placeholder only).

- [ ] **Step 3: Verify least privilege** — connect as `sensory_web`; assert: `SELECT/INSERT/UPDATE` on `sensory_sessions` OK; `SELECT * FROM data_rows` → `permission denied`; `SELECT dve_commit_cell(...)` → `permission denied`; `dve_append_sensory_sample(...)` OK. Encode as `pytest` in Task 3.3.

- [ ] **Step 4: Commit** — `feat(db): least-privilege sensory_web role for the web form (DATAVIEWER-11)`.

### Task 2.3 (OWNER-GATED): DV-15 re-key already-forked rows

**Files:** create `deploy/postgres/migrations/2026-06-25-dv15-rekey-forked-sensory.sql` — **do NOT auto-apply**; run manually after a backup.

- [ ] **Step 1:** Write a **dry-run SELECT first** that lists rows whose `session_name <> json_data->>'test_title'` (forked by the old import paths), and which target key `(test_title, tester_name, date)` they'd move to, flagging any that would collide with an existing canonical row.
- [ ] **Step 2:** Write the guarded UPDATE that sets `session_name = json_data->>'test_title'` **only** where it does not violate `idx_sensory_sessions_key` (skip + report collisions for manual resolution). Wrap in a transaction; idempotent.
- [ ] **Step 3:** Flag to owner: run after a DB backup, review the dry-run output, then apply. **Commit** the SQL (not run) — `chore(db): DV-15 re-key migration for forked sensory rows (owner-gated) (DATAVIEWER-15)`.

---

# PHASE 3 — Flask phone-form service (NAS track)

> Net-new stack at `deploy/sensory-collect/`. Mirrors the NAS `bug-form-app` gunicorn precedent (NAS-only, not in-repo). Connects to the DB by **internal service name `dataviewer-db:5432`** as `sensory_web`. Ships independently — **not** part of the v2.6.0 wrap.

### Task 3.1: Service scaffold

**Files:** create `deploy/sensory-collect/{app.py, requirements.txt, Dockerfile, docker-compose.yml, .env.example, templates/form.html, README.md, tests/test_append.py}` (create `.py` via Python delete-and-rewrite if MIP labels them).

- [ ] **Step 1: `requirements.txt`** — `Flask==3.0.*`, `gunicorn==22.0.*`, `psycopg2-binary==2.9.*`, `Flask-Limiter==3.*`.
- [ ] **Step 2: `Dockerfile`**:

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8000
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:8000", "app:app"]
```

- [ ] **Step 3: `docker-compose.yml`** (shared external network so it reaches the DB by name):

```yaml
services:
  sensory-collect:
    build: .
    image: dve-sensory-collect:1
    container_name: sensory-collect
    restart: unless-stopped
    environment:
      DVE_DB_HOST: dataviewer-db
      DVE_DB_PORT: "5432"
      DVE_DB_NAME: dataviewer
      DVE_SENSORY_WEB_USER: sensory_web
      DVE_SENSORY_WEB_PASSWORD: ${DVE_SENSORY_WEB_PASSWORD}
      DVE_OFFICE_TZ: America/New_York
    ports:
      - "8000:8000"   # behind the existing NAS reverse proxy + TLS
    networks: [dve-net]
networks:
  dve-net:
    external: true
```

- [ ] **Step 4: `.env.example`** — `DVE_SENSORY_WEB_PASSWORD=change-me` (placeholder; real secret only in the NAS `.env`).
- [ ] **Step 5: Commit** — `feat(sensory-collect): service scaffold (Dockerfile/compose/deps) (DATAVIEWER-11)`.

### Task 3.2: Form + validated append handler

**Files:** `deploy/sensory-collect/app.py`, `templates/form.html`.

- [ ] **Step 1: `app.py`** — one GET (form) + one POST (validate → call the stored function):

```python
import os, json, uuid
import psycopg2
from flask import Flask, request, render_template, jsonify
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

app = Flask(__name__)
limiter = Limiter(get_remote_address, app=app, default_limits=["60 per hour"])
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024  # request-size cap

METRICS = ["Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness", "Overall Liking"]
OFFICE_TZ = os.environ.get("DVE_OFFICE_TZ", "America/New_York")

def _conn():
    return psycopg2.connect(
        host=os.environ["DVE_DB_HOST"], port=os.environ.get("DVE_DB_PORT", "5432"),
        dbname=os.environ["DVE_DB_NAME"], user=os.environ["DVE_SENSORY_WEB_USER"],
        password=os.environ["DVE_SENSORY_WEB_PASSWORD"])

@app.get("/")
def form():
    return render_template("form.html", metrics=METRICS)

@app.post("/submit")
@limiter.limit("20 per minute")
def submit():
    f = request.form
    title  = (f.get("test_title")  or "").strip()
    tester = (f.get("tester")      or "").strip()
    if not title or not tester:                         # DATAVIEWER-8: required header
        return jsonify(error="Test Title and Tester are required."), 400
    rnd = (f.get("round") or "N/A").strip()             # '1' | '2' | 'N/A'
    # Build the sample object: numeric scores [1,9] + sample_uid.
    sample = {"name": (f.get("sample_name") or "").strip(),
              "comments": (f.get("comments") or "").strip(),
              "puff_length_sec": _num(f.get("puff_length_sec"), 3.0),
              "sample_uid": (f.get("sample_uid") or str(uuid.uuid4()))}
    for m in METRICS:
        v = _num(f.get(m), None)
        if v is None or not (1 <= v <= 9):
            return jsonify(error=f"{m} must be a number 1-9."), 400
        sample[m] = v                                   # JSON number, never a string
    try:
        with _conn() as cx, cx.cursor() as cur:
            cur.execute(
                "SELECT dve_append_sensory_sample(%s,%s,%s,%s,%s,%s::jsonb,%s)",
                (title, tester, rnd, (f.get("assessor") or "").strip(),
                 (f.get("media") or "").strip(), json.dumps(sample), OFFICE_TZ))
            sid = cur.fetchone()[0]
        return jsonify(ok=True, session_id=sid, sample_uid=sample["sample_uid"])
    except Exception as e:                              # least-priv role bounds the blast radius
        app.logger.exception("append failed")
        return jsonify(error="Could not save. Try again."), 502

def _num(s, default):
    try: return float(s)
    except (TypeError, ValueError): return default
```

- [ ] **Step 2: `templates/form.html`** — mobile-first, viewport meta, the header fields (Test Title, Assessor, Tester, Round{1,2,N/A}, Media), Sample Name, five `1–9` inputs (range slider + number echo), Puff length, Comments. A hidden `sample_uid` generated in JS and **regenerated only after a successful submit** (so a retried submit reuses the uid → idempotent). On success, keep the header fields, clear Sample Name + scores, regenerate `sample_uid`, show a confirmation. Keep it framework-free (vanilla JS `fetch`).
- [ ] **Step 3: Commit** — `feat(sensory-collect): mobile form + validated create-or-append POST (DATAVIEWER-11)`.

### Task 3.3: pytest (round-trip, append, idempotency, least-privilege)

**Files:** `deploy/sensory-collect/tests/test_append.py`.

- [ ] **Step 1:** Against an ephemeral DB (apply `init.sql` + the Phase-2 migrations; reuse the `tests\start-test-postgres.ps1` container or a `pytest` fixture spinning `postgres:16`), assert:
  - scores stored as JSON **numbers** (`jsonb_typeof = 'number'`), values 1–9;
  - **append-on-match:** two submits with identical title+tester+round → one row, two samples; a different header → a new row;
  - **idempotency:** re-POST with the same `sample_uid` → no duplicate;
  - **least privilege:** `sensory_web` gets `permission denied` on `data_rows` and on `dve_commit_cell*`; `dve_append_sensory_sample` works;
  - empty Test Title/Tester → HTTP 400; oversize body → 413; a string score → 400.
- [ ] **Step 2:** Repo-grep asserts **no secret** in the stack (no superuser credential, no committed password).
- [ ] **Step 3: Commit** — `test(sensory-collect): append/idempotency/least-privilege pytest (DATAVIEWER-11)`.

### Task 3.4: Network wiring + NAS docs

**Files:** `deploy/postgres/docker-compose.yml`, `deploy/sensory-collect/README.md`.

- [ ] **Step 1:** Add the shared network to the DB stack so both reach each other by name:

```yaml
# append to deploy/postgres/docker-compose.yml
    networks: [dve-net]      # under the dataviewer-db service
networks:
  dve-net:
    external: true
```

- [ ] **Step 2:** `README.md`: one-time `docker network create dve-net`; set `DVE_SENSORY_WEB_PASSWORD` in the NAS `.env`; `ALTER ROLE sensory_web PASSWORD …`; bring up `sensory-collect`; expose via the existing reverse proxy + TLS (confirm proxy details with owner — NAS-only, not in-repo); verify a phone off-network can submit. **Commit** `docs(deploy): wire sensory-collect to the DB network + NAS runbook (DATAVIEWER-11)`.

---

## Residual confirmations (flag to owner before/while building)

- **Reverse proxy + TLS** for `sensory-collect` is NAS-only (mirrors `bug-form-app`); nothing in-repo. Confirm the proxy hostname/route.
- **DB port** 5432 vs 5433: the internal Docker connection uses `dataviewer-db:5432` (container port) regardless of any host `5433` mapping — moot internally, but confirm the host mapping on the NAS.
- **Round "N/A" merge (§10c):** two `N/A` submits of the same title+tester+date are the same natural key → they **append** (same row). Confirm that's intended (vs distinguishable).
- **Office TZ** default `America/New_York` — confirm the office timezone for the server-side date pin.

## Sequencing & verification

1. **Phase 1** (desktop) is the only v2.6.0-train work; gated under MS-7 heavy-verification discipline (TDD + e2e + `/ponytail-review`); ships in an internal v2.5.x patch that wraps into v2.6.0. **DV-15 code (Task 1.4) before the re-key data migration (Task 2.3).**
2. **Phase 2 + 3** ship on the independent NAS track; not in the v2.6.0 wrap. Build order: 2.1 → 2.2 → 3.1 → 3.2 → 3.3 → 3.4; **Task 2.3 last, owner-gated, after a backup.**
3. Each desktop task: `-Werror` clean + `tests\run-tests.ps1` green (except known-flaky `tst_responsivelayout`). Each DB/service task: `pytest` green against an ephemeral DB.

## Self-review notes

- **Spec coverage:** §3 create-or-append → Task 2.1; §4.1 desktop merge fix → Tasks 1.2/1.3; §4.2 `sample_uid` → Task 1.1; §5 least-privilege/anonymous → Task 2.2 + 3.2 validation + 3.3; §6 network/online-only → Task 3.4; §7 JSON contract → Tasks 1.1/3.2; §8 form UX → Task 3.2; §9/§10d DV-15 → Tasks 1.4 + 2.3; §11 acceptance → Tasks 1.3/3.3.
- **Type consistency:** `sampleUid`/`sample_uid` across struct+serializer+SQL+form; `canonicalSensorySessionName(QString)` used at all 3 desktop sites; `dve_append_sensory_sample(TEXT,TEXT,TEXT,TEXT,TEXT,JSONB,TEXT)` signature identical in migration, GRANT, and the service call.
- **Data-safety floor:** scores are JSON numbers (validated in both the service and the SQL function); append-to-tail-only keeps existing positional paths stable for in-flight desktop edits; the desktop never silently drops a phone append (Tasks 1.2/1.3 + e2e).
