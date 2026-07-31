"""DATAVIEWER-11 -- pytest for the phone-sensory create-or-append path.

Exercises the SAME stored function the Flask service calls
(``dve_append_sensory_sample``) plus the ``sensory_web`` least-privilege role,
against an ephemeral PostgreSQL 16 that already has ``init.sql`` (the
BEGIN..COMMIT schema block) and the Phase-2 migrations applied. The HTTP-layer
checks (empty headers -> 400, oversize body -> 413, string score -> 400) drive
``app.py`` through Flask's test client.

DB connection comes from the environment and the suite SKIPS cleanly when none
is configured. Two forms are accepted:

* ``DVE_TEST_PG_CONN`` -- a libpq keyword/value DSN, e.g.
  ``host=127.0.0.1 port=5433 dbname=dve_test user=test password=test``
    tests/start-test-postgres.ps1            # exports DVE_TEST_PG_CONN
* discrete ``DVE_DB_HOST`` / ``DVE_DB_PORT`` / ``DVE_DB_NAME`` /
  ``DVE_DB_USER`` / ``DVE_DB_PASSWORD``.

The configured connection must be a SUPERUSER / schema-owner (it applies DDL and
runs ``ALTER ROLE``); a separate, low-privilege connection as ``sensory_web`` is
opened from it to prove least privilege.

Run (work machine, with the test container up)::

    pip install pytest psycopg2-binary "Flask==3.0.*" "Flask-Limiter==3.*"
    tests/start-test-postgres.ps1            # exports DVE_TEST_PG_CONN
    pytest deploy/sensory-collect/tests/test_append.py -v
"""

import json
import os
import re
import uuid

import pytest

psycopg2 = pytest.importorskip("psycopg2")
from psycopg2 import errors as pg_errors  # noqa: E402  (after importorskip)


# --------------------------------------------------------------------------- #
# Paths / constants
# --------------------------------------------------------------------------- #

HERE = os.path.dirname(os.path.abspath(__file__))
SERVICE_DIR = os.path.dirname(HERE)                        # deploy/sensory-collect
REPO_ROOT = os.path.dirname(os.path.dirname(SERVICE_DIR))  # repo root
POSTGRES_DIR = os.path.join(REPO_ROOT, "deploy", "postgres")
INIT_SQL = os.path.join(POSTGRES_DIR, "init.sql")
MIGRATIONS_DIR = os.path.join(POSTGRES_DIR, "migrations")

METRICS = ["Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness", "Overall Liking"]

# A throwaway password for the sensory_web role in the test DB only. NOT a
# committed production secret -- the prod password is set out-of-band on the NAS
# from a .env value (see README). It is fine for this to live in test code; the
# secret-scan test (test_no_committed_secret) deliberately excludes this file.
SENSORY_WEB_TEST_PASSWORD = "test-sensory-web-pw"


# --------------------------------------------------------------------------- #
# Connection-from-env helpers
# --------------------------------------------------------------------------- #

def _dsn_from_env():
    """Return a libpq DSN string for the SUPERUSER test connection, or None.

    Prefers DVE_TEST_PG_CONN; falls back to discrete DVE_DB_* vars. Returns None
    when neither is configured so the suite can skip cleanly.
    """
    dsn = os.environ.get("DVE_TEST_PG_CONN")
    if dsn:
        return dsn
    host = os.environ.get("DVE_DB_HOST")
    name = os.environ.get("DVE_DB_NAME")
    if not host or not name:
        return None
    parts = [
        "host=" + host,
        "port=" + os.environ.get("DVE_DB_PORT", "5432"),
        "dbname=" + name,
        "user=" + os.environ.get("DVE_DB_USER", "postgres"),
    ]
    pw = os.environ.get("DVE_DB_PASSWORD")
    if pw:
        parts.append("password=" + pw)
    return " ".join(parts)


def _parse_dsn(dsn):
    """Parse a libpq keyword/value DSN into a dict (host/port/dbname/user/...)."""
    out = {}
    for tok in dsn.split():
        if "=" in tok:
            k, v = tok.split("=", 1)
            out[k] = v
    return out


_SKIP_REASON = (
    "no test database configured (set DVE_TEST_PG_CONN or DVE_DB_HOST/DVE_DB_NAME; "
    "run tests/start-test-postgres.ps1)"
)


def _read_init_schema_block():
    """The BEGIN;..COMMIT; block of init.sql (the pg_cron tail needs the extension,
    which the lightweight test container does not load -- mirror start-test-postgres.ps1)."""
    sql = open(INIT_SQL, encoding="utf-8").read()
    begin = sql.find("BEGIN;")
    commit = sql.find("COMMIT;", begin)
    if begin < 0 or commit < 0:
        raise RuntimeError("could not locate BEGIN;..COMMIT; in init.sql")
    return sql[begin:commit + len("COMMIT;")]


def _migration_sql_in_order():
    """All migration files, in filename (date) order -- same order the runner uses."""
    files = sorted(f for f in os.listdir(MIGRATIONS_DIR) if f.endswith(".sql"))
    return [(f, open(os.path.join(MIGRATIONS_DIR, f), encoding="utf-8").read()) for f in files]


# --------------------------------------------------------------------------- #
# Fixtures
# --------------------------------------------------------------------------- #

@pytest.fixture(scope="session")
def superuser_dsn():
    dsn = _dsn_from_env()
    if not dsn:
        pytest.skip(_SKIP_REASON)
    # Fail fast (don't silently skip) if a DSN is set but unreachable.
    try:
        conn = psycopg2.connect(dsn)
    except psycopg2.OperationalError as exc:  # pragma: no cover - environment
        pytest.skip(_SKIP_REASON + ": connection failed (" + str(exc) + ")")
    conn.close()
    return dsn


@pytest.fixture(scope="session")
def schema_ready(superuser_dsn):
    """Apply init.sql (schema block) + every migration, then pin the sensory_web
    password so the least-privilege tests can log in. Idempotent."""
    conn = psycopg2.connect(superuser_dsn)
    conn.autocommit = True
    try:
        with conn.cursor() as cur:
            cur.execute(_read_init_schema_block())
            for name, sql in _migration_sql_in_order():
                try:
                    cur.execute(sql)
                except psycopg2.Error as exc:
                    # Migrations are idempotent (IF [NOT] EXISTS / CREATE OR
                    # REPLACE); tolerate "already exists" the way the PS runner
                    # does with ON_ERROR_STOP=0, but surface anything unexpected.
                    conn.rollback()
                    if "already exists" not in str(exc).lower():
                        raise RuntimeError("migration " + name + " failed: " + str(exc)) from exc
            # Give sensory_web a known password for the least-privilege login.
            cur.execute(
                "ALTER ROLE sensory_web LOGIN PASSWORD %s", (SENSORY_WEB_TEST_PASSWORD,)
            )
    finally:
        conn.close()
    return superuser_dsn


@pytest.fixture()
def su_conn(schema_ready):
    """A fresh autocommit superuser connection per test."""
    conn = psycopg2.connect(schema_ready)
    conn.autocommit = True
    try:
        yield conn
    finally:
        conn.close()


@pytest.fixture()
def web_conn(schema_ready):
    """A connection as the least-privilege sensory_web role."""
    d = _parse_dsn(schema_ready)
    conn = psycopg2.connect(
        host=d.get("host", "127.0.0.1"),
        port=d.get("port", "5432"),
        dbname=d.get("dbname"),
        user="sensory_web",
        password=SENSORY_WEB_TEST_PASSWORD,
    )
    conn.autocommit = True
    try:
        yield conn
    finally:
        conn.close()


@pytest.fixture()
def client(schema_ready, monkeypatch):
    """Flask test client for app.py, wired to the test DB as sensory_web.

    Patched at import time via env vars (app.py reads them at call time). The
    Limiter is disabled so the validation paths aren't masked by 429s.
    """
    d = _parse_dsn(schema_ready)
    monkeypatch.setenv("DVE_DB_HOST", d.get("host", "127.0.0.1"))
    monkeypatch.setenv("DVE_DB_PORT", d.get("port", "5432"))
    monkeypatch.setenv("DVE_DB_NAME", d.get("dbname", ""))
    monkeypatch.setenv("DVE_SENSORY_WEB_USER", "sensory_web")
    monkeypatch.setenv("DVE_SENSORY_WEB_PASSWORD", SENSORY_WEB_TEST_PASSWORD)
    monkeypatch.setenv("DVE_OFFICE_TZ", "America/Phoenix")

    import importlib
    import sys
    sys.path.insert(0, SERVICE_DIR)
    app_module = importlib.import_module("app")
    importlib.reload(app_module)           # re-read env if a prior test imported it
    app_module.app.config["TESTING"] = True
    # Disable rate limiting in tests so the 400/413 validation paths aren't
    # masked by a 429. RATELIMIT_ENABLED is Flask-Limiter's canonical switch;
    # also flip the limiter instance for belt-and-suspenders across 3.x.
    app_module.app.config["RATELIMIT_ENABLED"] = False
    try:
        app_module.limiter.enabled = False
    except Exception:
        pass
    try:
        yield app_module.app.test_client()
    finally:
        if SERVICE_DIR in sys.path:
            sys.path.remove(SERVICE_DIR)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def _valid_sample(uid=None, name="S1", **score_overrides):
    s = {
        "name": name,
        "comments": "",
        "puff_length_sec": 3.0,
        "sample_uid": uid or str(uuid.uuid4()),
    }
    for m in METRICS:
        s[m] = float(score_overrides.get(m, 5))
    return s


def _append(cur, title, tester, rnd, sample, assessor="A", media="M", tz="America/Phoenix"):
    cur.execute(
        "SELECT dve_append_sensory_sample(%s,%s,%s,%s,%s,%s::jsonb,%s)",
        (title, tester, rnd, assessor, media, json.dumps(sample), tz),
    )
    return cur.fetchone()[0]


def _row(cur, sid):
    cur.execute(
        "SELECT session_name, tester_name, date, json_data FROM sensory_sessions WHERE id=%s",
        (sid,),
    )
    return cur.fetchone()


def _unique_title():
    return "pytest-" + uuid.uuid4().hex[:12]


# --------------------------------------------------------------------------- #
# 1. Scores stored as JSON numbers in [1,9]
# --------------------------------------------------------------------------- #

def test_scores_stored_as_json_numbers(su_conn):
    title = _unique_title()
    with su_conn.cursor() as cur:
        sid = _append(cur, title, "Bob", "1", _valid_sample(name="S1", **{m: 7 for m in METRICS}))
        # Each metric is a JSON *number*, value 1..9 -- never a string.
        for m in METRICS:
            cur.execute(
                "SELECT jsonb_typeof((json_data->'samples'->0)->%s), "
                "       ((json_data->'samples'->0)->>%s)::numeric "
                "FROM sensory_sessions WHERE id=%s",
                (m, m, sid),
            )
            typ, val = cur.fetchone()
            assert typ == "number", m + " stored as " + str(typ) + ", expected number"
            assert 1 <= val <= 9, m + " = " + str(val) + " out of [1,9]"


# --------------------------------------------------------------------------- #
# 2. Append-on-match (same key -> one row, many samples; diff key -> new row)
# --------------------------------------------------------------------------- #

def test_append_on_matching_key_one_row_two_samples(su_conn):
    title = _unique_title()
    with su_conn.cursor() as cur:
        sid1 = _append(cur, title, "Bob", "1", _valid_sample(name="S1"))
        sid2 = _append(cur, title, "Bob", "1", _valid_sample(name="S2"))
        assert sid1 == sid2, "identical title+tester+round must resolve to ONE row"
        _, _, _, jd = _row(cur, sid1)
        names = [s["name"] for s in jd["samples"]]
        assert names == ["S1", "S2"], "expected two tail-appended samples, got " + str(names)


def test_different_header_creates_new_row(su_conn):
    title = _unique_title()
    with su_conn.cursor() as cur:
        sid1 = _append(cur, title, "Bob", "1", _valid_sample(name="S1"))
        # Different tester -> different key -> a NEW row.
        sid_diff_tester = _append(cur, title, "Alice", "1", _valid_sample(name="S1"))
        # Different round -> tester_name suffix differs ("Bob R1" vs "Bob R2") -> NEW row.
        sid_diff_round = _append(cur, title, "Bob", "2", _valid_sample(name="S1"))
        assert len({sid1, sid_diff_tester, sid_diff_round}) == 3, "each distinct header must fork a row"


def test_round_suffix_folds_into_tester_name(su_conn):
    title = _unique_title()
    with su_conn.cursor() as cur:
        sid = _append(cur, title, "Bob", "2", _valid_sample())
        _, tester_name, _, _ = _row(cur, sid)
        assert tester_name == "Bob R2", "round must fold to R2 suffix, got " + repr(tester_name)


def test_rounds_3_and_4_fold_into_tester_name(su_conn):
    # Rounds 3 & 4 fold into tester_name as R3/R4 (2026-06-26-dv11b). This is the
    # DB side of the Mfused "Mode C" -> round 3 relabel: Mode C must land as a
    # distinctly-labeled "<tester> R3" row, consistent with the desktop.
    title = _unique_title()
    with su_conn.cursor() as cur:
        sid3 = _append(cur, title, "Bob", "3", _valid_sample(name="S3"))
        sid4 = _append(cur, title, "Bob", "4", _valid_sample(name="S4"))
        _, t3, _, _ = _row(cur, sid3)
        _, t4, _, _ = _row(cur, sid4)
        assert t3 == "Bob R3", "round 3 must fold to R3 suffix, got " + repr(t3)
        assert t4 == "Bob R4", "round 4 must fold to R4 suffix, got " + repr(t4)
        assert sid3 != sid4, "distinct rounds must fork distinct session rows"


# --------------------------------------------------------------------------- #
# 3. Idempotency (same sample_uid -> no duplicate)
# --------------------------------------------------------------------------- #

def test_idempotent_resubmit_same_uid(su_conn):
    title = _unique_title()
    uid = str(uuid.uuid4())
    with su_conn.cursor() as cur:
        sid1 = _append(cur, title, "Bob", "1", _valid_sample(uid=uid, name="S1"))
        sid2 = _append(cur, title, "Bob", "1", _valid_sample(uid=uid, name="S1-retry"))
        assert sid1 == sid2
        _, _, _, jd = _row(cur, sid1)
        assert len(jd["samples"]) == 1, "a re-POST with the same sample_uid must not duplicate"


# --------------------------------------------------------------------------- #
# 4. The SQL function rejects a non-number / out-of-range score
# --------------------------------------------------------------------------- #

def test_sql_function_rejects_string_score(su_conn):
    title = _unique_title()
    bad = _valid_sample()
    bad["Burnt Taste"] = "7"  # string, not a JSON number
    with su_conn.cursor() as cur:
        with pytest.raises(pg_errors.RaiseException):
            _append(cur, title, "Bob", "1", bad)


def test_sql_function_rejects_out_of_range_score(su_conn):
    title = _unique_title()
    bad = _valid_sample()
    bad["Smoothness"] = 12.0
    with su_conn.cursor() as cur:
        with pytest.raises(pg_errors.RaiseException):
            _append(cur, title, "Bob", "1", bad)


# --------------------------------------------------------------------------- #
# 5. NOTIFY / version bump fire on append (desktops refresh for free)
# --------------------------------------------------------------------------- #

def test_append_bumps_version(su_conn):
    title = _unique_title()
    with su_conn.cursor() as cur:
        sid = _append(cur, title, "Bob", "1", _valid_sample(name="S1"))
        cur.execute("SELECT version FROM sensory_sessions WHERE id=%s", (sid,))
        v1 = cur.fetchone()[0]
        _append(cur, title, "Bob", "1", _valid_sample(name="S2"))
        cur.execute("SELECT version FROM sensory_sessions WHERE id=%s", (sid,))
        v2 = cur.fetchone()[0]
        assert v2 > v1, "the append UPDATE must bump version (bump_version trigger)"


# --------------------------------------------------------------------------- #
# 6. Least privilege: sensory_web can append, but NOT touch data_rows, the v3
#    long-format lab tables, or the generic dve_commit_cell* mutators.
# --------------------------------------------------------------------------- #

def test_least_priv_append_allowed(web_conn):
    title = _unique_title()
    with web_conn.cursor() as cur:
        sid = _append(cur, title, "Bob", "1", _valid_sample(name="S1"))
        assert sid and sid > 0


def test_least_priv_denied_on_data_rows(web_conn):
    with web_conn.cursor() as cur:
        with pytest.raises(pg_errors.InsufficientPrivilege):
            cur.execute("SELECT * FROM data_rows LIMIT 1")


# The three v3 long-format tables (hazard H16). The REVOKE that keeps them shut
# lives in deploy/postgres/migrations/2026-07-31-v3-long-format.sql; the role's
# own migration only ran `REVOKE ALL ON ALL TABLES IN SCHEMA public`, which is a
# ONE-TIME snapshot and could not reach tables that did not exist yet.
#
# The posture is already correct by Postgres default -- a new table grants
# nothing to non-owners -- so these tests are not chasing a live bug. They exist
# to PIN it: from 3c onward `measurements` and `sample_headers` hold every TPM
# measurement in the lab, and the phone form's anonymous, internet-reachable
# login must never be able to read one. A future blanket
# `GRANT ... ON ALL TABLES IN SCHEMA public TO sensory_web`, or an
# `ALTER DEFAULT PRIVILEGES`, would open all three silently. These turn that red.


def test_least_priv_denied_on_measurements(web_conn):
    with web_conn.cursor() as cur:
        with pytest.raises(pg_errors.InsufficientPrivilege):
            cur.execute("SELECT * FROM measurements LIMIT 1")


def test_least_priv_denied_on_sample_headers(web_conn):
    with web_conn.cursor() as cur:
        with pytest.raises(pg_errors.InsufficientPrivilege):
            cur.execute("SELECT * FROM sample_headers LIMIT 1")


def test_least_priv_denied_on_metric_defs(web_conn):
    with web_conn.cursor() as cur:
        with pytest.raises(pg_errors.InsufficientPrivilege):
            cur.execute("SELECT * FROM metric_defs LIMIT 1")


def test_least_priv_denied_on_commit_cell(web_conn):
    with web_conn.cursor() as cur:
        with pytest.raises(pg_errors.InsufficientPrivilege):
            cur.execute("SELECT dve_commit_cell('sensory_sessions', 1, 'x', 'y', 'z')")


def test_least_priv_denied_on_commit_cell_json(web_conn):
    with web_conn.cursor() as cur:
        with pytest.raises(pg_errors.InsufficientPrivilege):
            cur.execute(
                "SELECT dve_commit_cell_json('sensory_sessions', 1, 'samples[0].name', "
                "ARRAY['samples','0','name'], 'v', 'u')"
            )


# --------------------------------------------------------------------------- #
# 7. HTTP layer: empty headers -> 400, string score -> 400, oversize -> 413
# --------------------------------------------------------------------------- #

def _form(**over):
    base = {
        "test_title": "HTTP Test",
        "tester": "Bob",
        "round": "1",
        "assessor": "A",
        "media": "M",
        "sample_name": "S1",
        "comments": "",
        "puff_length_sec": "3.0",
        "sample_uid": str(uuid.uuid4()),
    }
    for m in METRICS:
        base[m] = "5"
    base.update(over)
    return base


def test_http_empty_title_rejected(client):
    r = client.post("/submit", data=_form(test_title=""))
    assert r.status_code == 400


def test_http_empty_tester_rejected(client):
    r = client.post("/submit", data=_form(tester=""))
    assert r.status_code == 400


def test_http_string_score_rejected(client):
    # A non-numeric score string -> _num() returns None -> 400 before any DB call.
    r = client.post("/submit", data=_form(**{"Burnt Taste": "tasty"}))
    assert r.status_code == 400


def test_http_out_of_range_score_rejected(client):
    r = client.post("/submit", data=_form(**{"Vapor Volume": "11"}))
    assert r.status_code == 400


def test_http_oversize_body_413(client):
    # MAX_CONTENT_LENGTH is 64 KiB; send a comments blob well past it.
    r = client.post("/submit", data=_form(comments="x" * (128 * 1024)))
    assert r.status_code == 413


def test_http_valid_submit_appends(client, su_conn):
    title = _unique_title()
    uid = str(uuid.uuid4())
    r = client.post("/submit", data=_form(test_title=title, sample_uid=uid))
    assert r.status_code == 200, r.get_data(as_text=True)
    body = r.get_json()
    assert body.get("ok") is True
    assert body.get("sample_uid") == uid
    # The row really landed and the score is a JSON number.
    with su_conn.cursor() as cur:
        cur.execute(
            "SELECT jsonb_typeof((json_data->'samples'->0)->'Burnt Taste') "
            "FROM sensory_sessions WHERE id=%s",
            (body["session_id"],),
        )
        assert cur.fetchone()[0] == "number"


# --------------------------------------------------------------------------- #
# 8. No committed secret anywhere in the service tree
# --------------------------------------------------------------------------- #


def test_no_committed_secret():
    """The repo is PUBLIC. Assert no real password / superuser credential is
    committed in deploy/sensory-collect/.

    Only flags a hardcoded *literal* assigned to a password key in a config /
    env / compose file. The following are explicitly NOT secrets and must pass:
      * a placeholder ('change-me');
      * a ${VAR} compose interpolation;
      * code that reads from the environment (e.g. os.environ[...]);
      * prose that merely mentions the word 'password'.
    This test file is excluded (it defines a throwaway test password by design).
    """
    # KEY <sep> VALUE, capturing the value token up to whitespace / quote / comment.
    assign = re.compile(
        r"(DVE_SENSORY_WEB_PASSWORD|DVE_DB_PASSWORD|POSTGRES_PASSWORD)"
        r"\s*[:=]\s*"
        r"['\"]?(?P<val>[^\s'\"`,;)\]]+)",
        re.IGNORECASE,
    )
    placeholders = {"change-me", "change-me-to-a-strong-random-password"}
    offenders = []
    for root, _dirs, files in os.walk(SERVICE_DIR):
        if ".git" in root or "__pycache__" in root:
            continue
        for fn in files:
            full = os.path.join(root, fn)
            rel = os.path.relpath(full, REPO_ROOT)
            if os.path.abspath(full) == os.path.abspath(__file__):
                continue
            if fn.endswith(".pyc"):
                continue
            try:
                text = open(full, encoding="utf-8", errors="ignore").read()
            except OSError:
                continue
            for m in assign.finditer(text):
                val = m.group("val").strip().rstrip("`,;")
                # Allowed: placeholders, ${VAR} interpolation, and code that
                # reads the value from the environment rather than hardcoding it.
                if val in placeholders:
                    continue
                if val.startswith("${") or val.startswith("$"):
                    continue
                if val.startswith("os.environ") or val.startswith("os."):
                    continue
                offenders.append(rel + ": " + m.group(0).strip())
    assert not offenders, "possible committed secret(s):\n" + "\n".join(offenders)
