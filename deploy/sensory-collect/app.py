import os, json, uuid
import psycopg2
from flask import Flask, request, render_template, jsonify
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

app = Flask(__name__)
limiter = Limiter(get_remote_address, app=app, default_limits=["60 per hour"])
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024  # request-size cap

METRICS = ["Burnt Taste", "Vapor Volume", "Overall Flavor", "Smoothness", "Overall Liking"]
OFFICE_TZ = os.environ.get("DVE_OFFICE_TZ", "America/Phoenix")  # Arizona: no DST

# Hosts that serve the custom-dropdown "Mfused" variant. Fixed option lists below
# are rendered straight into the form -- the dropdowns never query the DB (so the
# anonymous endpoint can't enumerate the test catalog). Same /submit handler.
MFUSED_HOSTS = set(h.strip().lower() for h in
                   os.environ.get("DVE_MFUSED_HOSTS", "mfused-sensory.ccell-sdr.com").split(",")
                   if h.strip())
MFUSED_OPTIONS = {
    "test_title":  ["Live Resin LVL 1", "Live Resin LVL 2", "Live Resin KO"],
    "assessor":    ["Isabel", "NA"],
    "round":       ["1", "2", "3", "4", "NA"],
    "media":       ["Blasted Pruno", "Cherry Haze"],
    "sample_name": ["BlastedPruno-1", "BlastedPruno-2", "BlastedPruno-3",
                    "CherryHaze-1", "CherryHaze-2", "CherryHaze-3"],
}


def _conn():
    return psycopg2.connect(
        host=os.environ["DVE_DB_HOST"], port=os.environ.get("DVE_DB_PORT", "5432"),
        dbname=os.environ["DVE_DB_NAME"], user=os.environ["DVE_SENSORY_WEB_USER"],
        password=os.environ["DVE_SENSORY_WEB_PASSWORD"])


@app.get("/")
def form():
    host = (request.host or "").split(":")[0].lower()
    mfused = host in MFUSED_HOSTS
    return render_template("form.html", metrics=METRICS,
                           mfused=mfused, opts=(MFUSED_OPTIONS if mfused else {}))


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
    except Exception:                              # least-priv role bounds the blast radius
        app.logger.exception("append failed")
        return jsonify(error="Could not save. Try again."), 502


def _num(s, default):
    try: return float(s)
    except (TypeError, ValueError): return default
