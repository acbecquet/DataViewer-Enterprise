"""DATAVIEWER-11 -- rendering tests for the Mfused customer form variant.

The Mfused host relabels "Round" as "Mode" with letter choices A/B/C that map to
round numbers 1/2/3 via the <option value>. GET / renders straight from fixed
in-code option lists and never touches the database, so these tests need NO DB
(nothing to skip). They import app.py directly and drive Flask's test client.
"""

import importlib
import os
import sys

import pytest

pytest.importorskip("flask")
pytest.importorskip("flask_limiter")

HERE = os.path.dirname(os.path.abspath(__file__))
SERVICE_DIR = os.path.dirname(HERE)                # deploy/sensory-collect

MFUSED_HOST = "mfused-sensory.ccell-sdr.com"
PLAIN_HOST = "sensory.example.com"


@pytest.fixture()
def client(monkeypatch):
    """Flask test client for app.py with the Mfused host set deterministically.

    Reloads the module under the patched env so MFUSED_HOSTS (read at import
    time) matches MFUSED_HOST regardless of the ambient environment. GET / never
    connects, so no DB env is needed.
    """
    monkeypatch.setenv("DVE_MFUSED_HOSTS", MFUSED_HOST)
    sys.path.insert(0, SERVICE_DIR)
    try:
        app_module = importlib.import_module("app")
        importlib.reload(app_module)
        app_module.app.config["TESTING"] = True
        app_module.app.config["RATELIMIT_ENABLED"] = False
        yield app_module.app.test_client()
    finally:
        if SERVICE_DIR in sys.path:
            sys.path.remove(SERVICE_DIR)


def _html(client, host):
    r = client.get("/", headers={"Host": host})
    assert r.status_code == 200, r.get_data(as_text=True)
    return r.get_data(as_text=True)


def test_mfused_relabels_round_as_mode(client):
    html = _html(client, MFUSED_HOST)
    assert '<label for="round">Mode</label>' in html
    assert '<label for="round">Round</label>' not in html


def test_mfused_mode_options_map_letters_to_round_numbers(client):
    html = _html(client, MFUSED_HOST)
    # Letter shown to the user; round number POSTed to /submit and stored in DB.
    assert '<option value="1">A</option>' in html
    assert '<option value="2">B</option>' in html
    assert '<option value="3">C</option>' in html
    # The old numeric / NA round choices must be gone from the Mfused variant.
    assert 'value="4"' not in html


def test_mfused_keeps_round_field_name(client):
    # Backend contract is unchanged: the field is still POSTed as name="round".
    html = _html(client, MFUSED_HOST)
    assert 'name="round"' in html


def test_plain_host_still_shows_round(client):
    html = _html(client, PLAIN_HOST)
    assert '<label for="round">Round</label>' in html
    assert '<label for="round">Mode</label>' not in html
    assert '<option value="1">1</option>' in html
    assert '<option value="4">4</option>' in html
