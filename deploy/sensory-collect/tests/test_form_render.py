"""DV-21 -- render tests for the mobile sensory form refinements.

GET / must include the new history/nav/plot affordances and the three script
includes for BOTH the default and the Mfused host, while the core submit fields
remain present (regression guard: the submit flow is untouched). DB-free.
"""
import importlib
import os
import sys

import pytest

pytest.importorskip("flask")
pytest.importorskip("flask_limiter")

HERE = os.path.dirname(os.path.abspath(__file__))
SERVICE_DIR = os.path.dirname(HERE)
MFUSED_HOST = "mfused-sensory.ccell-sdr.com"
PLAIN_HOST = "sensory.example.com"


@pytest.fixture()
def client(monkeypatch):
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


@pytest.mark.parametrize("host", [PLAIN_HOST, MFUSED_HOST])
def test_new_affordances_present(client, host):
    html = _html(client, host)
    for token in ('id="hist-btn"', 'id="hist-drawer"', 'id="hist-list"',
                  'id="hist-viewer"', 'id="hist-nav"',
                  'id="nav-prev"', 'id="nav-next"', 'id="nav-plot"',
                  'id="plot-page"', 'id="plot-canvas"', 'id="plot-test"',
                  'id="plot-toggles"', 'id="plot-back"',
                  "sensory_history.js", "sensory_ui.js", "sensory_plot.js"):
        assert token in html, token


@pytest.mark.parametrize("host", [PLAIN_HOST, MFUSED_HOST])
def test_core_submit_fields_untouched(client, host):
    html = _html(client, host)
    for token in ('id="f"', 'id="submit-btn"', 'name="test_title"', 'name="tester"',
                  'name="Burnt Taste"', 'name="Vapor Volume"', 'name="Overall Flavor"',
                  'name="Smoothness"', 'name="Overall Liking"'):
        assert token in html, token
