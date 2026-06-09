#!/usr/bin/env python3
"""Headless tests for build_clean_template helpers (no Excel needed).

    python excel-sidecar/test_build_helpers.py --source "C:\\...\\...v1.xlsm"
Exit 0 on PASS.
"""
import argparse
import os
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_clean_template as B   # noqa: E402


def make_stripped(src, dst):
    """Copy `src` minus its customUI/* parts and the customUI relationship from
    _rels/.rels (so inject_customui's ribbon add-path is exercised), but KEEP
    xl/webextensions/* so its web-add-in stripping is exercised too -> a faithful
    stand-in for the source copy that Excel re-saved."""
    import re
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            n = it.filename
            if n.startswith("customUI/"):
                continue
            data = zin.read(n)
            if n == "_rels/.rels":
                data = re.sub(r'<Relationship[^>]*customUI/customUI14\.xml[^>]*/>',
                              '', data.decode("utf-8")).encode("utf-8")
            zout.writestr(it, data)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", required=True)
    args = ap.parse_args()
    fails = []

    # Build-module constants for the Test Selection redesign.
    if B.UPLOAD_SHEET != "Test Selection":
        fails.append("UPLOAD_SHEET should be 'Test Selection', got %r" % B.UPLOAD_SHEET)
    if B.SETTINGS_SHEET not in B.KEEP:
        fails.append("SETTINGS_SHEET (%r) missing from KEEP" % B.SETTINGS_SHEET)
    if "Test Selection" not in B.KEEP:
        fails.append("'Test Selection' missing from KEEP")
    if "_Settings" not in B.NAMED["DV_FileName"]:
        fails.append("NAMED['DV_FileName'] should reference _Settings, got %r"
                     % B.NAMED["DV_FileName"])
    if "Test Selection" not in B.NAMED["DV_TestSelection"]:
        fails.append("NAMED['DV_TestSelection'] should reference Test Selection, got %r"
                     % B.NAMED["DV_TestSelection"])
    if len(B.SELECTION_ROWS) != 13:
        fails.append("SELECTION_ROWS should have 13 rows, got %d" % len(B.SELECTION_ROWS))
    if B.SELECTION_ROWS[0][0] != "Custom Test Template":
        fails.append("SELECTION_ROWS[0] should be 'Custom Test Template', got %r"
                     % (B.SELECTION_ROWS[0],))
    if B.SELECTION_ROWS[-1] != ("Test SOP's", True):
        fails.append("SELECTION_ROWS[-1] should be (\"Test SOP's\", True), got %r"
                     % (B.SELECTION_ROWS[-1],))

    # v1.1 Feature 2: Specify Test Name -> DV_OrigFileName named range + label.
    if not ("DV_OrigFileName" in B.NAMED and B.NAMED["DV_OrigFileName"] == "'_Settings'!$B$7"):
        fails.append("NAMED['DV_OrigFileName'] should be \"'_Settings'!$B$7\", got %r"
                     % B.NAMED.get("DV_OrigFileName"))
    if not (len(B.SETTINGS_LABELS) == 7 and B.SETTINGS_LABELS[-1] == "Original file name"):
        fails.append("SETTINGS_LABELS should have 7 entries ending 'Original file name', got %r"
                     % (B.SETTINGS_LABELS,))

    # v1.1 Feature 3: Clog auto-formula derived from same-row Draw Pressure.
    if B.clog_formula("D", 5) != '=IF(D5="","",IF(D5>=15,"Heavy Clog",IF(D5>5,"Light Clog","")))':
        fails.append("clog_formula('D', 5) wrong, got %r" % B.clog_formula("D", 5))
    if B.clog_formula("P", 5) != '=IF(P5="","",IF(P5>=15,"Heavy Clog",IF(P5>5,"Light Clog","")))':
        fails.append("clog_formula('P', 5) wrong, got %r" % B.clog_formula("P", 5))
    if not (B.BLOCK_COLS == 12 and B.CLOG_REL == 7 and B.DRAW_REL == 4):
        fails.append("Clog geometry constants wrong: BLOCK_COLS=%r CLOG_REL=%r DRAW_REL=%r"
                     % (B.BLOCK_COLS, B.CLOG_REL, B.DRAW_REL))

    # check_preconditions detects a missing source.
    if not B.check_preconditions("C:/does/not/exist.xlsm"):
        fails.append("check_preconditions should flag a missing source")

    # inject_customui produces a wired ribbon and strips no-add-in.
    tmpdir = os.environ.get("TEMP", ".")
    stripped = os.path.join(tmpdir, "_dv_stripped.xlsm")
    make_stripped(args.source, stripped)
    B.inject_customui(stripped, HERE, args.source)
    with zipfile.ZipFile(stripped) as z:
        names = z.namelist()
        out_ui = z.read("customUI/customUI14.xml").decode("utf-8")
        root_rels = z.read("_rels/.rels").decode("utf-8")
        ctypes = z.read("[Content_Types].xml").decode("utf-8")
        ui_rels = z.read("customUI/_rels/customUI14.xml.rels").decode("utf-8") \
            if "customUI/_rels/customUI14.xml.rels" in names else ""
    repo_ui = open(os.path.join(HERE, "customUI14.xml"), encoding="utf-8").read()
    if out_ui.replace("\r\n", "\n").strip() != repo_ui.replace("\r\n", "\n").strip():
        fails.append("injected customUI14.xml != repo copy")
    if "customUI/_rels/customUI14.xml.rels" not in names:
        fails.append("customUI rels missing")
    if "customUI/customUI14.xml" not in root_rels:
        fails.append("root _rels/.rels missing customUI relationship")
    if 'Extension="xml"' not in ctypes:
        fails.append("[Content_Types].xml missing xml Default (covers customUI14.xml)")
    if any(n.startswith("xl/webextensions/") for n in names):
        fails.append("web add-in parts survived stripping")
    if "webextension" in root_rels.lower():
        fails.append("_rels/.rels still references the web add-in")
    if "/xl/webextensions/" in ctypes:
        fails.append("[Content_Types].xml still has webextension overrides")
    for _icon_id, _fn in B.HELP_ICONS.items():
        if "customUI/images/%s" % _fn not in names:
            fails.append("help icon %s not injected" % _fn)
        if 'Id="%s"' % _icon_id not in ui_rels:
            fails.append("help icon relationship %s not injected" % _icon_id)
    try:
        os.remove(stripped)
    except OSError:
        pass

    for f in fails:
        print("FAIL", f)
    print("RESULT:", "ALL PASS" if not fails else "FAILURES")
    return 0 if not fails else 1


if __name__ == "__main__":
    sys.exit(main())
