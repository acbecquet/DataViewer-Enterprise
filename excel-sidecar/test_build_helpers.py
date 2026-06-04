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
    """Copy `src` minus its customUI/* and xl/webextensions/* PARTS, and drop the
    customUI relationship from _rels/.rels -> a faithful stand-in for a freshly
    built workbook with no ribbon, so inject_customui's add-path is exercised."""
    import re
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in zin.infolist():
            n = it.filename
            if n.startswith("customUI/") or n.startswith("xl/webextensions/"):
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
        fails.append("web add-in parts survived injection")
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
