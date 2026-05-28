#!/usr/bin/env python3
"""One-shot installer for the DataViewer upload sidecar.

Imports the canonical modules from this folder into a target `.xlsm` via Excel
COM automation, so re-deploying the macros is a single command instead of a
multi-step manual VBE import. Dry-run by default; `--apply` makes the changes
after taking a timestamped backup.

    # dry run (no changes; prints the plan + checks prerequisites)
    python install_sidecar.py --file "C:\\path\\to\\Automated Testing Template.xlsm"
    # apply
    python install_sidecar.py --file "C:\\path\\to\\Automated Testing Template.xlsm" --apply

Prerequisites for --apply:
  * pywin32  (pip install pywin32)
  * Excel Trust Center -> Macro Settings -> "Trust access to the VBA project
    object model" must be ENABLED, or COM cannot edit modules.

Safety:
  * Never touches the file in --dry-run.
  * --apply copies the workbook to <name>.bak-YYYYmmdd_HHMMSS first.
  * Closes Excel even on error.

NOTE: the ribbon (customUI14.xml) and its icons are NOT installed here - apply
those with the Office Custom UI Editor if they change (they rarely do).
"""
import argparse
import os
import shutil
import sys
import time

# repo file -> standard module name to (re)create. ThisWorkbook handled apart.
STD_MODULES = [
    ("DataViewerUpload.bas", "DataViewerUpload"),
    ("TestingTools.bas", "TestingTools"),
    ("SampleNav.bas", "SampleNav"),
]
# old module names to delete before importing (covers the Module1 -> TestingTools rename)
OLD_MODULES = ["DataViewerUpload", "TestingTools", "Module1", "SampleNav"]
THISWORKBOOK_FILE = "ThisWorkbook.cls.txt"

VBEXT_CT_DOCUMENT = 100  # vbext_ComponentType for document (ThisWorkbook/sheets)


def thisworkbook_body(repo: str) -> str:
    """Body to put in the ThisWorkbook module: drop the repo's pasted comment
    header, keep from `Option Explicit` onward."""
    text = open(os.path.join(repo, THISWORKBOOK_FILE), encoding="utf-8").read()
    lines = text.replace("\r\n", "\n").split("\n")
    start = next((i for i, l in enumerate(lines) if l.strip() == "Option Explicit"), 0)
    return "\n".join(lines[start:]).strip("\n")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--file", required=True, help="target .xlsm")
    ap.add_argument("--repo", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--apply", action="store_true", help="actually modify the workbook")
    args = ap.parse_args()

    # Validate inputs (both modes).
    missing = [f for f, _ in STD_MODULES if not os.path.isfile(os.path.join(args.repo, f))]
    if not os.path.isfile(os.path.join(args.repo, THISWORKBOOK_FILE)):
        missing.append(THISWORKBOOK_FILE)
    if missing:
        print("ERROR: canonical files missing from", args.repo, ":", ", ".join(missing))
        return 1
    if not os.path.isfile(args.file):
        print("ERROR: target not found:", args.file)
        return 1

    print("Target :", args.file)
    print("Plan   :")
    print("  1. backup workbook -> <name>.bak-<timestamp>")
    print("  2. remove modules :", ", ".join(OLD_MODULES))
    print("  3. import modules :", ", ".join(f for f, _ in STD_MODULES))
    print("  4. set ThisWorkbook code from", THISWORKBOOK_FILE)
    print("  (ribbon customUI14.xml / icons are NOT touched - use the Custom UI Editor)")

    if not args.apply:
        try:
            import win32com.client  # noqa: F401
            print("\npywin32: available")
        except Exception:
            print("\npywin32: NOT installed (needed for --apply): pip install pywin32")
        print("\nDRY RUN - nothing changed. Re-run with --apply to install.")
        return 0

    try:
        import win32com.client
    except Exception as e:
        print("ERROR: pywin32 required for --apply:", e)
        return 1

    backup = "%s.bak-%s" % (args.file, time.strftime("%Y%m%d_%H%M%S"))
    shutil.copy2(args.file, backup)
    print("\nBackup :", backup)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(os.path.abspath(args.file))
        try:
            vbproj = wb.VBProject
        except Exception:
            print("ERROR: cannot access the VBA project. Enable Excel Trust Center"
                  " -> 'Trust access to the VBA project object model' and retry.")
            wb.Close(SaveChanges=False)
            return 1

        comps = vbproj.VBComponents
        for name in OLD_MODULES:
            try:
                comps.Remove(comps(name))
                print("  removed", name)
            except Exception:
                pass  # not present

        for fname, _module in STD_MODULES:
            comps.Import(os.path.join(args.repo, fname))
            print("  imported", fname)

        # ThisWorkbook is a document component - set its code in place.
        tw = None
        for c in comps:
            if c.Type == VBEXT_CT_DOCUMENT and c.Name == "ThisWorkbook":
                tw = c
                break
        if tw is None:
            print("WARNING: ThisWorkbook component not found; skipped its code")
        else:
            cm = tw.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(thisworkbook_body(args.repo))
            print("  set ThisWorkbook code")

        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
        print("\nDONE. Reopen the workbook to fire Workbook_Open, then run"
              " verify_sidecar.py to confirm all modules match.")
        return 0
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()


if __name__ == "__main__":
    sys.exit(main())
