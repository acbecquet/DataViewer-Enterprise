#!/usr/bin/env python3
"""Build a CLEAN Automated Testing Template .xlsm from a (possibly messy)
source, on the WORK MACHINE via Excel COM (pywin32).

Follows this codebase's proven single-workbook pattern (cf. TrimSheetsInWorkbook
/ MakeTempXlsx): copy the source file, open the copy, delete the non-kept sheets
in place, stamp canonical sheets blank from their snapshots, remove the on-sheet
buttons, import the canonical VBA, then Save -- Excel rewrites the whole package
(regenerating calcChain, dropping orphaned parts). Finally, at the zip level, it
strips the embedded web add-in and swaps in the repo ribbon. There is NO
cross-workbook Worksheet.Copy (that fails with Excel error 1004 under COM).

    python excel-sidecar/build_clean_template.py --source "C:\\...\\Automated Testing Template v1.xlsm" --out "C:\\...\\Automated Testing Template v1 (clean).xlsm"

Requires Windows + Excel + pywin32, and Excel Trust Center ->
"Trust access to the VBA project object model" enabled.
"""
import argparse
import os
import re
import shutil
import sys
import zipfile

REPO = os.path.dirname(os.path.abspath(__file__))
CANON = [
    "Lifetime Test", "User Test Simulation", "Long Puff Lifetime Test",
    "Rapid Puff Lifetime Test", "Intense Test", "Big Headspace Serial Test",
    "Negative Pressure Test", "Temperature Cycling Test #2",
    "Viscosity Compatibility", "Various Oil Compatibility",
    "Custom Test Template", "Temperature Cycling Test #1",
]
SNAPSHOTS = ["_Template_%02d" % i for i in range(12)]
KEEP = ["Test SOP's", "DataViewer Upload"] + CANON + ["_Template_Master"] + SNAPSHOTS
NAMED = {  # name -> cell ref on the DataViewer Upload sheet
    "DV_FileName": "$I$6", "DV_SynologyPath": "$I$8", "DV_LocalPath": "$I$10",
    "DV_Status": "$H$16", "DV_Log": "$H$17", "DV_DataViewerExe": "$I$12",
    "DV_TestSelection": "$A$3:$B$16",
}
UPLOAD_SHEET = "DataViewer Upload"
XL_OPENXML_MACRO = 52   # xlOpenXMLWorkbookMacroEnabled (.xlsm)
XL_VERYHIDDEN = 2
XL_HIDDEN = 0
XL_VISIBLE = -1


def check_preconditions(source):
    problems = []
    if not os.path.isfile(source):
        problems.append("source not found: %s" % source)
    try:
        import win32com.client  # noqa: F401
    except Exception:
        problems.append("pywin32 not installed (run: pip install pywin32)")
    return problems


def _grab(pattern, text, label):
    m = re.search(pattern, text)
    if not m:
        raise RuntimeError("could not find %s in source package" % label)
    return m.group(0)


def inject_customui(target_xlsm, repo_dir, scaffold_src):
    """At the zip level: strip the embedded web add-in (xl/webextensions/* parts,
    its package relationship in _rels/.rels, and its content-type overrides), then
    make the ribbon the repo customUI14.xml -- lifting the relationship + icon
    wiring from `scaffold_src` only if the package doesn't already carry them.
    Pure zip surgery (no Excel). Idempotent; safe whether or not the package
    already has customUI."""
    new_ui = open(os.path.join(repo_dir, "customUI14.xml"), "rb").read()
    with zipfile.ZipFile(scaffold_src) as zs:
        sn = set(zs.namelist())
        src_root_rels = zs.read("_rels/.rels").decode("utf-8")
        src_ctypes = zs.read("[Content_Types].xml").decode("utf-8")
        ui_rels = zs.read("customUI/_rels/customUI14.xml.rels") \
            if "customUI/_rels/customUI14.xml.rels" in sn else None
        images = {n: zs.read(n) for n in sn if n.startswith("customUI/images/")}
    # The exact customUI <Relationship .../> + png <Default>, copied verbatim from
    # a workbook that already works. customUI14.xml needs NO content-type Override:
    # it is an .xml part, covered by the package's <Default Extension="xml"/>.
    rel = _grab(r'<Relationship[^>]*customUI/customUI14\.xml[^>]*/>',
                src_root_rels, "customUI relationship")
    png_default = _grab(r'<Default[^>]*Extension="png"[^>]*/>',
                        src_ctypes, "png content-type default")

    tmp = target_xlsm + ".tmp"
    written = set()
    with zipfile.ZipFile(target_xlsm) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            n = item.filename
            if n.startswith("xl/webextensions/"):
                continue                       # strip the web add-in parts
            if n == "customUI/customUI14.xml":
                continue                       # replaced with the repo copy below
            data = zin.read(n)
            if n == "_rels/.rels":
                s = data.decode("utf-8")
                s = re.sub(r'<Relationship[^>]*[Ww]ebextension[^>]*/>', '', s)
                if "customUI/customUI14.xml" not in s:
                    s = s.replace("</Relationships>", rel + "</Relationships>")
                data = s.encode("utf-8")
            elif n == "[Content_Types].xml":
                s = data.decode("utf-8")
                s = re.sub(r'<Override[^>]*[Ww]ebextension[^>]*/>', '', s)
                if 'Extension="png"' not in s:
                    s = s.replace("</Types>", png_default + "</Types>")
                data = s.encode("utf-8")
            zout.writestr(item, data)
            written.add(n)
        zout.writestr("customUI/customUI14.xml", new_ui)
        if "customUI/_rels/customUI14.xml.rels" not in written and ui_rels is not None:
            zout.writestr("customUI/_rels/customUI14.xml.rels", ui_rels)
        for n, b in images.items():
            if n not in written:
                zout.writestr(n, b)
    os.replace(tmp, target_xlsm)

    with zipfile.ZipFile(target_xlsm) as z:
        out = z.namelist()
        root_rels = z.read("_rels/.rels").decode("utf-8")
    assert "customUI/customUI14.xml" in out, "customUI not injected"
    assert not any(n.startswith("xl/webextensions/") for n in out), \
        "web add-in parts survived"
    assert "webextension" not in root_rels.lower(), \
        "_rels/.rels still references the web add-in"


def build(source, out):
    import win32com.client
    backup = source + ".bak"
    shutil.copyfile(source, backup)
    print("Backed up source ->", backup)
    if os.path.exists(out):
        os.remove(out)
    shutil.copyfile(source, out)          # work on a copy; never touch the source
    print("Working copy ->", out)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable (no macro prompts)
    wb = None
    try:
        # VBA object-model trust is required to import modules.
        try:
            _ = xl.VBE
        except Exception:
            raise RuntimeError(
                "Excel Trust Center -> Macro Settings -> 'Trust access to the "
                "VBA project object model' must be enabled.")
        xl.ScreenUpdating = False

        wb = xl.Workbooks.Open(out, UpdateLinks=0)   # open the copy in place
        print("Opened working copy:", wb.Worksheets.Count, "sheets")

        # 1) Delete every sheet not in KEEP (single-workbook; the proven pattern).
        keepset = set(KEEP)
        for s in list(wb.Worksheets):
            if s.Name not in keepset:
                print("  drop sheet:", s.Name)
                s.Visible = XL_VISIBLE          # a hidden sheet can't be the active one to delete
                s.Delete()

        # 2) Break any stray external links (safety net; sheets are self-contained).
        links = None
        try:
            links = wb.LinkSources(1)           # xlExcelLinks
        except Exception:
            links = None
        if links:
            if isinstance(links, str):
                links = [links]
            for lk in links:
                try:
                    wb.BreakLink(lk, 1)
                except Exception as e:
                    print("WARNING: could not break external link %r: %s" % (lk, e))

        # 3) Stamp each canonical sheet blank from its snapshot (single-workbook).
        present = set(s.Name for s in wb.Worksheets)
        for i, name in enumerate(CANON):
            tpl = "_Template_%02d" % i
            if name in present and tpl in present:
                d = wb.Worksheets(name)
                t = wb.Worksheets(tpl)
                d.Cells.Clear()
                t.UsedRange.Copy(d.Range("A1"))

        # 4) The DV_* names came with the copy (already workbook-scoped). Recreate
        #    them to be certain they point at the right cells; no sheet copy
        #    happened, so there are no sheet-local duplicates to clean up.
        up = wb.Worksheets(UPLOAD_SHEET)
        for nm, ref in NAMED.items():
            try:
                wb.Names(nm).Delete()
            except Exception:
                pass
            wb.Names.Add(nm, "='%s'!%s" % (UPLOAD_SHEET, ref))

        # 5) Remove the on-sheet upload buttons (the ribbon hosts them now).
        for shp in list(up.Shapes):
            try:
                if "Btn_" in (shp.OnAction or ""):
                    shp.Delete()
            except Exception:
                pass

        # 6) Visibility default + import the canonical VBA.
        _set_default_visibility(wb)
        _import_vba(wb)

        # 7) Save -- Excel rewrites the package cleanly (OUT is already .xlsm).
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = None
    finally:
        try:
            xl.ScreenUpdating = True
        except Exception:
            pass
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        xl.Quit()

    # 8) Strip the web add-in + swap in the repo ribbon at the zip level.
    inject_customui(out, REPO, source)
    print("Built clean workbook ->", out)


def _set_default_visibility(wb):
    # Sets the at-rest (pre-first-open) visibility only; Workbook_Open ->
    # ApplySheetVisibility re-derives it from DV_TestSelection thereafter.
    visible = {"DataViewer Upload", "Test SOP's", "Lifetime Test"}
    for s in wb.Worksheets:
        n = s.Name
        if n.startswith("_Template_") or n.startswith("_Macro"):
            s.Visible = XL_VERYHIDDEN
        elif n in visible:
            s.Visible = XL_VISIBLE
        else:
            s.Visible = XL_HIDDEN


def _import_vba(wb):
    proj = wb.VBProject
    # Remove any standard modules with our names, then import fresh.
    for comp in list(proj.VBComponents):
        if comp.Type == 1 and comp.Name in ("DataViewerUpload", "TestingTools",
                                             "SampleNav", "Module1"):
            proj.VBComponents.Remove(comp)
    for basfile in ("DataViewerUpload.bas", "TestingTools.bas", "SampleNav.bas"):
        proj.VBComponents.Import(os.path.join(REPO, basfile))
    # ThisWorkbook is a document module: replace its code (don't add a component).
    twb = proj.VBComponents("ThisWorkbook").CodeModule
    twb.DeleteLines(1, twb.CountOfLines)
    body = open(os.path.join(REPO, "ThisWorkbook.cls.txt"), encoding="utf-8").read()
    twb.AddFromString(body)


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--source", required=True, help="messy source .xlsm")
    ap.add_argument("--out", required=True, help="clean output .xlsm to write")
    args = ap.parse_args()
    problems = check_preconditions(args.source)
    if problems:
        print("Preconditions failed:")
        for p in problems:
            print("  -", p)
        return 1
    build(args.source, args.out)
    print("Done. Now run: python excel-sidecar/verify_sidecar.py --file \"%s\"" % args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
