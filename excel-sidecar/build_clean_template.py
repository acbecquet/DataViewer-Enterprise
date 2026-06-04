#!/usr/bin/env python3
"""Build a CLEAN Automated Testing Template .xlsm from a (possibly messy)
source, on the WORK MACHINE via Excel COM (pywin32).

Creates a brand-new workbook (fresh package: no inherited corruption, no web
add-in, no calcChain rot), copies the source's sheets in with full fidelity,
stamps canonical sheets blank from their snapshots, imports the canonical VBA
from this folder, removes the on-sheet upload buttons, saves, then injects the
ribbon at the zip level.

    python excel-sidecar/build_clean_template.py \
        --source "C:\\...\\Automated Testing Template v1.xlsm" \
        --out    "C:\\...\\Automated Testing Template v1 (clean).xlsm"

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
    """Add the ribbon to a saved .xlsm by lifting the proven customUI wiring
    from `scaffold_src` (a workbook that already has a working customUI14) and
    swapping in the repo customUI14.xml + icons. Pure zip surgery (no Excel).
    Idempotent. Raises if a web add-in survives in the output."""
    new_ui = open(os.path.join(repo_dir, "customUI14.xml"), "rb").read()
    with zipfile.ZipFile(scaffold_src) as zs:
        sn = set(zs.namelist())
        src_root_rels = zs.read("_rels/.rels").decode("utf-8")
        src_ctypes = zs.read("[Content_Types].xml").decode("utf-8")
        ui_rels = zs.read("customUI/_rels/customUI14.xml.rels") \
            if "customUI/_rels/customUI14.xml.rels" in sn else None
        images = {n: zs.read(n) for n in sn if n.startswith("customUI/images/")}
    # The exact <Relationship .../> for customUI14 + the png <Default>, copied
    # verbatim from a workbook that already works. customUI14.xml needs NO
    # content-type Override: it is an .xml part, covered by the package's
    # standard <Default Extension="xml"/> (the source workbook ships this way).
    rel = _grab(r'<Relationship[^>]*customUI/customUI14\.xml[^>]*/>',
                src_root_rels, "customUI relationship")
    png_default = _grab(r'<Default[^>]*Extension="png"[^>]*/>',
                        src_ctypes, "png content-type default")

    tmp = target_xlsm + ".tmp"
    with zipfile.ZipFile(target_xlsm) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            n = item.filename
            if n in ("customUI/customUI14.xml",):
                continue  # re-added below
            data = zin.read(n)
            if n == "_rels/.rels":
                s = data.decode("utf-8")
                if "customUI/customUI14.xml" not in s:
                    s = s.replace("</Relationships>", rel + "</Relationships>")
                data = s.encode("utf-8")
            elif n == "[Content_Types].xml":
                s = data.decode("utf-8")
                # customUI14.xml is covered by the standard <Default
                # Extension="xml"/>; only ensure the png icons are typed.
                if 'Extension="png"' not in s:
                    s = s.replace("</Types>", png_default + "</Types>")
                data = s.encode("utf-8")
            zout.writestr(item, data)
        zout.writestr("customUI/customUI14.xml", new_ui)
        if ui_rels is not None:
            zout.writestr("customUI/_rels/customUI14.xml.rels", ui_rels)
        for n, b in images.items():
            zout.writestr(n, b)
    os.replace(tmp, target_xlsm)

    with zipfile.ZipFile(target_xlsm) as z:
        out = z.namelist()
    assert "customUI/customUI14.xml" in out, "customUI not injected"
    assert not any(n.startswith("xl/webextensions/") for n in out), \
        "web add-in parts present in output"


def build(source, out):
    import win32com.client
    backup = source + ".bak"
    shutil.copyfile(source, backup)
    print("Backed up source ->", backup)

    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3   # msoAutomationSecurityForceDisable (no macro prompts)
    try:
        # VBA object-model trust is required to import modules.
        try:
            _ = xl.VBE
        except Exception:
            raise RuntimeError(
                "Excel Trust Center -> Macro Settings -> 'Trust access to the "
                "VBA project object model' must be enabled.")

        target = xl.Workbooks.Add()
        src = xl.Workbooks.Open(source, ReadOnly=True, UpdateLinks=0)

        # 1) Copy each kept sheet (full fidelity) in canonical order.
        src_names = [s.Name for s in src.Worksheets]
        keep_present = [n for n in KEEP if n in src_names]
        for n in keep_present:
            src.Worksheets(n).Copy(After=target.Worksheets(target.Worksheets.Count))
        # 2) Delete the workbook's original default sheet(s).
        for s in list(target.Worksheets):
            if s.Name not in keep_present:
                s.Delete()
        src.Close(SaveChanges=False)

        # 3) Break any cross-workbook links introduced by the copies.
        try:
            links = target.LinkSources(1)   # xlExcelLinks
            if links:
                for lk in links:
                    target.BreakLink(lk, 1)
        except Exception:
            pass

        # 4) Recreate workbook-scoped named ranges on the upload sheet.
        up = target.Worksheets(UPLOAD_SHEET)
        for nm, ref in NAMED.items():
            try:
                target.Names(nm).Delete()
            except Exception:
                pass
            target.Names.Add(nm, "='%s'!%s" % (UPLOAD_SHEET, ref))

        # 5) Stamp each canonical sheet blank from its snapshot (pristine template).
        for i, name in enumerate(CANON):
            tpl = "_Template_%02d" % i
            if name in [s.Name for s in target.Worksheets] and \
                    tpl in [s.Name for s in target.Worksheets]:
                t = target.Worksheets(tpl)
                d = target.Worksheets(name)
                d.Cells.Clear()
                t.UsedRange.Copy(d.Range("A1"))

        # 6) Remove the on-sheet upload buttons (the ribbon hosts them now).
        for shp in list(up.Shapes):
            try:
                if "Btn_" in (shp.OnAction or ""):
                    shp.Delete()
            except Exception:
                pass

        # 7) Visibility default + import the canonical VBA.
        _set_default_visibility(target)
        _import_vba(target)

        # 8) Save as .xlsm.
        if os.path.exists(out):
            os.remove(out)
        target.SaveAs(out, FileFormat=XL_OPENXML_MACRO)
        target.Close(SaveChanges=False)
    finally:
        xl.Quit()

    # 9) Inject the ribbon at the zip level using the source as scaffold.
    inject_customui(out, REPO, source)
    print("Built clean workbook ->", out)


def _set_default_visibility(wb):
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
