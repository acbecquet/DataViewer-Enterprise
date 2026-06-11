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
import time
import zipfile

REPO = os.path.dirname(os.path.abspath(__file__))
UPLOAD_SHEET = "Test Selection"
OLD_UPLOAD_SHEET = "DataViewer Upload"
SETTINGS_SHEET = "_Settings"
CANON = [
    "Lifetime Test", "User Test Simulation", "Long Puff Lifetime Test",
    "Rapid Puff Lifetime Test", "Intense Test", "Big Headspace Serial Test",
    "Negative Pressure Test", "Temperature Cycling Test #2",
    "Viscosity Compatibility", "Various Oil Compatibility",
    "Custom Test Template", "Temperature Cycling Test #1",
]
SNAPSHOTS = ["_Template_%02d" % i for i in range(12)]
KEEP = ["Test SOP's", UPLOAD_SHEET, SETTINGS_SHEET] + CANON + ["_Template_Master"] + SNAPSHOTS

# Display order on the Test Selection sheet: (sheet name, default checked)
SELECTION_ROWS = [
    ("Custom Test Template", False), ("Lifetime Test", True),
    ("Long Puff Lifetime Test", False), ("Rapid Puff Lifetime Test", False),
    ("Intense Test", False), ("User Test Simulation", False),
    ("Big Headspace Serial Test", False), ("Viscosity Compatibility", False),
    ("Various Oil Compatibility", False), ("Temperature Cycling Test #1", False),
    ("Temperature Cycling Test #2", False), ("Negative Pressure Test", False),
    ("Test SOP's", True),
]

# --- Test Selection table geometry (1-based, single-sourced) -------------------
# A top-row + left-column margin keeps the checkbox table off the sheet's
# top-left corner so it is clearly visible the moment the workbook opens. Bump the
# two margins to move the whole table: the named range, the value writes, the
# clears, and verify_sidecar.py all derive from these constants, so the layout
# can never drift between the build and its checker. (Accepted layout: a 1-row /
# 1-col margin -> title B2, hint B3, table B4:C16.)
TS_MARGIN_ROWS = 1
TS_MARGIN_COLS = 1
TS_TITLE_ROW = TS_MARGIN_ROWS + 1                       # "TEST SELECTION"
TS_HINT_ROW = TS_MARGIN_ROWS + 2                        # "Check the tests..."
TS_FIRST_ROW = TS_MARGIN_ROWS + 3                       # first checkbox/label row
TS_LAST_ROW = TS_FIRST_ROW + len(SELECTION_ROWS) - 1    # last checkbox/label row
TS_CHECK_COL = TS_MARGIN_COLS + 1                       # checkbox column
TS_LABEL_COL = TS_MARGIN_COLS + 2                       # label column

# P5: on-sheet banner under the selection table -- the only guidance channel
# that still works when macros are DISABLED.
TS_BANNER_FIRST = TS_LAST_ROW + 2
TS_BANNER_LINES = [
    "Macros required: if the ribbon buttons do nothing, click 'Enable Content' in the yellow bar.",
    "This is your reusable WORKING TEMPLATE - never email or upload this file.",
    "Upload All / Upload Checkpoint deliver your data to Synology + DataViewer automatically.",
]


def _col_letter(n):
    """1 -> 'A', 2 -> 'B', 27 -> 'AA' ..."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# v1.1 Feature 3: Clog auto-formula geometry (1-based column within a 12-col block)
BLOCK_COLS = 12
DRAW_REL = 4          # Draw Pressure (kpa)  -> block start + 3
CLOG_REL = 7          # Clog                 -> block start + 6
CLOG_ROW_FIRST = 5
CLOG_ROW_LAST = 115   # = BLOCK_ROWS canonical data extent


def clog_formula(draw_col_letter, row):
    """Clog formula keyed off the same-row Draw Pressure cell. ISNUMBER guard:
    text like 'n/a' sorts above all numbers in Excel and would otherwise
    classify as Heavy Clog (audit M-d)."""
    d = "%s%d" % (draw_col_letter, row)
    return ('=IF(NOT(ISNUMBER(%s)),"",IF(%s>=15,"Heavy Clog",IF(%s>5,"Light Clog","")))'
            % (d, d, d))


def apply_clog_formatting(wb):
    """Bake the Clog auto-formula + 2 conditional-format rules onto every 12-col
    block (Clog col = start+6) keyed off Draw Pressure (start+3), on every sheet
    whose contiguous block-starts read 'puffs' -- live data sheets, _Template_NN
    snapshots, and _Template_Master. Skips step-checklist / prose sheets. Idempotent."""
    YELLOW, BLACK, DARKRED, WHITE = 0x00FFFF, 0x000000, 0x0000C0, 0xFFFFFF  # BGR longs
    XL_CELLVALUE, XL_EQUAL = 1, 3
    for ws in wb.Worksheets:
        nblocks, col = 0, 1
        while col <= ws.Columns.Count:
            try:
                v = ws.Cells(4, col).Value
            except Exception:
                v = None
            if isinstance(v, str) and v.strip().lower() == "puffs":
                nblocks += 1
                col += BLOCK_COLS
            else:
                break
        if nblocks == 0:
            continue
        clog_ranges = []
        for b in range(nblocks):
            start = b * BLOCK_COLS + 1
            draw_letter = _col_letter(start + DRAW_REL - 1)
            clog_col = start + CLOG_REL - 1
            rng = ws.Range(ws.Cells(CLOG_ROW_FIRST, clog_col), ws.Cells(CLOG_ROW_LAST, clog_col))
            rng.Formula = clog_formula(draw_letter, CLOG_ROW_FIRST)  # Excel shifts row refs down
            clog_ranges.append(rng)
        union = clog_ranges[0]
        for rng in clog_ranges[1:]:
            union = wb.Application.Union(union, rng)
        union.FormatConditions.Delete()
        fc1 = union.FormatConditions.Add(XL_CELLVALUE, XL_EQUAL, '="Light Clog"')
        fc1.Interior.Color = YELLOW; fc1.Font.Color = BLACK; fc1.Font.Bold = True
        fc2 = union.FormatConditions.Add(XL_CELLVALUE, XL_EQUAL, '="Heavy Clog"')
        fc2.Interior.Color = DARKRED; fc2.Font.Color = WHITE; fc2.Font.Bold = True


def ts_selection_ref():
    """Sheet-qualified RefersTo for DV_TestSelection, derived from the geometry."""
    return "'%s'!$%s$%d:$%s$%d" % (
        UPLOAD_SHEET, _col_letter(TS_CHECK_COL), TS_FIRST_ROW,
        _col_letter(TS_LABEL_COL), TS_LAST_ROW)


# name -> full (sheet-qualified) RefersTo
NAMED = {
    "DV_TestSelection": ts_selection_ref(),
    "DV_FileName":      "'_Settings'!$B$1",
    "DV_SynologyPath":  "'_Settings'!$B$2",
    "DV_LocalPath":     "'_Settings'!$B$3",
    "DV_DataViewerExe": "'_Settings'!$B$4",
    "DV_Status":        "'_Settings'!$B$5",
    "DV_Log":           "'_Settings'!$B$6",
    "DV_OrigFileName":  "'_Settings'!$B$7",
    "DV_LastUpload":    "'_Settings'!$B$8",
}
SETTINGS_LABELS = ["File name (last used)", "Synology folder", "Local folder",
                   "DataViewer.exe", "Status", "Log", "Original file name",
                   "Last uploaded file"]
XL_OPENXML_MACRO = 52   # xlOpenXMLWorkbookMacroEnabled (.xlsm)
XL_VERYHIDDEN = 2
XL_HIDDEN = 0
XL_VISIBLE = -1

HELP_ICONS = {  # Help-button icons (repo-provided, injected into customUI/images)
    "imgGuide": "imgGuide.png", "imgTools": "imgTools.png", "imgUpload": "imgUpload.png",
}

# Base ribbon icons whose oversized scaffold copies are replaced by small
# repo-owned versions during inject_customui. The scaffold's imgPlus was
# 5120x5120 px, which makes Excel raise "image too large" -- the repo copies
# are 64x64. Their customUI relationships already exist in the scaffold, so we
# only swap the PNG bytes.
BASE_ICON_OVERRIDES = ("imgPlus.png", "imgMinus.png")


def assert_not_mip_ciphertext(path):
    """The deployed .xlsm is MIP-encrypted at rest; only the allowlisted Python
    reads plaintext. Fail loudly instead of propagating ciphertext (audit M-k)."""
    with open(path, "rb") as f:
        head = f.read(16)
    if head.startswith(b"%TSD-Header-"):
        raise SystemExit("FATAL: %s reads as MIP ciphertext under this interpreter.\n"
                         "Run with the allowlisted Python: "
                         "C:/Users/S1134987/AppData/Local/Programs/Python/Python313/python.exe" % path)
    if not head.startswith(b"PK\x03\x04"):
        raise SystemExit("FATAL: %s does not look like an OOXML (.xlsm/.xlsx) package." % path)


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
    wiring from `scaffold_src` only if the package doesn't already carry them, and
    injecting the repo's Help-button icons (HELP_ICONS) + their customUI
    relationships. Pure zip surgery (no Excel). Idempotent; safe whether or not the
    package already has customUI."""
    new_ui = open(os.path.join(repo_dir, "customUI14.xml"), "rb").read()
    # Repo-provided Help-button icons (id -> bytes), injected into customUI/images
    # with a matching relationship. The base ribbon images (imgPlus etc.) still
    # come from scaffold_src below.
    repo_icons = {}
    for icon_id, fn in HELP_ICONS.items():
        p = os.path.join(repo_dir, fn)
        if os.path.isfile(p):
            repo_icons[icon_id] = open(p, "rb").read()

    def _add_icon_rels(rels_text):
        add = ""
        for icon_id, fn in HELP_ICONS.items():
            if icon_id in repo_icons and ('Id="%s"' % icon_id) not in rels_text:
                add += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/'
                        'officeDocument/2006/relationships/image" Target="images/%s"/>'
                        % (icon_id, fn))
        return rels_text.replace("</Relationships>", add + "</Relationships>") \
            if add else rels_text

    with zipfile.ZipFile(scaffold_src) as zs:
        sn = set(zs.namelist())
        src_root_rels = zs.read("_rels/.rels").decode("utf-8")
        src_ctypes = zs.read("[Content_Types].xml").decode("utf-8")
        ui_rels = zs.read("customUI/_rels/customUI14.xml.rels").decode("utf-8") \
            if "customUI/_rels/customUI14.xml.rels" in sn else None
        images = {n: zs.read(n) for n in sn if n.startswith("customUI/images/")}
    # Replace oversized scaffold base icons (e.g. imgPlus 5120x5120 -> Excel
    # "image too large") with the small repo-owned copies, keeping the scaffold's
    # existing icon relationships intact (we only swap the PNG bytes).
    for _ovr in BASE_ICON_OVERRIDES:
        _op = os.path.join(repo_dir, _ovr)
        if os.path.isfile(_op):
            images["customUI/images/%s" % _ovr] = open(_op, "rb").read()
    # Add the Help-icon relationships to the scaffold rels (used only when the
    # target package does not already carry its own customUI rels).
    if ui_rels is not None:
        ui_rels = _add_icon_rels(ui_rels)
    # The exact customUI <Relationship .../> + png <Default>, copied verbatim from
    # a workbook that already works. customUI14.xml needs NO content-type Override:
    # it is an .xml part, covered by the package's <Default Extension="xml"/>.
    rel = _grab(r'<Relationship[^>]*customUI/customUI14\.xml[^>]*/>',
                src_root_rels, "customUI relationship")
    png_default = _grab(r'<Default[^>]*Extension="png"[^>]*/>',
                        src_ctypes, "png content-type default")

    tmp = target_xlsm + ".tmp"
    written = set()
    help_parts = {"customUI/images/%s" % fn for fn in HELP_ICONS.values()}
    override_parts = {"customUI/images/%s" % fn for fn in BASE_ICON_OVERRIDES
                      if os.path.isfile(os.path.join(repo_dir, fn))}
    with zipfile.ZipFile(target_xlsm) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            n = item.filename
            if n.startswith("xl/webextensions/"):
                continue                       # strip the web add-in parts
            if n == "customUI/customUI14.xml":
                continue                       # replaced with the repo copy below
            if n in help_parts:
                continue                       # Help icons are repo-owned (written below)
            if n in override_parts:
                continue                       # base icons overridden by small repo copies (written from `images`)
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
            elif n == "customUI/_rels/customUI14.xml.rels":
                data = _add_icon_rels(data.decode("utf-8")).encode("utf-8")
            zout.writestr(item, data)
            written.add(n)
        zout.writestr("customUI/customUI14.xml", new_ui)
        if "customUI/_rels/customUI14.xml.rels" not in written and ui_rels is not None:
            zout.writestr("customUI/_rels/customUI14.xml.rels", ui_rels.encode("utf-8"))
        for n, b in images.items():
            if n not in written and n not in help_parts:
                zout.writestr(n, b)
                written.add(n)
        for icon_id, b in repo_icons.items():
            part = "customUI/images/%s" % HELP_ICONS[icon_id]
            if part not in written:
                zout.writestr(part, b)
                written.add(part)
    # F12: verify the surgered package BEFORE swapping it in -- on any failure
    # the temp file is removed and the target is left untouched.
    try:
        with zipfile.ZipFile(tmp) as z:
            out = z.namelist()
            root_rels = z.read("_rels/.rels").decode("utf-8")
            ui_rels_out = z.read("customUI/_rels/customUI14.xml.rels").decode("utf-8") \
                if "customUI/_rels/customUI14.xml.rels" in out else ""
        assert "customUI/customUI14.xml" in out, "customUI not injected"
        assert not any(n.startswith("xl/webextensions/") for n in out), \
            "web add-in parts survived"
        assert "webextension" not in root_rels.lower(), \
            "_rels/.rels still references the web add-in"
        for icon_id, fn in HELP_ICONS.items():
            if icon_id in repo_icons:
                assert "customUI/images/%s" % fn in out, "help icon %s missing" % fn
                assert ('Id="%s"' % icon_id) in ui_rels_out, "icon rel %s missing" % icon_id
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    os.replace(tmp, target_xlsm)


def build(source, out, force=False):
    import win32com.client
    assert_not_mip_ciphertext(source)
    if os.path.exists(out):
        if os.path.samefile(source, out):
            raise SystemExit("FATAL: --out is the same file as --source; refusing (audit H6).")
        if not force:
            raise SystemExit("FATAL: --out already exists: %s  (pass --force to overwrite; "
                             "a timestamped backup of it will be kept)" % out)
        out_bak = out + time.strftime(".bak-%Y%m%d_%H%M%S")
        shutil.copyfile(out, out_bak)
        print("Backed up existing --out ->", out_bak)
    backup = source + time.strftime(".bak-%Y%m%d_%H%M%S")
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

        # 0) Read carried path/file values, then rename the upload sheet BEFORE the
        #    delete-non-KEEP step so the renamed 'Test Selection' survives KEEP.
        carried = _carry_over_values(wb)
        _rename_upload_sheet(wb)

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

        # 3a) v1.1 Feature 3: bake the auto-Clog formula + conditional format onto
        #     every 'puffs'-layout block (live sheets, snapshots, and the master).
        apply_clog_formatting(wb)

        # 3b) Build the very-hidden _Settings sheet (carries paths/file name) and
        #     re-lay the Test Selection sheet in place (checkboxes preserved).
        _build_settings_sheet(wb, carried)
        _relay_test_selection(wb)

        # 4) The DV_* names now span two sheets (Test Selection + _Settings).
        #    Recreate them from the full sheet-qualified RefersTo; no sheet copy
        #    happened, so there are no sheet-local duplicates to clean up.
        for nm, ref in NAMED.items():
            try:
                wb.Names(nm).Delete()
            except Exception:
                pass
            wb.Names.Add(nm, "=" + ref)

        # 5) Remove the on-sheet upload buttons (the ribbon hosts them now).
        up = wb.Worksheets(UPLOAD_SHEET)
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


def _carry_over_values(wb):
    """Read current path/file values before restructuring, so the operator does not
    re-pick folders after a rebuild."""
    carried = {}
    for nm in ("DV_FileName", "DV_SynologyPath", "DV_LocalPath", "DV_DataViewerExe",
               "DV_OrigFileName", "DV_LastUpload"):
        try:
            carried[nm] = wb.Names(nm).RefersToRange.Value
        except Exception:
            carried[nm] = ""
    return carried


def _rename_upload_sheet(wb):
    """Rename 'DataViewer Upload' -> 'Test Selection' (idempotent)."""
    for cand in (OLD_UPLOAD_SHEET, UPLOAD_SHEET):
        try:
            wb.Worksheets(cand).Name = UPLOAD_SHEET
            return
        except Exception:
            continue


def _build_settings_sheet(wb, carried):
    try:
        st = wb.Worksheets(SETTINGS_SHEET)
    except Exception:
        st = wb.Worksheets.Add()
        st.Name = SETTINGS_SHEET
    st.Visible = XL_VISIBLE        # very-hidden later by _set_default_visibility
    st.Cells.Clear()
    for i, label in enumerate(SETTINGS_LABELS, start=1):
        st.Cells(i, 1).Value = label
    st.Cells(1, 2).Value = carried.get("DV_FileName", "") or ""
    st.Cells(2, 2).Value = carried.get("DV_SynologyPath", "") or ""
    st.Cells(3, 2).Value = carried.get("DV_LocalPath", "") or ""
    st.Cells(4, 2).Value = carried.get("DV_DataViewerExe", "") or ""
    st.Cells(5, 2).Value = ""     # Status
    st.Cells(6, 2).Value = ""     # Log
    st.Cells(7, 2).Value = carried.get("DV_OrigFileName", "") or ""
    st.Cells(8, 2).Value = carried.get("DV_LastUpload", "") or ""
    st.Columns("A:B").AutoFit()


def _relay_test_selection(wb):
    """Re-lay the Test Selection sheet IN PLACE at the configured geometry (the
    TS_* constants -> a top/left margin so the table is visible on open). Preserve
    the native checkboxes on the checkbox column -- value writes only, never Clear
    that column, since its cell STYLE carries the checkbox control. Column widths
    and row heights are inherited from the source (operator-tuned sizing survives a
    rebuild); only structure + brand styling are (re)applied here."""
    ws = wb.Worksheets(UPLOAD_SHEET)
    cc, lc = TS_CHECK_COL, TS_LABEL_COL
    tr, hr, fr, lr = TS_TITLE_ROW, TS_HINT_ROW, TS_FIRST_ROW, TS_LAST_ROW

    # Clear clutter OUTSIDE the table only -- never the checkbox column cells, whose
    # style carries the native control. Clear() (not ClearContents) so any stray
    # checkbox style outside the table is reset, leaving no orphan checkbox.
    if cc > 1:                                                  # left margin col(s)
        ws.Range(ws.Cells(1, 1), ws.Cells(lr + 200, cc - 1)).Clear()
    if tr > 1:                                                  # top margin row(s)
        ws.Range(ws.Cells(1, cc), ws.Cells(tr - 1, lc)).Clear()
    ws.Range(ws.Cells(lr + 1, 1), ws.Cells(lr + 200, lc)).Clear()       # below table
    ws.Range(ws.Cells(1, lc + 1), ws.Cells(lr + 200, 52)).Clear()       # right of table

    # Title + hint text.
    ws.Cells(tr, cc).Value = "TEST SELECTION"
    ws.Cells(hr, cc).Value = "Check the tests you're running."

    # 13 rows: checkbox value (keeps the native control) + label.
    for i, (name, checked) in enumerate(SELECTION_ROWS):
        ws.Cells(fr + i, cc).Value = bool(checked)
        ws.Cells(fr + i, lc).Value = name

    # P5: the only guidance channel that works with macros DISABLED. Test
    # Selection never enters distributed copies, so it can be loud.
    for i, txt in enumerate(TS_BANNER_LINES):
        ws.Cells(TS_BANNER_FIRST + i, cc).Value = txt

    # ---- Cosmetic styling (guarded -- cosmetics must never abort a build) ----
    NAVY = 31 + 56 * 256 + 100 * 65536        # RGB(31,56,100)   brand header
    HINTBG = 221 + 227 * 256 + 240 * 65536    # RGB(221,227,240) hint band
    BAND = 244 + 247 * 256 + 252 * 65536      # RGB(244,247,252) alt-row band
    GRID = 200 + 210 * 256 + 225 * 65536      # RGB(200,210,225) inner border
    WHITE = 255 + 255 * 256 + 255 * 65536
    GREY = 89 + 89 * 256 + 89 * 65536
    XL_CENTER, XL_LEFT, XL_MEDIUM = -4108, -4131, -4138
    try:
        ws.Cells.Font.Name = "Calibri"
        for c1, c2 in ((ws.Cells(tr, cc), ws.Cells(tr, lc)),   # title spans cols
                       (ws.Cells(hr, cc), ws.Cells(hr, lc))):  # hint spans cols
            try:
                ws.Range(c1, c2).Merge()
            except Exception as _e:
                print("WARNING: merge skipped: %s" % _e)
        title = ws.Cells(tr, cc)
        title.Interior.Color = NAVY
        title.Font.Color = WHITE
        title.Font.Bold = True
        title.Font.Size = 16
        title.HorizontalAlignment = XL_CENTER
        title.VerticalAlignment = XL_CENTER
        hint = ws.Cells(hr, cc)
        hint.Interior.Color = HINTBG
        hint.Font.Italic = True
        hint.Font.Size = 10
        hint.Font.Color = GREY
        hint.HorizontalAlignment = XL_CENTER
        body = ws.Range(ws.Cells(fr, cc), ws.Cells(lr, lc))
        body.Font.Size = 11
        body.VerticalAlignment = XL_CENTER
        ws.Range(ws.Cells(fr, cc), ws.Cells(lr, cc)).HorizontalAlignment = XL_CENTER  # checkbox
        labels = ws.Range(ws.Cells(fr, lc), ws.Cells(lr, lc))
        labels.HorizontalAlignment = XL_LEFT
        labels.IndentLevel = 1
        for i in range(len(SELECTION_ROWS)):       # subtle banding on alternate rows
            if i % 2 == 1:
                ws.Range(ws.Cells(fr + i, cc), ws.Cells(fr + i, lc)).Interior.Color = BAND
        body.Borders.LineStyle = 1                 # thin inner gridlines
        body.Borders.Color = GRID
        box = ws.Range(ws.Cells(tr, cc), ws.Cells(lr, lc))     # title -> last row
        for _edge in (7, 8, 9, 10):                # xlEdgeLeft/Top/Bottom/Right
            b = box.Borders(_edge)
            b.LineStyle = 1
            b.Weight = XL_MEDIUM
            b.Color = NAVY
        DARKRED = 0x0000C0
        for i in range(len(TS_BANNER_LINES)):
            r = TS_BANNER_FIRST + i
            try:
                ws.Range(ws.Cells(r, cc), ws.Cells(r, lc + 6)).Merge()
            except Exception as _e:
                print("WARNING: banner merge skipped: %s" % _e)
            cell = ws.Cells(r, cc)
            cell.Font.Size = 10
            cell.Font.Bold = (i == 1)
            cell.Font.Color = DARKRED if i == 1 else GREY
            cell.HorizontalAlignment = XL_LEFT
        try:                                       # hide gridlines for a clean card look
            ws.Activate()
            ws.Application.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass
    except Exception as _e:
        print("WARNING: Test Selection styling partial:", _e)


def _set_default_visibility(wb):
    visible = {"Test Selection", "Test SOP's", "Lifetime Test"}
    for s in wb.Worksheets:
        n = s.Name
        if n.startswith("_Template_") or n.startswith("_Macro") or n == "_Settings":
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
    ap.add_argument("--force", action="store_true",
                    help="overwrite an existing --out")
    args = ap.parse_args()
    problems = check_preconditions(args.source)
    if problems:
        print("Preconditions failed:")
        for p in problems:
            print("  -", p)
        return 1
    build(args.source, args.out, force=args.force)
    print("Done. Now run: python excel-sidecar/verify_sidecar.py --file \"%s\"" % args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
