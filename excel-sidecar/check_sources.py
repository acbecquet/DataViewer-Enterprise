#!/usr/bin/env python3
"""Headless invariant checks for the excel-sidecar VBA/ribbon SOURCES.

Validates design invariants checkable without Excel: the crash-fix structure,
the Review/Delete-All procedures, the Review-name table, and that every ribbon
onAction has a matching VBA handler. Run with any Python.

    python excel-sidecar/check_sources.py
Exit 0 if all invariants hold, 1 otherwise.
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))


def rd(name):
    return open(os.path.join(HERE, name), encoding="utf-8").read()


results = []


def check(name, cond, detail=""):
    results.append((bool(cond), name, detail))


dvu = rd("DataViewerUpload.bas")
tt = rd("TestingTools.bas")
twb = rd("ThisWorkbook.cls.txt")
sn = rd("SampleNav.bas")
ui = rd("customUI14.xml")

# --- ThisWorkbook: crash / save-loop fixes ---
check("ThisWorkbook drops dead double-click handler",
      "Workbook_SheetBeforeDoubleClick" not in twb)
check("Workbook_Open sets ThisWorkbook.Saved=True",
      re.search(r"Workbook_Open\(\).*?ThisWorkbook\.Saved\s*=\s*True.*?End Sub",
                twb, re.S) is not None)
check("SheetChange dispatches picker then DataViewer handler",
      re.search(r"TestingTools\.TryPuffStepPicker\(Sh,\s*Target\)", twb) is not None
      and re.search(r"DataViewerUpload\.OnWorkbookSheetChange\s+Sh,\s*Target", twb)
      is not None)

# --- TestingTools: crash-safe picker ---
check("TestingTools defines TryPuffStepPicker",
      re.search(r"Public\s+Function\s+TryPuffStepPicker\b", tt) is not None)
m = re.search(r"Public\s+Function\s+TryPuffStepPicker\b.*?End Function", tt, re.S)
pbody = m.group(0) if m else ""
check("Picker guarantees EnableEvents restore (On Error GoTo PuffDone + label)",
      "On Error GoTo PuffDone" in pbody
      and re.search(r"PuffDone:\s*\r?\n\s*Application\.EnableEvents\s*=\s*savedEvents",
                    pbody) is not None)
check("Picker disables events exactly once and restores via saved state",
      pbody.count("Application.EnableEvents = False") == 1
      and "Application.EnableEvents = savedEvents" in pbody)

# --- DataViewerUpload: Review + reset + Delete-All + ribbon wrappers ---
for token in ["Public Sub DeleteAllReviewSheets",
              "Private Function ReviewBaseName",
              "Private Function UniqueReviewName",
              "Private Function IsReviewSheet",
              "Private Sub ResetSheetToBlankWithReview"]:
    check("DataViewerUpload defines `%s`" % token.split()[-1], token in dvu)
check("ResetLiveWorkbookAfterUpload calls ResetSheetToBlankWithReview",
      "ResetSheetToBlankWithReview ThisWorkbook" in dvu)
check("Old RestoreSheetFromTemplate removed",
      "Sub RestoreSheetFromTemplate" not in dvu)
for w in ["Ribbon_UploadAll", "Ribbon_DryRun", "Ribbon_PickSynology",
          "Ribbon_PickLocal", "Ribbon_DeleteReviewSheets"]:
    check("DataViewerUpload defines wrapper `%s`" % w,
          re.search(r"Public\s+Sub\s+%s\s*\(\s*control\s+As\s+IRibbonControl" % w,
                    dvu) is not None)

# --- Review base-name table: length + width vs CanonicalDataSheets ---
def vba_array(text, func):
    mm = re.search(r"Function\s+%s\b.*?Array\((.*?)\)" % func, text, re.S)
    return re.findall(r'"([^"]*)"', mm.group(1)) if mm else []

canon = vba_array(dvu, "CanonicalDataSheets")
rbase = vba_array(dvu, "ReviewBaseName")
check("ReviewBaseName has one entry per canonical sheet (12)",
      len(rbase) == len(canon) == 12, "canon=%d rbase=%d" % (len(canon), len(rbase)))
toolong = [b for b in rbase if len(b) + len(" - Review") > 31]
check("Every Review base fits '<base> - Review' in 31 chars", not toolong, str(toolong))

# --- Ribbon callback <-> VBA handler cross-check ---
# Collect every Public Sub name (any first-parameter name) across the VBA
# sources. Broader than the old `(\s*control` form so callbacks whose first
# argument is not `control` (e.g. Ribbon_OnLoad(ribbon As IRibbonUI)) resolve.
all_vba = "\n".join([dvu, tt, sn])
handlers = set(re.findall(r"Public\s+Sub\s+(\w+)\s*\(", all_vba))
# Every callback attribute the ribbon can reference, not just onAction.
callbacks = set()
for attr in ("onAction", "getText", "getSupertip", "onChange", "onLoad"):
    callbacks |= set(re.findall(r'%s="([^"]+)"' % attr, ui))
missing = sorted(callbacks - handlers)
check("Every ribbon callback has a matching Public Sub", not missing,
      "missing=%s" % missing)
# Spot-check the new redesign callbacks are present in the collected set so a
# silently-renamed attribute is caught even if some other sub happens to exist.
new_callbacks = ["Ribbon_OnLoad", "GetSynPathText", "GetLocPathText",
                 "GetExePathText", "GetSynPathTip", "GetLocPathTip",
                 "GetExePathTip", "PathReadOnly", "Ribbon_PickDataViewer",
                 "Ribbon_Instructions"]
new_missing = sorted(c for c in new_callbacks if c not in callbacks)
check("Ribbon references all new redesign callbacks", not new_missing,
      "absent from ribbon=%s" % new_missing)
new_unhandled = sorted(c for c in new_callbacks if c not in handlers)
check("All new redesign callbacks resolve to a Public Sub", not new_unhandled,
      "no handler=%s" % new_unhandled)
check("customUI has the 3rd 'DataViewer Upload' group",
      'label="DataViewer Upload"' in ui)

# --- Ribbon structure (Test Selection redesign) ---
check("customUI onLoad is Ribbon_OnLoad",
      re.search(r'<customUI\b[^>]*\bonLoad="Ribbon_OnLoad"', ui, re.S) is not None)
help_idx = ui.find('id="grpHelp"')
upload_idx = ui.find('id="grpDvUpload"')
check("grpHelp appears before grpDvUpload",
      help_idx != -1 and upload_idx != -1 and help_idx < upload_idx,
      "grpHelp=%d grpDvUpload=%d" % (help_idx, upload_idx))


def ribbon_element(tag, el_id):
    # Return the full opening tag (attributes) for <tag id="el_id" ...>.
    m = re.search(r'<%s\b[^>]*\bid="%s"[^>]*>' % (tag, re.escape(el_id)), ui, re.S)
    return m.group(0) if m else ""


upload_btn = ribbon_element("button", "btnUploadAll")
dryrun_btn = ribbon_element("button", "btnDryRun")
check("btnUploadAll has no image/imageMso",
      bool(upload_btn) and "image" not in upload_btn,
      repr(upload_btn))
check("btnDryRun has no image/imageMso",
      bool(dryrun_btn) and "image" not in dryrun_btn,
      repr(dryrun_btn))
delrev_btn = ribbon_element("button", "btnDelReviews")
check('btnDelReviews is size="large"',
      'size="large"' in delrev_btn, repr(delrev_btn))
for eb in ("ebSynPath", "ebLocPath", "ebExePath"):
    check("editBox %s present" % eb, ('id="%s"' % eb) in ui)
# Each <box> holds at most 3 child <button>s (the <=3-rows rule).
oversized = []
for m in re.finditer(r"<box\b[^>]*\bid=\"([^\"]+)\".*?</box>", ui, re.S):
    nbtn = len(re.findall(r"<button\b", m.group(0)))
    if nbtn > 3:
        oversized.append("%s=%d" % (m.group(1), nbtn))
check("Every <box> has at most 3 buttons", not oversized, str(oversized))


# --- VBA invariants (Test Selection redesign) ---
def vba_block(text, kind, name):
    # Non-greedy body of a Sub/Function up to its first matching End.
    m = re.search(r"\b%s\s+%s\b.*?End\s+%s" % (kind, re.escape(name), kind),
                  text, re.S | re.I)
    return m.group(0) if m else ""


check('UPLOAD_SHEET_NAME constant is "Test Selection"',
      re.search(r'UPLOAD_SHEET_NAME\s+As\s+String\s*=\s*"Test Selection"', dvu)
      is not None)

check("PromptForFileName defined",
      re.search(r"Function\s+PromptForFileName\b", dvu) is not None)
upload_body = vba_block(dvu, "Sub", "Btn_UploadAll")
check("Btn_UploadAll calls PromptForFileName",
      "PromptForFileName" in upload_body)

checklist_body = vba_block(dvu, "Function", "RunChecklist")
check("RunChecklist no longer checks DV_FileName is empty",
      bool(checklist_body) and "DV_FileName is empty" not in checklist_body)

dryrun_body = vba_block(dvu, "Sub", "Btn_DryRunChecklist")
check("Btn_DryRunChecklist shows a MsgBox", "MsgBox" in dryrun_body)
check("Btn_UploadAll shows a MsgBox", "MsgBox" in upload_body)
check("ShowFailures defined",
      re.search(r"Sub\s+ShowFailures\b", dvu) is not None)

isdefault_body = vba_block(dvu, "Function", "IsDefaultVisibleSheet")
check("IsDefaultVisibleSheet defined and references SOPS_SHEET_NAME",
      bool(isdefault_body) and "SOPS_SHEET_NAME" in isdefault_body)

keep_body = vba_block(dvu, "Function", "BuildKeepList")
check('BuildKeepList still adds Test SOP\'s (keep.Add SOPS_SHEET_NAME, True)',
      "keep.Add SOPS_SHEET_NAME, True" in keep_body)

check('EnsureVeryHidden "_Settings" present',
      'EnsureVeryHidden "_Settings"' in dvu)

check("Public gRibbon As IRibbonUI declared",
      re.search(r"Public\s+gRibbon\s+As\s+IRibbonUI", dvu) is not None)
check("RefreshPathLabels defined",
      re.search(r"Sub\s+RefreshPathLabels\b", dvu) is not None)
picksyn_body = vba_block(dvu, "Sub", "Btn_PickSynologyFolder")
pickloc_body = vba_block(dvu, "Sub", "Btn_PickLocalFolder")
check("Btn_PickSynologyFolder calls RefreshPathLabels",
      "RefreshPathLabels" in picksyn_body)
check("Btn_PickLocalFolder calls RefreshPathLabels",
      "RefreshPathLabels" in pickloc_body)

instr_body = vba_block(dvu, "Function", "InstructionsText")
check("InstructionsText defined and mentions 'always included in every upload'",
      bool(instr_body) and "always included in every upload" in instr_body)

# --- report ---
ok = True
for passed, name, detail in results:
    line = ("PASS " if passed else "FAIL ") + name
    if detail and not passed:
        line += "  [%s]" % detail
    print(line)
    ok = ok and passed
print("-" * 60)
print("RESULT:", "ALL PASS" if ok else "FAILURES")
sys.exit(0 if ok else 1)
