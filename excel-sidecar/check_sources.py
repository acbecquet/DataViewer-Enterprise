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

# --- Ribbon onAction <-> VBA handler cross-check ---
handlers = set(re.findall(r"Public\s+Sub\s+(\w+)\s*\(\s*control\b",
                          "\n".join([dvu, tt, sn])))
actions = set(re.findall(r'onAction="([^"]+)"', ui))
missing = sorted(actions - handlers)
check("Every ribbon onAction has a matching VBA handler", not missing,
      "missing=%s" % missing)
check("customUI has the 3rd 'DataViewer Upload' group",
      'label="DataViewer Upload"' in ui)

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
