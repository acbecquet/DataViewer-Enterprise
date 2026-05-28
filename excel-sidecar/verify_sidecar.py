#!/usr/bin/env python3
"""Drift detector for the DataViewer upload sidecar.

Extracts the VBA from a deployed `Automated Testing Template.xlsm` and diffs
each module against the canonical `.bas` / `ThisWorkbook.cls.txt` in this
folder. This is the anti-divergence poka-yoke: the repo is the source of truth,
and this tells you whether the workbook actually matches it.

    python verify_sidecar.py --file "C:\\path\\to\\Automated Testing Template.xlsm"

Exit code 0 if every module matches, 1 if any differ / are missing.

Must be run with the MIP-allowlisted Python so the encrypted .xlsm reads as
plaintext. Requires `oletools` (olevba).
"""
import argparse
import difflib
import os
import sys

# deployed VBA module name -> repo file. Module1 is the pre-rename alias of
# TestingTools, so accept either name for that slot.
MODULE_MAP = {
    "DataViewerUpload": "DataViewerUpload.bas",
    "TestingTools": "TestingTools.bas",
    "Module1": "TestingTools.bas",
    "SampleNav": "SampleNav.bas",
    "ThisWorkbook": "ThisWorkbook.cls.txt",
}


def normalize(code: str) -> str:
    """Compare code from the first `Option Explicit` onward, ignoring the
    Attribute header, an explanatory comment header, trailing whitespace, and
    trailing blank lines. This makes the Module1->TestingTools rename and the
    repo's pasted ThisWorkbook header invisible to the diff."""
    lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    lines = [l for l in lines if not l.startswith("Attribute ")]
    start = next((i for i, l in enumerate(lines)
                  if l.strip() == "Option Explicit"), 0)
    lines = [l.rstrip() for l in lines[start:]]
    while lines and lines[-1] == "":
        lines.pop()
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--file", required=True, help="path to the .xlsm to check")
    ap.add_argument("--repo", default=os.path.dirname(os.path.abspath(__file__)),
                    help="folder with the canonical files (default: this folder)")
    args = ap.parse_args()

    if not os.path.isfile(args.file):
        print("ERROR: file not found:", args.file)
        return 1
    try:
        from oletools.olevba import VBA_Parser
    except Exception as e:  # pragma: no cover
        print("ERROR: oletools (olevba) not available:", e)
        return 1

    # Extract deployed modules.
    deployed = {}
    vp = VBA_Parser(args.file)
    try:
        for _fn, _stream, vba_name, code in vp.extract_macros():
            deployed[os.path.splitext(vba_name)[0]] = code
    finally:
        vp.close()

    print("Deployed modules:", ", ".join(sorted(deployed)) or "(none)")
    print("-" * 60)

    any_diff = False
    checked_files = set()
    for dep_name, code in sorted(deployed.items()):
        repo_file = MODULE_MAP.get(dep_name)
        if repo_file is None:
            # Sheet code modules (Sheet1.cls etc.) are stubs - skip unless they
            # carry real code.
            if normalize(code):
                print(f"[?] {dep_name}: in workbook but unmanaged by the repo "
                      f"(has code) - review")
                any_diff = True
            continue
        repo_path = os.path.join(args.repo, repo_file)
        checked_files.add(repo_file)
        if not os.path.isfile(repo_path):
            print(f"[!] {dep_name}: repo file {repo_file} missing")
            any_diff = True
            continue
        dep_n = normalize(code)
        repo_n = normalize(open(repo_path, encoding="utf-8").read())
        if dep_n == repo_n:
            print(f"[OK] {dep_name:18} == {repo_file}")
        else:
            any_diff = True
            print(f"[DIFFERS] {dep_name:18} != {repo_file}")
            diff = difflib.unified_diff(dep_n.split("\n"), repo_n.split("\n"),
                                        "deployed", "repo", lineterm="", n=0)
            shown = [l for l in diff if l.startswith(("+", "-"))
                     and not l.startswith(("+++", "---"))]
            for l in shown[:24]:
                print("    " + l)
            if len(shown) > 24:
                print(f"    ... ({len(shown) - 24} more changed lines)")

    # repo files with no matching deployed module
    for repo_file in sorted(set(MODULE_MAP.values()) - checked_files):
        print(f"[!] {repo_file}: no matching module found in the workbook")
        any_diff = True

    print("-" * 60)
    print("RESULT:", "DRIFT DETECTED" if any_diff else "all modules match")
    return 1 if any_diff else 0


if __name__ == "__main__":
    sys.exit(main())
