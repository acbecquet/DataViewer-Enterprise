# Excel Upload Sidecar — Clean Rebuild & Hardening

**Date:** 2026-06-04
**Branch:** `feature/excel-sidecar-rebuild` (git worktree off `main` @ v2.2.4)
**Status:** Approved design — ready for an implementation plan
**Relates to:** `excel-sidecar/` (single source of truth), prior reset-from-template design
(`docs/superpowers/specs/2026-05-28-sidecar-reset-from-template-design.md`)

---

## 1. Problem & motivation

The operator's live data-collection workbook —
`C:\Users\S1134987\Documents\Templates\Automated Testing Template v1.xlsm` — has three
problems:

1. **It randomly crashes** and then traps the user in a *save → cancel → close → repeat*
   loop.
2. **Upload All clears the live data sheets** with no in-workbook way to review what was
   just uploaded/cleared.
3. **It has drifted** from the repo `excel-sidecar/` source of truth (modules hand-edited
   only inside the `.xlsm`), so the repo no longer describes the deployed reality.

Goal: a **clean, robust template** that multiple users across multiple networks can run
without crashes or data-loss surprises, with the repo as the provable single source of
truth.

---

## 2. Forensic findings (live file, read-only pass 2026-06-04)

Inspected with the MIP-allowlisted Python (olevba + zipfile) — the live file was never
opened in Excel.

**Structure:** 28 sheets, 2.16 MB. Visible: `Lifetime Test`, `DataViewer Upload`,
`Test SOP's`. Canonical data sheets hidden; `_Template_00`..`_Template_11` +
`_Template_Master` + `_Macro_Install` very-hidden. 7 named ranges
(`DV_FileName`, `DV_SynologyPath`, `DV_LocalPath`, `DV_Status`, `DV_Log`,
`DV_TestSelection` = `'DataViewer Upload'!$A$3:$B$16`, `DV_DataViewerExe`). Upload
actions are **shape buttons on the DataViewer Upload sheet** (`Btn_UploadAll`,
`Btn_DryRunChecklist`, `Btn_PickSynologyFolder`, `Btn_PickLocalFolder`); the ribbon
("TPM Testing" tab) hosts only Sample Blocks + Sample Navigation.

**Drift vs repo:** `SampleNav` matches. `DataViewerUpload`, `TestingTools`, and
`ThisWorkbook` all DIFFER (hand-edited live variants). The live `DataViewerUpload` is
still the snapshot-based safe-reset architecture (good) but with cosmetic/structural
edits. The live `ThisWorkbook` has an extra **Puffs step-picker** not in the repo.

**Crash / save-loop root causes (with confidence):**

1. **Events silently die — HIGH.** `ThisWorkbook.Workbook_SheetChange`'s Puffs
   step-picker does `Application.EnableEvents = False`, writes formulas, then `= True`.
   Its error handler `On Error GoTo AfterPuffs` **never restores `EnableEvents`**, so any
   mid-edit error kills all events for the rest of the session (toggles + picker silently
   stop working; later handlers run with events disabled). This is the "random" flakiness.
2. **Dirty-on-open → save nag — LIKELY.** Open-time churn (re-asserting visibility,
   recalculation, and the embedded add-in re-syncing) marks the workbook modified every
   session, so Excel prompts "Save?" on every close even with no user edits → the
   cancel→close→repeat loop.
3. **Corruption-prone reset — LIKELY contributor to "random crash".** The post-upload
   reset **deletes each live data sheet and recreates it** from a snapshot. Repeated VBA
   sheet delete/insert churn is a classic cause of gradual `.xlsm` corruption → phantom
   "we repaired this file, save?" prompts.
4. **Embedded "Claude for Excel" web add-in — cruft.** `xl/webextensions/` references
   store add-in `WA200010725` (`store="Omex"`, property `claude.fileId`) as a hidden
   task pane. It rides into the distributed copies and can trigger add-in behavior on
   other machines.
5. **Dead code.** `Workbook_SheetBeforeDoubleClick` toggles rows **20–33**, but the real
   selection is rows **3–16** (TRUE/FALSE data-validation dropdowns). The handler does
   nothing but add risk; the workbook also carries a leftover native-checkbox
   `featurePropertyBag` that nothing uses.

**Not data-loss today:** Upload All copies to Synology + Local + DataViewer *before*
clearing, and the reset restores from internal blank snapshots. The "irreversible clear"
is simply that the live sheets are reset with no in-workbook review copy.

---

## 3. Locked decisions

| Decision | Choice |
|---|---|
| Approach | **Rebuild clean from scratch** (Excel-driven, full-fidelity) |
| Review sheets | **Persist until manually deleted** (no auto-clear on upload) |
| Buttons | **Consolidate every action onto the TPM Testing ribbon** |
| Robustness scope | **Stability + clean structure** — paths stay as named ranges |

---

## 4. Goals / non-goals

**Goals**
- Eliminate the crash + save-loop sources (events leak, dirty-on-open, sheet-churn, add-in,
  dead code).
- Make Upload All non-destructive: a `<name> - Review` copy of each cleared sheet,
  persisting until the user bulk-deletes them.
- One ribbon home for every action incl. a new **Delete All Review Sheets**.
- Repo `excel-sidecar/` is the provable single source of truth; `verify_sidecar.py`
  reports all-MATCH against the rebuilt file.
- Distributed copies stay macro-free `.xlsx`, now also free of the web add-in.

**Non-goals**
- Per-machine path auto-detection / per-user path memory (explicitly out of scope).
- Changing the test data model, formulas, or DataViewer ingestion.
- Any Synology drop (the user is the human ship checkpoint — never automated).

---

## 5. Constraints

- **No Excel / no VBA compile in the dev environment.** Final assembly runs on the work
  machine via `pywin32` + Excel (same pattern as the existing `install_sidecar.py`). The
  dev session authors all sources; the user runs one build script + eyeball-tests.
- **MIP/AIP** encrypts Office + source files at rest in `Documents`; the allowlisted
  Python reads plaintext. (The worktree lives outside `Documents`, so authored sources
  avoid MIP churn.)
- **Excel 31-char sheet-name limit** governs Review sheet naming.
- **"Trust access to the VBA project object model"** must be enabled for the build
  script's VBA import (script checks + instructs).
- Distributed copies **must** be macro-free `.xlsx` (standing decision, recorded twice).

---

## 6. Design

### 6.1 Repo as single source of truth (`excel-sidecar/`)

| File | Change |
|---|---|
| `DataViewerUpload.bas` | Reconcile to live + add churn-free reset, Review copies, `DeleteAllReviewSheets`, ribbon wrappers |
| `TestingTools.bas` | Reconcile (one canonical version) |
| `SampleNav.bas` | Unchanged (already matches) |
| `ThisWorkbook.cls.txt` | Hardened Puffs picker (guaranteed `EnableEvents` restore), drop dead double-click handler, add no-edit `Saved=True` guard |
| `customUI14.xml` | Add 3rd ribbon group "DataViewer Upload" (5 buttons) |
| `build_clean_template.py` | **NEW** — pywin32 Excel-driven clean rebuild |
| `verify_sidecar.py` | Extend: cover ribbon + the new procedures |
| `README.md`, `RUNBOOK-*.md` | Update for the rebuild flow + Review feature |

### 6.2 Clean rebuild pipeline (`build_clean_template.py`)

Runs on the work machine. Produces a brand-new workbook → fresh package → **no inherited
corruption, no web add-in, no calcChain rot**, with 100% formatting fidelity because Excel
does the copying.

1. **Preconditions:** `pywin32` import OK; Excel COM available; VBA-object-model trust
   enabled; `--source` exists. Fail with actionable messages otherwise.
2. **Backup** the source to a timestamped `.bak`.
3. Launch Excel hidden (`DisplayAlerts=False`, `AutomationSecurity=ForceDisable`).
4. **Create a NEW workbook** (the clean target).
5. Open the source **read-only**.
6. **Copy** these sheets source→target (Excel `Sheet.Copy`, preserving order + formatting):
   `Test SOP's`, `DataViewer Upload`, the 12 canonical data sheets, `_Template_Master`,
   `_Template_00..11`. (`_Macro_Install` is a legacy install helper — evaluate dropping it
   in the clean rebuild rather than carrying it forward.) Delete the default blank sheet.
   Verify no external links were created (break if any).
7. **Stamp each canonical data sheet from its `_Template_NN` snapshot** so the template
   ships pristine/blank regardless of the source's current contents.
8. **Recreate the 7 named ranges** explicitly (workbook-scoped) at their known cells.
9. **Delete the on-sheet `Btn_*` shapes** from DataViewer Upload (the ribbon hosts them
   now) — by matching shapes whose `.OnAction` contains `Btn_`.
10. **Import VBA** from repo: clear existing components; import the 3 `.bas`; set
    `ThisWorkbook` code from `ThisWorkbook.cls.txt`.
11. **Save** as `.xlsm` (FileFormat 52); quit Excel.
12. **Inject the ribbon** at the zip level (Python `zipfile`): add `customUI/customUI14.xml`
    + icon PNGs + rels + `[Content_Types]` entries. (Excel COM can't set customUI.) Confirm
    no `xl/webextensions/*` parts exist (fresh workbook → none).
13. **Verify:** run `verify_sidecar.py` against the output → expect all-MATCH; report.
14. Output the clean `.xlsm` path. **User eyeball-tests and ships** — never auto-dropped.

A **manual fallback** (new workbook → Move/Copy sheets → import VBA in the VBE → apply
ribbon in the Custom UI Editor → save) is documented in the RUNBOOK, per the
keep-a-reliable-manual-path principle.

### 6.3 Crash & save-loop fixes

- **Puffs step-picker (ThisWorkbook):** restructure so `Application.EnableEvents` is
  restored on **every** exit path via a single cleanup label that always runs (save prior
  state, `... On Error GoTo PuffCleanup ... PuffCleanup: Application.EnableEvents = saved`).
  Never leave events disabled. Keep the feature (the user added it deliberately) but make
  it crash-safe and re-entrancy-guarded.
- **Save-loop:** (a) **strip the web add-in** (a known dirtier); (b) keep `Workbook_Open`
  side-effect-free where state already matches; (c) at the end of `Workbook_Open`, set
  `ThisWorkbook.Saved = True` **only** when the open made no substantive change — so a
  no-edit session closes with no nag, while real edits still prompt normally.
- **Dead code:** delete `Workbook_SheetBeforeDoubleClick` (rows 20–33) and ensure the
  leftover native-checkbox feature bag is absent in the rebuilt file.
- **Churn-free reset** (§6.4) removes the repeated sheet delete/insert corruption vector.

### 6.4 Churn-free reset + Review copies (the core behavior change)

On Upload All, after the `.xlsx` has been distributed, for **each uploaded canonical data
sheet `X`** (selected + populated; `Test SOP's` is never reset):

1. **Rename `X` → its Review name** (§6.6). The original sheet — entered data, formatting,
   formulas all intact — *becomes* the review copy. (A rename, not copy+delete.)
2. **Copy the internal blank snapshot `_Template_NN`** into the workbook at the renamed
   sheet's former tab position, named `X` → a pristine blank canonical sheet for the next
   campaign, full fidelity.
3. **Move the Review sheet to the end** of the workbook (keeps the working area
   uncluttered); leave it `xlSheetVisible`.
4. Re-hide the snapshot; `ApplySheetVisibility` then sets fresh `X`'s visibility from the
   selection.

**Why this is better:** churn drops to **+1 sheet add per uploaded sheet** (the fresh
blank) with **no deletion of the user's data sheet** — versus today's add **and** delete.
The Review copy is free and 100%-fidelity (it *is* the original). The data sheets are
self-contained (within-sheet relative formulas; no cross-sheet refs), so the renamed
Review sheet stays valid without flattening to values.

**Poka-yokes (transactional):**
- If the snapshot `_Template_NN` is missing, or the copy adds no sheet, **abort that sheet**:
  rename the Review sheet back to `X` and leave it fully intact + log. Never finish without
  the canonical sheet present.
- Only the rename is "destructive," and it's reversible until step 2 succeeds.

**Integration (everything keys off the canonical list, so Review sheets are auto-ignored):**
- `BuildKeepList` / `RunChecklist` / `SheetHasPopulatedSamples` iterate `CanonicalDataSheets`
  → Review sheets are never selected, validated, or uploaded.
- `ApplySheetVisibility` only touches canonical + `_Template_*`/`_Macro_` → Review sheets
  stay visible, untouched.
- Review sheets are created on the **live** workbook *after* staging/trim, so they never
  appear in the distributed `.xlsx`. On the next upload, the staging copy's Trim step drops
  them (not in keep-list) — copies stay clean.

### 6.5 Ribbon consolidation + Delete All Review Sheets

Add a third group to the "TPM Testing" tab (`customUI14.xml`), after Sample Navigation:

```
Group "DataViewer Upload":
  Upload All           onAction=Ribbon_UploadAll            imageMso=UploadCenterManageUploads (or PublishToWebSite)
  Dry-Run Checklist    onAction=Ribbon_DryRun               imageMso=FileValidation
  Pick Synology Folder onAction=Ribbon_PickSynology         imageMso=FolderOpen
  Pick Local Folder    onAction=Ribbon_PickLocal            imageMso=FolderOpen
  Delete All Reviews   onAction=Ribbon_DeleteReviewSheets   imageMso=TableDelete (or Delete)
```

- Built-in `imageMso` icons (no new PNG assets needed).
- New ribbon wrappers in `DataViewerUpload.bas` (`control As IRibbonControl`) calling the
  existing `Btn_*` subs; plus the new core sub.
- **On-sheet buttons removed** (build step 6.2-#9) — the ribbon is the single home.

**`DeleteAllReviewSheets`:**
- Counts every Review sheet (name matches the Review pattern, §6.6), shows a confirm with
  the count, deletes them in one pass, then activates a safe sheet (`DataViewer Upload`).
- **Hard guard:** never deletes a sheet that is canonical, `_Template_*`, `_Macro_Install`,
  `DataViewer Upload`, or `Test SOP's` — only true Review sheets.
- Wrapped in `EnableEvents/DisplayAlerts/ScreenUpdating` save+restore.

### 6.6 Review-sheet naming (31-char limit + accumulation)

- Each canonical sheet has a **curated short base label ≤ 20 chars** (parallel to
  `CanonicalDataSheets`), so `"<base> - Review"` ≤ 29 and `"<base> - Review N"` ≤ 31 fit:

  | Canonical | Review base |
  |---|---|
  | Lifetime Test | `Lifetime Test` |
  | User Test Simulation | `User Test Simulation` |
  | Long Puff Lifetime Test | `Long Puff Lifetime` |
  | Rapid Puff Lifetime Test | `Rapid Puff Lifetime` |
  | Intense Test | `Intense Test` |
  | Big Headspace Serial Test | `Big Headspace Serial` |
  | Negative Pressure Test | `Negative Pressure` |
  | Temperature Cycling Test #2 | `Temp Cycling Test #2` |
  | Viscosity Compatibility | `Viscosity Compat` |
  | Various Oil Compatibility | `Various Oil Compat` |
  | Custom Test Template | `Custom Test Template` |
  | Temperature Cycling Test #1 | `Temp Cycling Test #1` |

- Name = `"<base> - Review"`. Because reviews **persist and can repeat**, if that name
  exists, append a counter: `"<base> - Review 2"`, `3`, … A helper trims the base further
  if a (rare) multi-digit counter would exceed 31 chars. Identification for
  Delete-All is by the `" - Review"` marker (no canonical name contains it), with the §6.5
  hard guard as backstop.

### 6.7 Drift checker (`verify_sidecar.py`)

- Keep module diffing (`DataViewerUpload`, `TestingTools`/`Module1`, `SampleNav`,
  `ThisWorkbook`).
- Add a **ribbon check**: extract `customUI/customUI14.xml` from the workbook and compare
  (normalized) to the repo `customUI14.xml` → MATCH/DIFFERS.
- Optional: assert no `xl/webextensions/*` parts exist (regression guard against the
  add-in creeping back).

---

## 7. Data flow — Upload All, end to end (post-rebuild)

1. Clear log; status "Starting upload".
2. Run checklist on selected canonical sheets; abort + report on failure.
3. Save workbook (persist in-memory edits).
4. Resolve `DV_SynologyPath`, `DV_LocalPath`; verify both reachable.
5. Build keep-list (Test SOP's + selected populated canonical); abort if nothing real.
6. Stage a trimmed `.xlsm` in TEMP; trim non-keep sheets + sample-block tails.
7. Materialize one clean macro-free `.xlsx` (FileFormat 51 strips VBA).
8. Copy that `.xlsx` → Synology + Local; launch DataViewer on it (Postgres ingest).
9. **Reset live workbook (§6.4):** per uploaded canonical sheet → rename to Review, stamp
   fresh blank from snapshot, move Review to end.
10. Reset `DV_TestSelection` to default; `ApplySheetVisibility`; save; status "OK".

Failure after dispatch (step 9–10) → status "OK (live reset partial — see log)"; the upload
already succeeded.

---

## 8. Error handling / poka-yokes (summary)

- Reset is transactional per sheet; a missing snapshot or failed copy leaves the data sheet
  **intact** (renamed back) — never lost.
- Events/alerts/screen-updating saved + restored on every path (picker, reset, delete-all).
- Delete-All hard-guards canonical/utility/template sheets.
- Build script backs up first, verifies trust + COM, breaks stray external links, and
  finishes with a drift check.
- `verify_sidecar.py` is the standing anti-drift gate (modules + ribbon + no-add-in).

---

## 9. Testing & acceptance (work machine)

Author-side (dev): `verify_sidecar.py` against the rebuilt file → **all MATCH**; static
review of every changed module.

Operator-side checklist (Excel, work machine):
1. Open the rebuilt file, change nothing, close → **no save prompt**.
2. Toggle a test TRUE/FALSE → matching sheet shows/hides; Add/Remove Sample, Reset
   Formulas, First/Prev/Next/Last + Ctrl+Shift+.,/ all work.
3. Puffs step-picker works; force an error mid-edit → events still alive afterward
   (toggles still respond).
4. Dry-Run passes on a populated sheet.
5. Upload All on ≥1 populated test → the `.xlsx` lands in Synology + Local + opens in
   DataViewer; each uploaded sheet now has a `… - Review` copy at the end; the canonical
   sheet is blank/fresh; selection reset to Lifetime-only.
6. Upload a second test without deleting reviews → a second Review sheet appears (counter if
   the same test); distributed copies contain **no** Review sheets.
7. Delete All Review Sheets → all `… - Review` gone in one click; canonical/utility sheets
   untouched.
8. Inspect the file (or re-run forensics) → no `xl/webextensions/*`, no dead double-click
   handler.

---

## 10. Risks & mitigations

| Risk | Mitigation |
|---|---|
| Cross-workbook sheet copy creates external links | Copy full set together; break links; verify in step 6 |
| VBA import blocked (trust setting off) | Script checks first + prints the exact Trust Center steps; manual fallback in RUNBOOK |
| `pywin32` not installed | Script detects + instructs `pip install pywin32` |
| Build produces a subtly different layout | Excel does the copying (full fidelity); operator checklist §9 catches regressions |
| Review sheets accumulate → bloat | Expected per the decision; Delete-All is the release valve; documented |
| Source has data in canonical sheets | Step 6.2-#7 stamps each from its snapshot → template ships blank |
| A data sheet has a cross-sheet formula ref → the rename could misroute it | None observed (test logs are self-contained); rebuild + §9 checklist confirm; if found, flatten that sheet's Review copy to values instead of renaming |
| Can't verify crash/loop fix in dev env | Verified on the work machine via §9; diagnosis confidence stated in §2 |

---

## 11. Deliverables

- `excel-sidecar/DataViewerUpload.bas` (reconciled + reset/Review/Delete-All + ribbon wrappers)
- `excel-sidecar/TestingTools.bas` (reconciled)
- `excel-sidecar/ThisWorkbook.cls.txt` (hardened picker, drop dead handler, Saved guard)
- `excel-sidecar/customUI14.xml` (3rd ribbon group)
- `excel-sidecar/build_clean_template.py` (NEW)
- `excel-sidecar/verify_sidecar.py` (ribbon + no-add-in checks)
- `excel-sidecar/README.md`, `RUNBOOK-*.md` (rebuild flow + Review feature)
- This spec; an implementation plan to follow.

Output artifact (work machine, not committed): the rebuilt
`Automated Testing Template v1.xlsm`.

---

## 12. Open items

- Final `imageMso` icon picks (cosmetic; confirm during implementation).
- Whether to also move the rebuilt file's default landing sheet to `DataViewer Upload`
  vs `Lifetime Test` (minor).
- Position of accumulated Review sheets (end-of-workbook assumed; trivial to change).
