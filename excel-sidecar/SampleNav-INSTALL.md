# Sample Navigation — install guide

Installs sample-block navigation into `Automated Testing Template.xlsm`:
four buttons on the **TPM Testing** ribbon tab + two hotkeys.

You do this manually because the `.xlsm` is actively in use for data
collection — nothing should auto-modify it.

## What gets installed

| File from this folder | Goes into | How |
|---|---|---|
| `SampleNav.bas` | A new VBA standard module named `SampleNav` | VBE: File → Import |
| `SampleNav-ThisWorkbook.txt` | The `ThisWorkbook` code module | VBE: paste |
| `SampleNav-ribbon-snippet.xml` | The workbook's `customUI14.xml` | Custom UI Editor: paste-merge |

## Prerequisites

- Workbook saved as **macro-enabled** (`.xlsm`). If it's currently `.xlsx`,
  save-as `.xlsm` first.
- One-time tool install: [Office RibbonX Editor](https://github.com/fernandreu/office-custom-ui-editor/releases)
  — pick the latest `OfficeRibbonXEditor.*.exe` release and install. Used
  only for editing ribbon XML inside the `.xlsm`. Skip if you've used it
  before.

## Step-by-step

### 1. Close the workbook in Excel

Save your work and close Excel completely. The Custom UI Editor and Excel
can't have the same file open simultaneously.

### 2. Add the ribbon group with the Custom UI Editor

1. Open **Office RibbonX Editor**.
2. **File → Open** → select your `.xlsm`.
3. If the right-hand pane is empty, right-click the file in the left tree
   → **Office 2010+ Custom UI Part** (this creates `customUI14.xml`).
   If it already has a customUI14.xml, click it to view.
4. Find your existing `<tab ... label="TPM Testing">` element in the XML.
5. Open `SampleNav-ribbon-snippet.xml` in any text editor. Copy the entire
   `<group id="grpSampleNav">…</group>` block (the instructional
   `<!-- ... -->` comment at the top can be skipped).
6. Paste the group inside the `<tab ... label="TPM Testing">` element,
   after the existing `<group>` elements.
7. Click **Validate** (green check icon in the toolbar). Should say
   "Custom UI XML is well-formed."
8. **File → Save.**
9. Close the Custom UI Editor.

### 3. Import the VBA module

1. Open the `.xlsm` in Excel. **Enable macros** when prompted.
2. Press `Alt+F11` to open the VBA editor (VBE).
3. In Project Explorer (left pane), verify you're under:
   `VBAProject (Automated Testing Template.xlsm)`.
4. **File → Import File...** → select `SampleNav.bas` from this folder.
5. The new module appears under **Modules → SampleNav**.

### 4. Wire the hotkeys via ThisWorkbook

1. Still in the VBE, in Project Explorer, double-click **ThisWorkbook**
   (under Microsoft Excel Objects).
2. The code window opens — likely empty or with a couple of existing
   handlers.
3. Open `SampleNav-ThisWorkbook.txt` in any text editor. Copy everything
   from `Option Explicit` down (the comment header above is optional).
4. Paste at the bottom of the ThisWorkbook code module.
   - **If `Option Explicit` is already at the top:** delete the duplicate
     from the paste.
   - **If `Workbook_Open` or `Workbook_BeforeClose` already exist:** do
     NOT overwrite them. Instead, copy just the two `Application.OnKey`
     lines from each of my handlers into the existing handlers.
5. Save the workbook (`Ctrl+S`). Must remain `.xlsm`.

### 5. Reload and test

1. Close the workbook entirely.
2. Reopen it (enable macros). This fires `Workbook_Open` and registers
   the hotkeys.
3. Click any cell in row 5 or below on a data sheet (e.g. "Lifetime Test").
4. Press `Ctrl+Shift+.` — selection jumps 12 columns right.
5. Press `Ctrl+Shift+,` — selection jumps 12 columns left.
6. On the **TPM Testing** ribbon tab, look for the new **Sample
   Navigation** group with four buttons:
   - **First** → jumps to column C on the current row.
   - **Prev Sample** → jumps 12 cols left (clamps at C).
   - **Next Sample** → jumps 12 cols right (clamps at CI / col 87).
   - **Last** → jumps to the rightmost sample column with a non-empty
     header in row 4.

If anything doesn't work, see **Troubleshooting** below.

## Troubleshooting

### Hotkeys do nothing

- Hotkeys only register when `Workbook_Open` runs. Did you fully close
  and reopen the workbook after pasting the ThisWorkbook code?
- In the VBE Immediate window (`Ctrl+G`), type `JumpRight12` and press
  Enter. If the cursor moves → the macro works and the issue is OnKey
  registration. Check the exact spelling in `Workbook_Open`:
  `"+^."` and `"+^,"` — `+` is Shift, `^` is Ctrl, the literal `.` and
  `,` must match.

### Ribbon buttons missing

- Did you save the `.xlsm` in the Custom UI Editor *before* opening it in
  Excel? Excel only loads ribbon XML at file-open time.
- In the Custom UI Editor, click **Validate**. Any XML error stops the
  whole ribbon from loading.
- If the "TPM Testing" tab is there but without the new group, the
  `<group>` was pasted outside the `<tab>` element. Re-check the nesting.

### Ribbon buttons present but clicking does nothing

- The `onAction` value in the XML must match a public Sub name in
  SampleNav.bas exactly. The snippet uses the `_Ribbon` wrappers
  (`JumpFirstSample_Ribbon`, etc.) — these are required because customUI
  `onAction` passes an `IRibbonControl` parameter that the no-arg
  hotkey-targeted Subs don't accept. Don't change `onAction` to point at
  the no-arg names.
- In the VBE, place the cursor in `JumpRight12_Ribbon` and press F5. If
  it errors with "argument not optional", VBE is trying to run it
  manually with no `control` arg — that's expected; ignore. Try running
  `JumpRight12` (no `_Ribbon` suffix) instead.

### "Compile error: User-defined type not defined"

- One macro (the helper) uses `Range` and `Worksheet`. These are part of
  the Excel object library, which is referenced by default. If you see
  this error, check **VBE → Tools → References** and make sure
  "Microsoft Excel xx.x Object Library" is checked.

### Macros raise "ActiveCell is Nothing" or similar

- Should be impossible — `SafeActiveCell()` returns `Nothing` and the
  macros early-exit. If you see this, the `.bas` import didn't take.
  Re-import `SampleNav.bas` (delete the existing module first).

## Unzip fallback (no Custom UI Editor)

If you can't install the Custom UI Editor, edit the ribbon XML by hand:

1. Make a backup: copy your `.xlsm` to `.xlsm.bak`.
2. Rename `.xlsm` → `.zip`.
3. Open the zip with 7-Zip or Windows Explorer.
4. Look for `customUI/customUI14.xml`.
   - **If present:** edit it to insert the `<group>` from
     `SampleNav-ribbon-snippet.xml` inside the existing
     `<tab ... label="TPM Testing">` element.
   - **If absent:** you'd also need to create `customUI/customUI14.xml`
     AND add a `<Relationship>` entry to `_rels/.rels`. At that point
     the Custom UI Editor is much easier — install it.
5. Save the edited XML back into the zip.
6. Rename `.zip` → `.xlsm`.
7. Open in Excel.

## Uninstall

1. VBE → right-click `SampleNav` in Project Explorer → **Remove SampleNav**.
   Decline the "export before removing?" prompt.
2. VBE → ThisWorkbook → delete the two `Application.OnKey ...` lines from
   `Workbook_Open` and `Workbook_BeforeClose`. (Or delete the whole
   handlers if they were added by this install and contain nothing else.)
3. Custom UI Editor → open `.xlsm` → delete the
   `<group id="grpSampleNav">` element → Save.
4. Save the workbook.
