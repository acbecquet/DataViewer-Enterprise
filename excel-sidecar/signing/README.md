# VBA code-signing kit

Signing the template's VBA project is what gets testers to **zero macro
prompts on every delivery channel** — including copies that arrive by
email/Teams/download and therefore carry Mark of the Web. Excel checks
the trusted-publisher signature *before* applying the MOTW macro block,
so a signed workbook from a trusted publisher opens clean where an
unsigned one would be blocked or prompt.

## Owner flow (template maintainer)

**Once ever:**

1. Run `make_cert.ps1` (right-click -> Run with PowerShell). It creates a
   10-year code-signing certificate `CN=SDR DataViewer Templates` in your
   user store (re-running reuses the existing one) and exports the public
   half as `DataViewerTemplates.cer` next to the script. Commit the `.cer`.

**After EVERY build of the template:**

1. Open the built `.xlsm` in Excel.
2. `Alt+F11` -> **Tools** -> **Digital Signature...** -> **Choose** ->
   pick **SDR DataViewer Templates** -> **OK**.
3. Save the workbook and close Excel.
4. Gate it:

   ```
   python excel-sidecar/verify_sidecar.py --file <out> --require-signature
   ```

Note: **any** edit to the VBA project inside the workbook strips the
signature. That is a feature — a build that fails `--require-signature`
either was never signed or was modified after signing, so the check
doubles as a drift alarm.

## Tester flow (once per Windows account)

1. Get `tester-setup.ps1` and `DataViewerTemplates.cer` (same folder).
2. Right-click `tester-setup.ps1` -> **Run with PowerShell**.

No admin rights needed — it writes only to the current user's Root and
TrustedPublisher stores. After that, close and reopen the template:
signed builds run with no macro prompts.

## Security notes

- **Never commit a `.pfx`** (or export the private key at all). The
  private key lives only in the owner's user certificate store; the repo
  carries only the public `.cer`.
- Anything signed with this certificate **auto-runs for every tester**
  who ran the setup script. Guard the owner's Windows account accordingly.
- The signature covers **only the VBA project**, not workbook content.
  Cell values, sheets, and named ranges can change without invalidating
  it — use `verify_sidecar.py` for structural integrity.
