# Deployment self-test

Smoke test for an installed copy of DataViewer Enterprise. Run after every
install on the work machine before placing the new build on the Synology
share.

## What it covers

| Phase | Check                              | Why                                                                 |
|-------|------------------------------------|---------------------------------------------------------------------|
| 1     | Install tree                       | Every Qt DLL, plugin, bundled Python, and resource file present     |
| 2     | `appdata_writable`                 | `QStandardPaths::AppLocalDataLocation` is writable                  |
| 2     | `registry`                         | `QSettings("SDR","DataViewerEnterprise")` round-trip                |
| 2     | `sqlite`                           | Qt's `QSQLITE` driver loads + creates a temp DB                     |
| 2     | `python_bundle`                    | Bundled `python\python.exe` runs and `import openpyxl` succeeds     |
| 2     | `excel_roundtrip`                  | openpyxl writes + reads a temp `.xlsx` (covers temp-file perms)     |
| 2     | `synology_root`                    | The folder `UpdateChecker::updateRoot()` returns actually exists    |
| 2     | `latest_version`                   | `UpdateChecker::latestAvailable()` finds a valid version + setup.exe |
| 3     | Synology folder probe              | Independent of the binary — lists every version subdir and flags ones missing `DataViewer-setup.exe` |

Phases 2 and 3 verify the same thing two ways. Phase 2 runs the production
code path inside `DataViewer.exe` (so any code change to `UpdateChecker` is
exercised). Phase 3 walks the directory directly with PowerShell, so a
regression in the in-process code can be distinguished from a missing
folder or absent installer.

## Run it

After installing, from any PowerShell prompt:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\deployment\Test-Deployment.ps1
```

Or with explicit paths:

```powershell
.\Test-Deployment.ps1 `
    -InstallDir "C:\Users\S1134987\AppData\Local\Programs\DataViewer Enterprise" `
    -ReportPath "$env:TEMP\dataviewer_selftest.json"
```

Exit code is `0` only if every phase passes.

## What's not tested

- The auto-update **download + replace** flow. The script stops at "the
  binary saw the new version on Synology." Run an actual install when you
  want to validate the replace step.
- Drag-and-drop file loading.
- `SingleInstance` IPC (would require launching twice; left out for now).

If any of the above start mattering, add them here rather than as separate
ad-hoc tests.

## Adding a new check

1. Add a `TestResult testWhatever()` function in `src/utils/SelfTest.cpp`.
2. Append it to the `results` list in `runSelfTest`.
3. Document it in the table above.

Each test must be self-contained, leave no on-disk side effects, and finish
in a few seconds. Anything slower belongs in a separate diagnostic.
