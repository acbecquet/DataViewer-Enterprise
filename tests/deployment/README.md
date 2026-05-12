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

## Full Report manual checklist

Run after every install on the work machine that touches reporting code.

### Single-file Full Report
- [ ] Generates without error on a normal file (3+ sheets)
- [ ] Title slide unchanged
- [ ] Test Protocol slide shows 6 cols, only tests present in the file
- [ ] Test Overview shows auto-templated line + bullets + empty trailing textbox
- [ ] Each data slide: TPM Trend has markers, Y-axis matches the rule
- [ ] Image slides unchanged
- [ ] Conclusions slide: blank textbox

### Multi-file Combined Report (3 files)
- [ ] File picker accepts multi-select
- [ ] Output folder dialog appears once
- [ ] 3 individual reports + 1 combined report appear in chosen folder
- [ ] Combined order: Title -> Protocol -> Overview -> Lifetime Comparison ->
      Section divider -> File 1 overview -> File 1 data -> ... -> Conclusions
- [ ] Section divider slides: cover style, no date, filename centered
- [ ] Lifetime Comparison: bars colored by file, shaded by sample within file
- [ ] Files without "Lifetime Test" sheet are silently skipped on the comparison slide

### Edge cases
- [ ] Long Puff Lifetime Test: Y-axis 0-25 default, fallback to maxTPM+1 out of range
- [ ] Sheet with avgTPM > 7: Y-axis bumps to maxTPM+1
- [ ] Puffing regime "200mL/10s/60s": detected as Long Puff even on a non-Long-Puff sheet name
- [ ] File picked that's already loaded: reused (no re-parse)
- [ ] One report failing in batch: other reports still generate, summary dialog shows failures
- [ ] Filename collision in output dir: appends (2), no overwrite

## Phase 4 - Migration verification

Compares the renamed `<name>.pre-migration.sqlite` file (kept on Synology
after a successful migration) against the live PostgreSQL database. For
every editable table in both, row counts must match. If they diverge,
the phase fails and prints which table.

This phase requires:
- The pre-migration SQLite file at its rename location
  (`<name>.pre-migration.sqlite`).
- A working Postgres connection (host/port/db/user/password).
- `sqlite3.exe` and `psql.exe` available on `PATH`.

Invocation example:

```powershell
.\Test-Deployment.ps1 -PreMigrationSqlite "Z:\SynologyDrive\dve.sqlite.pre-migration.sqlite"
```

If `-PreMigrationSqlite` is omitted, Phase 4 is skipped with a warning.

## Manual checklist (cannot be automated end-to-end)

Verify on the work machine after a fresh v2 install:

- [ ] First-launch identity prompt appears, accepts a name + color, and
      does NOT appear on the second launch.
- [ ] `%PROGRAMDATA%\DataViewer\db.conf` exists and is readable only to
      administrators.
- [ ] `DataViewer.exe --self-test` reports `postgres_connection: passed`.
- [ ] Opening a TPM file from the migrated database displays the same
      sheet/sample/row data as the pre-migration SQLite did on v1.3.x.
- [ ] Opening a sensory session shows the same metric values and any
      saved layout JSON renders correctly.
- [ ] Opening a detailed sensory session shows the same Q1-Q14 responses.
- [ ] Embedded images render in TPM samples and sensory sessions.
- [ ] No `<dbPath>.lock` sidecar files are created when the app is
      running (the file-lock code path was deleted in Plan C).
