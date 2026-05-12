# libpq 16 Windows runtime DLLs

These DLLs are the runtime dependencies of Qt's QPSQL driver and the
`DataViewer.exe` Postgres-facing code paths. They get copied into
`release/` by `build_installer.bat` and bundled into the installer
via `installer.iss` (Task 23 / 24).

## Provenance

- **Source:** Official PostgreSQL Windows binaries ZIP from EnterpriseDB
  at https://get.enterprisedb.com/postgresql/postgresql-16.10-1-windows-x64-binaries.zip
- **Version:** PostgreSQL 16.10 (x86_64)
- **Files extracted from `pgsql/bin/`:**
  - libpq.dll              — PostgreSQL client library
  - libcrypto-3-x64.dll    — OpenSSL crypto (libpq dependency)
  - libssl-3-x64.dll       — OpenSSL TLS (libpq dependency)
  - libintl-9.dll          — GNU gettext (libpq dependency)
  - libiconv-2.dll         — iconv (libpq dependency)

Total weight: ~8 MB.

## Updating

To upgrade to a new Postgres minor version, fetch the corresponding
Windows binaries ZIP from EnterpriseDB and replace all 5 DLLs.
Verify the test suite still passes against an ephemeral
`postgres:16` (or upgraded version) container.

Major version upgrades (16 → 17) are a deliberate event and must be
planned alongside the server-side Postgres upgrade — do NOT bump the
client DLLs ahead of the server.

## Why committed (vs downloaded at build time)

- Reproducible builds: `build_installer.bat` works without external
  network access.
- Hash-locked: changing the DLLs requires a deliberate commit.
- Small: ~8 MB total, within reason for repo bloat.

## Running tests locally that hit Postgres

The Qt QPSQL plugin (`qsqlpsql.dll` in `C:\Qt\6.10.x\mingw_64\plugins\sqldrivers\`)
loads `libpq.dll` at runtime via the OS DLL search path. To make these
DLLs findable when running tests directly:

```bash
export PATH="<repo>/vendor/libpq-16:$PATH"
./tests/tst_postgresconnection/release/tst_postgresconnection.exe
```

Or copy the 5 DLLs next to the test executable. The installer puts
them next to `DataViewer.exe` in production.
