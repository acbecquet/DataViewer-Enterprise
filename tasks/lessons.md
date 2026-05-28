# Lessons Learned

## 2026-03-19: Sensory layout should be horizontal, not vertical
- **Mistake:** Set sensory panel splitter to `Qt::Vertical` (cards top, chart bottom) matching the TPM layout
- **Correction:** Sensory mode uses a **horizontal** layout — cards on the left, radar chart on the right. This is intentionally different from TPM mode (table top, plot bottom).
- **Rule:** Don't assume new UI modes should mirror existing layout orientations. When the user says "place data and plot in the space where table and chart are," the spatial arrangement can still differ.

## 2026-05-11: QSqlDatabase reference release on Windows
- **Discovery:** Calling `db.close()` alone is insufficient on Windows — the file handle isn't released until `QSqlDatabase::removeDatabase(name)` is called. When the Qt documentation warns "connection still in use", assigning `db = QSqlDatabase()` first prevents the warning.
- **Applied in:** `PostgresConnection::close()`, `PostgresConnection::open()` (connection cycling), and `MigrationTool::finalizeSource()` (SQLite close after reading source).
- **Rule:** Always call `QSqlDatabase::removeDatabase()` when closing a database for the last time, or clear the object reference first.

## 2026-05-11: QTemporaryFile blocks QFile::rename on Windows
- **Discovery:** Even after explicit `close()`, `QTemporaryFile` retains a file-share flag that prevents `QFile::rename()` on the same object. The file is not fully released to other processes until it goes out of scope.
- **Applied in:** `tst_migrationtool` — tests that rename files should use a plain path under `QDir::tempPath()` instead of `QTemporaryFile`.
- **Rule:** For tests that need atomic file operations (create temp, rename to final), avoid `QTemporaryFile`; use explicit temp paths and manual cleanup.

## 2026-05-11: MIP/AIP source-file mitigation
- **Discovery:** Every new `.cpp`/`.h` file in this repo may inherit Microsoft Information Protection (MIP) sensitivity labels at rest, causing build tools (`g++`, `head`, `cat`) to see ciphertext (`%TSD-Header-###%...`). The Python interpreter is on the MIP allowlist and reads plaintext.
- **Applied in:** All file creation in Plan A — write new source files via Python's delete-and-rewrite pattern (see CLAUDE.md section "File creation convention"), not via the Write tool. Before any C++ build, run `python tools/decrypt_via_copy.py --apply` to strip labels.
- **Rule:** New `.cpp`/`.h` files must be written via Python. Subagents that produce source code must include the deletion pattern. The decrypt script is idempotent and runs as a build pre-step until labels stabilize.

## 2026-05-11: postgres:16 doesn't bundle pg_cron
- **Discovery:** The official Docker image for `postgres:16` does not include the `pg_cron` extension. Container startup fails on `CREATE EXTENSION pg_cron` with "could not access file pg_cron".
- **Applied in:** `deploy/postgres/Dockerfile` — derived image installs `postgresql-16-cron` from the PostgreSQL Development Group (PGDG) repository before initializing the database.
- **Rule:** When using a postgres Docker image, always derive a custom image for any non-bundled extensions (`pg_cron`, `pg_partman`, etc.). The Dockerfile pattern is: `FROM postgres:16 && apt update && apt install -y postgresql-16-<extension> && ...`.
## 2026-05-28: VBA/COM worksheet copy — capture by name-diff, never by position/ActiveSheet
- **Discovery:** Seeding the post-upload reset snapshots failed two ways that silently corrupted the workbook. (1) `Worksheet.Copy After:=Worksheets(Count)` when the last sheet is **very-hidden** (`_Template_Master`) misplaces the copy, so capturing the new sheet by position grabbed a pre-existing sheet and `snap.Name = ...` renamed it 12 times (only `_Template_11` survived; `_Template_Master` was clobbered). (2) In pywin32, `ws.Copy(After=anchor)` as a **named** arg doesn't bind in late-bound COM and silently spawns a throwaway workbook (0 sheets added to the target).
- **Applied in:** `excel-sidecar/DataViewerUpload.bas` — `AddSnapshotFrom` and `RestoreSheetFromTemplate` copy after/before a **visible** anchor (the "DataViewer Upload" sheet / the live sheet), find the new sheet by **name-diff** (record sheet names before the copy), and never `Delete` the original until the replacement is confirmed present. Recovery COM scripts pass the anchor **positionally** as `Before` (`ws.Copy(beforeSheet)`), not `After=`.
- **Rule:** When copying sheets in VBA/COM: (a) never anchor Before/After on a very-hidden sheet; (b) capture the new sheet by name-diff, not `ActiveSheet` or `Worksheets(Count)`; (c) make any replace transactional — copy first, confirm, then delete; (d) in pywin32 pass `Copy`/`Move` anchors positionally, not as `After=`. Separately: an over-strict guard (snapshot `A1` must equal the live sheet's `A1`) false-skipped a sheet whose template title legitimately differed ("Lifetime Test 2") — prefer a robust structural invariant over brittle equality.
