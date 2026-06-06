# Hotfix: schema drift — `data_rows.puffing_regime` / `tests.raw_grid` missing on live DB

**Reported (work machine, on Ctrl+U):**
```
tryWriteFile(prepare data_rows): ERROR: column "puffing_regime" of relation
"data_rows" does not exist (42703) QPSQL: Unable to prepare statement
```

## Root cause (confirmed by tracing)
- `data_rows.puffing_regime` (migration `2026-05-29-v2.2.1-per-row-regime.sql`) and
  `tests.raw_grid` (migration `2026-06-04-v2.x-tests-raw-grid.sql`) are additive
  columns that exist inline in `deploy/postgres/init.sql` **and** as migration files.
- Postgres runs `init.sql` only at **first container init**. The live NAS DB was
  created earlier; the two migrations were never applied to it → columns absent.
- `DatabaseManager::tryWriteFile` references both columns *unconditionally* in its
  prepared `tests`/`data_rows` statements → every file upload fails at prepare-time.
- This is the real cause of "nothing saved since ~6-1"; DV-3 (v2.3.2) only made the
  previously-swallowed error visible.
- **No corruption:** the whole write is one transaction and rolls back on the prepare
  failure — failed uploads are absent, not partial. Re-upload / `--backfill` recovers.

## Fix (additive, idempotent, self-healing — no manual NAS step)
Add `DatabaseManager::ensureSchema()`, called on every successful connect
(`open()` + `reopen()`):
- For each post-baseline additive column, do a cheap `pg_attribute` catalog check;
  only when missing, run `ALTER TABLE … ADD COLUMN IF NOT EXISTS …`.
- Catalog-check-first avoids taking `ACCESS EXCLUSIVE` on the hot path (multi-user safe).
- List = exactly the two migration-added columns; documented to stay in sync with
  `deploy/postgres/migrations/*.sql` + `init.sql`.

## Tasks
- [ ] **T1 (TDD red):** add regression test `ensureSchema_reAddsDroppedAdditiveColumns`
      to `tests/tst_databasemanager` — drop both columns out-of-band, call
      `ensureSchema()`, assert both re-added + idempotent + `tryWriteFile` returns
      Success. Run → fails to compile (method missing).
- [ ] **T2 (TDD green):** implement `ensureSchema()` in `DatabaseManager.cpp`, declare
      in `.h`, call from `open()` and `reopen()`. Run test → PASS. Run full
      `tst_databasemanager` → all PASS.
- [ ] **T3:** bump `VERSION` 2.3.2 → **2.3.3** in `DataViewerEnterprise.pro`;
      `mingw32-make clean` + clean release rebuild; `build_installer.bat` →
      `dist\DataViewer-setup.exe` (do **not** touch Synology).
- [ ] **T4:** comment on DATAVIEWER-3 in Plane (root cause + this fix); report to user
      with recovery guidance (re-upload / backfill the 6-1→now files).

## Versioning
Internal patch **v2.3.3** (no deploy). Shifts the batch: DV-4 → v2.3.4, DV-2 → v2.3.5,
wrap into deployable **v2.4.0**.
