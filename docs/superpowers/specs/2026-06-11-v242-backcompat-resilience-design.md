# v2.4.2 — Backwards-Compatibility + Adversarial Resilience — Design

**Date:** 2026-06-11 (impl-status updated 2026-06-17) · **Status:** SP1 (v2.4.2) + SP2 (v2.4.3) IMPLEMENTED + adversarially reviewed on `feature/v2.4.0-bugfix-batch` (not pushed/merged). Tier-3 durability (SP3) + triage UI (SP4) PENDING. Backwards-compat with live v2.4.1 VERIFIED. See "Implementation status" below.
**Branch:** continues `feature/v2.4.0-bugfix-batch` as the next internal patch(es) after v2.4.1.
**Sources:** the v2.4.0 regression dossier (`2026-06-10-v24-save-sync-regressions-evidence.md`), the v2.4.1 fix batch (commits `0c21100..1a31c11`), and a 7-agent adversarial failure-mode map of the connection / LiveSync / NOTIFY / offline-snapshot / crash-recovery surfaces (59 failure modes + 5 cross-cutting + 8 missing scenarios + a ranked hardening list).

---

## Implementation status (as built — updated 2026-06-17)

Executed subagent-driven on `feature/v2.4.0-bugfix-batch` (not pushed/merged), each task gated by a 3-lens adversarial review (spec / quality / domain).

| Item | Status | Where |
|---|---|---|
| A1 version stamping | ✅ done | SP1 · v2.4.2 |
| A2 6-arg `dve_commit_cell_json` numeric heal | ✅ done (defensive-only — see corrections) | SP1 · v2.4.2 |
| A2 `dve_normalize_legacy_json` + one-time + nightly cron | ✅ done | SP2 · v2.4.3 |
| A3 CompatClassifier | ⬜ pending | SP4 |
| A4 DB-Browser version/health multi-select filter | ⬜ pending | SP4 |
| A5 manual repair / delete | ⬜ pending | SP4 |
| R1 query deadlines + dead-socket + 25P02 | ✅ done | SP1 |
| R1b bounded liveness ping | ✅ done (bounded via `statement_timeout`) | SP1 |
| R2 `application_name` in connect string | ✅ done | SP1 |
| R3 post-reconnect catch-up + presence re-activate | ✅ done | SP2 · T4+T5 |
| R4 NOTIFY maintenance-suppression + advisory-locked one-time | ✅ done | SP2 · T1+T3 |
| R4b listen-socket liveness + per-channel re-subscribe | ✅ done | SP2 · T6 |
| R5 MIP-resilience for durability files | ⬜ pending | SP3 |
| R6 atomic Excel write-back | ✅ atomic `deleteRowFromExcel` (SP1); ⬜ off-thread write | SP1 + SP3 |
| R7 crash-safe snapshot promotion + clock discipline | ⬜ pending | SP3 |
| R7b snapshot `app_version` columns + count assertion | ⬜ pending | SP3 |

**Backwards-compat with live v2.4.1 (commit `bf80bcd`): VERIFIED COMPATIBLE** (3-lens audit, 2026-06-17). v2.4.1's `jsonToDouble` reads numeric scores via an `isDouble()`-first branch (proven by compiling the frozen reader against Qt 6.10.1: number `7.5→7.5`, string `"7.5"→7.5`, reset-to-5 path structurally unreachable). All schema/function/NOTIFY changes are additive/nullable/byte-identical; the 7-arg OCC function v2.4.1 calls is unchanged. Adversarial breaker found no v2.4.1 malfunction on 6 vectors. (Live-DB round-trip deferred — container was offline; static + compiled proof is decisive.)

### Corrections folded back from implementation
- **A2 normalizer** — the rewrite MUST add `ORDER BY elem.ord` inside `jsonb_agg`; the original body left sample order unspecified (Postgres does not guarantee aggregate input order without it). Lossless + idempotent + order-preserving confirmed by adversarial SQL tests.
- **R3 catch-up (the keystone)** — "re-SELECT + dirty merge" hid a critical wiring trap: a `sensorySessionFromJson(merged)` round-trip drops images + resets `id`/`version` (breaks OCC), AND `refreshNavigator()→buildSession()` reverts the merge from stale on-screen widgets (reset-to-5 returns). Correct impl: **overlay merged scores onto the in-memory struct** (preserve `id`/`version`/images), **re-render via `applySession` BEFORE any navigator flush**, via a shared `overlayMergedScores()` helper that both panels AND the e2e test call. The adversarial review caught this — the unit test passed while production was broken.
- **A2 6-arg heal is defensive-only** — the 6-arg `dve_commit_cell_json` is uncallable in prod (ambiguous with the 7-arg `DEFAULT NULL` form, SQLSTATE 42725). v2.4.1 and v2.4.3 both call the byte-identical 7-arg; the heal is single-source-of-truth insurance, verified at catalog level.
- **R4 one-time normalize** — lock-then-gate ordering (advisory key `4242002`) empirically proven race-safe under concurrent first-connects; `SET LOCAL statement_timeout=0` in the function body is required because R1 sets a 10 s cap on every connection.
- **Test infra** — `start-test-postgres.ps1` now applies `deploy/postgres/migrations/*` after `init.sql` (the OCC overloads live only in migrations); the throwaway container has no `pg_cron`, so the nightly-cron line is verified only via the one-time heal path and is a NAS-admin record.

---

## Goal

Make every v2-era client's data converge to a clean, current shape in the shared Postgres DB, give the owner a Database-Browser version/health filter to triage and manually repair/remove bad rows, and harden the whole sync/offline/crash stack so it survives the worst real-world conditions: a spotty Synology NAS, Chinese networks behind the GFW (silent RST, half-open sockets, multi-second latency), MIP file encryption that can render local files unreadable the instant a handle closes, clock skew on NTP-blocked machines, and hard crashes.

**Hard cutoff:** compatibility target is **v2.x only** (≥ v2.0.0). Released v2 versions to support: v2.0.0, v2.0.2, v2.0.5, v2.0.8, v2.0.10, v2.1.0, v2.2.0, v2.2.1, v2.2.3, v2.2.4, v2.2.5, v2.3.1, v2.4.1. Pre-v2 and known-broken builds are out of scope.

## The two pillars (and why they're interlocking)

The investigation's headline finding: **the backwards-compat features cannot be correct on a flaky network without specific resilience fixes landing first.**

- Version stamping (Pillar A, feature 1) is **shipped-dead** unless `application_name` rides in the *connect string* — QPSQL never forwards it, and any post-connect `SET` is blanked by the first reconnect (`DatabaseManager::reopen` and `LiveSyncWorker` stop/start both rebuild the backend). → requires **R2**.
- The nightly normalizer (Pillar A, feature 3) **re-introduces the DATAVIEWER-4 "reset-to-5" bug** through a three-way seam: it rewrites scores to numeric and fires a per-row NOTIFY; a client that was briefly offline during the run never catches up (LISTEN/NOTIFY has no server buffering and `onConnectionCameOnline` does **no inbound re-SELECT**); the next whole-file save adopts fresh versions under `FOR UPDATE` (row-level last-writer-wins, OCC disabled in prod per `MainWindow.cpp:229-242`) and writes the stale in-memory value back over the normalized one. → requires **R3** (catch-up) + **R4** (NOTIFY suppression).
- The classifier (Pillar A, feature 4) must not trust `app_version` alone — every historical row and every old-client write is NULL-stamped. → classifier infers era/health from **observable data shape**.

So v2.4.2 is **Pillar A (backwards-compat features)** sitting on **Pillar B (resilience, 3 tiers)**, with the dependencies above.

---

## Pillar A — Backwards-compatibility features

### A1 · Version stamping (going forward)
- New nullable `app_version TEXT` column on `files`, `sensory_sessions`, `detailed_sensory_sessions`.
- Stamped **server-side** by a `BEFORE INSERT` (and `BEFORE UPDATE … WHEN app_version IS NULL`) trigger reading `current_setting('application_name', true)`. **Fail-safe**: nullable, no CHECK, never RAISE, never blank a previously-good stamp.
- Client advertises `application_name = "DataViewer <DVE_APP_VERSION>"` **in the connect string** at all three sites (see R2). Old clients send no name → rows land NULL → displayed as **"pre-v2.4.2"**.
- Healed by `ensureSchema` (catalog-guarded `ADD COLUMN` + trigger `CREATE OR REPLACE`); mirrored in `init.sql` + a canonical migration file in lockstep.

### A2 · Server-side cleanup — one-time + continuous (lossless only)
- **6-arg `dve_commit_cell_json` numeric-CASE fix.** Apply the same numeric-CASE body the 7-arg got in v2.4.1, via the proven `ensureSchema` catalog-guard (`prosrc` NOT containing `to_jsonb($2::numeric`, `pronargs=6`). This closes the path through which any old/alt caller still writes string scores.
- **`dve_normalize_legacy_json()`** — losslessly rewrites numeric-looking *string* score values under known score keys to JSON numbers. **Provably idempotent**: only touches paths where `jsonb_typeof = 'string'` AND value matches `^-?[0-9]+(\.[0-9]+)?$`.
  - **One-time heal** during `ensureSchema`, guarded by a `schema_meta` key, **wrapped in `pg_advisory_xact_lock(<const>)` + a single transaction** with the gate re-checked after acquiring the lock (defeats the concurrent-first-connect thundering herd + the partial-run-with-committed-marker hole).
  - **Nightly `pg_cron`** so anything an old client writes converges within a day. If the DB role can't register cron jobs, log it and ship the `cron.schedule` line in the canonical migration for the NAS admin to run.
  - Both the one-time and nightly runs set `dve.maintenance='1'` so `notify_row_change` early-returns (see R4) — a bulk normalize must not masquerade as thousands of user edits.

### A3 · CompatClassifier (client-side C++, unit-testable)
- **Era buckets:** static table of the 13 released v2 versions + their release dates (from git history); a row's bucket = newest release that existed at its creation time (`loaded_at` / session `timestamp`, fallback `updated_at`). Labeled *(approx.)*. **Server-clock basis** — never mix client and server time (see R7).
- **Health flags** derived from **data shape, not app_version**:
  - sensory/detailed → `Legacy string scores` (any `jsonb_typeof='string'` numeric score), `Junk candidate` (unnamed / "New Session" / zero samples), `Healthy`.
  - TPM files → `No samples`, `Missing puff regimes`, `Healthy`.
  - After the heal, `Legacy string scores` should read **zero** — it doubles as a verification signal.
- Degrades gracefully when `app_version` is NULL: NULL → `pre-v2.4.2 / unknown` era bucket; health still computed from shape.

### A4 · Database Browser version/health filter (all three tabs)
- A checkable **"Version ▾"** dropdown beside each tab's existing text filter: one checkbox per released-version bucket, plus "pre-v2.4.2 (unstamped)" and the health flags. Multi-select, AND-combined with the text filter, with live per-bucket counts.
- Classification computed client-side on open via a few aggregate queries per tab (office-scale row counts → negligible cost); rules live in C++ where tests pin them and changes need no migration.

### A5 · Manual repair / delete (never automatic destruction)
- Lossless repairs are automatic (A2). **Deletion is always manual**: filter → select rows → Repair/Delete, with a confirmation that lists exactly what will be removed. No automatic row deletion, ever.

---

## Pillar B — Adversarial resilience (3 tiers, all approved for this batch)

Each item lists the failure modes it closes (file:line from the map). Full 59-mode dossier retained in the plan.

### Tier 1 — Transport (foundation; feature-enabling)

- **R1 · Query deadlines + dead-socket detection, COUPLED with 25P02 handling.** Add to the connect string at `PostgresConnection::openOne` (`PostgresConnection.cpp:30`), `LiveSyncWorker::openConnection` (`LiveSyncWorker.cpp:58`), and `MigrationTool.cpp:65`: `keepalives=1;keepalives_idle=…;keepalives_interval=…;keepalives_count=…;tcp_user_timeout=…` and a `statement_timeout` (via `options=-c statement_timeout=…` or `SET`). **In the same change**, add SQLSTATE `25P02` (in_failed_sql_transaction) to `LiveSyncWorker::isConnectionError` and issue an explicit `ROLLBACK`/`DISCARD ALL` before any replay. *Coupling is mandatory:* `statement_timeout` aborts a query mid-transaction, leaving the connection in `25P02`; without 25P02 handling the fix converts an indefinite hang into a permanently-wedged connection.
  - Closes: half-open socket hangs GUI ping (critical), whole-save blocks indefinitely on `m_db` (critical), poison `25P02` never recovers (high).
- **R2 · `application_name` in the connect string**, sourced from `-DDVE_APP_VERSION` (can't drift from `.pro` VERSION), at all three sites — the only reconnect-durable form. Plus the fail-safe stamp trigger (A1).
  - Closes: app_version NULL on first connect (high), app_version NULL after reconnect (high), trigger-RAISE aborts writes (medium).
- **R1b · Bounded/off-thread liveness ping** so `ConnectionMonitor::ping` and the Retry button can't freeze the UI on the very failure they detect.

### Tier 2 — Convergence (keeps the normalizer from reintroducing DATAVIEWER-4)

- **R3 · Post-reconnect inbound catch-up** — at the end of `onConnectionCameOnline` (after re-subscribe + outbound drain): re-SELECT/reload the currently-open resource (`loadFile`/`loadFileByPath`) merged dirty-aware (same rule as `handleRemoteRowChange`), then `refreshAllPresence`, then **re-activate local presence** so the heartbeat restarts. Keystone fix.
  - Closes: reconnect never catches up missed rows → LWW resurrection (critical), missed-NOTIFY + dirty-merge reverts remote edit (high), local user becomes a ghost after reconnect (high).
- **R4 · Maintenance-write NOTIFY suppression** — `notify_row_change` early-returns when `current_setting('dve.maintenance', true)='1'`; the one-time heal and nightly cron set it. One-time normalize also gets `pg_advisory_xact_lock` + single-transaction + idempotent rewrite (A2).
  - Closes: nightly NOTIFY storm freezes left-open clients (high), concurrent-first-connect double full-table rewrite (high), partial-run-with-committed-marker (high).
- **R4b · Listen-socket liveness + per-channel re-subscribe completeness** — implement the "signal 1" documented in `ConnectionMonitor.h:28` but never built (probe the listen connection, or NotificationListener self-watchdog); change the `cameOnline` re-subscribe guard from all-or-nothing `!isSubscribed()` to a per-channel `isSubscribedTo()` check that retries missing channels.
  - Closes: listen socket dies silently, query socket lives → permanent silent loss of all live updates (critical), half-open listen socket (critical), partial re-subscribe leaves a dead channel (high).

### Tier 3 — Durability (crash / MIP / offline hardening)

- **R5 · MIP-resilience for the three durability files** — recovery store (`index.json`/blobs), `snapshot.sqlite`, `pending_edits.sqlite`. On open-failure detect the `%TSD-Header-###%` marker and fall back to reading via the **allowlisted bundled python** (the ExcelReader pattern); surface a **loud** user-visible warning when a store exists but won't decode. **Stop swallowing `enqueueCellEdit` failures** — bump the unsynced counter, emit a distinct signal, retain in-memory.
  - Closes (cross-cutting #3): one crash-while-offline collapses crash-recovery + offline-read + offline-write fallbacks simultaneously and silently.
- **R6 · Atomic Excel write-back** — fix `deleteRowFromExcel` (`MainWindow.cpp:6320-6328`) to use `tmp + os.replace` like the sibling `writeCellsToExcel` (today a mid-write kill **truncates the user's source workbook to zero** — a real critical data-loss bug); factor a shared save-helper so the two scripts can't drift again. Move Excel write-back off the UI thread with a tiered timeout (5–8 s interactive, 30 s batch) to kill the up-to-30 s "Not Responding" freeze that invites force-kills.
  - Closes: `deleteRowFromExcel` truncation (critical), 30 s UI hang per Excel write (high).
- **R7 · Crash-safe snapshot promotion + clock discipline** — replace delete-then-`QFile::rename` (`OfflineSnapshot.cpp:787-802`) with `MoveFileEx`/`ReplaceFileW(MOVEFILE_REPLACE_EXISTING|WRITE_THROUGH)`, fsync the tmp file + directory, `synchronous=FULL` for the regenerate connection; **wire up the dead `source_schema_version` validation** in `openReadOnly`; add a max-staleness warning; and **pin staleness + era math to a server-supplied timestamp** (snapshot currently stamps client UTC while the DB uses server `now()` — multi-hour skew on NTP-blocked machines makes "data is fresh" a lie).
  - Closes: crash/MIP between remove and rename → no snapshot (medium), schema-drift snapshot opens silently (medium), clock-skew staleness lie (cross-cutting #5).
- **R7b · Update the snapshot SELECT/INSERT column lists** for the new `app_version` columns on the three session tables, with a debug-build assertion that per-table SELECT column count == INSERT bind count (the hand-maintained lists are a known footgun).

### Cross-cutting failures this batch must defend against (from the critique)
1. Nightly normalizer storm + missing catch-up + LWW → re-corrupts/reverts normalized scores (the reset-to-5 seam) → **R3 + R4**.
2. Half-open GFW socket defeats the detector AND the whole-save on the same tick; three connections split-brain on "online" → **R1 + R1b** (and coordinate reconnect authority where cheap).
3. MIP collapses all three durability fallbacks at once → **R5**.
4. `app_version` NULL after every reconnect, worker poisons era-bucketing mid-session → **R2**.
5. Clock skew breaks staleness honesty + era inference → **R7**.

---

## Data model & schema changes (kept in lockstep: `ensureSchema` + `init.sql` + canonical migration)
- `ADD COLUMN app_version TEXT` (nullable) on `files`, `sensory_sessions`, `detailed_sensory_sessions`.
- Stamp trigger function + `BEFORE INSERT/UPDATE` triggers on those three tables (fail-safe, never RAISE).
- `notify_row_change` gains the `dve.maintenance` early-return.
- `dve_normalize_legacy_json()` function; one-time call guarded by a `schema_meta` key under `pg_advisory_xact_lock` in a single transaction; nightly `cron.schedule` entry.
- 6-arg `dve_commit_cell_json` numeric-CASE body (heal).
- Snapshot SQLite schema gains `app_version` on the three session tables; SELECT/INSERT lists updated with the count assertion.
- Connect strings gain `application_name` + keepalives + `tcp_user_timeout` + `statement_timeout`.

## Error-handling philosophy
- Never block the user; never silently drop a save or an edit. Heals are best-effort + logged; classification failure degrades to "filter unavailable" while browsing still works; the normalizer never deletes. The one place that is currently silent (failed offline-edit enqueue, MIP-unreadable stores) becomes **loud and recoverable** (R5).

## Testing strategy
- **Unit (C++):** `CompatClassifier` era boundaries (each release-date edge), health-flag rules (string-score detection, junk detection, missing-regime), NULL-stamp degradation.
- **DB-container (ephemeral `postgres:16`, `DVE_TEST_PG_CONN`):**
  - stamp trigger fills `app_version` from `application_name`; **survives a forced reconnect** (the gap no current test covers — existing tests only filter pg_cron *out* by application_name);
  - `dve_normalize_legacy_json` fixes seeded string scores, is a no-op on already-numeric rows, idempotent across re-runs;
  - one-time heal under `pg_advisory_xact_lock` runs exactly once under concurrent connects;
  - `dve.maintenance` GUC suppresses `notify_row_change`;
  - 6-arg numeric-CASE heal.
- **E2E (extend `tst_saveintegrity_e2e`):**
  - old client writes a string score → nightly/normalizer convergence → v2.4.2 reader gets a number;
  - **post-reconnect catch-up prevents the LWW clobber** (offline-window remote edit survives a local save) — the reset-to-5 regression guard;
  - `statement_timeout` makes a hung query fail-fast (no indefinite block) and 25P02 doesn't wedge the connection;
  - MIP-unreadable store (`%TSD-Header-###%`) triggers the bundled-python fallback / loud warning rather than silent loss.
- **Deployment self-test:** add a `SelfTest` check that the app's own connection reports a non-empty `application_name`.

## Out of scope (YAGNI)
- Pre-v2 compatibility. NOTIFY sequence/version tokens (full reload on reconnect is correct without them). A general write-ahead journal of individual keystrokes (the 2 s debounced snapshot is the accepted design). Re-enabling OCC. The minimum-version write gate (rejected — violates "never block").

## Versioning / rollout
- Lands as consecutive internal patches (v2.4.2 → v2.4.x), each clean-rebuilt + harness-verified. Wraps to deployable **v2.5.0** on the user's install-test approval. User performs the Synology drop manually. No automatic Synology interaction at any point.
