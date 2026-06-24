# DATAVIEWER-11 — Remote Phone Sensory Web Form: Design Spike (sub-spec)

**Date:** 2026-06-24 · **Status:** SPIKE COMPLETE — awaiting owner sign-off (Phase 5a gate) · **Parent:** MS-8 in `docs/superpowers/specs/2026-06-23-backlog-master-spec.md` · **Plane:** DATAVIEWER-11 (High).

> This sub-spec is the output of Phase 5a (the MS-8 design spike). It **supersedes the original MS-8 assumptions** ("INSERT-only, office-WiFi-acceptable, zero C++"). Where this spike and the master spec disagree, this spike is the newer source for MS-8 and the master spec has been updated to match. It was produced from a code-grounded investigation + an adversarial stress pass (12 agents) whose every robustness claim against the naive design **failed at HIGH severity** — those failures are the reason for the design below.

---

## 1. Purpose & owner decisions (locked 2026-06-24)

Build a **mobile-first web form** that lets a tester (often **offsite, e.g. with a customer**) record a **5-metric Sensory** sample and, on **"Submit sample,"** write it into the same PostgreSQL `sensory_sessions` schema the desktop uses — so every desktop sees it live via NOTIFY. The desktop app is **not** an HTTP server and never becomes one; the form is a standalone NAS Docker stack.

Owner answers to the four Phase-0 blockers:

1. **Reachable from anywhere** over the internet (offsite customers). **Office-Wi-Fi-only is REJECTED.**
2. **HTTPS at transport; NO user-facing auth** (owner decision 2026-06-24 — anonymous, frictionless). All access control is backend: a least-privilege DB role + server-side validation + invisible proxy rate-limit (§5).
3. **Online-only** — no offline capture/queue. "If offline, it doesn't work" is acceptable.
4. **Create-or-append by the session's identity** — a new sample whose headers match an existing session is **appended to that session**, not rejected and not forked. *(The owner's mental model "name = test title - tester - R#" is the on-disk filename label; the DB natural key works differently — see §2.)*

The form UX (owner's words): **one sample at a time**; the headers (Test Title, Assessor, Tester, Round, Media) **persist after the first submit and stay editable**; only Sample Name changes for the next sample; loop sample→round→test smoothly.

---

## 2. CRITICAL correction — the DB natural key is not "title - tester - R#"

The owner's identity model describes the **filename**, not the DB unique key. Grounded in code:

- **`session_name` = the Test Title ALONE** on the live save path — no tester, no round (`src/ui/SensoryPanel.cpp:1029-1030`). `sess.testTitle` is stored separately (`:1037`).
- **Round lives in `tester_name`**, folded as a trailing `" R1"`/`" R2"` suffix via `combineTesterRound` (`src/ui/SensoryPanel.cpp:1039`; `src/ui/TesterRound.h:31-37`); split back out in `applySession` (`:1102-1104`). Round **"N/A"** (a real combo option, `:815`) appends **nothing**.
- **The unique index** is `idx_sensory_sessions_key ON sensory_sessions(session_name, tester_name, date)` — three **raw TEXT** columns, **no** `lower()`/trim/normalization (`deploy/postgres/init.sql:145-146`).
- **`date` is a local-time `yyyy-MM-dd` TEXT string** (`src/ui/SensoryPanel.cpp:905,1041`), distinct from the separate ISO-UTC `timestamp` column. So two rounds *are* distinct rows — but only because `tester_name` differs (`"Bob R1"` vs `"Bob R2"`), **not** because `session_name` differs.
- **Three inconsistent `session_name` derivations exist** in the desktop today: live save = title-only (`:1030`); saved-Excel import = `title + " - " + tester` (`:2183`); panelist-template import = `title + " - " + tester` (`:2314`). The migration comment confirms the canonical convention keys on `json_data->>'test_title'`, **not** `session_name` (`deploy/postgres/migrations/2026-06-07-dv2-sample-name-by-test.sql:23`). **This is a pre-existing desktop bug** — re-saving an imported session via the UI silently changes its key from `"title - tester"` to `"title"`, forking it. See §9.
- **Collision policy today = FORK, not append.** On a natural-key collision at close the desktop calls `tryWriteSensorySessionAutoSuffix`, which bumps `session_name`/`testTitle` with `_1`/`_2` and re-INSERTs a **second row** (`src/MainWindow.cpp:2954-2956`; `src/database/DatabaseManager.cpp:2067-2093`). A fresh-struct INSERT on collision returns `UniqueViolation` (`:1879-1881`). **There is no "append a sample to an existing session" code path today.**

**Consequence:** the web service must **not** try to "derive the same `session_name` string and INSERT." It must **resolve the existing row by the three key columns and append**, creating only on a true miss (§3).

---

## 3. Write contract — create-or-append by natural key (supersedes MS-8 "INSERT-only")

A new **purpose-built stored function** (`dve_append_sensory_sample`) the web service calls. It hard-codes the `sensory_sessions` table (it is **not** the generic `dve_commit_cell*` primitive — see §5) and does, in one transaction:

1. **Compute the key columns exactly as the desktop live path does:**
   - `session_name` := **trimmed Test Title only** (matches `SensoryPanel.cpp:1030`; *not* the import shape).
   - `tester_name` := `combineTesterRound(trimmed tester, round)` — append exactly `" R1"`/`" R2"` for round 1/2, nothing otherwise. Re-implement `TesterRound.h:31-37` / split regex `^(.*\S)\s+R([12])$` verbatim.
   - `date` := **office-timezone** `yyyy-MM-dd` computed **server-side on the NAS** (NOT the phone's local date — an offsite phone in another TZ near midnight would otherwise produce a different calendar date and fork the key).
2. **Resolve or create:** `INSERT INTO sensory_sessions(session_name, tester_name, date, json_data, …) VALUES(…, '{"samples":[]}'::jsonb …) ON CONFLICT (session_name, tester_name, date) DO NOTHING;` then `SELECT id FROM sensory_sessions WHERE <key> FOR UPDATE;`
3. **Append at the TAIL** (never reorder/insert/delete existing elements): `UPDATE sensory_sessions SET json_data = jsonb_set(json_data, '{samples}', COALESCE(json_data->'samples','[]'::jsonb) || $sample::jsonb) WHERE id = $id;` The `FOR UPDATE` serializes concurrent appends (mirrors the desktop's own `FOR UPDATE` merge at `DatabaseManager.cpp:1783`); `bump_version`/`notify_row_change` triggers auto-bump `version`/`updated_at` and emit `NOTIFY dataviewer_changes` (`init.sql:257-286,314-366`), so desktops refresh for free. The service does **not** manage `version` or call `pg_notify`.
4. **Idempotency:** each sample object carries a client-generated `sample_uid`; the append is a **no-op** if a sample with that uid already exists in the array (guards a retried POST over a flaky connection — samples otherwise have **no** stable id, only array position).

Scores **must be JSON numbers in [1,9]**, never strings (string scores are the DATAVIEWER-4 "reset-to-5" data-loss class; `src/pipeline/SensoryData.h:13-32`).

---

## 4. Data-safety blockers the spike found (MUST be fixed before DV-11 ships)

The adversarial pass found the naive "just append" design is **not** safe. Two of these are required fixes.

### 4.1 — REQUIRED (C++): the desktop whole-session save silently DROPS phone-appended samples

`mergeSensoryPreservingDbScores` seeds `merged = inMemory`, overlays DB scores only over the index-overlap `i < min(mem, db)`, and returns `merged["samples"] = memSamples` (`src/pipeline/SensoryData.cpp:107,110,124`). A DB-only sample (the phone's append, at an index past the desktop's in-memory count) is **never copied out** and is overwritten away. The version guard does not save it: the whole-session save is **row-level last-writer-wins by design** and re-adopts the fresh DB version (`src/database/DatabaseManager.cpp:1767-1797`). The live NOTIFY path is score-only and sensory-inert (`onRemoteCellChanged` returns for `table != "data_rows"`, `src/MainWindow.cpp:3657`), so the desktop's in-memory sample count never grows while the session is open.

**Result:** any desktop **save of an open session** (interactive or auto-save-at-close) **permanently deletes** a sample the phone appended while it was open — silently, no conflict. This is the exact inverse of the owner's locked requirement #4.

**Fix (surgical, ~3 lines + a card-rebuild):** in `mergeSensoryPreservingDbScores`, when `dbSamples.size() > memSamples.size()`, **append the trailing DB-only samples** to `merged["samples"]` before returning (the desktop never had them in memory, so there is no local edit to reconcile). Pair with growing the card list in `SensoryPanel::applyMergedScoresToCurrentSession` (rebuild cards when the merged blob has more samples than the struct). **Add a regression test** (none exists): in-memory N samples + DB row N+1 → save → reload still shows N+1. → **This makes DV-11 carry a desktop-side C++ change in the v2.6.0 train (it is NOT zero-C++).** See §10 decision (a).

### 4.2 — REQUIRED: stable per-sample `sample_uid`

Samples are addressed purely by array index (per-cell LiveSync paths `samples[i].<metric>`; the merge loops are index-bounded). A stable `sample_uid` (a) gives the §3 idempotency key, and (b) lets the merge reconcile by id instead of fragile position (defends against a future concurrent insert/delete). It is additive — `sensorySessionFromJson` ignores unknown keys. The minimal-viable invariant if we defer uid-keyed merging: **append-to-tail only** (the web form may never edit/remove existing samples), which keeps existing indices valid for in-flight desktop edits.

### 4.3 — Append-to-tail-only invariant

The web form only ever pushes to the end of `samples[]`; it never edits or removes an existing sample. This keeps absolute positional paths (`samples[N]`) of existing samples stable so in-flight desktop per-cell edits can't cross wires.

---

## 5. Security — least privilege, not the superuser app role

The only DB role is **`dataviewer_app`**, which the official Postgres image makes a **superuser + schema owner**; there is **no** least-privilege role anywhere (`deploy/postgres/docker-compose.yml:8-11`; zero `CREATE ROLE/GRANT/REVOKE/POLICY` in `init.sql` + all migrations). The generic `dve_commit_cell` / `dve_commit_cell_json` are "mutate **any** table by id" primitives whose SQL-safety is **delegated to the C++ client's allowlist** (`init.sql:424-425`) — a web caller bypasses that. A single shared passcode + this role would gate **full write to the entire prod DB** (a web-service bug → superuser → `COPY … TO PROGRAM` RCE).

**Required controls (new `deploy/postgres` migration + service config):**
1. **Dedicated login role `sensory_web`** with **only** `SELECT, INSERT, UPDATE` on `sensory_sessions` (+ `USAGE` on its sequence), and **`REVOKE EXECUTE`** on `dve_commit_cell*` and all other tables. The service connects as `sensory_web`, never `dataviewer_app`. Its password is a **second** NAS-env secret (mirror `.env.example`; repo is public — never commit).
2. **The service never calls the generic `dve_commit_cell*`** — only `dve_append_sensory_sample` (table name hard-coded).
3. **No user-facing auth (owner decision 2026-06-24): anonymous, frictionless.** The one-time customer submits with **no password or token**. Protection is **backend / invisible**: the `sensory_web` least-privilege role bounds the blast radius (append sensory rows only); strict input validation (required fields, scores 1–9, exact JSON shape); invisible reverse-proxy abuse controls (per-IP rate-limit, request-size cap; honeypot/CAPTCHA only if needed). The NAS backend is the sole gatekeeper (validate-then-write); `updated_by = "web/<tester-typed-name>"` is an unauthenticated label. **Residual risk (accepted):** an open anonymous endpoint can post junk sensory rows — bounded by the least-priv role, deletable by a desktop user — the deliberate trade for a frictionless one-time-customer UX.
4. The DB stays firewalled to the office VLAN; the service reaches it by **internal Docker service name** `dataviewer-db:5432` (note the doc/code 5432-vs-5433 ambiguity — moot for an internal connection). **Flag to owner.**

---

## 6. Network / hosting

- Phones reach the service **over the internet** via the **existing NAS reverse proxy + TLS** (the same precedent as the `bug-form-app` that files these Plane reports — **NAS-only, confirm with owner**; nothing about the proxy/bug-form-app is in-repo). New stack at `deploy/sensory-collect/` on a **shared Docker network** with the DB (net-new wiring — the committed compose has no `networks:` block).
- **Online-only.** No offline/PWA.
- Secrets (the `sensory_web` DB password) live only in a NAS `.env` (repo is public — never commit).

---

## 7. Exact JSON contract (byte-compatible with `sensorySessionToJson`/`FromJson`)

Canonical for all three desktop persistence paths (Postgres JSONB, offline snapshot, .json export); any new field must be added to both serializers + `tst_sensorydataplaceholder` (`src/pipeline/SensoryData.h:194-201`).

**Root (session) keys (18 + `samples`):** `session_name`, `test_title`, `assessor_name`, `tester_name`, `media`, `date`, `timestamp`, `control`, `is_blind` (bool), `primary_differences`, `puff_length` (legacy session string), `burn_status`, `clog_status`, `leak_status`, `resistance`, `voltage`, `power` (legacy session numbers), `heating_technology`, and `samples` (array). (`src/pipeline/SensoryData.cpp:14-48`.)

**Per-sample keys (13):** `name`, `comments`, the **five literal metric keys** `"Burnt Taste"`, `"Vapor Volume"`, `"Overall Flavor"`, `"Smoothness"`, `"Overall Liking"` (flat, **not** nested under `scores`; `SensoryData.h:48-50`, `SensoryData.cpp:38-39`), `voltage`, `resistance`, `power`, `heating_technology`, `power_type`, `puff_length_sec`. Plus the spike's added `sample_uid` (§4.2).

**Types/defaults:** scores = JSON numbers, default 5.0, clamped [1,9] on read; `power_type` default `"Constant Voltage"`; `puff_length_sec` default 3.0; `voltage/resistance/power` default 0.0; strings default empty; `is_blind` default false. The reader tolerates absent/optional fields, so the phone form may omit device fields.

**Do NOT emit** (loader/UI anchors, not in the contract): `id`, `version`, `sourceFilePath`, `sourceImagePath`, `imagePaths`/`imageIds`/`imageLayouts`/`imageVersions`, `excelLayoutJson`, `originalSessionName`, `dirtyCells` (`SensoryData.h:99-138`).

---

## 8. Form UX (owner's flow) → contract mapping

One sample/screen: header fields (Test Title, Assessor, Tester, Round{1,2,N/A}, Media) + Sample Name + five 1–9 inputs + Puff length + Comments + optional device fields. On Submit: validate (passcode/token; non-empty Test Title + Tester per `isSensorySessionSavable`, `SensoryData.cpp:182-187`; scores 1–9), POST → `dve_append_sensory_sample`. After the first submit, headers persist client-side and stay editable; Sample Name clears for the next sample. Changing a header just changes which natural-key session the next append targets (resolve-or-create handles it).

---

## 9. Pre-existing desktop bug to flag (separate item)

The three inconsistent `session_name` derivations (§2) mean **re-saving an imported sensory session via the UI changes its natural key** (`"title - tester"` → `"title"`), forking it into a duplicate row — independent of DV-11, and it undermines web/desktop coexistence. **Recommend a separate backlog item** to unify all three paths on the title-only convention (or pick one and migrate). Do not silently fix it inside DV-11.

---

## 10. Owner decisions — RESOLVED 2026-06-24 (Phase 5a sign-off)

a. **Do the §4.1 desktop merge fix in the v2.6.0 train (recommended)** so DV-11 is data-safe end-to-end — vs. ship the web service alone and accept that a desktop save of an open session deletes phone appends (NOT recommended; that is active data loss). *This is the one that flips MS-8 from "zero C++" to "one required C++ fix."*
b. **Auth: NO user-facing auth — anonymous & frictionless** (decided). A one-time customer carries no password/token; all controls are backend (least-priv `sensory_web` role + validation + invisible proxy rate-limit). The NAS backend validates-then-writes. See §5.
c. **Round "N/A" semantics:** two N/A sessions of the same title+tester+date are the *same* natural key — confirm they should merge (append) rather than be distinguishable.
d. **Unify the 3 session_name derivations (§9): YES** — created as **DATAVIEWER-15** (a DV-11 prerequisite). One canonical title-only derivation across live-save + both Excel-import paths; audit Detailed for the same; re-key already-forked rows.

---

## 11. Acceptance criteria

- A non-developer submits a 5-metric session from a phone **off the office network**; scores stored as JSON **numbers** 1–9 (pytest round-trip).
- **Append-on-match:** a second submit with identical Test Title + Tester + Round + (office-TZ) date **appends a sample to the same row** (no fork, no `UniqueViolation`); a different header set creates a new row.
- **No data loss across the desktop:** with the §4.1 fix, a desktop that had the session open and then saves it **retains** the phone-appended sample (regression test green); a running desktop receives the append live via NOTIFY, renders the radar, re-saves cleanly.
- **Idempotency:** a re-POSTed sample (same `sample_uid`) does not duplicate.
- **Least privilege:** the web role can write **only** `sensory_sessions`; a pytest asserts `permission denied` on other tables and `EXECUTE` denied on `dve_commit_cell*`. A repo grep finds **no** secret (incl. no superuser credential in the web stack).
- Empty Test Title/Tester rejected with a clear message; the endpoint is anonymous (no user password/token) and abuse is bounded by per-IP rate-limit + the least-privilege role.
- The service runs as a NAS Docker stack reachable over the internet via the reverse proxy + TLS, reaches the DB by internal service name, survives a restart.

---

## 12. Build sequence (revises master-spec Phase 5)

> **Numbering map to the companion plan:** the desktop merge fix below = plan **Phase 5.0** (ships in v2.6.0); this spike = plan **Phase 5a**; the DB migration + service = plan **Phase 5b** (independent NAS track). The "5b/5c/5d" labels below are this doc's internal sequence only.

- **5a (this doc) — DONE pending owner sign-off** on §10.
- **5b — desktop merge-safety fix (C++, in the v2.6.0 train):** §4.1 fix to `mergeSensoryPreservingDbScores` + `applyMergedScoresToCurrentSession` + regression test. Gated behind the MS-7 heavy-verification discipline (it touches the load-bearing whole-session path). **This is the only DV-11 work that ships in v2.6.0.**
- **5c — DB migration:** `dve_append_sensory_sample` (hard-coded table, FOR UPDATE, jsonb tail-append, uid idempotency) + the `sensory_web` least-privilege role.
- **5d — the service:** `deploy/sensory-collect/` (Flask + Dockerfile + compose on the shared network), anonymous endpoint + backend validation + invisible per-IP rate-limit, office-TZ date, `pytest`. Independent NAS versioning track — **not** part of the v2.6.0 wrap.
