# Post-Wave Bug-fix Batch — Design

- **Date:** 2026-05-15
- **Status:** Draft (pending user review)
- **Context:** Postgres multi-user sprint (v2.0) shipped. A laundry list of
  small issues surfaced during real-world multi-user use. This spec covers
  the batch of fixes — most one-shot, two larger — needed to smooth out
  the v2.0 experience.

## Scope

Seven issues identified by the user. Triaged into two tracks:

**Track A — Bug-fix batch (this spec):** Items #1, #3, #4, #5, #6, #7.
Ship as one focused branch with small reviewable commits. Estimated 1–2
days of focused work.

**Track B — Live collaborative editing (separate spec, follows this one):**
Item #2 escalated from "smarter save-time merge" to "Google-Sheets-style
live cell sync" per user preference. This is a 1–2 week initiative
that replaces the v2.0 optimistic-lock conflict layer with cell-commit
live sync. Gets its own design doc immediately after this batch ships.

## Non-goals (this spec)

- Live collaborative editing. See Track B preview at the end.
- Schema redesign beyond the two new columns required for #7.
- Replacing existing conflict-dialog UI — stays in place until Track B
  retires it.
- Reworking the Excel write-back pipeline.

## Per-item design

### #1 — "New session" collision (Sensory + Detailed Sensory)

**Symptom:** When two users each create a new sensory session, both
sessions appear in each other's navigator as a shared "new session"
row. They are not isolated until first save.

**Root cause:** Presence activation happens on session creation, before
the session has a real DB id. A placeholder id (or fixed sentinel) is
used, so multiple users' presence rows collide on the same resource.

**Fix:** In `SensoryPanel` and `DetailedSensoryPanel`, defer
`PresenceManager::activate()` until `session.id > 0` (after first
successful save). For unsaved sessions, presence is local-only — the
user does not appear in any other client's presence list, and no other
user shows up in this client's presence list for an unsaved session.

**Files touched:**
- `src/ui/SensoryPanel.cpp` — gate `m_presence->activate(...)` on `id > 0`.
- `src/ui/DetailedSensoryPanel.cpp` — same.
- After first successful save returns the new id, call `activate()` then.

**Test:** Existing `tst_presencemanager` plus a new manual two-client
scenario documented in `tests/deployment/README.md` (two installs, both
create a new sensory session, verify no cross-pollination until either
saves).

---

### #3 — Color picker shows taken colors

**Symptom:** New users pick a color already in use by a coworker, with
no visual indication.

**Fix:** In `IdentityPromptDialog`, on dialog open, query the
`presence` table for distinct `user_color` values whose `last_heartbeat`
is within the 30-second freshness window. For each color swatch in the
picker:
- If color is in the taken set: draw a 2px solid ring around the swatch
  in a neutral gray, plus `setToolTip("Taken by <name>")` showing the
  current user(s) using it.
- User can still select a taken color (not blocked, just visually
  discouraged).

**Files touched:**
- `src/database/IdentityPromptDialog.cpp` — add `populateTakenColors()`
  helper called from constructor; modify swatch widget paint to honor
  the taken flag.

**Test:** Existing dialog tests + manual verification on the work machine.

---

### #4 — Smooth presence updates

**Symptom:** Presence dots in the navigator do not refresh in real-time;
they only update after the local user takes an action (clicks something,
selects a file, etc.).

**Investigation needed first:** Trace `NotificationListener` →
expected presence-changed signal → navigator refresh path. Current
hypothesis: the LISTEN payload arrives but is not wired to a slot that
re-queries `PresenceManager::activeFor()` and invalidates the
`QTreeWidgetItem` data role.

**Likely fix:** Add a `MainWindow` slot connected to
`NotificationListener::presenceChanged(QString resourceType, qint64
resourceId)`. The slot:
1. Calls `m_presence->activeFor(resourceType, resourceId)`.
2. Finds the matching `QTreeWidgetItem` in the navigator (or
   `QListWidgetItem` for sensory sessions).
3. Updates `kColorsRole` and `kIntentsRole` data.
4. Calls `update()` on the view so `PresenceDotsDelegate` repaints.

**Files touched:**
- `src/MainWindow.cpp` — new slot + connection.
- Possibly `src/database/NotificationListener.cpp` if the signal does
  not exist or carries wrong payload.

**Test:** Update `tst_notificationlistener` to assert the
presence-changed signal fires within 500ms of a remote presence row
change. Manual two-client test.

---

### #5 — Presence dots position when filename truncates

**Symptom:** When a filename in the navigator is too long, Qt elides
the text with `…`. The presence dots paint *after* the text, so
the ellipsis eats the dots. Users cannot tell who else is on a
long-named file.

**Fix:** In `PresenceDotsDelegate::paint()`, restructure the layout:
1. Compute total available width from `option.rect`.
2. Compute dots width = `N * (dotSize + spacing)` where N is the count
   from `kColorsRole`.
3. Reserve dots region at the right edge of `option.rect`.
4. Compute remaining text region = total minus dots minus padding.
5. Elide the text into the remaining region using
   `QFontMetrics::elidedText(..., Qt::ElideRight)`.
6. Paint text, then dots in the reserved right region.

`sizeHint` must include the dots width so the layout reserves space.

**Files touched:**
- `src/widgets/PresenceDotsDelegate.cpp` — rewrite paint and sizeHint.

**Test:** Add a `tst_presencedotsdelegate` test class (currently no
unit test exists for it) covering: short text + dots, long text +
dots (elision triggered), no dots (renders as plain text), one dot,
many dots.

---

### #6 — Draw pressure y-axis scaling

**Symptom:** Draw pressure plots auto-scale their y-axis based on the
data range. For low-pressure draws, the axis might top out at 0.8 or
1.2, which makes small variations look dramatic. Users want a fixed
floor.

**Fix:** In the plot engine, for the draw-pressure series only:
```
yMax = max(2.0, ceil(seriesMax))
```
Where `seriesMax` is the largest value across all visible series on
that chart. `ceil()` rounds up to the next integer (so 2.7 → 3,
3.1 → 4). y-axis floor stays at 0.

Applied identically in the UI plot widget and the PPTX report image
plot so the report matches the on-screen view.

**Files touched:**
- `src/pipeline/PlotEngine.cpp` (or wherever the draw-pressure axis
  computation lives — verify during implementation).
- Same logic invoked by report generation in `ReportGenerator` /
  `PptxWriter`.

**Test:** Extend `tst_plotengine` with cases: max=0.5 → axis=2,
max=2.0 → axis=2, max=2.7 → axis=3, max=5.1 → axis=6.

---

### #7 — Sensory mode: power_type + puff_length

**Symptom (feature request):** Sensory data collection lacks two
important fields:
- **Power type** — categorizes the device under test as one of
  Constant Voltage, Constant Power, Variable Voltage, Variable Power.
- **Puff length** — how long each puff was, in seconds.

Both need to surface in the data-collection UI and in generated reports.

**Schema change:** Add two columns to the sensory sample table
(verify exact table name during implementation):
```sql
ALTER TABLE sensory_sample
  ADD COLUMN power_type TEXT NOT NULL DEFAULT 'Constant Voltage',
  ADD COLUMN puff_length_sec REAL NOT NULL DEFAULT 3.0;
```
Existing rows backfill to those defaults per user direction.

**Pipeline struct change:** In `src/pipeline/SensoryData.h`, add to the
sample-level struct:
```cpp
QString powerType    = "Constant Voltage";
double  puffLengthSec = 3.0;
```
Scope is **Sensory mode only**. `DetailedSensoryData` is explicitly out
of scope for this batch — the user requested these fields for the
sensory data-collection workflow specifically. If parity is wanted
later, that's a separate small fix.

**UI change in `SensoryPanel`:**
- Add `QComboBox` per sample card with the four options. Placement:
  below the existing voltage/resistance/power row.
- Add `QDoubleSpinBox` per sample card. Suffix `" s"`, range
  0.1–60.0, step 0.5, decimals 1. Placement: between the scoring
  section and the comments text edit. (User-specified: "above
  comments but below the scoring".)
- Both widgets bind to the underlying `SensorySample` struct fields
  and trigger the standard "modified" flag on edit.

**Report integration:**
- In the sensory report table, add a `Puff Length (s)` column placed
  immediately before the `Notes` column.
- In the existing voltage/resistance/power cell, append the power_type
  on a new line in parentheses: `3.7V / 1.2Ω / 11W\n(Variable Voltage)`.
- Apply to both the legacy direct-to-PPTX path and the
  `ReportPreviewDialog` path (since the preview replaced direct-to-PPTX
  for sensory).

**Files touched:**
- `deploy/postgres/init.sql` — schema additions.
- New migration script under `deploy/postgres/migrations/`.
- `src/pipeline/SensoryData.h` — struct fields.
- `src/database/DatabaseManager.cpp` — read/write of new columns.
- `src/ui/SensoryPanel.cpp` and `.h` — UI widgets, signal wiring.
- `src/reporting/SensoryReportSource.cpp` — expose new columns to
  the preview/PPTX.
- `src/reporting/PptxWriter.cpp` / `ReportGenerator.cpp` — column
  layout, voltage cell formatting.

**Test:** Extend `tst_sensoryreportsource` to verify both fields
round-trip from `SensorySession` → PPTX bytes. Extend the schema
test (`tst_postgres_*`) to verify migration applies cleanly. Extend
`tst_sensorypanel` (or add one if absent) to verify widget binding.

---

## Track B preview — #2 live collaborative editing

Spec'd separately in a follow-on `live-collab-design.md` (next
initiative after this batch ships). High-level direction the user has
already agreed to:

- Every cell commit (focus leaves the cell, or Enter pressed) writes
  directly to Postgres and fires NOTIFY.
- Every other open client applies the change in real time via the
  existing `NotificationListener` infrastructure.
- Cells currently being edited by another user show a colored border
  (the user's presence color) so collisions are visible *before*
  they happen.
- Last-writer-wins on a same-cell collision — no operational
  transform, no CRDT.
- "Save" button repurposed: it stops being a conflict boundary and
  becomes "flush in-flight cell writes + write back to Excel".
- The v2.0 optimistic-lock layer (SaveCoordinator / ConflictResolver /
  VersionMismatchDialog / RowDeletedDialog) becomes dead code and is
  removed. `UniqueViolationDialog` stays — unique-key violations
  are still a real failure mode.

Track A items #1 and #4 are explicitly designed to integrate cleanly
with Track B: #1's "defer presence activation until saved id exists"
holds true regardless of whether saves are batched (today) or
per-cell (Track B). #4's "NOTIFY-driven navigator refresh" is the
same machinery Track B will extend.

## Execution plan

After this spec is reviewed and approved, transition to writing-plans
to produce `2026-05-15-postwave-bugfix-batch-plan.md` covering the six
items above as ordered, checkable steps. Track B gets a separate
brainstorm → spec → plan cycle afterward.
