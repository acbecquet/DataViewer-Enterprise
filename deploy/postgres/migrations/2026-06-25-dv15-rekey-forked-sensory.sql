-- DATAVIEWER-15: re-key already-forked sensory rows onto the canonical
-- title-only session_name convention.
--
-- *** do NOT auto-apply -- run manually after a backup. ***
--
-- The desktop historically had three divergent session_name derivations: the
-- live-save path used the Test Title ALONE, while the two Excel/panelist import
-- paths used "title - tester". Re-saving an imported session via the UI thus
-- silently changed its natural key (session_name, tester_name, date), forking
-- it into a duplicate row. DATAVIEWER-15 (desktop Task 1.4) unified all three
-- code paths on the title-only convention, so NEW saves are consistent. This
-- migration repairs the rows that were already forked under the old paths.
--
-- The unique key is idx_sensory_sessions_key ON sensory_sessions(session_name,
-- tester_name, date). Re-keying a forked row to session_name = test_title can
-- collide with an already-canonical row that shares (test_title, tester_name,
-- date). The UPDATE below is guarded to SKIP such collisions (they need manual
-- resolution) and is idempotent: a re-run is a no-op once the safe rows are
-- already re-keyed.
--
-- USAGE:
--   1. Take a DB backup (pg_dump).
--   2. Run STEP 1 (dry-run) and review the output. Rows flagged would_collide=t
--      must be reconciled by hand (the canonical row already exists).
--   3. Run STEP 2 (the guarded UPDATE) to re-key the safe rows.

-- ── STEP 1: DRY-RUN -- list forked rows + their target key + collision flag ──
-- Run this on its own first; it mutates nothing.
SELECT
    f.id,
    f.session_name                        AS current_session_name,
    f.json_data->>'test_title'            AS target_session_name,
    f.tester_name,
    f.date,
    EXISTS (
        SELECT 1 FROM sensory_sessions c
         WHERE c.id          <> f.id
           AND c.session_name = f.json_data->>'test_title'
           AND c.tester_name  = f.tester_name
           AND c.date         = f.date
    )                                     AS would_collide
FROM sensory_sessions f
WHERE f.json_data->>'test_title' IS NOT NULL
  AND f.session_name IS DISTINCT FROM f.json_data->>'test_title'
ORDER BY would_collide DESC, f.id;

-- ── STEP 2: GUARDED RE-KEY -- only the rows that will NOT violate the index ──
-- Transaction-wrapped; idempotent; collisions are skipped + must be resolved
-- manually (see the STEP 1 dry-run output, would_collide = true).
--
-- Intra-batch safety: if TWO forked rows both target the same key (same
-- test_title + tester_name + date) and neither currently holds it, a naive
-- "NOT EXISTS against current session_name" guard would clear for BOTH and the
-- UPDATE would try to set both to the same key -> idx_sensory_sessions_key
-- unique violation, aborting the whole batch. The row_number() tiebreak re-keys
-- AT MOST ONE forked row per target key per run; a re-run then sees that row as
-- the canonical holder (would_collide = true for the rest) and leaves the
-- genuine duplicates for manual resolution. Idempotent: a re-run is a no-op once
-- every safely-movable row is re-keyed.
BEGIN;

WITH forked AS (
    SELECT f.id,
           f.json_data->>'test_title' AS target,
           f.tester_name,
           f.date,
           row_number() OVER (
               PARTITION BY f.json_data->>'test_title', f.tester_name, f.date
               ORDER BY f.id
           ) AS rn
      FROM sensory_sessions f
     WHERE f.json_data->>'test_title' IS NOT NULL
       AND f.session_name IS DISTINCT FROM f.json_data->>'test_title'
)
UPDATE sensory_sessions s
   SET session_name = fk.target
  FROM forked fk
 WHERE s.id = fk.id
   AND fk.rn = 1                       -- one forked row per target key per run
   AND NOT EXISTS (                    -- skip if a canonical row already holds it
        SELECT 1 FROM sensory_sessions c
         WHERE c.id          <> fk.id
           AND c.session_name = fk.target
           AND c.tester_name  = fk.tester_name
           AND c.date         = fk.date
   );

COMMIT;
