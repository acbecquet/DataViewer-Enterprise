-- DATAVIEWER-11 follow-up (v2.5.14): backfill json_data header keys for
-- sensory_sessions created by the phone web form BEFORE 2026-06-26-dv11b
-- started seeding a complete blob. Those rows have json_data = '{"samples":[…]}'
-- with NO header keys, so the desktop -- which reads test title / tester /
-- assessor / media / date out of json_data (the Database Browser listing AND
-- the whole-session load) -- showed them with an empty tester that fell back to
-- the assessor ("the assessor field fills both assessor and tester"). The
-- natural-key COLUMNS are authoritative and already correct, so rebuild the
-- header keys from them and merge into json_data.
--
-- `jsonb_build_object(...) || json_data` keeps anything already in json_data
-- (the samples array) on top of the rebuilt headers. The WHERE clause targets
-- only rows missing the tester_name key (web-created), so desktop rows -- which
-- always carry it -- are untouched and re-applying is a no-op (idempotent).
-- The ::text casts make this safe whether `date`/`timestamp` are DATE/TIMESTAMP
-- or TEXT columns; either way the desktop reads them back as 'YYYY-MM-DD' style
-- strings, exactly as sensorySessionToJson emits them.
UPDATE sensory_sessions
   SET json_data = jsonb_build_object(
           'session_name',  session_name,
           'test_title',    session_name,
           'assessor_name', coalesce(assessor_name, ''),
           'tester_name',   coalesce(tester_name, ''),
           'media',         coalesce(media, ''),
           'date',          coalesce("date"::text, ''),
           'timestamp',     coalesce("timestamp"::text, '')
       ) || coalesce(json_data, '{}'::jsonb)
 WHERE NOT (coalesce(json_data, '{}'::jsonb) ? 'tester_name');
