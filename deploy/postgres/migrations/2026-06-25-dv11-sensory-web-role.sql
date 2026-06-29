-- DATAVIEWER-11: dedicated least-privilege role for the phone web form.
-- The password is set out-of-band on the NAS from a .env secret (never committed).
DO $$ BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_roles WHERE rolname = 'sensory_web') THEN
        CREATE ROLE sensory_web LOGIN;
    END IF;
END $$;

-- Strip everything the role might have, then grant ONLY what the form needs.
REVOKE ALL ON ALL TABLES    IN SCHEMA public FROM sensory_web;
REVOKE ALL ON ALL FUNCTIONS IN SCHEMA public FROM sensory_web;
REVOKE ALL ON ALL SEQUENCES IN SCHEMA public FROM sensory_web;
GRANT  USAGE ON SCHEMA public TO sensory_web;
GRANT  SELECT, INSERT, UPDATE ON sensory_sessions TO sensory_web;
GRANT  USAGE, SELECT ON SEQUENCE sensory_sessions_id_seq TO sensory_web;
GRANT  EXECUTE ON FUNCTION dve_append_sensory_sample(TEXT,TEXT,TEXT,TEXT,TEXT,JSONB,TEXT) TO sensory_web;

-- CRITICAL least-privilege: PostgreSQL grants EXECUTE on every function to PUBLIC
-- by default, and sensory_web is a PUBLIC member -- so "REVOKE EXECUTE ... FROM
-- sensory_web" alone leaves the generic mutators executable (verified:
-- has_function_privilege returned TRUE for both dve_commit_cell overloads). Revoke
-- EXECUTE on every dve_commit_cell* overload FROM PUBLIC (and the role) so the web
-- caller cannot reach the "mutate any table by id" primitives whose SQL-safety is
-- delegated to the C++ client allowlist a web caller bypasses. The schema owner
-- (dataviewer_app) keeps EXECUTE regardless -- owners bypass GRANT checks -- so the
-- desktop is unaffected. The signature-agnostic loop covers the 5/6-arg
-- dve_commit_cell and 6/7-arg dve_commit_cell_json overloads (and any future one).
DO $$
DECLARE r record;
BEGIN
    FOR r IN SELECT p.oid::regprocedure AS sig
               FROM pg_proc p JOIN pg_namespace n ON n.oid = p.pronamespace
              WHERE n.nspname = 'public' AND p.proname LIKE 'dve_commit_cell%'
    LOOP
        EXECUTE format('REVOKE EXECUTE ON FUNCTION %s FROM PUBLIC, sensory_web', r.sig);
    END LOOP;
END $$;
