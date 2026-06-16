-- v2.4.2 R4 migration: dve_normalize_legacy_json()
--
-- Canonical record of the legacy-score normalizer (also defined in
-- deploy/postgres/init.sql and healed onto live DBs by DatabaseManager::
-- ensureSchema()'s create-or-replace). Provisioner/production parity.
--
-- Losslessly rewrite numeric-looking STRING score values to JSON numbers in
-- sensory_sessions + detailed_sensory_sessions. Idempotent: only rows that
-- still have a string-typed score matching the writer's numeric regex are
-- touched (so clean rows do not version-bump or NOTIFY). Sets dve.maintenance
-- so the per-row UPDATE does not storm clients, and SET LOCAL statement_timeout
-- = 0 so the full-table scan is not aborted by the 10s connection cap (v2.4.2
-- sub-plan 1). jsonb_agg ... ORDER BY elem.ord preserves sample order. Returns
-- the number of rows rewritten.
--
-- NOTE: the one-time advisory-locked run + the nightly pg_cron schedule live
-- elsewhere (DatabaseManager::ensureSchema one-time gate + init.sql cron tail);
-- this migration installs the function definition only.
CREATE OR REPLACE FUNCTION dve_normalize_legacy_json() RETURNS INTEGER AS $$
DECLARE
    keys TEXT[] := ARRAY[
        'Burnt Taste','Vapor Volume','Overall Flavor','Smoothness','Overall Liking',
        'Burn Taste','Flavor Intensity','Throat Irritation','Nasal Irritation',
        'Vapor Quality Overall','Cough','Volume Consistency',
        'Performance Consistency','Vapor Temperature','Vapor vs Oil'];
    total INTEGER := 0;
    n INTEGER;
    tbl TEXT;
BEGIN
    SET LOCAL statement_timeout = 0;
    PERFORM set_config('dve.maintenance', '1', true);  -- tx-local; suppresses NOTIFY
    FOREACH tbl IN ARRAY ARRAY['sensory_sessions','detailed_sensory_sessions'] LOOP
        EXECUTE format($q$
            UPDATE %I s SET json_data = jsonb_set(
                s.json_data, '{samples}',
                (SELECT jsonb_agg(
                    COALESCE(
                      (SELECT jsonb_object_agg(kv.key,
                          CASE WHEN kv.key = ANY($1)
                                AND jsonb_typeof(kv.value) = 'string'
                                AND (kv.value #>> '{}') ~ '^-?[0-9]+(\.[0-9]+)?$'
                               THEN to_jsonb((kv.value #>> '{}')::numeric)
                               ELSE kv.value END)
                       FROM jsonb_each(elem.value) kv),
                      elem.value)
                    ORDER BY elem.ord)
                 FROM jsonb_array_elements(s.json_data->'samples')
                      WITH ORDINALITY AS elem(value, ord)))
            WHERE jsonb_typeof(s.json_data->'samples') = 'array'
              AND EXISTS (
                SELECT 1 FROM jsonb_array_elements(s.json_data->'samples') e,
                              jsonb_each(e.value) kv
                WHERE kv.key = ANY($1)
                  AND jsonb_typeof(kv.value) = 'string'
                  AND (kv.value #>> '{}') ~ '^-?[0-9]+(\.[0-9]+)?$')
        $q$, tbl) USING keys;
        GET DIAGNOSTICS n = ROW_COUNT;
        total := total + n;
    END LOOP;
    RETURN total;
END;
$$ LANGUAGE plpgsql;
