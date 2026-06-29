-- DATAVIEWER-11 (v2.5.13 follow-up): create-or-append a single sensory sample
-- by natural key. Extends the round suffix from R1/R2 to R1-R4 (rounds 3 & 4
-- universal). CREATE OR REPLACE -- safe to re-apply; supersedes the 2026-06-25
-- definition. R1/R2 behaviour is unchanged.
-- DATAVIEWER-11: create-or-append a single sensory sample by natural key.
-- Called ONLY by the sensory_web role from the phone web form. Hard-codes the
-- sensory_sessions table (it is NOT the generic dve_commit_cell* primitive).
CREATE OR REPLACE FUNCTION dve_append_sensory_sample(
    p_session_name TEXT,   -- trimmed Test Title ONLY (canonical key)
    p_tester       TEXT,   -- trimmed tester, no round suffix
    p_round        TEXT,   -- '1'..'4' => ' R1'..' R4'; anything else => no suffix
    p_assessor     TEXT,
    p_media        TEXT,
    p_sample       JSONB,  -- one sample object: numeric scores [1,9] + sample_uid
    p_office_tz    TEXT DEFAULT 'America/Phoenix'   -- Arizona: MST/UTC-7 all year, NO DST
) RETURNS BIGINT
LANGUAGE plpgsql AS $$
DECLARE
    v_tester TEXT := btrim(coalesce(p_tester, ''));
    v_date   TEXT;
    v_id     BIGINT;
    v_uid    TEXT := p_sample->>'sample_uid';
    v_metric TEXT;
BEGIN
    -- Defense in depth: scores must be JSON numbers in [1,9] (string scores are
    -- the DATAVIEWER-4 reset-to-5 data-loss class). The service validates too.
    FOREACH v_metric IN ARRAY ARRAY['Burnt Taste','Vapor Volume','Overall Flavor',
                                    'Smoothness','Overall Liking'] LOOP
        IF jsonb_typeof(p_sample->v_metric) IS DISTINCT FROM 'number'
           OR (p_sample->>v_metric)::numeric < 1
           OR (p_sample->>v_metric)::numeric > 9 THEN
            RAISE EXCEPTION 'sensory score "%" must be a JSON number in [1,9]', v_metric;
        END IF;
    END LOOP;

    -- Rounds 1-4 fold into tester_name as ' R1'..' R4' (mirror combineTesterRound
    -- in TesterRound.h; v2.5.13 extended R1/R2 -> R1-R4). An EMPTY tester, or a
    -- round outside 1-4 (e.g. 'N/A'), appends nothing.
    IF v_tester <> '' AND p_round IN ('1','2','3','4') THEN
        v_tester := v_tester || ' R' || p_round;
    END IF;
    v_date := to_char((now() AT TIME ZONE p_office_tz)::date, 'YYYY-MM-DD');  -- office-TZ date

    -- Resolve-or-create, then lock the row (serializes concurrent appends).
    INSERT INTO sensory_sessions(session_name, tester_name, assessor_name, media,
                                 date, timestamp, json_data, updated_by)
    VALUES (btrim(p_session_name), v_tester, p_assessor, p_media, v_date,
            to_char(now(), 'YYYY-MM-DD"T"HH24:MI:SS'),
            '{"samples":[]}'::jsonb, 'web/' || v_tester)
    ON CONFLICT (session_name, tester_name, date) DO NOTHING;

    SELECT id INTO v_id FROM sensory_sessions
     WHERE session_name = btrim(p_session_name)
       AND tester_name  = v_tester
       AND date         = v_date
     FOR UPDATE;

    -- Defensive: the INSERT ... ON CONFLICT DO NOTHING + this SELECT always
    -- resolve a row on the normal path; fail loudly rather than silently no-op
    -- the tail-append UPDATE if resolution ever returns nothing.
    IF v_id IS NULL THEN
        RAISE EXCEPTION 'dve_append_sensory_sample: could not resolve or create session for (%, %, %)',
            btrim(p_session_name), v_tester, v_date;
    END IF;

    -- Idempotency: a retried POST with the same sample_uid is a no-op.
    IF v_uid IS NOT NULL THEN
        PERFORM 1 FROM sensory_sessions s,
                      jsonb_array_elements(coalesce(s.json_data->'samples','[]'::jsonb)) e
         WHERE s.id = v_id AND e->>'sample_uid' = v_uid;
        IF FOUND THEN RETURN v_id; END IF;
    END IF;

    -- Tail-append only (never reorder/remove existing elements).
    UPDATE sensory_sessions
       SET json_data = jsonb_set(json_data, '{samples}',
                                 coalesce(json_data->'samples','[]'::jsonb) || p_sample)
     WHERE id = v_id;   -- bump_version + notify_row_change fire automatically

    RETURN v_id;
END;
$$;
