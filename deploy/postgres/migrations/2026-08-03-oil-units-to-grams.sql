-- ============================================================================
-- v3 unit normalization: oil quantities become GRAMS in the database
-- ============================================================================
--
-- OWNER RULING (2026-08-03): "everything in grams for oil ... make sure units
-- are consistent throughout."
--
-- WHAT WAS INCONSISTENT
-- ---------------------
-- Grams was already canonical almost everywhere:
--   * the Excel template's own column header is "OilCum(g)", and its formula
--     =IF(I5="","", (I$5*A$5)/1000) divides milligrams down to grams itself;
--   * MetricRegistry declares oil_consumed and initial_oil_mass with unit "g";
--   * samples.initial_oil_mass has always been stored in grams.
-- The sole dissenter was SheetProcessor::calculateMetrics, which scaled the
-- gram-valued weight difference UP by 1000 into milligrams. Every display site
-- then divided by 1000 again on the way out, so the app LOOKED right while
-- data_rows.oil_consumed and samples.total_oil_consumed held milligrams under
-- a column whose registry unit said grams - a latent 1000x error for anything
-- that trusted metric_defs.unit (Phase 4 UI, Phase 5 template builder, exports).
--
-- The C++ side now computes grams. This migration brings stored data in line.
--
-- WHY THIS RECOMPUTES INSTEAD OF DIVIDING BY 1000
-- -----------------------------------------------
-- A naive `SET oil_consumed = oil_consumed / 1000` is NOT safe here, because
-- between the code change and this migration a client can save a row that is
-- ALREADY in grams. Scaling would then silently divide that row a second time
-- (0.175 -> 0.000175) and no later migration could tell the two apart.
--
-- oil_consumed and total_oil_consumed are DERIVED (MetricDef Role::Derived):
-- both are functions of before_weight and after_weight, which are stored in
-- GRAMS and are untouched by any of this. So this migration recomputes them
-- from those weights instead. That is idempotent, immune to a half-converted
-- column, and reproduces exactly what the application would compute on its own
-- next load - the definition below mirrors SheetProcessor::calculateMetrics
-- statement for statement.
--
-- WHY IT IS A SEPARATE FILE FROM THE LONG-FORMAT MIGRATION
-- --------------------------------------------------------
-- dve_migrate_to_long_format() is a pure SHAPE transform, and the 3b gate that
-- guards it asserts data_rows_v is row-for-row VALUE-identical to data_rows.
-- Folding a unit change into it would break that parity gate - destroying the
-- referee that proves the long-format migration correct. This runs FIRST, as
-- its own step; the long-format migration then stays value-preserving.
--
-- WHEN TO RUN IT (v3.0.0 runbook, index D7/D8)
-- ---------------------------------------------
-- By hand, against the NAS, in the supervised v3.0.0 migration window, BEFORE
-- 2026-07-31-v3-long-format.sql's dve_migrate_to_long_format(). Nothing applies
-- it automatically: ensureSchema() delivers additive DDL only (index D6), and
-- docker-compose mounts only init.sql.
--
-- DO NOT APPLY IT EARLY. Clients still running v2.10.x compute milligrams and
-- would write them straight back into the converted columns. It is safe only
-- once every client is on the grams build.
--
-- Triggers are disabled for the duration: data_rows and samples both carry
-- notify_row_change, and a whole-table UPDATE would emit one NOTIFY per row -
-- the DV-23 storm class the v2.9 fix exists to prevent - plus a bump_version
-- churn on every row. This is a supervised offline window with no live clients,
-- so suppressing both is correct rather than merely convenient.
-- ============================================================================

BEGIN;

DO $mig$
DECLARE
    v_rows_updated  BIGINT := 0;
    v_samps_updated BIGINT := 0;
BEGIN
    -- Serialize against a second concurrent applier. Not needed for
    -- correctness (the recompute is idempotent) but it keeps the reported
    -- counts honest and matches the v242_legacy_score_normalize precedent.
    PERFORM pg_advisory_xact_lock(hashtext('dve_oil_units_grams'));

    ALTER TABLE data_rows DISABLE TRIGGER USER;
    ALTER TABLE samples   DISABLE TRIGGER USER;

    -- data_rows.oil_consumed
    -- Mirrors calculateMetrics: a row counts only when BOTH weights are > 0
    -- ("hasData"); oil is the running total over data-bearing rows within the
    -- sample; a row without data is 0; a cumulative total above 10 g is treated
    -- as corrupt and clamped to 0 (the same physical bound the old code spelled
    -- as 10000 mg).
    --
    -- Samples with no weight data at all are deliberately SKIPPED, not zeroed.
    -- On such a sample (an inferred-schema sheet with no weight columns) the
    -- app computes 0 anyway, so a stored non-zero value came from somewhere
    -- else - MigrationTool copying a legacy row verbatim - and overwriting it
    -- would be data loss, not normalization.
    WITH weighed_samples AS (
        SELECT DISTINCT sample_id
          FROM data_rows
         WHERE before_weight > 0 AND after_weight > 0
    ),
    running AS (
        SELECT
            d.id,
            d.before_weight > 0 AND d.after_weight > 0 AS has_data,
            SUM(CASE WHEN d.before_weight > 0 AND d.after_weight > 0
                     THEN d.before_weight - d.after_weight
                     ELSE 0 END)
                OVER (PARTITION BY d.sample_id
                      ORDER BY d.sort_order, d.id
                      ROWS UNBOUNDED PRECEDING) AS cum_g
          FROM data_rows d
          JOIN weighed_samples w ON w.sample_id = d.sample_id
    ),
    recomputed AS (
        SELECT id,
               CASE WHEN NOT has_data THEN 0.0
                    WHEN cum_g > 10.0 THEN 0.0
                    ELSE cum_g
               END AS oil_g
          FROM running
    )
    UPDATE data_rows d
       SET oil_consumed = r.oil_g
      FROM recomputed r
     WHERE d.id = r.id
       AND d.oil_consumed IS DISTINCT FROM r.oil_g;
    GET DIAGNOSTICS v_rows_updated = ROW_COUNT;

    -- samples.total_oil_consumed
    -- Sum over data-bearing rows only, and NO clamp - calculateMetrics does not
    -- clamp the sample total, only the per-row running value.
    WITH totals AS (
        SELECT d.sample_id,
               SUM(d.before_weight - d.after_weight) AS total_g
          FROM data_rows d
         WHERE d.before_weight > 0 AND d.after_weight > 0
         GROUP BY d.sample_id
    )
    UPDATE samples s
       SET total_oil_consumed = t.total_g
      FROM totals t
     WHERE s.id = t.sample_id
       AND s.total_oil_consumed IS DISTINCT FROM t.total_g;
    GET DIAGNOSTICS v_samps_updated = ROW_COUNT;

    -- samples.efficiency_percent
    -- efficiency = total_oil_consumed / initial_oil_mass * 100, clamped 0-100.
    -- Both operands are now grams, so the old * 1000 that cancelled the
    -- milligram scaling is gone from the C++ and must not reappear here.
    --
    -- Restricted to the SAME weighed-sample set as the two updates above. A
    -- sample we declined to normalize still holds an un-normalized total, and
    -- recomputing efficiency from that would change a value on a row this
    -- migration promised to leave alone - turning a stale 42 into a confident
    -- "100%" instead of the 0 the application itself computes for a sample with
    -- no weight data. Skipping means skipping the whole derived chain.
    UPDATE samples s
       SET efficiency_percent =
               LEAST(100.0, GREATEST(0.0,
                   (s.total_oil_consumed / s.initial_oil_mass) * 100.0))
     WHERE s.initial_oil_mass > 0
       AND EXISTS (SELECT 1 FROM data_rows d
                    WHERE d.sample_id = s.id
                      AND d.before_weight > 0 AND d.after_weight > 0)
       AND s.efficiency_percent IS DISTINCT FROM
               LEAST(100.0, GREATEST(0.0,
                   (s.total_oil_consumed / s.initial_oil_mass) * 100.0));

    ALTER TABLE data_rows ENABLE TRIGGER USER;
    ALTER TABLE samples   ENABLE TRIGGER USER;

    -- Recorded for auditability only. This migration is idempotent, so unlike
    -- v242_legacy_score_normalize the key is NOT a correctness gate - re-running
    -- the file is harmless and will report 0 rows updated.
    INSERT INTO schema_meta (key, value)
         VALUES ('v300_oil_units_grams', now()::text)
    ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value;

    RAISE NOTICE 'oil units -> grams: % data_rows, % samples updated',
                 v_rows_updated, v_samps_updated;
END
$mig$;

COMMIT;
