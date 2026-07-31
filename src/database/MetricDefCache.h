// MetricDefCache.h -- (kind, key) -> metric_defs.id resolution, plus the
// value_type <-> (value_num, value_text) storage contract, for the v3
// long-format tables. TPM template v3, Phase 3c.
//
// Plan:  docs/superpowers/plans/2026-07-31-tpm-v3-phase3c-open-metrics-plan.md
// Index: docs/superpowers/plans/2026-07-30-tpm-v3-phase3-INDEX.md
//
// WHY THIS EXISTS
//   Every write of DataRow::extra / SampleResult::extra into measurements /
//   sample_headers needs a metric_defs.id. The 35 keys with a wide column are
//   pre-seeded by deploy/postgres/migrations/2026-07-31-v3-long-format.sql, and
//   the whole compiled MetricRegistry is upserted by
//   DatabaseManager::ensureSchema. A CUSTOM column - `coil_temp` from the
//   manifest demo workbook, say - is in NEITHER, so it has to be registered on
//   first save or its data has nowhere to go.
//
// WHAT IT MUST NEVER DO
//   Update an existing metric_defs row. ensureSchema owns the registry-sourced
//   fields, metric_defs carries bump_version() and is deliberately NOT in its
//   no-op suppression list (hazard H19), and
//   tst_databasemanager::v3LongFormat_metricDefsSeedIsNotRewritten asserts the
//   35 seeded rows are still at version 1 with updated_by='migration'. So
//   ensureMetric() is INSERT ... ON CONFLICT DO NOTHING RETURNING id followed by
//   a re-select on conflict - which also makes two clients registering the same
//   new key converge on one row.
#pragma once

#include <QHash>
#include <QSqlDatabase>
#include <QString>
#include <QVariant>

namespace DVE {

class MetricDefCache {
public:
    // `who` lands in updated_by on an auto-registered row - the writer identity
    // that first saw the key. Registry rows keep 'registry', seeded ones
    // 'migration'.
    MetricDefCache(const QSqlDatabase& db, const QString& who);

    // Reads the whole (kind, key) -> id / value_type map in ONE query, behind a
    // probe for ALL THREE long-format tables - metric_defs, measurements and
    // sample_headers. Returns false when they are not all reachable, in which
    // case ensureMetric() refuses to resolve anything and the caller is expected
    // to skip the open-metric work entirely rather than fail the save. Probing
    // only metric_defs would let a half-created schema through and the save
    // would then die on the first `measurements` statement inside the
    // transaction, which is the outcome this exists to prevent.
    //
    // CALL THIS BEFORE OPENING THE SAVE TRANSACTION: a failed statement poisons
    // a Postgres transaction, so probing from inside one would turn "this
    // database has no long tables" into "this file cannot be saved at all".
    bool load(QString* outError = nullptr);
    bool isAvailable() const { return m_available; }

    // Resolves (kind, key), registering the key if it is absent. Returns -1 on
    // failure (with *outError set) and never throws. `valueHint` is used ONLY
    // when a new row is created; for an existing key the stored value_type wins,
    // because that is what the read path will route on.
    //
    // An auto-registered row gets display_name = the key itself and
    // role = 'measured'. That is DELIBERATE, not a placeholder to be fixed
    // opportunistically: the human-facing display name lives in the workbook's
    // `_dve_schema` manifest, which is not plumbed down to the database layer,
    // and custom columns are not displayed anywhere until Phase 4. Enriching
    // metric_defs from the manifest is a Phase 4 task with its own plumbing, not
    // a reason to thread a schema through the save path now.
    qint64 ensureMetric(const QString& kind, const QString& key,
                        const QVariant& valueHint, QString* outError = nullptr);

    // metric_defs.value_type for a resolved id; empty for an unknown id.
    QString valueType(qint64 id) const;

    // Manifest.cpp kTypeNames spelling ("number"/"text"/"bool"/"mixed"/
    // "numberlist"/"image") for a QVariant that has no metric_defs row yet.
    //
    // The discrimination ORDER mirrors the recovery JSON envelope's
    // extraValueFromJson (src/pipeline/ReportDataJson.cpp:124-140): bytes, then
    // list, then integral/floating, then everything else as text. The two are
    // the only serializations DataRow::extra has, and they must agree about what
    // a value IS. bool is the one deliberate divergence - the envelope has no
    // boolean case and coerces it to its string form, while metric_defs has a
    // real 'bool' type, so this returns it.
    static QString inferValueType(const QVariant& v);

    // ---- the storage contract, both directions ----------------------------
    // The writer's half and the reader's half live in ONE place so they cannot
    // drift. Exactly one of outNum / outText is populated; never both, and
    // (sparse rule, index D2) never neither - a value that has nothing to say is
    // not written at all, which is the caller's job to decide before calling.
    //
    //   number      -> value_num
    //   mixed       -> value_num when numeric, else value_text
    //   bool        -> value_num, 0 or 1
    //   image       -> value_text, base64 (same encoding as the envelope's "b")
    //   numberlist  -> value_text, compact JSON array (the envelope's "a")
    //   anything    -> value_text, QVariant::toString()
    //
    // THE COERCION RULE, and it governs EVERY typed slot, not just the numeric
    // one: a value is written into a slot only when it can honestly occupy it,
    // and otherwise falls through to value_text. Never guessed at. Once a
    // guessed value is stored it is indistinguishable from a measured one, and
    // after one save the original is gone - stable at the wrong value, which is
    // strictly worse than a value the reader has to interpret as text.
    //
    //   bool        real bools and numerics coerce; so do the Y / YES / N / NO
    //               spellings LegacyAdapter::normaliseBCL accepts. Anything else
    //               is TEXT. QVariant::toBool() is emphatically NOT the test -
    //               it is true for every string except "", "0" and "false", so
    //               it turns a did_burn cell reading "no" into TRUE.
    //   number      toDouble() must succeed. Round-trips an integral extra as a
    //   / mixed     double: metric_defs has no integer type, and the registry's
    //               vocabulary is the authority on what a metric is.
    //   image       bytes only.
    //   numberlist  a genuine sequence whose every element is a number.
    //
    // The text fallback itself keeps a lossless rendering (JSON / base64) for
    // the container types, whose QVariant::toString() is empty or mangled -
    // demoting a value to a string is the rule; destroying it is not.
    static void encodeValue(const QString& valueType, const QVariant& v,
                            QVariant& outNum, QVariant& outText);

    // Inverse of encodeValue: whatever encodeValue put in value_text comes back
    // out, including a string it fell back to under a typed key (so an `image`
    // holding "no photo" is NOT base64-decoded into garbage). Returns an invalid
    // QVariant when both columns are NULL (which encodeValue never produces).
    static QVariant decodeValue(const QString& valueType,
                                const QVariant& num, const QVariant& text);

private:
    static QString cacheKey(const QString& kind, const QString& key);

    QSqlDatabase           m_db;
    QString                m_who;
    QHash<QString, qint64> m_ids;     // "kind|key" -> metric_defs.id
    QHash<qint64, QString> m_types;   // metric_defs.id -> value_type
    bool                   m_available = false;
};

} // namespace DVE
