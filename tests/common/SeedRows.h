#pragma once
// v3 Phase 3d test helper (plan Task 4): seed samples and data rows in
// whichever shape the connected database speaks.
//
// PRE-cutover (data_rows is a real table) it INSERTs wide rows exactly as the
// fixtures always did. POST-cutover (data_rows is the name-holder view) it
// INSERTs the narrow samples_core row plus per-key sample_headers /
// measurements, resolving ids and num-vs-text routing through metric_defs -
// the same single authority the production code uses. Fixtures switch to
// these helpers BEFORE the container flips, so the flip itself changes no
// test code (index D-3d-8).
//
// Values are passed as key -> value maps. Pre-cutover the key IS the wide
// column name; post-cutover it IS the metric_defs key. Those are identical
// for every standard column by construction (the migration's mapping is an
// identity map), which is what makes a single map serve both worlds.

#include <QtSql>
#include <QMap>
#include <QString>
#include <QVariant>

namespace DVE {
namespace TestSeed {

inline bool isPostCutover(QSqlDatabase& db)
{
    QSqlQuery q(db);
    if (!q.exec(QStringLiteral(
            "SELECT c.relkind FROM pg_class c WHERE c.oid = to_regclass('data_rows')")))
        return false;
    return q.next() && q.value(0).toString() == QLatin1String("v");
}

// One long-format row (sample_headers or measurements), routed num/text by
// the resolved metric_defs.value_type - resolved CLIENT-SIDE first, because a
// server-side CASE over a casted parameter can be constant-folded before the
// untaken branch is pruned and error out at plan time. Returns false on SQL
// failure or an unknown key - fixtures only seed standard keys, which are
// always present.
inline bool putLongValue(QSqlDatabase& db, const char* kind, qint64 sampleId,
                         const QString& key, const QVariant& value,
                         int sortOrder, const QString& updatedBy)
{
    const bool isMeasurement = qstrcmp(kind, "metric") == 0;

    QSqlQuery look(db);
    if (!look.prepare(QStringLiteral(
            "SELECT id, value_type = 'number' FROM metric_defs "
            "WHERE kind = ? AND key = ?")))
        return false;
    look.addBindValue(QLatin1String(kind));
    look.addBindValue(key);
    if (!look.exec() || !look.next()) {
        qWarning().noquote() << "TestSeed::putLongValue: no metric_defs row for"
                             << kind << key << look.lastError().text();
        return false;
    }
    const qint64 defId = look.value(0).toLongLong();
    const bool   isNum = look.value(1).toBool();

    QSqlQuery q(db);
    const QString sql = isMeasurement
        ? QStringLiteral(
              "INSERT INTO measurements (sample_id, metric_id, sort_order, "
              " value_num, value_text, updated_by) VALUES (?, ?, ?, ?, ?, ?) "
              "ON CONFLICT (sample_id, metric_id, sort_order) DO UPDATE "
              "SET value_num = EXCLUDED.value_num, value_text = EXCLUDED.value_text")
        : QStringLiteral(
              "INSERT INTO sample_headers (sample_id, field_id, "
              " value_num, value_text, updated_by) VALUES (?, ?, ?, ?, ?) "
              "ON CONFLICT (sample_id, field_id) DO UPDATE "
              "SET value_num = EXCLUDED.value_num, value_text = EXCLUDED.value_text");
    if (!q.prepare(sql)) return false;
    q.addBindValue(static_cast<qlonglong>(sampleId));
    q.addBindValue(static_cast<qlonglong>(defId));
    if (isMeasurement) q.addBindValue(sortOrder);
    q.addBindValue(isNum ? QVariant(value.toDouble())
                         : QVariant(QMetaType(QMetaType::Double)));
    q.addBindValue(isNum ? QVariant(QMetaType(QMetaType::QString))
                         : QVariant(value.toString()));
    q.addBindValue(updatedBy);
    if (!q.exec()) {
        qWarning().noquote() << "TestSeed::putLongValue failed:"
                             << q.lastError().text();
        return false;
    }
    return true;
}

// Seeds one sample. Returns the new sample id, or -1 on failure.
// `values` maps samples column names (== header metric keys) to values.
inline qint64 seedSample(QSqlDatabase& db, qint64 testId,
                         const QMap<QString, QVariant>& values,
                         int sortOrder = 0,
                         const QString& updatedBy = QStringLiteral("test-seed"))
{
    QSqlQuery q(db);
    if (!isPostCutover(db)) {
        QStringList cols{ QStringLiteral("test_id"), QStringLiteral("sort_order"),
                          QStringLiteral("updated_by") };
        QStringList marks{ QStringLiteral("?"), QStringLiteral("?"), QStringLiteral("?") };
        for (auto it = values.constBegin(); it != values.constEnd(); ++it) {
            cols << it.key();
            marks << QStringLiteral("?");
        }
        if (!q.prepare(QStringLiteral(
                "INSERT INTO samples (%1) VALUES (%2) RETURNING id")
                    .arg(cols.join(QLatin1String(", ")),
                         marks.join(QLatin1String(", ")))))
            return -1;
        q.addBindValue(static_cast<qlonglong>(testId));
        q.addBindValue(sortOrder);
        q.addBindValue(updatedBy);
        for (auto it = values.constBegin(); it != values.constEnd(); ++it)
            q.addBindValue(it.value());
        if (!q.exec() || !q.next()) {
            qWarning().noquote() << "TestSeed::seedSample (wide) failed:"
                                 << q.lastError().text();
            return -1;
        }
        return q.value(0).toLongLong();
    }

    if (!q.prepare(QStringLiteral(
            "INSERT INTO samples_core (test_id, sort_order, updated_by) "
            "VALUES (?, ?, ?) RETURNING id")))
        return -1;
    q.addBindValue(static_cast<qlonglong>(testId));
    q.addBindValue(sortOrder);
    q.addBindValue(updatedBy);
    if (!q.exec() || !q.next()) {
        qWarning().noquote() << "TestSeed::seedSample (core) failed:"
                             << q.lastError().text();
        return -1;
    }
    const qint64 sid = q.value(0).toLongLong();
    for (auto it = values.constBegin(); it != values.constEnd(); ++it) {
        if (!putLongValue(db, "header", sid, it.key(), it.value(),
                          /*sortOrder=*/0, updatedBy))
            return -1;
    }
    return sid;
}

// Seeds one data row. Pre-cutover returns the new data_rows.id; post-cutover
// there is no such identity (a row IS its measurements, keyed by
// (sample_id, sort_order)) and 0 is returned on success. -1 on failure.
//
// Post-cutover a row with no values at all cannot exist (sparse rule / H18),
// so `puffs` is seeded as 0 when the map does not provide it - which is what
// a real parse always writes anyway.
inline qint64 seedDataRow(QSqlDatabase& db, qint64 sampleId, int sortOrder,
                          const QMap<QString, QVariant>& values = {},
                          const QString& updatedBy = QStringLiteral("test-seed"))
{
    QSqlQuery q(db);
    if (!isPostCutover(db)) {
        QStringList cols{ QStringLiteral("sample_id"), QStringLiteral("sort_order"),
                          QStringLiteral("updated_by") };
        QStringList marks{ QStringLiteral("?"), QStringLiteral("?"), QStringLiteral("?") };
        for (auto it = values.constBegin(); it != values.constEnd(); ++it) {
            cols << it.key();
            marks << QStringLiteral("?");
        }
        if (!q.prepare(QStringLiteral(
                "INSERT INTO data_rows (%1) VALUES (%2) RETURNING id")
                    .arg(cols.join(QLatin1String(", ")),
                         marks.join(QLatin1String(", ")))))
            return -1;
        q.addBindValue(static_cast<qlonglong>(sampleId));
        q.addBindValue(sortOrder);
        q.addBindValue(updatedBy);
        for (auto it = values.constBegin(); it != values.constEnd(); ++it)
            q.addBindValue(it.value());
        if (!q.exec() || !q.next()) {
            qWarning().noquote() << "TestSeed::seedDataRow (wide) failed:"
                                 << q.lastError().text();
            return -1;
        }
        return q.value(0).toLongLong();
    }

    if (!values.contains(QStringLiteral("puffs"))
        && !putLongValue(db, "metric", sampleId, QStringLiteral("puffs"),
                         0.0, sortOrder, updatedBy))
        return -1;
    for (auto it = values.constBegin(); it != values.constEnd(); ++it) {
        if (!putLongValue(db, "metric", sampleId, it.key(), it.value(),
                          sortOrder, updatedBy))
            return -1;
    }
    return 0;
}

} // namespace TestSeed
} // namespace DVE
