#include "LiveSync.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "NotificationListener.h"
#include "OfflineSnapshot.h"

#include <QHash>
#include <QSet>
#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QStringList>
#include <QRegularExpression>
#include <QUuid>

namespace DVE {

LiveSync::LiveSync(PostgresConnection* conn, IdentityManager* identity,
                   QObject* parent)
    : QObject(parent), m_conn(conn), m_identity(identity) {}

// Single source of truth for the set of tables that v2.0.1 live-syncs.
static bool isLiveSyncTable(const QString& t)
{
    return t == QLatin1String("data_rows")
        || t == QLatin1String("samples")
        || t == QLatin1String("tests")
        || t == QLatin1String("files")
        || t == QLatin1String("sensory_sessions")
        || t == QLatin1String("detailed_sensory_sessions");
}

// Per-table allowlist of columns that can be live-sync'd. Used as the
// final gate before interpolating the column name into the UPDATE
// statement -- without this, a malicious or buggy caller passing a
// column like "x = 0; DROP TABLE files; --" would execute arbitrary
// SQL. The table itself is already gated by isLiveSyncTable.
//
// Lists are derived from deploy/postgres/init.sql. Columns excluded on
// purpose: id, version, updated_at, updated_by, sort_order, computed
// aggregates owned by the recalc pipeline (samples.average_tpm,
// samples.stddev_tpm, samples.avg_power_density, samples.efficiency_percent,
// samples.total_oil_consumed, samples.total_puffs, samples.normalized_tpm,
// tests.overall_avg_tpm, tests.overall_stddev_tpm, tests.is_raw_table,
// files.loaded_at, files.sheet_count, files.sample_count), and
// per-resource JSONB layout columns (sensory_sessions.layout_json,
// sensory_sessions.timestamp, detailed_sensory_sessions.timestamp).
//
// For JSONB tables (sensory_sessions, detailed_sensory_sessions) the
// only column we touch is `json_data`; the actual field path lives in
// the "json_path:..." prefix and is parsed separately. Listing
// json_data here lets runJsonPathUpdate share the validation path.
static bool isLiveSyncColumn(const QString& table, const QString& column)
{
    static const QHash<QString, QSet<QString>> kAllowed = {
        { QStringLiteral("data_rows"), {
            QStringLiteral("puffs"), QStringLiteral("before_weight"),
            QStringLiteral("after_weight"), QStringLiteral("draw_pressure"),
            QStringLiteral("resistance"), QStringLiteral("smell"),
            QStringLiteral("clog"), QStringLiteral("notes"),
            QStringLiteral("tpm"), QStringLiteral("tpm_power_density"),
            QStringLiteral("variation_tpm"), QStringLiteral("oil_consumed")
        }},
        { QStringLiteral("samples"), {
            QStringLiteral("sample_name"), QStringLiteral("sample_id"),
            QStringLiteral("date"), QStringLiteral("tester"),
            QStringLiteral("media"), QStringLiteral("viscosity"),
            QStringLiteral("resistance"), QStringLiteral("voltage"),
            QStringLiteral("power"), QStringLiteral("heating_technology"),
            QStringLiteral("puffing_regime"), QStringLiteral("initial_oil_mass"),
            QStringLiteral("burn_status"), QStringLiteral("clog_status"),
            QStringLiteral("leak_status")
        }},
        { QStringLiteral("tests"), {
            QStringLiteral("sheet_name"), QStringLiteral("template_version")
        }},
        { QStringLiteral("files"), {
            QStringLiteral("file_path"), QStringLiteral("file_name"),
            QStringLiteral("template_version")
        }},
        { QStringLiteral("sensory_sessions"), {
            QStringLiteral("session_name"), QStringLiteral("tester_name"),
            QStringLiteral("assessor_name"), QStringLiteral("media"),
            QStringLiteral("puff_length"), QStringLiteral("date"),
            QStringLiteral("json_data")
        }},
        { QStringLiteral("detailed_sensory_sessions"), {
            QStringLiteral("session_name"), QStringLiteral("tester_name"),
            QStringLiteral("assessor_name"), QStringLiteral("media"),
            QStringLiteral("date"), QStringLiteral("json_data")
        }}
    };
    auto it = kAllowed.constFind(table);
    if (it == kAllowed.constEnd()) return false;
    return it.value().contains(column);
}

bool LiveSync::commitCell(const QString& table, qint64 rowId,
                          const QString& column, const QVariant& value)
{
    if (!isLiveSyncTable(table)) {
        qWarning() << "LiveSync::commitCell unknown table" << table;
        return false;
    }
    // Offline fallback: queue the edit for replay on reconnect. The
    // column gate (isLiveSyncColumn) is enforced at the per-path layer
    // below; on replay each entry routes back through commitCell so
    // the same allowlist applies before the actual UPDATE is run.
    if (!m_conn || !m_conn->isOpen()) {
        if (m_snapshot) {
            return m_snapshot->enqueueCellEdit(table, rowId, column, value);
        }
        return false;
    }
    if (column.startsWith(QLatin1String("json_path:"))) {
        const QString path = column.mid(QStringLiteral("json_path:").size());
        return runJsonPathUpdate(table, rowId, path, value);
    }
    return runScalarUpdate(table, rowId, column, value);
}

void LiveSync::setOfflineSnapshot(OfflineSnapshot* snap)
{
    m_snapshot = snap;
}

int LiveSync::flushPending()
{
    if (!m_snapshot) return 0;
    if (!m_conn || !m_conn->isOpen()) return 0;
    return m_snapshot->drainPendingEdits(
        [this](const QString& t, qint64 r, const QString& c, const QVariant& v) {
            return this->commitCell(t, r, c, v);
        });
}

bool LiveSync::runScalarUpdate(const QString& table, qint64 rowId,
                               const QString& column, const QVariant& value)
{
    if (!isLiveSyncColumn(table, column)) {
        qWarning() << "LiveSync::commitCell column not allowed for"
                   << table << ":" << column;
        return false;
    }

    QSqlQuery q(m_conn->queryDb());
    if (!q.exec(QStringLiteral("BEGIN"))) {
        qWarning() << "LiveSync::commitCell BEGIN failed:" << q.lastError().text();
        return false;
    }

    // Tag the session so the AFTER trigger emits a column-aware payload.
    // set_config(..., true) makes the GUC transaction-local so it can't
    // leak into a subsequent non-LiveSync UPDATE on the same connection.
    q.prepare("SELECT set_config('dve.live_column', ?, true), "
              "       set_config('dve.live_value',  ?, true)");
    q.addBindValue(column);
    q.addBindValue(value.toString());
    if (!q.exec()) {
        qWarning() << "LiveSync::commitCell set_config failed:" << q.lastError().text();
        q.exec(QStringLiteral("ROLLBACK"));
        return false;
    }

    const QString uuid = m_identity ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString();
    const QString sql = QStringLiteral(
        "UPDATE %1 SET %2 = ?, version = version + 1, "
        "updated_at = now(), updated_by = ? WHERE id = ?")
        .arg(table, column);
    q.prepare(sql);
    q.addBindValue(value);
    q.addBindValue(uuid);
    q.addBindValue(rowId);
    if (!q.exec()) {
        qWarning() << "LiveSync::commitCell UPDATE failed:" << q.lastError().text();
        q.exec(QStringLiteral("ROLLBACK"));
        return false;
    }
    const int affected = q.numRowsAffected();
    if (!q.exec(QStringLiteral("COMMIT"))) {
        qWarning() << "LiveSync::commitCell COMMIT failed:" << q.lastError().text();
        return false;
    }
    return affected > 0;
}

bool LiveSync::runJsonPathUpdate(const QString& table, qint64 rowId,
                                 const QString& jsonPath, const QVariant& value)
{
    if (!isLiveSyncColumn(table, QStringLiteral("json_data"))) {
        qWarning() << "LiveSync::commitCell json_path not allowed for"
                   << table;
        return false;
    }

    // Parse "samples[2].scores.smoothness" into a Postgres text-array
    // path: '{samples,2,scores,smoothness}'. Brackets become array
    // index entries, dots become segment separators.
    QStringList parts;
    static const QRegularExpression re(QStringLiteral(
        R"(([^.\[\]]+)|\[(\d+)\])"));
    auto it = re.globalMatch(jsonPath);
    while (it.hasNext()) {
        const auto m = it.next();
        parts << (m.captured(1).isEmpty() ? m.captured(2) : m.captured(1));
    }
    if (parts.isEmpty()) {
        qWarning() << "LiveSync json path parse failed:" << jsonPath;
        return false;
    }
    const QString pgPath = QStringLiteral("{%1}").arg(parts.join(QLatin1Char(',')));

    QSqlQuery q(m_conn->queryDb());
    if (!q.exec(QStringLiteral("BEGIN"))) {
        qWarning() << "LiveSync jsonb BEGIN failed:" << q.lastError().text();
        return false;
    }

    q.prepare("SELECT set_config('dve.live_column', ?, true), "
              "       set_config('dve.live_value',  ?, true)");
    q.addBindValue(QStringLiteral("json_path:") + jsonPath);
    q.addBindValue(value.toString());
    if (!q.exec()) {
        qWarning() << "LiveSync jsonb set_config failed:" << q.lastError().text();
        q.exec(QStringLiteral("ROLLBACK"));
        return false;
    }

    const QString uuid = m_identity ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString();
    const QString sql = QStringLiteral(
        "UPDATE %1 SET json_data = jsonb_set(json_data, ?::text[], "
        "to_jsonb(?::text)::jsonb, true), "
        "version = version + 1, updated_at = now(), updated_by = ? "
        "WHERE id = ?").arg(table);
    q.prepare(sql);
    q.addBindValue(pgPath);
    q.addBindValue(value.toString());
    q.addBindValue(uuid);
    q.addBindValue(rowId);
    if (!q.exec()) {
        qWarning() << "LiveSync jsonb_set failed:" << q.lastError().text();
        q.exec(QStringLiteral("ROLLBACK"));
        return false;
    }
    const int affected = q.numRowsAffected();
    if (!q.exec(QStringLiteral("COMMIT"))) {
        qWarning() << "LiveSync jsonb COMMIT failed:" << q.lastError().text();
        return false;
    }
    return affected > 0;
}

bool LiveSync::focusCell(const QString& table, qint64 rowId, const QString& column)
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    QSqlQuery q(m_conn->queryDb());

    // One focus per user. Clear prior, then insert. Two statements are
    // safer than ON CONFLICT against a 4-column PK.
    q.prepare("DELETE FROM cell_focus WHERE user_uuid = ?::uuid");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    if (!q.exec()) return false;

    q.prepare(
        "INSERT INTO cell_focus(user_uuid, table_name, row_id, "
        "column_name, user_name, user_color) "
        "VALUES(?::uuid, ?, ?, ?, ?, ?)");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(m_identity->displayName());
    q.addBindValue(m_identity->color());
    if (!q.exec()) {
        qWarning() << "LiveSync::focusCell INSERT failed:" << q.lastError().text();
        return false;
    }
    return true;
}

bool LiveSync::blurCell()
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("DELETE FROM cell_focus WHERE user_uuid = ?::uuid");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    return q.exec();
}

void LiveSync::onRowChanged(const RowChange& c)
{
    if (c.column.isEmpty()) return;       // multi-column UPDATE - skip
    emit cellChanged(c.table, c.id, c.column, c.newValue);
}

void LiveSync::onCellFocusChanged(const CellFocusChange& f)
{
    if (f.op == QLatin1String("DELETE"))
        emit cellBlurred(f.tableName, f.rowId, f.columnName);
    else
        emit cellFocused(f.tableName, f.rowId, f.columnName,
                         f.userName, f.userColor);
}

} // namespace DVE
