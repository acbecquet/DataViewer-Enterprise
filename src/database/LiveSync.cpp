#include "LiveSync.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"
#include "NotificationListener.h"

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

bool LiveSync::commitCell(const QString& table, qint64 rowId,
                          const QString& column, const QVariant& value)
{
    if (!m_conn || !m_conn->isOpen()) return false;
    if (!isLiveSyncTable(table)) {
        qWarning() << "LiveSync::commitCell unknown table" << table;
        return false;
    }
    if (column.startsWith(QLatin1String("json_path:"))) {
        const QString path = column.mid(QStringLiteral("json_path:").size());
        return runJsonPathUpdate(table, rowId, path, value);
    }
    return runScalarUpdate(table, rowId, column, value);
}

bool LiveSync::runScalarUpdate(const QString& table, qint64 rowId,
                               const QString& column, const QVariant& value)
{
    QSqlQuery q(m_conn->queryDb());

    // Tag the session so the AFTER trigger emits a column-aware payload.
    q.prepare("SELECT set_config('dve.live_column', ?, false), "
              "       set_config('dve.live_value',  ?, false)");
    q.addBindValue(column);
    q.addBindValue(value.toString());
    if (!q.exec()) {
        qWarning() << "LiveSync::commitCell set_config failed:" << q.lastError().text();
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
        return false;
    }
    return q.numRowsAffected() > 0;
}

bool LiveSync::runJsonPathUpdate(const QString& table, qint64 rowId,
                                 const QString& jsonPath, const QVariant& value)
{
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

    q.prepare("SELECT set_config('dve.live_column', ?, false), "
              "       set_config('dve.live_value',  ?, false)");
    q.addBindValue(QStringLiteral("json_path:") + jsonPath);
    q.addBindValue(value.toString());
    if (!q.exec()) return false;

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
        return false;
    }
    return q.numRowsAffected() > 0;
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
