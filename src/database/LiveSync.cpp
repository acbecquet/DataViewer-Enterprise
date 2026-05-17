#include "LiveSync.h"
#include "LiveSyncWorker.h"
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
#include <QTimer>
#include <QThread>

namespace DVE {

LiveSync::LiveSync(PostgresConnection* conn, IdentityManager* identity,
                   QObject* parent)
    : QObject(parent), m_conn(conn), m_identity(identity)
{
    m_throttleTimer = new QTimer(this);
    m_throttleTimer->setSingleShot(true);
    m_throttleTimer->setInterval(200);
    connect(m_throttleTimer, &QTimer::timeout, this, &LiveSync::onThrottleTick);
}

LiveSync::~LiveSync()
{
    if (m_workerThread) {
        if (m_worker)
            QMetaObject::invokeMethod(m_worker, "stop", Qt::BlockingQueuedConnection);
        m_workerThread->quit();
        m_workerThread->wait(2000);
    }
}

void LiveSync::setWorkerConfig(const DbConfig& cfg)
{
    if (m_workerThread) return;
    m_workerThread = new QThread(this);
    m_worker       = new LiveSyncWorker(cfg);
    m_worker->moveToThread(m_workerThread);
    connect(m_workerThread, &QThread::finished, m_worker, &QObject::deleteLater);
    connect(m_worker, &LiveSyncWorker::commitFailed,
            this,     &LiveSync::onWorkerCommitFailed);
    m_workerThread->start();
    QMetaObject::invokeMethod(m_worker, "start", Qt::QueuedConnection);
}

bool LiveSync::isLiveSyncTable(const QString& t)
{
    return t == QLatin1String("data_rows")
        || t == QLatin1String("samples")
        || t == QLatin1String("tests")
        || t == QLatin1String("files")
        || t == QLatin1String("sensory_sessions")
        || t == QLatin1String("detailed_sensory_sessions");
}

bool LiveSync::isLiveSyncColumn(const QString& table, const QString& column)
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
                          const QString& column, const QVariant& value,
                          bool allowQueue)
{
    if (!isLiveSyncTable(table)) {
        qWarning() << "LiveSync::commitCell unknown table" << table;
        return false;
    }
    const bool isJsonPath = column.startsWith(QLatin1String("json_path:"));
    const QString gateCol = isJsonPath ? QStringLiteral("json_data") : column;
    if (!isLiveSyncColumn(table, gateCol)) {
        qWarning() << "LiveSync::commitCell column not allowed for"
                   << table << ":" << column;
        return false;
    }

    if (!m_conn || !m_conn->isOpen()) {
        if (allowQueue && m_snapshot) {
            return m_snapshot->enqueueCellEdit(table, rowId, column, value);
        }
        return false;
    }

    m_pendingCommits.insert({table, rowId, column}, value);
    if (!m_throttleTimer->isActive()) m_throttleTimer->start();
    return true;
}

void LiveSync::onThrottleTick()
{
    if (m_pendingCommits.isEmpty()) return;
    const auto drained = m_pendingCommits;
    m_pendingCommits.clear();
    for (auto it = drained.constBegin(); it != drained.constEnd(); ++it) {
        dispatchCommit(it.key().table, it.key().rowId,
                       it.key().column, it.value());
    }
}

void LiveSync::dispatchCommit(const QString& table, qint64 rowId,
                              const QString& column, const QVariant& value)
{
    const bool isJsonPath = column.startsWith(QLatin1String("json_path:"));
    if (!m_conn || !m_conn->isOpen()) {
        if (m_snapshot) m_snapshot->enqueueCellEdit(table, rowId, column, value);
        return;
    }
    const QString uuid = m_identity
        ? m_identity->uuid().toString(QUuid::WithoutBraces)
        : QString();

    if (m_worker) {
        if (isJsonPath) {
            const QString path = column.mid(QStringLiteral("json_path:").size());
            QMetaObject::invokeMethod(m_worker, "commitJson", Qt::QueuedConnection,
                Q_ARG(QString, table), Q_ARG(qint64, rowId),
                Q_ARG(QString, path),  Q_ARG(QVariant, value),
                Q_ARG(QString, uuid));
        } else {
            QMetaObject::invokeMethod(m_worker, "commitScalar", Qt::QueuedConnection,
                Q_ARG(QString, table),  Q_ARG(qint64, rowId),
                Q_ARG(QString, column), Q_ARG(QVariant, value),
                Q_ARG(QString, uuid));
        }
        return;
    }

    if (isJsonPath) {
        runJsonPathUpdateSync(table, rowId,
            column.mid(QStringLiteral("json_path:").size()), value);
    } else {
        runScalarUpdateSync(table, rowId, column, value);
    }
}

void LiveSync::onWorkerCommitFailed(QString table, qint64 rowId,
                                    QString column, QVariant value)
{
    if (m_snapshot) {
        m_snapshot->enqueueCellEdit(table, rowId, column, value);
    }
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
            return this->commitCell(t, r, c, v, /*allowQueue=*/false);
        });
}

bool LiveSync::runScalarUpdateSync(const QString& table, qint64 rowId,
                                   const QString& column, const QVariant& value)
{
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("SELECT dve_commit_cell(?, ?, ?, ?, ?)"));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(value.toString());
    q.addBindValue(m_identity
        ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString());
    if (!q.exec()) {
        qWarning() << "LiveSync sync scalar failed:" << q.lastError().text();
        return false;
    }
    return q.next() && q.value(0).toBool();
}

bool LiveSync::runJsonPathUpdateSync(const QString& table, qint64 rowId,
                                     const QString& jsonPath, const QVariant& value)
{
    QStringList parts;
    static const QRegularExpression re(QStringLiteral(R"(([^.\[\]]+)|\[(\d+)\])"));
    auto it = re.globalMatch(jsonPath);
    while (it.hasNext()) {
        const auto m = it.next();
        parts << (m.captured(1).isEmpty() ? m.captured(2) : m.captured(1));
    }
    if (parts.isEmpty()) return false;
    const QString pgPath = QStringLiteral("{%1}").arg(parts.join(QLatin1Char(',')));

    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("SELECT dve_commit_cell_json(?, ?, ?, ?::text[], ?, ?)"));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(jsonPath);
    q.addBindValue(pgPath);
    q.addBindValue(value.toString());
    q.addBindValue(m_identity
        ? m_identity->uuid().toString(QUuid::WithoutBraces) : QString());
    if (!q.exec()) {
        qWarning() << "LiveSync sync jsonb failed:" << q.lastError().text();
        return false;
    }
    return q.next() && q.value(0).toBool();
}

bool LiveSync::focusCell(const QString& table, qint64 rowId, const QString& column)
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    const QString uuid = m_identity->uuid().toString(QUuid::WithoutBraces);
    if (m_worker) {
        QMetaObject::invokeMethod(m_worker, "focusCell", Qt::QueuedConnection,
            Q_ARG(QString, uuid),  Q_ARG(QString, table),
            Q_ARG(qint64,  rowId), Q_ARG(QString, column),
            Q_ARG(QString, m_identity->displayName()),
            Q_ARG(QString, m_identity->color()));
        return true;
    }
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)"));
    q.addBindValue(uuid);
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(m_identity->displayName());
    q.addBindValue(m_identity->color());
    return q.exec();
}

bool LiveSync::blurCell()
{
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;
    const QString uuid = m_identity->uuid().toString(QUuid::WithoutBraces);
    if (m_worker) {
        QMetaObject::invokeMethod(m_worker, "blurCell", Qt::QueuedConnection,
            Q_ARG(QString, uuid));
        return true;
    }
    QSqlQuery q(m_conn->queryDb());
    q.prepare(QStringLiteral("DELETE FROM cell_focus WHERE user_uuid = ?::uuid"));
    q.addBindValue(uuid);
    return q.exec();
}

void LiveSync::onRowChanged(const RowChange& c)
{
    if (c.column.isEmpty()) return;
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
