#include "LiveSyncWorker.h"

#include <QSqlQuery>
#include <QSqlError>
#include <QDebug>
#include <QUuid>
#include <QRegularExpression>
#include <QStringList>
#include <QThread>

namespace DVE {

LiveSyncWorker::LiveSyncWorker(const DbConfig& cfg, QObject* parent)
    : QObject(parent), m_cfg(cfg)
{
    // Unique connection name per worker instance so multiple LiveSyncs (in
    // tests, or app + tests) don't collide on QSqlDatabase's global registry.
    m_connName = QStringLiteral("dve_livesync_worker_%1")
                     .arg(QUuid::createUuid().toString(QUuid::WithoutBraces));
}

LiveSyncWorker::~LiveSyncWorker()
{
    if (m_open) stop();
}

void LiveSyncWorker::start()
{
    if (m_open) return;
    m_open = openConnection();
    if (!m_open) {
        qWarning() << "LiveSyncWorker::start failed:" << m_db.lastError().text();
    }
}

void LiveSyncWorker::stop()
{
    if (m_db.isOpen()) m_db.close();
    if (QSqlDatabase::contains(m_connName)) {
        m_db = QSqlDatabase();
        QSqlDatabase::removeDatabase(m_connName);
    }
    m_open = false;
}

bool LiveSyncWorker::openConnection()
{
    if (QSqlDatabase::contains(m_connName)) {
        QSqlDatabase::removeDatabase(m_connName);
    }
    m_db = QSqlDatabase::addDatabase(QStringLiteral("QPSQL"), m_connName);
    m_db.setHostName(m_cfg.host);
    m_db.setPort(m_cfg.port);
    m_db.setDatabaseName(m_cfg.database);
    m_db.setUserName(m_cfg.user);
    m_db.setPassword(m_cfg.password);
    // Short connect timeout so a slow NAS doesn't hang the worker startup.
    m_db.setConnectOptions(QStringLiteral("connect_timeout=5"));
    return m_db.open();
}

QString LiveSyncWorker::parsePathToPgArray(const QString& jsonPath) const
{
    // "samples[2].scores.smoothness" -> "{samples,2,scores,smoothness}"
    QStringList parts;
    static const QRegularExpression re(QStringLiteral(R"(([^.\[\]]+)|\[(\d+)\])"));
    auto it = re.globalMatch(jsonPath);
    while (it.hasNext()) {
        const auto m = it.next();
        parts << (m.captured(1).isEmpty() ? m.captured(2) : m.captured(1));
    }
    return QStringLiteral("{%1}").arg(parts.join(QLatin1Char(',')));
}

void LiveSyncWorker::commitScalar(QString table, qint64 rowId,
                                  QString column, QVariant value, QString uuid,
                                  qint64 expectedVersion)
{
    if (!m_open || !m_db.isOpen()) {
        emit commitFailed(table, rowId, column, value);
        return;
    }
    QSqlQuery q(m_db);
    q.prepare(QStringLiteral("SELECT dve_commit_cell(?, ?, ?, ?, ?, ?)"));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(value.toString());
    q.addBindValue(uuid);
    // expectedVersion < 0 → bind NULL so the stored function takes its
    // pre-v2.0.2 no-OCC path (DEFAULT NULL). Otherwise we send the version
    // for the server-side AND version = $6 check.
    if (expectedVersion < 0) {
        q.addBindValue(QVariant(QMetaType(QMetaType::Int)));
    } else {
        q.addBindValue(static_cast<int>(expectedVersion));
    }
    if (!q.exec()) {
        qWarning() << "LiveSyncWorker::commitScalar failed:" << q.lastError().text()
                   << "for" << table << column << rowId;
        emit commitFailed(table, rowId, column, value);
        return;
    }
    // The stored function returns BOOLEAN. FALSE means the WHERE matched
    // zero rows: either expectedVersion didn't match OR the row was deleted.
    // The local edit did NOT land — bubble it up as a conflict so LiveSync
    // does NOT replay-enqueue a stale value.
    if (!q.next() || !q.value(0).toBool()) {
        emit commitConflict(table, rowId, column, value, expectedVersion);
    }
}

void LiveSyncWorker::commitJson(QString table, qint64 rowId,
                                QString jsonPath, QVariant value, QString uuid,
                                qint64 expectedVersion)
{
    if (!m_open || !m_db.isOpen()) {
        emit commitFailed(table, rowId,
                          QStringLiteral("json_path:") + jsonPath, value);
        return;
    }
    const QString pgArr = parsePathToPgArray(jsonPath);
    QSqlQuery q(m_db);
    q.prepare(QStringLiteral(
        "SELECT dve_commit_cell_json(?, ?, ?, ?::text[], ?, ?, ?)"));
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(jsonPath);
    q.addBindValue(pgArr);
    q.addBindValue(value.toString());
    q.addBindValue(uuid);
    if (expectedVersion < 0) {
        q.addBindValue(QVariant(QMetaType(QMetaType::Int)));
    } else {
        q.addBindValue(static_cast<int>(expectedVersion));
    }
    if (!q.exec()) {
        qWarning() << "LiveSyncWorker::commitJson failed:" << q.lastError().text()
                   << "for" << table << jsonPath << rowId;
        emit commitFailed(table, rowId,
                          QStringLiteral("json_path:") + jsonPath, value);
        return;
    }
    if (!q.next() || !q.value(0).toBool()) {
        emit commitConflict(table, rowId,
                            QStringLiteral("json_path:") + jsonPath,
                            value, expectedVersion);
    }
}

void LiveSyncWorker::focusCell(QString uuid, QString table, qint64 rowId,
                                QString column, QString userName, QString userColor)
{
    if (!m_open || !m_db.isOpen()) return;
    QSqlQuery q(m_db);
    q.prepare(QStringLiteral("SELECT dve_focus_cell(?, ?, ?, ?, ?, ?)"));
    q.addBindValue(uuid);
    q.addBindValue(table);
    q.addBindValue(rowId);
    q.addBindValue(column);
    q.addBindValue(userName);
    q.addBindValue(userColor);
    if (!q.exec()) {
        qWarning() << "LiveSyncWorker::focusCell failed:" << q.lastError().text();
    }
}

void LiveSyncWorker::blurCell(QString uuid)
{
    if (!m_open || !m_db.isOpen()) return;
    QSqlQuery q(m_db);
    q.prepare(QStringLiteral("DELETE FROM cell_focus WHERE user_uuid = ?::uuid"));
    q.addBindValue(uuid);
    if (!q.exec()) {
        qWarning() << "LiveSyncWorker::blurCell failed:" << q.lastError().text();
    }
}

void LiveSyncWorker::sync()
{
    // DATAVIEWER-4 drain barrier. By the time this slot runs, the worker
    // thread has already processed every commit slot queued before it (Qt
    // event queues are FIFO), so all preceding writes have hit Postgres.
    // No DB work needed — just signal completion back to the UI thread.
    emit synced();
}

} // namespace DVE
