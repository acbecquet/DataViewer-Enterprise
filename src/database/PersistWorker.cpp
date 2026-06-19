#include "PersistWorker.h"
#include "DatabaseOps.h"
#include "ConfigLoader.h"

#include <QSqlError>
#include <QUuid>
#include <QDebug>

namespace DVE {

PersistWorker::PersistWorker(DbConfig cfg, QObject* parent)
    : QObject(parent), m_cfg(std::move(cfg))
{
    m_connName = QStringLiteral("dve_persist_worker_%1")
                     .arg(QUuid::createUuid().toString(QUuid::WithoutBraces));
}

void PersistWorker::start() { openConnection(); }
void PersistWorker::stop()  { closeConnection(); }

void PersistWorker::enqueue(DVE::PersistJob job)
{
    m_queue.enqueue(std::move(job));
    processQueue();
}

void PersistWorker::drainAndSignal()
{
    processQueue();
    emit drainFinished();
}

void PersistWorker::processQueue()
{
    while (!m_queue.isEmpty()) {
        PersistJob job = m_queue.dequeue();
        if (!m_open) {
            // Connection never opened (NAS unreachable at startup). Report
            // OfflineReadOnly so the UI keeps the file dirty -- identical to the
            // synchronous offline path. Never block.
            emit persistFinished(job, DVE::WriteResult::OfflineReadOnly);
            continue;
        }
        QString err;
        // persistFileOnConnection touches ONLY m_pgDb + the job's writerUuid +
        // the job's UI-thread-captured online flag. It never reaches a
        // DatabaseManager. m_pgDb.isOpen() is the live-socket gate.
        DVE::WriteResult wr = DVE::persistFileOnConnection(
            m_pgDb, job.writerUuid, job.online, m_pgDb.isOpen(), job.snapshot, &err);
        if (wr != DVE::WriteResult::Success && !err.isEmpty())
            qWarning() << "[PersistWorker] persist failed:" << err;
        emit persistFinished(job, wr);
    }
}

void PersistWorker::openConnection()
{
    if (QSqlDatabase::contains(m_connName))
        QSqlDatabase::removeDatabase(m_connName);
    m_pgDb = QSqlDatabase::addDatabase(QStringLiteral("QPSQL"), m_connName);
    m_pgDb.setHostName(m_cfg.host);
    m_pgDb.setPort(m_cfg.port);
    m_pgDb.setDatabaseName(m_cfg.database);
    m_pgDb.setUserName(m_cfg.user);       // DbConfig::user (NOT username)
    m_pgDb.setPassword(m_cfg.password);
    // Match LiveSyncWorker: short connect timeout + shared options
    // (application_name + keepalives + tcp_user_timeout, v2.4.2 hardening).
    m_pgDb.setConnectOptions(QStringLiteral("connect_timeout=5;") + pgSharedConnectOptions());
    if (!m_pgDb.open()) {
        qWarning() << "[PersistWorker] PG open failed:" << m_pgDb.lastError().text();
        m_open = false;
        return;
    }
    applyPgSessionSettings(m_pgDb);       // statement_timeout etc. (best-effort)
    m_open = true;
    qInfo() << "[PersistWorker] PG connection open";
}

void PersistWorker::closeConnection()
{
    if (m_pgDb.isOpen()) m_pgDb.close();
    m_pgDb = QSqlDatabase();
    if (QSqlDatabase::contains(m_connName))
        QSqlDatabase::removeDatabase(m_connName);
    m_open = false;
}

} // namespace DVE
