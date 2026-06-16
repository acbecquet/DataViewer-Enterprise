#include "PostgresConnection.h"

#include <QSqlError>
#include <QSqlQuery>
#include <QUuid>
#include <QElapsedTimer>
#include <QThread>
#include <QRandomGenerator>

namespace DVE {

PostgresConnection::PostgresConnection(QObject* parent) : QObject(parent) {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_queryName  = "dve_query_"  + tag;
    m_listenName = "dve_listen_" + tag;
}

PostgresConnection::~PostgresConnection() {
    close();
}

static bool openOne(QSqlDatabase& db, const QString& name, const DbConfig& cfg,
                    QString& err) {
    db = QSqlDatabase::addDatabase("QPSQL", name);
    db.setHostName(cfg.host);
    db.setPort(cfg.port);
    db.setDatabaseName(cfg.database);
    db.setUserName(cfg.user);
    db.setPassword(cfg.password);
    // connect_timeout bounds the INITIAL connect; the shared options add
    // application_name + keepalives + tcp_user_timeout (v2.4.2 Tier-1).
    db.setConnectOptions(QStringLiteral("connect_timeout=3;") + pgSharedConnectOptions());
    if (!db.open()) {
        err = db.lastError().text();
        db = QSqlDatabase();          // drop ref before removeDatabase
        QSqlDatabase::removeDatabase(name);
        return false;
    }
    applyPgSessionSettings(db);       // statement_timeout (best-effort)
    return true;
}

bool PostgresConnection::open(const DbConfig& cfg) {
    if (m_open) close();

    if (!openOne(m_queryDb, m_queryName, cfg, m_lastError)) {
        return false;
    }
    if (!openOne(m_listenDb, m_listenName, cfg, m_lastError)) {
        m_queryDb.close();
        m_queryDb = QSqlDatabase();   // drop ref before removeDatabase
        QSqlDatabase::removeDatabase(m_queryName);
        return false;
    }
    m_open = true;
    m_lastError.clear();
    return true;
}

void PostgresConnection::close() {
    if (m_open) {
        m_queryDb.close();
        m_listenDb.close();
    }
    // Clear the member QSqlDatabase handles BEFORE removeDatabase so the
    // reference-counted database isn't "still in use" when we drop it.
    // Otherwise Qt emits a QSqlDatabasePrivate "connection still in use"
    // warning even though the connection is functionally closed.
    m_queryDb  = QSqlDatabase();
    m_listenDb = QSqlDatabase();
    if (QSqlDatabase::contains(m_queryName))  QSqlDatabase::removeDatabase(m_queryName);
    if (QSqlDatabase::contains(m_listenName)) QSqlDatabase::removeDatabase(m_listenName);
    m_open = false;
}

bool PostgresConnection::ping() {
    if (!m_open) return false;
    QSqlQuery q(m_queryDb);
    return q.exec("SELECT 1");
}

bool PostgresConnection::pingListen() {
    // Mirror ping() but on the LISTEN socket. Same m_open guard; a failed
    // exec("SELECT 1") here means the NOTIFY connection is dead even if the
    // query socket is still answering (the half-open / GFW case).
    if (!m_open) return false;
    QSqlQuery q(m_listenDb);
    return q.exec("SELECT 1");
}

bool PostgresConnection::tryOpenWithRetry(const DbConfig& cfg, int totalTimeoutMs) {
    QElapsedTimer t;
    t.start();
    int delayMs = 250;
    while (true) {
        if (open(cfg)) return true;
        if (t.elapsed() >= totalTimeoutMs) return false;
        const int jitter = QRandomGenerator::global()->bounded(5000);
        const int sleep  = qMin(delayMs + jitter, totalTimeoutMs - int(t.elapsed()));
        if (sleep <= 0) return false;
        QThread::msleep(sleep);
        delayMs = qMin(delayMs * 2, 5000);
    }
}

} // namespace DVE
