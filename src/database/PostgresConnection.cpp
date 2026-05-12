#include "PostgresConnection.h"

#include <QSqlError>
#include <QSqlQuery>
#include <QUuid>

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
    db.setConnectOptions("connect_timeout=3");
    if (!db.open()) {
        err = db.lastError().text();
        QSqlDatabase::removeDatabase(name);
        return false;
    }
    return true;
}

bool PostgresConnection::open(const DbConfig& cfg) {
    if (m_open) close();

    if (!openOne(m_queryDb, m_queryName, cfg, m_lastError)) {
        return false;
    }
    if (!openOne(m_listenDb, m_listenName, cfg, m_lastError)) {
        m_queryDb.close();
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
    if (QSqlDatabase::contains(m_queryName))  QSqlDatabase::removeDatabase(m_queryName);
    if (QSqlDatabase::contains(m_listenName)) QSqlDatabase::removeDatabase(m_listenName);
    m_open = false;
}

bool PostgresConnection::ping() {
    if (!m_open) return false;
    QSqlQuery q(m_queryDb);
    return q.exec("SELECT 1");
}

} // namespace DVE
