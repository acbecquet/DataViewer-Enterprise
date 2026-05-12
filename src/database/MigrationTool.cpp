#include "MigrationTool.h"

#include <QFile>
#include <QSqlError>
#include <QSqlQuery>
#include <QUuid>
#include <QCryptographicHash>

namespace DVE {

MigrationTool::MigrationTool(QObject* parent) : QObject(parent) {
    const QString tag = QUuid::createUuid().toString(QUuid::WithoutBraces).left(8);
    m_pgConnName = "dve_mig_" + tag;
}

MigrationTool::~MigrationTool() {
    if (m_sqlite.isOpen()) m_sqlite.close();
    if (m_pg.isOpen())     m_pg.close();
    m_sqlite = QSqlDatabase();
    m_pg     = QSqlDatabase();
    const QString srcName = "dve_mig_src_" + m_pgConnName;
    if (QSqlDatabase::contains(srcName))     QSqlDatabase::removeDatabase(srcName);
    if (QSqlDatabase::contains(m_pgConnName)) QSqlDatabase::removeDatabase(m_pgConnName);
}

bool MigrationTool::open(const QString& sqlitePath, const DbConfig& pg) {
    m_sqlitePath = sqlitePath;
    m_report.setSourcePath(sqlitePath);

    if (!QFile::exists(sqlitePath)) {
        m_lastError = "Source SQLite not found at " + sqlitePath;
        return false;
    }

    const QString srcName = "dve_mig_src_" + m_pgConnName;
    m_sqlite = QSqlDatabase::addDatabase("QSQLITE", srcName);
    m_sqlite.setDatabaseName(sqlitePath);
    m_sqlite.setConnectOptions("QSQLITE_OPEN_READONLY");
    if (!m_sqlite.open()) {
        m_lastError = "SQLite open failed: " + m_sqlite.lastError().text();
        return false;
    }

    QFile f(sqlitePath);
    if (f.open(QIODevice::ReadOnly)) {
        QCryptographicHash h(QCryptographicHash::Sha256);
        if (h.addData(&f)) m_sourceSha256 = h.result().toHex();
    }

    m_pg = QSqlDatabase::addDatabase("QPSQL", m_pgConnName);
    m_pg.setHostName(pg.host);
    m_pg.setPort(pg.port);
    m_pg.setDatabaseName(pg.database);
    m_pg.setUserName(pg.user);
    m_pg.setPassword(pg.password);
    m_pg.setConnectOptions("connect_timeout=5");
    if (!m_pg.open()) {
        m_lastError = "Postgres open failed: " + m_pg.lastError().text();
        return false;
    }
    return true;
}

bool MigrationTool::checkSchemaMetaEmpty() {
    QSqlQuery q(m_pg);
    if (!q.exec("SELECT value FROM schema_meta WHERE key = 'migrated_at'")) {
        m_lastError = "Could not read schema_meta: " + q.lastError().text();
        return false;
    }
    return !q.next();
}

// Stubs (implemented in tasks 16-20):
bool MigrationTool::wipePostgresData()                 { return false; }
int  MigrationTool::sqliteRowCount(const QString&)     { return 0; }
int  MigrationTool::postgresRowCount(const QString&)   { return 0; }
bool MigrationTool::migrateTable(const QString&)       { return false; }
bool MigrationTool::bumpSequence(const QString&)       { return false; }
bool MigrationTool::writeSchemaMeta()                  { return false; }
bool MigrationTool::finalizeSource()                   { return false; }
bool MigrationTool::run(bool)                          { return false; }

} // namespace DVE
