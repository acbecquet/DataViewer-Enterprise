#pragma once

#include <QObject>
#include <QString>
#include <QSqlDatabase>
#include "ConfigLoader.h"
#include "MigrationReport.h"

namespace DVE {

class MigrationTool : public QObject {
    Q_OBJECT
public:
    explicit MigrationTool(QObject* parent = nullptr);
    ~MigrationTool() override;

    bool open(const QString& sqlitePath, const DbConfig& pg);
    bool run(bool force = false);
    bool finalizeSource();

    const MigrationReport& report() const { return m_report; }
    QString lastError() const { return m_lastError; }

    // Exposed for direct test use (Task 16 onwards):
    bool migrateTable(const QString& name);

private:
    bool checkSchemaMetaEmpty();
    bool wipePostgresData();
    int  sqliteRowCount(const QString& table);
    int  postgresRowCount(const QString& table);
    bool bumpSequence(const QString& table);
    bool writeSchemaMeta();

    QSqlDatabase    m_sqlite;
    QSqlDatabase    m_pg;
    QString         m_sqlitePath;
    QString         m_pgConnName;
    QString         m_lastError;
    MigrationReport m_report;
    QString         m_sourceSha256;
};

} // namespace DVE
