#pragma once

#include <QString>
#include <QStringList>
#include <QVector>
#include <QPair>

namespace DVE {

class MigrationReport {
public:
    void addTable(const QString& name, int sqliteCount, int postgresCount);
    void addError(const QString& msg);
    void setDuration(qint64 ms);
    void setSourcePath(const QString& path);
    void setStatus(const QString& s);

    bool    writeJson(const QString& path) const;
    QString summary() const;

private:
    QString m_status;
    QString m_sourcePath;
    qint64  m_durationMs = 0;
    QStringList m_errors;
    QVector<QPair<QString, QPair<int,int>>> m_tables;
};

} // namespace DVE
