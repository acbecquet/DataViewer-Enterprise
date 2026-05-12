#include "MigrationReport.h"

#include <QFile>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QDateTime>

namespace DVE {

void MigrationReport::addTable(const QString& name, int sqliteCount, int postgresCount) {
    m_tables.push_back({name, {sqliteCount, postgresCount}});
}

void MigrationReport::addError(const QString& msg) { m_errors << msg; }
void MigrationReport::setDuration(qint64 ms)         { m_durationMs = ms; }
void MigrationReport::setSourcePath(const QString& p){ m_sourcePath = p; }
void MigrationReport::setStatus(const QString& s)    { m_status = s; }

bool MigrationReport::writeJson(const QString& path) const {
    QJsonObject root;
    root["status"]      = m_status;
    root["source"]      = m_sourcePath;
    root["duration_ms"] = m_durationMs;
    root["timestamp"]   = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    QJsonArray tables;
    for (const auto& t : m_tables) {
        QJsonObject o;
        o["name"]           = t.first;
        o["sqlite_count"]   = t.second.first;
        o["postgres_count"] = t.second.second;
        o["match"]          = (t.second.first == t.second.second);
        tables.append(o);
    }
    root["tables"] = tables;

    QJsonArray errs;
    for (const QString& e : m_errors) errs.append(e);
    root["errors"] = errs;

    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Truncate)) return false;
    f.write(QJsonDocument(root).toJson(QJsonDocument::Indented));
    return true;
}

QString MigrationReport::summary() const {
    int total = 0;
    bool allMatch = true;
    for (const auto& t : m_tables) {
        total += t.second.first;
        if (t.second.first != t.second.second) allMatch = false;
    }
    return QString("status=%1 tables=%2 rows=%3 match=%4 errors=%5 took=%6ms")
        .arg(m_status)
        .arg(m_tables.size())
        .arg(total)
        .arg(allMatch ? "yes" : "no")
        .arg(m_errors.size())
        .arg(m_durationMs);
}

} // namespace DVE
