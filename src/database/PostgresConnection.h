#pragma once

#include <QObject>
#include <QString>
#include <QSqlDatabase>
#include "ConfigLoader.h"

namespace DVE {

class PostgresConnection : public QObject {
    Q_OBJECT
public:
    explicit PostgresConnection(QObject* parent = nullptr);
    ~PostgresConnection() override;

    bool open(const DbConfig& cfg);
    bool tryOpenWithRetry(const DbConfig& cfg, int totalTimeoutMs);
    void close();
    bool isOpen() const { return m_open; }

    bool ping();

    QSqlDatabase& queryDb()  { return m_queryDb; }
    QSqlDatabase& listenDb() { return m_listenDb; }

    QString lastError() const { return m_lastError; }

private:
    bool         m_open = false;
    QSqlDatabase m_queryDb;
    QSqlDatabase m_listenDb;
    QString      m_lastError;
    QString      m_queryName;
    QString      m_listenName;
};

} // namespace DVE
