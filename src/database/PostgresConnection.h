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
    // v2.4.2 R4b: probe the LISTEN socket (m_listenDb) independently of the
    // query socket. A half-open NOTIFY connection (query socket still
    // answering, listen socket dead) silently stops live updates on flaky /
    // GFW networks; ConnectionMonitor pings BOTH so a dead listen socket also
    // flips offline -> reconnect.
    bool pingListen();

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
