#include "SingleInstance.h"
#include <QLocalSocket>
#include <QDataStream>

static const int TIMEOUT_MS = 3000;

SingleInstance::SingleInstance(const QString& key, QObject* parent)
    : QObject(parent), m_key(key)
{
}

SingleInstance::~SingleInstance()
{
    if (m_server) {
        m_server->close();
    }
}

bool SingleInstance::sendToRunning(const QString& message)
{
    QLocalSocket socket;
    socket.connectToServer(m_key);
    if (!socket.waitForConnected(TIMEOUT_MS))
        return false;

    QByteArray data = message.toUtf8();
    socket.write(data);
    socket.waitForBytesWritten(TIMEOUT_MS);
    socket.disconnectFromServer();
    return true;
}

bool SingleInstance::startServer()
{
    m_server = new QLocalServer(this);
    connect(m_server, &QLocalServer::newConnection, this, &SingleInstance::onNewConnection);

    // Remove stale server (e.g. after crash)
    QLocalServer::removeServer(m_key);
    return m_server->listen(m_key);
}

void SingleInstance::onNewConnection()
{
    QLocalSocket* socket = m_server->nextPendingConnection();
    if (!socket) return;

    socket->waitForReadyRead(TIMEOUT_MS);
    QString message = QString::fromUtf8(socket->readAll());
    socket->disconnectFromServer();
    socket->deleteLater();

    if (!message.isEmpty()) {
        emit messageReceived(message);
    }
}
