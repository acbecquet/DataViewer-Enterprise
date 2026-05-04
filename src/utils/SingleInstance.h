#pragma once

#include <QObject>
#include <QLocalServer>
#include <QLocalSocket>

/**
 * Ensures only one instance of the application runs at a time.
 * If another instance is already running, sends the file path to it and returns true.
 * Otherwise, starts a local server to listen for future instances.
 */
class SingleInstance : public QObject
{
    Q_OBJECT

public:
    explicit SingleInstance(const QString& key, QObject* parent = nullptr);
    ~SingleInstance() override;

    /// Try to send a message to an existing instance. Returns true if sent (caller should exit).
    bool sendToRunning(const QString& message);

    /// Start listening for messages from future instances.
    bool startServer();

signals:
    /// Emitted when another instance sends a file path.
    void messageReceived(const QString& message);

private slots:
    void onNewConnection();

private:
    QString       m_key;
    QLocalServer* m_server = nullptr;
};
