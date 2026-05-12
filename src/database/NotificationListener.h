#pragma once

#include <QObject>
#include <QString>
#include <QUuid>

class QSqlDriver;

namespace DVE {

class PostgresConnection;

struct RowChange {
    QString table;
    QString op;          // "INSERT" | "UPDATE" | "DELETE"
    qint64  id   = -1;
    QString updatedBy;   // UUID string (matches IdentityManager::uuid().toString(WithoutBraces))
};

struct PresenceChange {
    QString op;
    QUuid   userUuid;
    QString resourceType;
    qint64  resourceId = -1;
    QString intent;
};

class NotificationListener : public QObject {
    Q_OBJECT
public:
    explicit NotificationListener(PostgresConnection* conn, QObject* parent = nullptr);
    ~NotificationListener() override;

    bool subscribe();
    void unsubscribe();
    bool isSubscribed() const { return m_subscribed; }

signals:
    void rowChanged(const DVE::RowChange& change);
    void presenceChanged(const DVE::PresenceChange& change);

private slots:
    void onNotification(const QString& name, int source, const QVariant& payload);

private:
    PostgresConnection* m_conn;
    bool                m_subscribed = false;
};

} // namespace DVE

Q_DECLARE_METATYPE(DVE::RowChange)
Q_DECLARE_METATYPE(DVE::PresenceChange)
