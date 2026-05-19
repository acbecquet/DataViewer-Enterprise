#pragma once

#include <QObject>
#include <QString>
#include <QSet>
#include <QUuid>
#include <QVariant>

class QSqlDriver;

namespace DVE {

class PostgresConnection;

struct RowChange {
    QString  table;
    QString  op;          // "INSERT" | "UPDATE" | "DELETE"
    qint64   id   = -1;
    QString  updatedBy;   // UUID string (matches IdentityManager::uuid().toString(WithoutBraces))
    QString  column;      // empty unless single-column UPDATE
    QVariant newValue;    // empty unless single-column UPDATE
};

struct PresenceChange {
    QString op;
    QUuid   userUuid;
    QString resourceType;
    qint64  resourceId = -1;
    QString intent;
};

struct CellFocusChange {
    QString op;          // "INSERT" | "UPDATE" | "DELETE"
    QUuid   userUuid;
    QString userName;
    QString userColor;
    QString tableName;
    qint64  rowId = -1;
    QString columnName;
};

class NotificationListener : public QObject {
    Q_OBJECT
public:
    explicit NotificationListener(PostgresConnection* conn, QObject* parent = nullptr);
    ~NotificationListener() override;

    // v2.0.2 H5: subscribe attempts all three channels independently.
    // Returns true if at least one channel was subscribed; the caller
    // can use isSubscribedTo() to verify a specific channel was attached.
    // This replaces the previous all-or-nothing bool which left the
    // listener with no subscriptions when a single LISTEN failed (e.g.,
    // permission revoke on one channel after restart).
    bool subscribe();
    void unsubscribe();
    bool isSubscribed() const { return !m_subscribedChannels.isEmpty(); }
    bool isSubscribedTo(const QString& channel) const {
        return m_subscribedChannels.contains(channel);
    }

signals:
    void rowChanged(const DVE::RowChange& change);
    void presenceChanged(const DVE::PresenceChange& change);
    void cellFocusChanged(const DVE::CellFocusChange& change);

private slots:
    void onNotification(const QString& name, int source, const QVariant& payload);

private:
    PostgresConnection* m_conn;
    QSet<QString>       m_subscribedChannels;
};

} // namespace DVE

Q_DECLARE_METATYPE(DVE::RowChange)
Q_DECLARE_METATYPE(DVE::PresenceChange)
Q_DECLARE_METATYPE(DVE::CellFocusChange)
