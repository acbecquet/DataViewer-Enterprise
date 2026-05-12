#pragma once

#include <QObject>
#include <QTimer>
#include <QString>
#include <QUuid>
#include <QVector>

namespace DVE {

class PostgresConnection;
class IdentityManager;

struct PresenceRow {
    QUuid   userUuid;
    QString userName;
    QString userColor;
    QString resourceType;
    qint64  resourceId = -1;
    QString intent;
};

class PresenceManager : public QObject {
    Q_OBJECT
public:
    PresenceManager(PostgresConnection* conn, IdentityManager* identity,
                    QObject* parent = nullptr);
    ~PresenceManager() override;

    // Switch the active resource for this user. Inserts a new presence row
    // and deletes any previous active row in the same resourceType. Starts
    // the heartbeat timer if not already running.
    bool activate(const QString& resourceType, qint64 resourceId,
                  const QString& intent = "viewing");

    // Update intent for the current active resource ("viewing" → "editing").
    bool setIntent(const QString& intent);

    // Remove this user's presence for the current active resource.
    bool deactivate();

    // List active presence rows for a resource (for UI presence dots).
    // Filters by 30s last_heartbeat freshness window.
    QVector<PresenceRow> activeFor(const QString& resourceType, qint64 resourceId) const;

    QString activeResourceType() const { return m_activeType; }
    qint64  activeResourceId()   const { return m_activeId; }
    QString activeIntent()       const { return m_activeIntent; }

private slots:
    void heartbeat();

private:
    PostgresConnection* m_conn;
    IdentityManager*    m_identity;
    QTimer              m_timer;
    QString             m_activeType;
    qint64              m_activeId = -1;
    QString             m_activeIntent;
};

} // namespace DVE
