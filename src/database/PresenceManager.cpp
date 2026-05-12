#include "PresenceManager.h"
#include "PostgresConnection.h"
#include "IdentityManager.h"

#include <QSqlDatabase>
#include <QSqlQuery>
#include <QSqlError>

namespace DVE {

PresenceManager::PresenceManager(PostgresConnection* conn,
                                 IdentityManager* identity, QObject* parent)
    : QObject(parent), m_conn(conn), m_identity(identity) {
    m_timer.setInterval(10000);
    connect(&m_timer, &QTimer::timeout, this, &PresenceManager::heartbeat);
}

PresenceManager::~PresenceManager() {
    if (m_activeId >= 0 && m_conn && m_conn->isOpen()) deactivate();
}

bool PresenceManager::activate(const QString& resourceType, qint64 resourceId,
                                const QString& intent) {
    if (!m_conn || !m_conn->isOpen() || !m_identity) return false;

    QSqlDatabase& db = m_conn->queryDb();
    if (!db.transaction()) return false;

    const QString uuidStr = m_identity->uuid().toString(QUuid::WithoutBraces);

    // Delete any prior active row in the same resourceType for this user.
    QSqlQuery del(db);
    del.prepare("DELETE FROM presence WHERE user_uuid = ? AND resource_type = ?");
    del.addBindValue(uuidStr);
    del.addBindValue(resourceType);
    if (!del.exec()) {
        db.rollback();
        return false;
    }

    // Insert (or upsert) the new active row.
    QSqlQuery ins(db);
    ins.prepare("INSERT INTO presence(user_uuid, user_name, user_color, "
                "resource_type, resource_id, intent, last_heartbeat) "
                "VALUES (?, ?, ?, ?, ?, ?, now()) "
                "ON CONFLICT (user_uuid, resource_type, resource_id) "
                "DO UPDATE SET intent = EXCLUDED.intent, last_heartbeat = now()");
    ins.addBindValue(uuidStr);
    ins.addBindValue(m_identity->displayName());
    ins.addBindValue(m_identity->color());
    ins.addBindValue(resourceType);
    ins.addBindValue(qlonglong(resourceId));
    ins.addBindValue(intent);
    if (!ins.exec()) {
        db.rollback();
        return false;
    }
    if (!db.commit()) return false;

    m_activeType   = resourceType;
    m_activeId     = resourceId;
    m_activeIntent = intent;
    if (!m_timer.isActive()) m_timer.start();
    return true;
}

bool PresenceManager::setIntent(const QString& intent) {
    if (m_activeId < 0 || !m_conn || !m_conn->isOpen()) return false;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("UPDATE presence SET intent = ?, last_heartbeat = now() "
              "WHERE user_uuid = ? AND resource_type = ? AND resource_id = ?");
    q.addBindValue(intent);
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(m_activeType);
    q.addBindValue(qlonglong(m_activeId));
    if (!q.exec()) return false;
    m_activeIntent = intent;
    return true;
}

bool PresenceManager::deactivate() {
    if (m_activeId < 0) { m_timer.stop(); return true; }
    if (!m_conn || !m_conn->isOpen()) {
        m_timer.stop();
        return false;
    }
    QSqlQuery q(m_conn->queryDb());
    q.prepare("DELETE FROM presence WHERE user_uuid = ? AND resource_type = ? "
              "AND resource_id = ?");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(m_activeType);
    q.addBindValue(qlonglong(m_activeId));
    const bool ok = q.exec();
    m_activeType.clear();
    m_activeId = -1;
    m_activeIntent.clear();
    m_timer.stop();
    return ok;
}

void PresenceManager::heartbeat() {
    if (m_activeId < 0 || !m_conn || !m_conn->isOpen()) return;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("UPDATE presence SET last_heartbeat = now() "
              "WHERE user_uuid = ? AND resource_type = ? AND resource_id = ?");
    q.addBindValue(m_identity->uuid().toString(QUuid::WithoutBraces));
    q.addBindValue(m_activeType);
    q.addBindValue(qlonglong(m_activeId));
    q.exec();
}

QVector<PresenceRow> PresenceManager::activeFor(const QString& resourceType,
                                                 qint64 resourceId) const {
    QVector<PresenceRow> result;
    if (!m_conn || !m_conn->isOpen()) return result;
    QSqlQuery q(m_conn->queryDb());
    q.prepare("SELECT user_uuid::text, user_name, user_color, intent "
              "FROM presence WHERE resource_type = ? AND resource_id = ? "
              "AND last_heartbeat > now() - INTERVAL '30 seconds'");
    q.addBindValue(resourceType);
    q.addBindValue(qlonglong(resourceId));
    if (!q.exec()) return result;
    while (q.next()) {
        PresenceRow r;
        r.userUuid     = QUuid(q.value(0).toString());
        r.userName     = q.value(1).toString();
        r.userColor    = q.value(2).toString();
        r.resourceType = resourceType;
        r.resourceId   = resourceId;
        r.intent       = q.value(3).toString();
        result.push_back(r);
    }
    return result;
}

} // namespace DVE
