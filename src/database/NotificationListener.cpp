#include "NotificationListener.h"
#include "PostgresConnection.h"

#include <QSqlDriver>
#include <QJsonDocument>
#include <QJsonObject>
#include <mutex>

namespace DVE {

NotificationListener::NotificationListener(PostgresConnection* conn, QObject* parent)
    : QObject(parent), m_conn(conn) {
    static std::once_flag once;
    std::call_once(once, []() {
        qRegisterMetaType<RowChange>();
        qRegisterMetaType<PresenceChange>();
    });
}

NotificationListener::~NotificationListener() { unsubscribe(); }

bool NotificationListener::subscribe() {
    if (!m_conn || !m_conn->isOpen()) return false;
    QSqlDriver* drv = m_conn->listenDb().driver();
    if (!drv) return false;
    bool ok = drv->subscribeToNotification("dataviewer_changes");
    ok = drv->subscribeToNotification("dataviewer_presence") && ok;
    if (ok) {
        connect(drv, &QSqlDriver::notification,
                this, &NotificationListener::onNotification);
        m_subscribed = true;
    }
    return ok;
}

void NotificationListener::unsubscribe() {
    if (!m_subscribed) return;
    if (m_conn && m_conn->isOpen()) {
        QSqlDriver* drv = m_conn->listenDb().driver();
        if (drv) {
            drv->unsubscribeFromNotification("dataviewer_changes");
            drv->unsubscribeFromNotification("dataviewer_presence");
            disconnect(drv, &QSqlDriver::notification, this, nullptr);
        }
    }
    m_subscribed = false;
}

void NotificationListener::onNotification(const QString& name, int, const QVariant& payload) {
    const QJsonDocument doc = QJsonDocument::fromJson(payload.toString().toUtf8());
    if (!doc.isObject()) return;
    const QJsonObject o = doc.object();

    if (name == "dataviewer_changes") {
        RowChange c;
        c.table     = o.value("table").toString();
        c.op        = o.value("op").toString();
        c.id        = o.value("id").toVariant().toLongLong();
        c.updatedBy = o.value("updated_by").toString();
        emit rowChanged(c);
    } else if (name == "dataviewer_presence") {
        PresenceChange p;
        p.op           = o.value("op").toString();
        p.userUuid     = QUuid(o.value("user_uuid").toString());
        p.resourceType = o.value("resource_type").toString();
        p.resourceId   = o.value("resource_id").toVariant().toLongLong();
        p.intent       = o.value("intent").toString();
        emit presenceChanged(p);
    }
}

} // namespace DVE
