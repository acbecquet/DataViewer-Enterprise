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
        qRegisterMetaType<CellFocusChange>();
    });
}

NotificationListener::~NotificationListener() { unsubscribe(); }

bool NotificationListener::subscribe() {
    if (!m_conn || !m_conn->isOpen()) return false;
    QSqlDriver* drv = m_conn->listenDb().driver();
    if (!drv) return false;
    static const QStringList kChannels = {
        QStringLiteral("dataviewer_changes"),
        QStringLiteral("dataviewer_presence"),
        QStringLiteral("dataviewer_cell_focus"),
    };
    // H5: attempt each channel independently and record successes. If
    // one LISTEN fails (e.g. permission denied on a subset), we still
    // get notifications on the channels that did succeed instead of
    // ending up with no subscriptions at all.
    for (const QString& ch : kChannels) {
        if (drv->subscribeToNotification(ch)) {
            m_subscribedChannels.insert(ch);
        }
    }
    if (!m_subscribedChannels.isEmpty()) {
        connect(drv, &QSqlDriver::notification,
                this, &NotificationListener::onNotification);
    }
    return !m_subscribedChannels.isEmpty();
}

void NotificationListener::unsubscribe() {
    if (m_subscribedChannels.isEmpty()) return;
    if (m_conn && m_conn->isOpen()) {
        QSqlDriver* drv = m_conn->listenDb().driver();
        if (drv) {
            // Iterate the recorded set so partial subscriptions are
            // unwound correctly — calling unsubscribeFromNotification on
            // a channel we never subscribed to is a Qt-level warning.
            for (const QString& ch : m_subscribedChannels) {
                drv->unsubscribeFromNotification(ch);
            }
            disconnect(drv, &QSqlDriver::notification, this, nullptr);
        }
    }
    m_subscribedChannels.clear();
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
        c.column    = o.value("column").toString();
        if (o.contains("new_value")) c.newValue = o.value("new_value").toVariant();
        emit rowChanged(c);
    } else if (name == "dataviewer_presence") {
        PresenceChange p;
        p.op           = o.value("op").toString();
        p.userUuid     = QUuid(o.value("user_uuid").toString());
        p.resourceType = o.value("resource_type").toString();
        p.resourceId   = o.value("resource_id").toVariant().toLongLong();
        p.intent       = o.value("intent").toString();
        emit presenceChanged(p);
    } else if (name == "dataviewer_cell_focus") {
        CellFocusChange f;
        f.op         = o.value("op").toString();
        f.userUuid   = QUuid(o.value("user_uuid").toString());
        f.userName   = o.value("user_name").toString();
        f.userColor  = o.value("user_color").toString();
        f.tableName  = o.value("table").toString();
        f.rowId      = o.value("row_id").toVariant().toLongLong();
        f.columnName = o.value("column").toString();
        emit cellFocusChanged(f);
    }
}

} // namespace DVE
