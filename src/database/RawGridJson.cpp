#include "database/RawGridJson.h"
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
namespace DVE {
QString rawGridToJson(const QStringList& headers, const QVector<QStringList>& rows)
{
    if (headers.isEmpty() && rows.isEmpty())
        return QString();
    QJsonArray hdr;
    for (const QString& h : headers) hdr.append(h);
    QJsonArray rws;
    for (const QStringList& r : rows) {
        QJsonArray cells;
        for (const QString& c : r) cells.append(c);
        rws.append(cells);
    }
    QJsonObject obj;
    obj.insert(QStringLiteral("headers"), hdr);
    obj.insert(QStringLiteral("rows"), rws);
    return QString::fromUtf8(QJsonDocument(obj).toJson(QJsonDocument::Compact));
}
void rawGridFromJson(const QString& json, QStringList& headers, QVector<QStringList>& rows)
{
    headers.clear();
    rows.clear();
    if (json.isEmpty()) return;
    const QJsonDocument doc = QJsonDocument::fromJson(json.toUtf8());
    if (!doc.isObject()) return;
    const QJsonObject obj = doc.object();
    const QJsonArray hdr = obj.value(QStringLiteral("headers")).toArray();
    for (const QJsonValue& h : hdr) headers.append(h.toString());
    const QJsonArray rws = obj.value(QStringLiteral("rows")).toArray();
    for (const QJsonValue& rv : rws) {
        QStringList cells;
        const QJsonArray r = rv.toArray();
        for (const QJsonValue& c : r) cells.append(c.toString());
        rows.append(cells);
    }
}
} // namespace DVE
