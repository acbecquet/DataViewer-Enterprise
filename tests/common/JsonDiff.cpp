#include "JsonDiff.h"
#include <QJsonArray>
#include <QJsonObject>

namespace DVE { namespace testutil {

static QString shortRepr(const QJsonValue& v)
{
    switch (v.type()) {
    case QJsonValue::Object: return QStringLiteral("{object}");
    case QJsonValue::Array:  return QStringLiteral("[array:%1]").arg(v.toArray().size());
    default:                 return v.toVariant().toString().left(60);
    }
}

QStringList diffJson(const QJsonValue& a, const QJsonValue& b, const QString& path)
{
    QStringList out;
    if (a.type() != b.type()) {
        out << QStringLiteral("%1: type %2 != %3").arg(path).arg(static_cast<int>(a.type())).arg(static_cast<int>(b.type()));
        return out;
    }
    if (a.isObject()) {
        const QJsonObject oa = a.toObject(), ob = b.toObject();
        QStringList keys = oa.keys() + ob.keys();
        keys.removeDuplicates();
        keys.sort();
        for (const QString& k : keys) {
            if (!oa.contains(k)) { out << QStringLiteral("%1/%2: missing on left").arg(path, k);  continue; }
            if (!ob.contains(k)) { out << QStringLiteral("%1/%2: missing on right").arg(path, k); continue; }
            out += diffJson(oa.value(k), ob.value(k), path + '/' + k);
        }
        return out;
    }
    if (a.isArray()) {
        const QJsonArray aa = a.toArray(), ab = b.toArray();
        if (aa.size() != ab.size()) {
            out << QStringLiteral("%1: array size %2 != %3").arg(path).arg(aa.size()).arg(ab.size());
            return out;
        }
        for (int i = 0; i < aa.size(); ++i)
            out += diffJson(aa.at(i), ab.at(i), QStringLiteral("%1[%2]").arg(path).arg(i));
        return out;
    }
    if (a != b)
        out << QStringLiteral("%1: %2 != %3").arg(path, shortRepr(a), shortRepr(b));
    return out;
}

}} // namespace DVE::testutil
