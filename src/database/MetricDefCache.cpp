#include "MetricDefCache.h"

#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonValue>
#include <QMetaType>
#include <QSqlError>
#include <QSqlQuery>

namespace DVE {

namespace {
// Manifest.cpp kTypeNames spelling. Kept as literals rather than reaching for
// model::ValueType: this file is linked by suites that deliberately do NOT link
// the model layer (tst_v3longformat is an SQL-first harness), and metric_defs
// stores the NAME, not the enum.
const QLatin1String kNumber("number");
const QLatin1String kText("text");
const QLatin1String kBool("bool");
const QLatin1String kNumberList("numberlist");
const QLatin1String kImage("image");
} // namespace

MetricDefCache::MetricDefCache(const QSqlDatabase& db, const QString& who)
    : m_db(db), m_who(who)
{
}

QString MetricDefCache::cacheKey(const QString& kind, const QString& key)
{
    // UNIQUE is (kind, key), never (key): `resistance` and `puffing_regime`
    // each exist as BOTH a data_rows metric and a samples header field, and the
    // registry keeps the two namespaces separate.
    return kind + QLatin1Char('|') + key;
}

bool MetricDefCache::load(QString* outError)
{
    m_ids.clear();
    m_types.clear();
    m_available = false;

    QSqlQuery q(m_db);
    if (!q.exec(QStringLiteral("SELECT id, kind, key, value_type FROM metric_defs"))) {
        if (outError) *outError = q.lastError().text();
        return false;
    }
    while (q.next()) {
        const qint64 id = q.value(0).toLongLong();
        m_ids.insert(cacheKey(q.value(1).toString(), q.value(2).toString()), id);
        m_types.insert(id, q.value(3).toString());
    }
    m_available = true;
    return true;
}

qint64 MetricDefCache::ensureMetric(const QString& kind, const QString& key,
                                    const QVariant& valueHint, QString* outError)
{
    if (!m_available) {
        if (outError)
            *outError = QStringLiteral("MetricDefCache: load() has not succeeded");
        return -1;
    }

    const QString ck = cacheKey(kind, key);
    const auto hit = m_ids.constFind(ck);
    if (hit != m_ids.constEnd()) return hit.value();

    // DO NOTHING, never DO UPDATE - see the header. On conflict no row comes
    // back and the re-select below resolves the row the other writer committed.
    QSqlQuery ins(m_db);
    if (!ins.prepare(QStringLiteral(
            "INSERT INTO metric_defs (kind, key, display_name, value_type, role, updated_by) "
            "VALUES (?, ?, ?, ?, 'measured', ?) "
            "ON CONFLICT (kind, key) DO NOTHING RETURNING id, value_type"))) {
        if (outError) *outError = ins.lastError().text();
        return -1;
    }
    ins.addBindValue(kind);
    ins.addBindValue(key);
    ins.addBindValue(key);                       // display_name IS the key
    ins.addBindValue(inferValueType(valueHint));
    ins.addBindValue(m_who);
    if (!ins.exec()) {
        if (outError)
            *outError = QStringLiteral("MetricDefCache(register %1/%2): ")
                            .arg(kind, key) + ins.lastError().text();
        return -1;
    }

    qint64  id = -1;
    QString type;
    if (ins.next()) {
        id   = ins.value(0).toLongLong();
        type = ins.value(1).toString();
    } else {
        QSqlQuery sel(m_db);
        sel.prepare(QStringLiteral(
            "SELECT id, value_type FROM metric_defs WHERE kind = ? AND key = ?"));
        sel.addBindValue(kind);
        sel.addBindValue(key);
        if (!sel.exec() || !sel.next()) {
            if (outError)
                *outError = QStringLiteral("MetricDefCache(resolve %1/%2): ")
                                .arg(kind, key) + sel.lastError().text();
            return -1;
        }
        id   = sel.value(0).toLongLong();
        type = sel.value(1).toString();
    }

    m_ids.insert(ck, id);
    m_types.insert(id, type);
    return id;
}

QString MetricDefCache::valueType(qint64 id) const
{
    return m_types.value(id);
}

QString MetricDefCache::inferValueType(const QVariant& v)
{
    switch (v.typeId()) {
    case QMetaType::QByteArray:
        return kImage;
    case QMetaType::QVariantList:
        // {"a": [...]} in the envelope - a numeric list, e.g. the assembled
        // PV1..PV5 draw_pressure_per_puff metric (registry 8.2). A QStringList
        // is a DIFFERENT QMetaType and deliberately does not land here, exactly
        // as in the envelope.
        return kNumberList;
    case QMetaType::Bool:
        return kBool;
    case QMetaType::Int:
    case QMetaType::UInt:
    case QMetaType::LongLong:
    case QMetaType::ULongLong:
    case QMetaType::Double:
    case QMetaType::Float:
        return kNumber;
    default:
        return kText;
    }
}

void MetricDefCache::encodeValue(const QString& valueType, const QVariant& v,
                                 QVariant& outNum, QVariant& outText)
{
    outNum  = QVariant();
    outText = QVariant();

    if (valueType == kBool) {
        outNum = v.toBool() ? 1.0 : 0.0;
        return;
    }
    if (valueType == kNumber) {
        bool ok = false;
        const double d = v.toDouble(&ok);
        if (ok) {
            outNum = d;
            return;
        }
        // Not representable as a number. Fall through to the text encoders
        // rather than storing a coerced 0.0 - the value survives, and the reader
        // takes whichever column is populated.
    }
    if (valueType == kImage || v.typeId() == QMetaType::QByteArray) {
        outText = QString::fromLatin1(v.toByteArray().toBase64());
        return;
    }
    if (valueType == kNumberList || v.typeId() == QMetaType::QVariantList) {
        QJsonArray arr;
        const QVariantList list = v.toList();
        for (const QVariant& e : list) arr.append(e.toDouble());
        outText = QString::fromUtf8(QJsonDocument(arr).toJson(QJsonDocument::Compact));
        return;
    }
    outText = v.toString();
}

QVariant MetricDefCache::decodeValue(const QString& valueType,
                                     const QVariant& num, const QVariant& text)
{
    if (!num.isNull()) {
        if (valueType == kBool) return QVariant(num.toDouble() != 0.0);
        return QVariant(num.toDouble());
    }
    if (text.isNull()) return QVariant();

    const QString s = text.toString();
    if (valueType == kImage)
        return QVariant(QByteArray::fromBase64(s.toLatin1()));
    if (valueType == kNumberList) {
        QVariantList list;
        const QJsonArray arr = QJsonDocument::fromJson(s.toUtf8()).array();
        for (const QJsonValue& e : arr) list.append(e.toDouble());
        return QVariant(list);
    }
    return QVariant(s);
}

} // namespace DVE
