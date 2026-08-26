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
const QLatin1String kMixed("mixed");
const QLatin1String kNumberList("numberlist");
const QLatin1String kImage("image");

// Can this value honestly occupy a BOOLEAN slot? A real bool and a numeric both
// can; so does a recognised yes/no word. Anything else CANNOT, and saying so is
// the whole point: QVariant::toBool() on a string is true for everything except
// "", "0" and "false", so "no" / "n" / "N" / "No" all read as TRUE - and
// did_burn / did_clog / did_leak are Bool in the registry while
// SchemaDrivenReader hands their header cells over as raw QStrings.
//
// The accepted spellings are LegacyAdapter::normaliseBCL's
// (src/model/LegacyAdapter.cpp), which is where the project already decides what
// a Y/N answer looks like. Restated rather than shared on purpose - this file is
// linked by suites that deliberately do NOT link the model layer, see above -
// so the two must be kept in step: a spelling added there belongs here too.
bool asBool(const QVariant& v, bool* out)
{
    if (v.typeId() == QMetaType::Bool) { *out = v.toBool(); return true; }
    bool numeric = false;
    const double d = v.toDouble(&numeric);
    if (numeric) { *out = (d != 0.0); return true; }
    const QString u = v.toString().trimmed().toUpper();
    if (u == QLatin1String("Y") || u == QLatin1String("YES")) { *out = true;  return true; }
    if (u == QLatin1String("N") || u == QLatin1String("NO"))  { *out = false; return true; }
    // TRUE/FALSE are unambiguously boolean spellings, so by this function's own
    // rule - coerce only what can honestly occupy the slot - they belong here.
    // They are also the two string forms QVariant::toBool() got right before the
    // did_burn="no" inversion fix, so accepting them keeps that behaviour.
    if (u == QLatin1String("TRUE"))  { *out = true;  return true; }
    if (u == QLatin1String("FALSE")) { *out = false; return true; }
    return false;
}

// Compact JSON array of doubles, or false when the value is not a genuine
// sequence of numbers. QVariant::toList() returns an EMPTY list for anything
// that is not a sequence, so without the type test a scalar under a
// `numberlist` key would be destroyed and stored as the literal "[]"; without
// the per-element test a text cell inside the list ("N/A") would be flattened
// to 0.0, which is indistinguishable from a measured zero once stored.
bool asNumberListJson(const QVariant& v, QString* out)
{
    const int t = v.typeId();
    if (t != QMetaType::QVariantList && t != QMetaType::QStringList) return false;
    QJsonArray arr;
    const QVariantList list = v.toList();
    for (const QVariant& e : list) {
        bool ok = false;
        const double d = e.toDouble(&ok);
        if (!ok) return false;
        arr.append(d);
    }
    *out = QString::fromUtf8(QJsonDocument(arr).toJson(QJsonDocument::Compact));
    return true;
}

// Compact JSON array preserving each element's own JSON form. The text
// fallback's rendering of a list - lossless, and never the empty string
// QVariant::toString() would have produced.
QString listAsJson(const QVariant& v)
{
    QJsonArray arr;
    const QVariantList list = v.toList();
    for (const QVariant& e : list) arr.append(QJsonValue::fromVariant(e));
    return QString::fromUtf8(QJsonDocument(arr).toJson(QJsonDocument::Compact));
}
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

    // "Available" means ALL THREE long tables exist, not merely metric_defs.
    // DatabaseOps gates its prune pre-image SELECTs against `measurements` and
    // `sample_headers` INSIDE the save transaction, where one failed statement
    // poisons the transaction and takes the whole save down with it - which is
    // exactly what this pre-transaction probe exists to prevent. ensureSchema
    // creates the three independently and best-effort, `continue`-ing past a
    // failure, so "metric_defs but not measurements" is a reachable state and a
    // metric_defs-only probe would report it available.
    //
    // to_regclass() yields NULL for an absent relation instead of raising, and
    // resolves through search_path exactly as the app's own unqualified
    // statements do. Same probe shape as snapshotContentFingerprint's, for the
    // same reason.
    QSqlQuery probe(m_db);
    if (!probe.exec(QStringLiteral(
            "SELECT to_regclass('metric_defs')    IS NOT NULL "
            "   AND to_regclass('measurements')   IS NOT NULL "
            "   AND to_regclass('sample_headers') IS NOT NULL"))
        || !probe.next()) {
        if (outError) *outError = probe.lastError().text();
        return false;
    }
    if (!probe.value(0).toBool()) {
        if (outError)
            *outError = QStringLiteral("the long-format tables (metric_defs, "
                                       "measurements, sample_headers) are not all present");
        return false;
    }

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
        if (!sel.exec()) {
            if (outError)
                *outError = QStringLiteral("MetricDefCache(resolve %1/%2): ")
                                .arg(kind, key) + sel.lastError().text();
            return -1;
        }
        if (!sel.next()) {
            // Distinct from the branch above: the query RAN. There is simply no
            // row, which means the INSERT's ON CONFLICT fired against something
            // this SELECT cannot see. lastError() is empty here, so saying so
            // explicitly is the difference between a diagnosable message and a
            // dangling colon.
            if (outError)
                *outError = QStringLiteral("MetricDefCache(resolve %1/%2): the INSERT "
                                           "reported a conflict but no row is there")
                                .arg(kind, key);
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

qint64 MetricDefCache::lookup(const QString& kind, const QString& key) const
{
    // Pure cache read: load() pulled the whole table, and every STANDARD key
    // is guaranteed present there by the migration seed + ensureSchema. A miss
    // is a broken-contract signal for the caller, never a cue to register.
    return m_ids.value(cacheKey(kind, key), -1);
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

    // ONE rule, applied to every typed slot: coerce only when the value can
    // HONESTLY occupy it, and otherwise fall through to the text fallback.
    // Never guess. A guessed value is indistinguishable from a measured one once
    // it is stored, and after one save the original is gone - a confident lie is
    // strictly worse than a value the reader has to interpret as text.
    if (valueType == kBool) {
        bool b = false;
        if (asBool(v, &b)) { outNum = b ? 1.0 : 0.0; return; }
    } else if (valueType == kNumber || valueType == kMixed) {
        // `mixed` is the registry's "a number where it can be, text otherwise"
        // type (live: metric|voltage), which IS this policy - so a numeric
        // per-row voltage reaches value_num and comes back a double.
        bool ok = false;
        const double d = v.toDouble(&ok);
        if (ok) { outNum = d; return; }
    } else if (valueType == kImage) {
        // Only bytes can occupy the base64 slot. QVariant::toByteArray() would
        // happily hand back an empty array for a list, or a number's digits.
        if (v.typeId() == QMetaType::QByteArray) {
            outText = QString::fromLatin1(v.toByteArray().toBase64());
            return;
        }
    } else if (valueType == kNumberList) {
        QString json;
        if (asNumberListJson(v, &json)) { outText = json; return; }
    }

    // ---- text fallback ---------------------------------------------------
    // Everything the branches above declined lands here, and decodeValue hands
    // this string straight back, so the two halves always agree on what is
    // stored.
    //
    // The text slot has an honesty test of its own. QVariant::toString() is the
    // EMPTY string for a list and mangles non-UTF-8 bytes, so those two keep a
    // lossless textual rendering instead - compact JSON and base64. That DEMOTES
    // the value to a string rather than destroying it, which is the same rule
    // one level down. inferValueType already gives an auto-registered key
    // holding bytes or a list the matching value_type, so this only fires when a
    // pre-existing metric_defs row disagrees with the workbook.
    if (v.typeId() == QMetaType::QByteArray) {
        outText = QString::fromLatin1(v.toByteArray().toBase64());
        return;
    }
    if (v.typeId() == QMetaType::QVariantList || v.typeId() == QMetaType::QStringList) {
        outText = listAsJson(v);
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
    // Both decoders below have to tolerate a string encodeValue's TEXT FALLBACK
    // wrote: a value that could not occupy the typed slot lands in value_text
    // under its own type name, and decoding it blindly would turn "N/A" into an
    // empty list and "no photo" into garbage bytes - re-introducing on the read
    // side exactly the destruction the write side just refused to commit.
    if (valueType == kImage) {
        const QByteArray::FromBase64Result decoded = QByteArray::fromBase64Encoding(
            s.toLatin1(), QByteArray::AbortOnBase64DecodingErrors);
        if (decoded.decodingStatus == QByteArray::Base64DecodingStatus::Ok)
            return QVariant(decoded.decoded);
        return QVariant(s);
    }
    if (valueType == kNumberList) {
        const QJsonDocument doc = QJsonDocument::fromJson(s.toUtf8());
        if (doc.isArray()) return QVariant(doc.array().toVariantList());
        return QVariant(s);
    }
    return QVariant(s);
}

} // namespace DVE
