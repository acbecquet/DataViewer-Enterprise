#include "pipeline/ReportDataJson.h"

#include <QJsonArray>
#include <QJsonValue>

namespace DVE {

// JSON contract for the TPM FileResult tree. Mirrors the sensorySessionToJson
// style. Keep the field set + key names symmetric across the To/From pair, and
// in sync with the round-trip test in tst_reportdatajson. See ReportDataJson.h
// for the list of intentionally-omitted (re-derivable / transient) fields.
//
// Reader convention: a missing/absent key decodes to the struct's zero-default
// (QJsonValue::Undefined -> toDouble()/toInt() == 0, toString() == "", toBool()
// == false, toArray()/toObject() empty). This is deliberate — it lets old blobs
// forward-decode. A future field whose *correct* default is non-zero must pass
// that default explicitly to the toX() call (e.g. toDouble(defaultValue)), since
// the implicit fallback is always zero.

namespace {

// qint64 ids are written as JSON numbers via double. JSON has no integer type
// and a double's 53-bit mantissa exactly represents every id Postgres BIGSERIAL
// will realistically hand out, so this is lossless for the values we persist.
QJsonValue idToJson(qint64 id)
{
    return QJsonValue(static_cast<double>(id));
}

qint64 idFromJson(const QJsonValue& v)
{
    // Missing id -> -1 ("not yet persisted"), per ReportData.h. This is one of
    // the few readers that overrides the implicit zero-default on purpose: an id
    // of 0 is a valid persisted row, so absence must map to the sentinel instead.
    return static_cast<qint64>(v.toDouble(-1));
}

QJsonObject rectToJson(const QRectF& r)
{
    QJsonObject o;
    o["x"] = r.x();
    o["y"] = r.y();
    o["w"] = r.width();
    o["h"] = r.height();
    return o;
}

QRectF rectFromJson(const QJsonValue& v)
{
    const QJsonObject o = v.toObject();
    return QRectF(o["x"].toDouble(), o["y"].toDouble(),
                  o["w"].toDouble(), o["h"].toDouble());
}

QJsonArray stringsToJson(const QStringList& list)
{
    QJsonArray arr;
    for (const QString& s : list)
        arr.append(s);
    return arr;
}

QStringList stringsFromJson(const QJsonValue& v)
{
    QStringList list;
    const QJsonArray arr = v.toArray();
    for (const QJsonValue& e : arr)
        list.append(e.toString());
    return list;
}

// ---- DataRow::extra typed envelope -------------------------------------------
//
// DataRow::extra holds open per-row values from inferred (non-standard)
// schemas (e.g. PV1..PV5, chronology) - unlike SampleResult::extra, whose
// plain fromVariantMap()/toVariantMap() round-trip is fine because it only
// ever holds JSON-native types. DataRow::extra additionally needs to carry
// QByteArray (and preserve integer-vs-double-ness) losslessly, so each value
// gets a one-key typed envelope: {"n": double} | {"i": qint64-as-double} |
// {"s": string} | {"b": base64 bytes}. Any other QVariant type (e.g. bool,
// QDateTime) coerces to its toString() form under "s" - Phase 3 broadens the
// envelope if a schema ever needs one of those types to round-trip exactly.
QJsonValue extraValueToJson(const QVariant& v)
{
    QJsonObject o;
    switch (v.typeId()) {
    case QMetaType::QByteArray:
        o["b"] = QString::fromLatin1(v.toByteArray().toBase64());
        break;
    case QMetaType::Int:
    case QMetaType::UInt:
    case QMetaType::LongLong:
    case QMetaType::ULongLong:
        // Same double-mantissa caveat as idToJson: exact for every value this
        // batch's inferred schemas realistically produce.
        o["i"] = static_cast<double>(v.toLongLong());
        break;
    case QMetaType::Double:
    case QMetaType::Float:
        o["n"] = v.toDouble();
        break;
    case QMetaType::QString:
        o["s"] = v.toString();
        break;
    default:
        o["s"] = v.toString();
        break;
    }
    return o;
}

QVariant extraValueFromJson(const QJsonValue& v)
{
    const QJsonObject o = v.toObject();
    if (o.contains("b"))
        return QVariant(QByteArray::fromBase64(o["b"].toString().toLatin1()));
    if (o.contains("i"))
        return QVariant(static_cast<qint64>(o["i"].toDouble()));
    if (o.contains("n"))
        return QVariant(o["n"].toDouble());
    return QVariant(o["s"].toString());
}

QJsonObject extraMapToJson(const QMap<QString, QVariant>& m)
{
    QJsonObject o;
    for (auto it = m.constBegin(); it != m.constEnd(); ++it)
        o[it.key()] = extraValueToJson(it.value());
    return o;
}

QMap<QString, QVariant> extraMapFromJson(const QJsonValue& v)
{
    QMap<QString, QVariant> m;
    const QJsonObject o = v.toObject();
    for (auto it = o.constBegin(); it != o.constEnd(); ++it)
        m.insert(it.key(), extraValueFromJson(it.value()));
    return m;
}

// ---- DataRow ----------------------------------------------------------------

QJsonObject rowToJson(const DataRow& r)
{
    QJsonObject o;
    o["puffs"]            = r.puffs;
    o["before_weight"]    = r.beforeWeight;
    o["after_weight"]     = r.afterWeight;
    o["draw_pressure"]    = r.drawPressure;
    o["resistance"]       = r.resistance;
    o["puffing_regime"]   = r.puffingRegime;
    o["smell"]            = r.smell;
    o["clog"]             = r.clog;
    o["notes"]            = r.notes;
    o["tpm"]              = r.tpm;
    o["tpm_power_density"] = r.tpmPowerDensity;
    o["variation_tpm"]    = r.variationTPM;
    o["oil_consumed"]     = r.oilConsumed;
    // Omit the key entirely when empty - the vast majority of rows have no
    // open per-row values, and existing recovery snapshots (written before
    // this field existed) must stay byte-stable.
    if (!r.extra.isEmpty())
        o["extra"] = extraMapToJson(r.extra);
    o["id"]               = idToJson(r.id);
    o["version"]          = r.version;
    return o;
}

DataRow rowFromJson(const QJsonObject& o)
{
    DataRow r;
    r.puffs           = o["puffs"].toDouble();
    r.beforeWeight    = o["before_weight"].toDouble();
    r.afterWeight     = o["after_weight"].toDouble();
    r.drawPressure    = o["draw_pressure"].toDouble();
    r.resistance      = o["resistance"].toDouble();
    r.puffingRegime   = o["puffing_regime"].toString();
    r.smell           = o["smell"].toString();
    r.clog            = o["clog"].toString();
    r.notes           = o["notes"].toString();
    r.tpm             = o["tpm"].toDouble();
    r.tpmPowerDensity = o["tpm_power_density"].toDouble();
    r.variationTPM    = o["variation_tpm"].toDouble();
    r.oilConsumed     = o["oil_consumed"].toDouble();
    if (o.contains("extra"))
        r.extra = extraMapFromJson(o["extra"]);
    r.id              = idFromJson(o["id"]);
    r.version         = o["version"].toInt();
    return r;
}

// ---- SampleResult -----------------------------------------------------------

QJsonObject sampleToJson(const SampleResult& s)
{
    QJsonObject o;
    o["sample_name"]           = s.sampleName;
    o["sample_id"]             = s.sampleID;
    o["date"]                  = s.date;
    o["tester"]                = s.tester;
    o["media"]                 = s.media;
    o["viscosity"]             = s.viscosity;
    o["resistance"]            = s.resistance;
    o["voltage"]               = s.voltage;
    o["power"]                 = s.power;
    o["heating_technology"]    = s.heatingTechnology;
    o["puffing_regime"]        = s.puffingRegime;
    o["initial_oil_mass"]      = s.initialOilMass;
    o["average_tpm"]           = s.averageTPM;
    o["std_dev_tpm"]           = s.stdDevTPM;
    o["average_power_density"] = s.averagePowerDensity;
    o["efficiency_percent"]    = s.efficiencyPercent;
    o["total_oil_consumed"]    = s.totalOilConsumed;
    o["total_puffs"]           = s.totalPuffs;
    o["normalized_tpm"]        = s.normalizedTPM;
    o["burn_status"]           = s.burnStatus;
    o["clog_status"]           = s.clogStatus;
    o["leak_status"]           = s.leakStatus;

    QJsonArray rowsArr;
    for (const DataRow& r : s.rows)
        rowsArr.append(rowToJson(r));
    o["rows"] = rowsArr;

    // extra is a QVariantMap (QMap<QString,QVariant>). fromVariantMap/toVariantMap
    // round-trips losslessly only for JSON-native value types (number/string/bool/
    // nested map/list) — which is all `extra` ever holds. A non-native QVariant
    // (e.g. QDateTime, QColor) would be coerced/dropped, so don't start stashing one.
    o["extra"] = QJsonObject::fromVariantMap(s.extra);

    o["image_paths"] = stringsToJson(s.imagePaths);

    QJsonArray layoutsArr;
    for (const QRectF& r : s.imageLayouts)
        layoutsArr.append(rectToJson(r));
    o["image_layouts"] = layoutsArr;

    QJsonArray cropsArr;
    for (const QRectF& r : s.imageCrops)
        cropsArr.append(rectToJson(r));
    o["image_crops"] = cropsArr;

    QJsonArray imageIdsArr;
    for (qint64 id : s.imageIds)
        imageIdsArr.append(idToJson(id));
    o["image_ids"] = imageIdsArr;

    QJsonArray imageVersionsArr;
    for (int v : s.imageVersions)
        imageVersionsArr.append(v);
    o["image_versions"] = imageVersionsArr;

    o["id"]      = idToJson(s.id);
    o["version"] = s.version;
    return o;
}

SampleResult sampleFromJson(const QJsonObject& o)
{
    SampleResult s;
    s.sampleName          = o["sample_name"].toString();
    s.sampleID            = o["sample_id"].toString();
    s.date                = o["date"].toString();
    s.tester              = o["tester"].toString();
    s.media               = o["media"].toString();
    s.viscosity           = o["viscosity"].toDouble();
    s.resistance          = o["resistance"].toDouble();
    s.voltage             = o["voltage"].toDouble();
    s.power               = o["power"].toDouble();
    s.heatingTechnology   = o["heating_technology"].toString();
    s.puffingRegime       = o["puffing_regime"].toString();
    s.initialOilMass      = o["initial_oil_mass"].toDouble();
    s.averageTPM          = o["average_tpm"].toDouble();
    s.stdDevTPM           = o["std_dev_tpm"].toDouble();
    s.averagePowerDensity = o["average_power_density"].toDouble();
    s.efficiencyPercent   = o["efficiency_percent"].toDouble();
    s.totalOilConsumed    = o["total_oil_consumed"].toDouble();
    s.totalPuffs          = o["total_puffs"].toInt();
    s.normalizedTPM       = o["normalized_tpm"].toDouble();
    s.burnStatus          = o["burn_status"].toString();
    s.clogStatus          = o["clog_status"].toString();
    s.leakStatus          = o["leak_status"].toString();

    const QJsonArray rowsArr = o["rows"].toArray();
    for (const QJsonValue& rv : rowsArr)
        s.rows.append(rowFromJson(rv.toObject()));

    s.extra = o["extra"].toObject().toVariantMap();

    s.imagePaths = stringsFromJson(o["image_paths"]);

    const QJsonArray layoutsArr = o["image_layouts"].toArray();
    for (const QJsonValue& lv : layoutsArr)
        s.imageLayouts.append(rectFromJson(lv));

    const QJsonArray cropsArr = o["image_crops"].toArray();
    for (const QJsonValue& cv : cropsArr)
        s.imageCrops.append(rectFromJson(cv));

    const QJsonArray imageIdsArr = o["image_ids"].toArray();
    for (const QJsonValue& iv : imageIdsArr)
        s.imageIds.append(idFromJson(iv));

    const QJsonArray imageVersionsArr = o["image_versions"].toArray();
    for (const QJsonValue& vv : imageVersionsArr)
        s.imageVersions.append(vv.toInt());

    s.id      = idFromJson(o["id"]);
    s.version = o["version"].toInt();
    return s;
}

// ---- SheetResult ------------------------------------------------------------

QJsonObject sheetToJson(const SheetResult& sh)
{
    QJsonObject o;
    o["sheet_name"]          = sh.sheetName;
    o["template_version"]    = sh.templateVersion;
    o["has_per_row_regime"]  = sh.hasPerRowRegime;
    // Persisted into the recovery blob on purpose (not to Postgres): a recovered
    // inference-path sheet must stay write-protected. Absent in old blobs -> false.
    o["from_inferred_schema"] = sh.fromInferredSchema;
    o["column_headers"]      = stringsToJson(sh.columnHeaders);
    o["overall_avg_tpm"]     = sh.overallAvgTPM;
    o["overall_std_dev_tpm"] = sh.overallStdDevTPM;
    o["is_raw_table"]        = sh.isRawTable;
    o["raw_headers"]         = stringsToJson(sh.rawHeaders);

    QJsonArray rawRowsArr;
    for (const QStringList& row : sh.rawRows)
        rawRowsArr.append(stringsToJson(row));
    o["raw_rows"] = rawRowsArr;

    QJsonArray samplesArr;
    for (const SampleResult& s : sh.samples)
        samplesArr.append(sampleToJson(s));
    o["samples"] = samplesArr;

    o["id"]      = idToJson(sh.id);
    o["version"] = sh.version;

    // OMITTED (re-derived/transient): tpmTrend, puffCounts, images,
    // dbDataIncomplete. See ReportDataJson.h.
    return o;
}

SheetResult sheetFromJson(const QJsonObject& o)
{
    SheetResult sh;
    sh.sheetName        = o["sheet_name"].toString();
    sh.templateVersion  = o["template_version"].toString();
    sh.hasPerRowRegime  = o["has_per_row_regime"].toBool();
    sh.fromInferredSchema = o["from_inferred_schema"].toBool();
    sh.columnHeaders    = stringsFromJson(o["column_headers"]);
    sh.overallAvgTPM    = o["overall_avg_tpm"].toDouble();
    sh.overallStdDevTPM = o["overall_std_dev_tpm"].toDouble();
    sh.isRawTable       = o["is_raw_table"].toBool();
    sh.rawHeaders       = stringsFromJson(o["raw_headers"]);

    const QJsonArray rawRowsArr = o["raw_rows"].toArray();
    for (const QJsonValue& rv : rawRowsArr)
        sh.rawRows.append(stringsFromJson(rv));

    const QJsonArray samplesArr = o["samples"].toArray();
    for (const QJsonValue& sv : samplesArr)
        sh.samples.append(sampleFromJson(sv.toObject()));

    sh.id      = idFromJson(o["id"]);
    sh.version = o["version"].toInt();
    return sh;
}

} // namespace

// ---- FileResult (public) ----------------------------------------------------

QJsonObject fileResultToJson(const FileResult& f)
{
    QJsonObject root;
    root["file_path"]        = f.filePath;
    root["file_name"]        = f.fileName;
    root["template_version"] = f.templateVersion;
    root["sheet_names"]      = stringsToJson(f.sheetNames);

    QJsonArray sheetsArr;
    for (const SheetResult& sh : f.sheets)
        sheetsArr.append(sheetToJson(sh));
    root["sheets"] = sheetsArr;

    root["id"]      = idToJson(f.id);
    root["version"] = f.version;
    return root;
}

FileResult fileResultFromJson(const QJsonObject& obj)
{
    FileResult f;
    f.filePath        = obj["file_path"].toString();
    f.fileName        = obj["file_name"].toString();
    f.templateVersion = obj["template_version"].toString();
    f.sheetNames      = stringsFromJson(obj["sheet_names"]);

    const QJsonArray sheetsArr = obj["sheets"].toArray();
    for (const QJsonValue& sv : sheetsArr)
        f.sheets.append(sheetFromJson(sv.toObject()));

    f.id      = idFromJson(obj["id"]);
    f.version = obj["version"].toInt();
    return f;
}

} // namespace DVE
