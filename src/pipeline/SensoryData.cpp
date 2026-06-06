#include "SensoryData.h"

#include <QJsonArray>
#include <QJsonObject>
#include <QJsonValue>

namespace DVE {

QJsonObject sensorySessionToJson(const SensorySession& s)
{
    QJsonObject root;
    root["session_name"]        = s.sessionName;
    root["test_title"]          = s.testTitle;
    root["assessor_name"]       = s.assessorName;
    root["tester_name"]         = s.testerName;
    root["media"]               = s.media;
    root["date"]                = s.date;
    root["timestamp"]           = s.timestamp;
    root["control"]             = s.control;
    root["is_blind"]            = s.isBlind;
    root["primary_differences"] = s.primaryDifferences;
    root["puff_length"]         = s.puffLength;
    root["burn_status"]         = s.burnStatus;
    root["clog_status"]         = s.clogStatus;
    root["leak_status"]         = s.leakStatus;
    root["resistance"]          = s.resistance;
    root["voltage"]             = s.voltage;
    root["power"]               = s.power;
    root["heating_technology"]  = s.heatingTechnology;

    QJsonArray samplesArr;
    for (const SensorySample& sample : s.samples) {
        QJsonObject sObj;
        sObj["name"]     = sample.name;
        sObj["comments"] = sample.comments;
        for (const QString& metric : kSensoryMetrics)
            sObj[metric] = sample.scores.value(metric, 5.0);
        sObj["voltage"]            = sample.voltage;
        sObj["resistance"]         = sample.resistance;
        sObj["power"]              = sample.power;
        sObj["heating_technology"] = sample.heatingTechnology;
        sObj["power_type"]         = sample.powerType;
        sObj["puff_length_sec"]    = sample.puffLengthSec;
        samplesArr.append(sObj);
    }
    root["samples"] = samplesArr;
    return root;
}

SensorySession sensorySessionFromJson(const QJsonObject& root)
{
    SensorySession sess;
    sess.sessionName        = root["session_name"].toString();
    sess.testTitle          = root["test_title"].toString();
    sess.assessorName       = root["assessor_name"].toString();
    sess.testerName         = root["tester_name"].toString();
    sess.media              = root["media"].toString();
    sess.date               = root["date"].toString();
    sess.timestamp          = root["timestamp"].toString();
    sess.control            = root["control"].toString();
    sess.isBlind            = root["is_blind"].toBool(false);
    sess.primaryDifferences = root["primary_differences"].toString();
    sess.puffLength         = root["puff_length"].toString();
    sess.burnStatus         = root["burn_status"].toString();
    sess.clogStatus         = root["clog_status"].toString();
    sess.leakStatus         = root["leak_status"].toString();
    sess.resistance         = root["resistance"].toDouble();
    sess.voltage            = root["voltage"].toDouble();
    sess.power              = root["power"].toDouble();
    sess.heatingTechnology  = root["heating_technology"].toString();

    const QJsonArray samples = root["samples"].toArray();
    for (const QJsonValue& sv : samples) {
        const QJsonObject sObj = sv.toObject();
        SensorySample sample;
        sample.name              = sObj["name"].toString();
        sample.comments          = sObj["comments"].toString();
        sample.voltage           = sObj["voltage"].toDouble();
        sample.resistance        = sObj["resistance"].toDouble();
        sample.power             = sObj["power"].toDouble();
        sample.heatingTechnology = sObj["heating_technology"].toString();
        // #7: backward-compatible defaults preserved when reading older
        // rows / files that pre-date these fields.
        sample.powerType         = sObj.contains("power_type")
            ? sObj["power_type"].toString()
            : QStringLiteral("Constant Voltage");
        sample.puffLengthSec     = sObj.contains("puff_length_sec")
            ? sObj["puff_length_sec"].toDouble(3.0) : 3.0;
        for (const QString& metric : kSensoryMetrics)
            sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(5.0), 9.0);
        sess.samples.append(sample);
    }
    return sess;
}

QJsonObject mergeSensoryPreservingDbScores(const QJsonObject& inMemory,
                                           const QJsonObject& dbCurrent)
{
    QJsonObject merged = inMemory;
    const QJsonArray dbSamples  = dbCurrent.value("samples").toArray();
    QJsonArray       memSamples = merged.value("samples").toArray();
    for (int i = 0; i < memSamples.size() && i < dbSamples.size(); ++i) {
        QJsonObject       memSample = memSamples[i].toObject();
        const QJsonObject dbSample  = dbSamples[i].toObject();
        for (const QString& metric : kSensoryMetrics) {
            if (dbSample.contains(metric))
                memSample[metric] = dbSample.value(metric);
        }
        memSamples[i] = memSample;
    }
    merged["samples"] = memSamples;
    return merged;
}

} // namespace DVE
