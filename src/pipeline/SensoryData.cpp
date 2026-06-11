#include "SensoryData.h"
#include "ui/TesterRound.h"

#include <QJsonArray>
#include <QJsonObject>
#include <QJsonValue>
#include <QRegularExpression>

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
    // Tolerant numeric reads — see jsonToDouble in SensoryData.h. Live-streamed
    // values may be stored as JSON strings by the (pre-fix) commit function.
    sess.resistance         = jsonToDouble(root["resistance"], 0.0);
    sess.voltage            = jsonToDouble(root["voltage"], 0.0);
    sess.power              = jsonToDouble(root["power"], 0.0);
    sess.heatingTechnology  = root["heating_technology"].toString();

    const QJsonArray samples = root["samples"].toArray();
    for (const QJsonValue& sv : samples) {
        const QJsonObject sObj = sv.toObject();
        SensorySample sample;
        sample.name              = sObj["name"].toString();
        sample.comments          = sObj["comments"].toString();
        sample.voltage           = jsonToDouble(sObj["voltage"], 0.0);
        sample.resistance        = jsonToDouble(sObj["resistance"], 0.0);
        sample.power             = jsonToDouble(sObj["power"], 0.0);
        sample.heatingTechnology = sObj["heating_technology"].toString();
        // #7: backward-compatible defaults preserved when reading older
        // rows / files that pre-date these fields.
        sample.powerType         = sObj.contains("power_type")
            ? sObj["power_type"].toString()
            : QStringLiteral("Constant Voltage");
        sample.puffLengthSec     = sObj.contains("puff_length_sec")
            ? jsonToDouble(sObj["puff_length_sec"], 3.0) : 3.0;
        // Scores: tolerant read THEN clamp — a string-typed score ("7.0") from
        // the live per-cell stream must parse to 7.0, not collapse to the 5.0
        // default (the reset-to-5 revert this fix kills).
        for (const QString& metric : kSensoryMetrics)
            sample.scores[metric] = qBound(1.0, jsonToDouble(sObj[metric], 5.0), 9.0);
        sess.samples.append(sample);
    }
    return sess;
}

QJsonObject mergeSensoryPreservingDbScores(const QJsonObject& inMemory,
                                           const QJsonObject& dbCurrent,
                                           const QSet<QString>& dirtyCells)
{
    QJsonObject merged = inMemory;
    const QJsonArray dbSamples  = dbCurrent.value("samples").toArray();
    QJsonArray       memSamples = merged.value("samples").toArray();
    for (int i = 0; i < memSamples.size() && i < dbSamples.size(); ++i) {
        QJsonObject       memSample = memSamples[i].toObject();
        const QJsonObject dbSample  = dbSamples[i].toObject();
        for (const QString& metric : kSensoryMetrics) {
            // v2.5.0 Task 3 (RC2): a cell the user edited this run stays
            // in-memory-authoritative; only untouched cells take the DB value.
            const QString path =
                QStringLiteral("samples[%1].%2").arg(i).arg(metric);
            if (dirtyCells.contains(path)) continue;
            if (dbSample.contains(metric))
                memSample[metric] = dbSample.value(metric);
        }
        memSamples[i] = memSample;
    }
    merged["samples"] = memSamples;
    return merged;
}

QSet<QString> remapDirtyCellsAfterSampleRemoval(const QSet<QString>& dirty,
                                                int removedIdx)
{
    // Parse the leading "samples[<idx>]." of each path. Anything that doesn't
    // match that prefix is not index-bearing and passes through verbatim.
    static const QRegularExpression re(
        QStringLiteral("^samples\\[(\\d+)\\]\\.(.*)$"));
    QSet<QString> out;
    out.reserve(dirty.size());
    for (const QString& path : dirty) {
        const QRegularExpressionMatch m = re.match(path);
        if (!m.hasMatch()) {            // non-sample path: keep as-is
            out.insert(path);
            continue;
        }
        const int idx = m.captured(1).toInt();
        if (idx == removedIdx)          // the removed sample: drop
            continue;
        if (idx > removedIdx) {         // later sample: shift index down by one
            out.insert(QStringLiteral("samples[%1].%2")
                           .arg(idx - 1)
                           .arg(m.captured(2)));
        } else {                        // earlier sample: index unchanged
            out.insert(path);
        }
    }
    return out;
}

QSet<QString> adoptedDirtyCellsAfterSave(int savedId,
                                         const QSet<QString>& savedDirty,
                                         const QSet<QString>& panelDirty)
{
    // savedId <= 0 means the write never landed (placeholder / version mismatch
    // / hard error) — leave the panel's set exactly as it was.
    if (savedId <= 0) return panelDirty;
    // savedId > 0: adopt the caller's set. The caller cleared it iff this tick's
    // WriteResult == Success, so a failed (but previously-persisted) session
    // still carries its dirty cells here and keeps its protection.
    return savedDirty;
}

bool isSensorySessionSavable(const SensorySession& s)
{
    if (s.testTitle.trimmed().isEmpty()) return false;
    const QString tester = splitTesterRound(s.testerName).tester;   // strip round
    return !tester.trimmed().isEmpty();
}

} // namespace DVE
