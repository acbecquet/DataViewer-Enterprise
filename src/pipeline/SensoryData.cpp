#include "SensoryData.h"
#include "ui/TesterRound.h"

#include <QJsonArray>
#include <QJsonObject>
#include <QJsonValue>
#include <QHash>
#include <QRegularExpression>
#include <QtGlobal>

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
        // DATAVIEWER-11: emit sample_uid ONLY when set, so desktop-authored
        // samples keep the exact prior JSON shape (additive, tolerant on read).
        if (!sample.sampleUid.isEmpty())
            sObj["sample_uid"]     = sample.sampleUid;
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
        // DATAVIEWER-11: tolerant read -- absent on every pre-existing row/file.
        sample.sampleUid         = sObj.value("sample_uid").toString();
        // Scores: tolerant read THEN clamp — a string-typed score ("7.0") from
        // the live per-cell stream must parse to 7.0, not collapse to the 5.0
        // default (the reset-to-5 revert this fix kills).
        for (const QString& metric : kSensoryMetrics)
            sample.scores[metric] = qBound(1.0, jsonToDouble(sObj[metric], 5.0), 9.0);
        sess.samples.append(sample);
    }
    return sess;
}

const QStringList& sensoryArbitratedSampleKeys()
{
    static const QStringList keys = QStringList()
        << kSensoryMetrics
        << QStringLiteral("name") << QStringLiteral("comments")
        << QStringLiteral("voltage") << QStringLiteral("resistance")
        << QStringLiteral("power")
        << QStringLiteral("heating_technology")
        << QStringLiteral("power_type")
        << QStringLiteral("puff_length_sec");
    return keys;
}

QJsonObject mergeSensoryPreservingDbScores(const QJsonObject& inMemory,
                                           const QJsonObject& dbCurrent,
                                           const QSet<QString>& dirtyCells,
                                           const QSet<QString>& removedUids)
{
    QJsonObject merged = inMemory;
    const QJsonArray dbSamples  = dbCurrent.value("samples").toArray();
    QJsonArray       memSamples = merged.value("samples").toArray();

    // DATAVIEWER-19 (audit fix): a sample was removed this run iff removedUids is
    // non-empty (SensoryPanel::onRemoveCard records the removed sample's uid, or a
    // sentinel for a uid-less desktop sample). The two regimes need DIFFERENT
    // merges:
    //  * No removal -> the in-memory and DB sample arrays are prefix-aligned (the
    //    desktop may have ADDED a sample locally, or the DB may carry a concurrent
    //    phone append the desktop lacks). The original positional overlay + the
    //    size-driven DB-tail re-adopt are correct and preserve a phone append.
    //  * After a removal -> array indices have SHIFTED, so index-based matching
    //    would smear a deleted sample's DB scores onto the survivor that slid into
    //    its slot, and the size-driven tail re-adopt would RESURRECT the removed
    //    sample (the audited corruption). So overlay DB scores ONLY by sample_uid
    //    (index-independent) and re-adopt ONLY a genuine concurrent append -- a DB
    //    uid the desktop never had AND did not remove. A desktop sample has no uid
    //    and is therefore never resurrected.
    const bool removalThisRun = !removedUids.isEmpty();

    // DV-25: DB-authoritative unless locally dirty, for EVERY per-cell-
    // committed sample key (was: score metrics only - which is why a stale
    // whole-save could revert a co-open client's committed rename). "power" is
    // UI-derived from voltage/resistance and has no per-cell dirty path of its
    // own, so it rides voltage-or-resistance dirtiness.
    const auto fieldDirty = [&dirtyCells](int idx, const QString& key) {
        const auto path = [idx](const QString& k) {
            return QStringLiteral("samples[%1].%2").arg(idx).arg(k);
        };
        if (dirtyCells.contains(path(key))) return true;
        if (key == QLatin1String("power"))
            return dirtyCells.contains(path(QStringLiteral("voltage")))
                || dirtyCells.contains(path(QStringLiteral("resistance")));
        return false;
    };

    if (!removalThisRun) {
        for (int i = 0; i < memSamples.size() && i < dbSamples.size(); ++i) {
            QJsonObject       memSample = memSamples[i].toObject();
            const QJsonObject dbSample  = dbSamples[i].toObject();
            for (const QString& key : sensoryArbitratedSampleKeys()) {
                // v2.5.0 Task 3 (RC2) + DV-25: a cell the user edited this run
                // stays in-memory-authoritative; only untouched cells take the
                // DB value - now for every arbitrated field, not just scores.
                if (fieldDirty(i, key)) continue;
                if (dbSample.contains(key))
                    memSample[key] = dbSample.value(key);
            }
            memSamples[i] = memSample;
        }
        // DATAVIEWER-11: keep a sample that exists ONLY in the DB -- a phone append
        // that arrived while this session was open (no local edit to reconcile).
        for (int i = memSamples.size(); i < dbSamples.size(); ++i)
            memSamples.append(dbSamples.at(i));
    } else {
        QHash<QString, QJsonObject> dbByUid;
        for (const QJsonValue& dv : dbSamples) {
            const QString u = dv.toObject().value(QStringLiteral("sample_uid")).toString();
            if (!u.isEmpty()) dbByUid.insert(u, dv.toObject());
        }
        QSet<QString> memUids;
        for (const QJsonValue& mv : memSamples) {
            const QString u = mv.toObject().value(QStringLiteral("sample_uid")).toString();
            if (!u.isEmpty()) memUids.insert(u);
        }
        for (int i = 0; i < memSamples.size(); ++i) {
            QJsonObject   memSample = memSamples[i].toObject();
            const QString uid       = memSample.value(QStringLiteral("sample_uid")).toString();
            if (!uid.isEmpty() && dbByUid.contains(uid)) {        // identity match (shift-proof)
                const QJsonObject dbSample = dbByUid.value(uid);
                for (const QString& key : sensoryArbitratedSampleKeys()) {
                    if (fieldDirty(i, key)) continue;   // DV-25: full field set
                    if (dbSample.contains(key))
                        memSample[key] = dbSample.value(key);
                }
            }
            memSamples[i] = memSample;
        }
        for (const QJsonValue& dv : dbSamples) {                 // re-adopt genuine appends only
            const QJsonObject d = dv.toObject();
            const QString u = d.value(QStringLiteral("sample_uid")).toString();
            if (u.isEmpty()) continue;                            // desktop sample -> never resurrect
            if (memUids.contains(u)) continue;                    // already present
            if (removedUids.contains(u)) continue;                // user removed it -> honor removal
            memSamples.append(d);
        }
    }
    merged["samples"] = memSamples;
    return merged;
}

bool sensorySessionNeedsSave(const SensorySession& s)
{
    if (s.id <= 0) return true;                        // never persisted
    if (s.dirty) return true;                          // any local edit this run
    if (!s.dirtyCells.isEmpty()) return true;          // per-cell edits recorded
    if (!s.removedSampleUids.isEmpty()) return true;   // pending removals
    if (!s.originalSessionName.isEmpty()
        && s.originalSessionName != s.sessionName)     // pending rename -> INSERT
        return true;
    return false;
}

void overlayMergedScores(SensorySession& s, const QJsonObject& merged)
{
    const QJsonArray mergedSamples = merged.value("samples").toArray();
    for (int i = 0; i < s.samples.size() && i < mergedSamples.size(); ++i) {
        const QJsonObject ms = mergedSamples[i].toObject();
        for (const QString& metric : kSensoryMetrics) {
            if (ms.contains(metric))
                s.samples[i].scores[metric] = ms.value(metric).toDouble();
        }
    }
}

void appendDbOnlyTailSamples(SensorySession& s, const QJsonObject& merged)
{
    const QJsonArray mergedSamples = merged.value("samples").toArray();
    if (mergedSamples.size() <= s.samples.size())
        return;
    // Parse the whole blob once and copy the tail samples verbatim (full fidelity
    // -- scores, name, comments, device props, sample_uid -- not just scores).
    const SensorySession parsed = sensorySessionFromJson(merged);
    for (int i = s.samples.size(); i < parsed.samples.size(); ++i)
        s.samples.append(parsed.samples.at(i));
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

double clampPuffSeconds(double rawSeconds)
{
    // Bound to the puff-length spin's range; the spin itself rounds to 1 decimal.
    return qBound(0.1, rawSeconds, 60.0);
}

} // namespace DVE
