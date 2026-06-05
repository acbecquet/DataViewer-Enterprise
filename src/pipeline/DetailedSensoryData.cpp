#include "DetailedSensoryData.h"

#include <QJsonArray>
#include <QJsonObject>
#include <QJsonValue>

namespace DVE {

// Canonical JSON contract for DetailedSensorySession. Promoted out of
// DatabaseManager (where it lived as a file-private anonymous-namespace pair)
// so the DB layer and the Plan C recovery snapshot share one implementation.
// Unlike the old DatabaseManager copies these return / accept a QJsonObject
// directly -- the QJsonDocument string (de)wrapping lives at the call sites.
//
// Image vectors and id/version are intentionally NOT serialized here, matching
// the historical DB helpers: images persist in their own *_images table and
// id/version are row identities the loader stamps back on after parsing.
QJsonObject detailedSensorySessionToJson(const DetailedSensorySession& s)
{
    QJsonObject root;
    root["session_name"]        = s.sessionName;
    root["test_title"]          = s.testTitle;
    root["assessor_name"]       = s.assessorName;
    root["tester_name"]         = s.testerName;
    root["facilitator_name"]    = s.facilitatorName;
    root["facilitator_comment"] = s.facilitatorComment;
    root["media"]               = s.media;
    root["date"]                = s.date;
    root["timestamp"]           = s.timestamp;
    root["oil_smell_liking"]    = s.oilSmellLiking;
    root["clog"]                = s.clog;
    root["clog_oil_level"]      = s.clogOilLevel;
    root["mouthpiece_notes"]    = s.mouthpieceNotes;
    root["device_return_date"]  = s.deviceReturnDate;
    root["viscosity"]           = s.viscosity;

    QJsonArray samplesArr;
    for (const DetailedSensorySample& sample : s.samples) {
        QJsonObject sObj;
        sObj["name"]     = sample.name;
        sObj["comments"] = sample.comments;
        for (const QString& metric : kDetailedAllMetrics) {
            sObj[metric] = sample.scores.value(metric, 0.0);
        }
        sObj["voltage"]            = sample.voltage;
        sObj["resistance"]         = sample.resistance;
        sObj["power"]              = sample.power;
        sObj["heating_technology"] = sample.heatingTechnology;
        samplesArr.append(sObj);
    }
    root["samples"] = samplesArr;
    return root;
}

DetailedSensorySession detailedSensorySessionFromJson(const QJsonObject& root)
{
    DetailedSensorySession sess;
    sess.sessionName        = root["session_name"].toString();
    sess.testTitle          = root["test_title"].toString();
    sess.assessorName       = root["assessor_name"].toString();
    sess.testerName         = root["tester_name"].toString();
    sess.facilitatorName    = root["facilitator_name"].toString();
    sess.facilitatorComment = root["facilitator_comment"].toString();
    sess.media              = root["media"].toString();
    sess.date               = root["date"].toString();
    sess.timestamp          = root["timestamp"].toString();
    sess.oilSmellLiking     = root["oil_smell_liking"].toInt(3);
    sess.clog               = root["clog"].toBool(false);
    sess.clogOilLevel       = root["clog_oil_level"].toString();
    sess.mouthpieceNotes    = root["mouthpiece_notes"].toString();
    sess.deviceReturnDate   = root["device_return_date"].toString();
    sess.viscosity          = root["viscosity"].toString();

    for (const QJsonValue& sv : root["samples"].toArray()) {
        const QJsonObject sObj = sv.toObject();
        DetailedSensorySample sample;
        sample.name              = sObj["name"].toString();
        sample.comments          = sObj["comments"].toString();
        sample.voltage           = sObj["voltage"].toDouble();
        sample.resistance        = sObj["resistance"].toDouble();
        sample.power             = sObj["power"].toDouble();
        sample.heatingTechnology = sObj["heating_technology"].toString();
        for (const QString& metric : kDetailedAllMetrics) {
            const double maxVal = kDetailedMetricMaxScore.value(metric, 9);
            sample.scores[metric] = qBound(1.0, sObj[metric].toDouble(1.0), maxVal);
        }
        sess.samples.append(sample);
    }
    return sess;
}

// True when the session is a freshly-created placeholder that the user has not
// meaningfully filled in yet. Plan C uses this to keep empty "New Session"
// cards out of the recovery snapshot (and, indirectly, out of the auto-save).
//
// Unlike SensorySession::isPlaceholderSession this does NOT short-circuit on
// id >= 0: the recovery snapshot is a pure "does this hold real work?" question
// independent of whether the row was ever persisted. A session counts as a
// placeholder only when ALL of the following hold:
//   1. The name is empty or still the default "New Session".
//   2. No user-editable session-level field has content. This covers the
//      header (test title / tester / assessor / media) AND the S2-1 extras
//      (facilitator name + comment, clog oil level, mouthpiece notes,
//      viscosity, device return date, oil-smell liking, clog flag).
//   3. Every sample (if any) is itself empty -- no name, comments, V/R/power,
//      heating tech, or any recorded score.
// NOTE: date/timestamp are intentionally NOT gated on -- newSession() auto-
// populates them, so a real placeholder always has them set; treating them as
// "content" would make every session non-placeholder.
bool isPlaceholderSession(const DetailedSensorySession& s)
{
    if (!s.sessionName.isEmpty()
        && s.sessionName != QStringLiteral("New Session")) {
        return false;
    }

    // The 3 / false literals are the struct (and newSession()) defaults for
    // oilSmellLiking and clog; any deviation means the user touched the field.
    const bool noSessionContent = s.testTitle.isEmpty()
                               && s.testerName.isEmpty()
                               && s.assessorName.isEmpty()
                               && s.media.isEmpty()
                               && s.facilitatorName.isEmpty()
                               && s.facilitatorComment.isEmpty()
                               && s.clogOilLevel.isEmpty()
                               && s.mouthpieceNotes.isEmpty()
                               && s.viscosity.isEmpty()
                               && s.deviceReturnDate.isEmpty()
                               && s.oilSmellLiking == 3
                               && s.clog == false;
    if (!noSessionContent) return false;

    for (const DetailedSensorySample& sample : s.samples) {
        if (!sample.name.isEmpty())              return false;
        if (!sample.comments.isEmpty())          return false;
        if (sample.voltage    > 0.0)             return false;
        if (sample.resistance > 0.0)             return false;
        if (sample.power      > 0.0)             return false;
        if (!sample.heatingTechnology.isEmpty()) return false;
        if (!sample.scores.isEmpty())            return false;
    }
    return true;
}

} // namespace DVE
