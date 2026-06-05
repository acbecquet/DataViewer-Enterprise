#include <QtTest>
#include <QJsonArray>
#include <QJsonObject>
#include "pipeline/DetailedSensoryData.h"

using namespace DVE;

class TstDetailedSensoryJson : public QObject
{
    Q_OBJECT
private slots:
    // Canonical JSON helpers (detailedSensorySessionToJson /
    // detailedSensorySessionFromJson). Promoted out of DatabaseManager so the
    // DB layer and the Plan C recovery snapshot share one implementation; a
    // regression here is a regression in both.
    void jsonRoundTripPreservesAllFields();
    void deserializeClampsScoresToMetricRange();

    // Placeholder predicate -- gates empty "New Session" cards out of the
    // recovery snapshot / auto-save.
    void freshNewSessionIsPlaceholder();
    void emptyNameIsPlaceholder();
    void typedTitleMakesItNonPlaceholder();
    void typedSampleNameMakesItNonPlaceholder();
    void typedHeaderFieldMakesItNonPlaceholder();
    void typedSessionLevelFieldMakesItNonPlaceholder();
    void renamedSessionIsNonPlaceholder();
};

void TstDetailedSensoryJson::jsonRoundTripPreservesAllFields()
{
    DetailedSensorySession orig;
    orig.sessionName        = "RT Detailed";
    orig.testTitle          = "Mint vs Berry";
    orig.assessorName       = "Reviewer 3";
    orig.testerName         = "Alice";
    orig.facilitatorName    = "Bob";
    orig.facilitatorComment = "went smoothly";
    orig.media              = "Tobacco";
    orig.date               = "2026-06-04";
    orig.timestamp          = "2026-06-04T12:34:56Z";
    orig.oilSmellLiking     = 4;
    orig.clog               = true;
    orig.clogOilLevel       = "half";
    orig.mouthpieceNotes    = "easy pull";
    orig.deviceReturnDate   = "2026-06-10";
    orig.viscosity          = "medium";
    // Fields that are intentionally NOT serialized by the helpers. They must
    // come back as struct-defaults after the round-trip.
    orig.id                 = 7;
    orig.version            = 2;
    orig.imagePaths         << "C:/scans/a.png";
    orig.imageLayouts       << QRectF(1.0, 2.0, 3.0, 4.0);

    DetailedSensorySample s1;
    s1.name              = "DeviceA";
    s1.comments          = "tasted clean";
    s1.voltage           = 3.7;
    s1.resistance        = 1.2;
    s1.power             = 11.4;
    s1.heatingTechnology = "Ceramic Coil";
    s1.scores["Burn Taste"]              = 2.0;
    s1.scores["Flavor Intensity"]        = 5.0;
    s1.scores["Throat Irritation"]       = 3.0;
    s1.scores["Cough"]                   = 2.0;   // max 4
    s1.scores["Performance Consistency"] = 3.0;   // max 3
    s1.scores["Vapor Volume"]            = 4.0;   // max 5
    orig.samples.append(s1);

    DetailedSensorySample s2;
    s2.name              = "DeviceB";
    s2.comments          = "slight burn";
    s2.voltage           = 4.0;
    s2.resistance        = 1.5;
    s2.power             = 10.6;
    s2.heatingTechnology = "Mesh";
    s2.scores["Burn Taste"]         = 6.0;
    s2.scores["Nasal Irritation"]   = 4.0;
    s2.scores["Vapor vs Oil"]       = 2.0;   // max 4
    s2.scores["Volume Consistency"] = 1.0;   // max 4
    orig.samples.append(s2);

    // Round-trip through the canonical QJsonObject pair (the helpers return /
    // accept the object directly -- no QJsonDocument string wrapping).
    const QJsonObject json = detailedSensorySessionToJson(orig);
    const DetailedSensorySession decoded = detailedSensorySessionFromJson(json);

    QCOMPARE(decoded.sessionName,        orig.sessionName);
    QCOMPARE(decoded.testTitle,          orig.testTitle);
    QCOMPARE(decoded.assessorName,       orig.assessorName);
    QCOMPARE(decoded.testerName,         orig.testerName);
    QCOMPARE(decoded.facilitatorName,    orig.facilitatorName);
    QCOMPARE(decoded.facilitatorComment, orig.facilitatorComment);
    QCOMPARE(decoded.media,              orig.media);
    QCOMPARE(decoded.date,               orig.date);
    QCOMPARE(decoded.timestamp,          orig.timestamp);
    QCOMPARE(decoded.oilSmellLiking,     orig.oilSmellLiking);
    QCOMPARE(decoded.clog,               orig.clog);
    QCOMPARE(decoded.clogOilLevel,       orig.clogOilLevel);
    QCOMPARE(decoded.mouthpieceNotes,    orig.mouthpieceNotes);
    QCOMPARE(decoded.deviceReturnDate,   orig.deviceReturnDate);
    QCOMPARE(decoded.viscosity,          orig.viscosity);
    QCOMPARE(decoded.samples.size(),     orig.samples.size());

    // Non-serialized fields default out. Seed both an imagePaths entry and an
    // imageLayouts rect above so more than one image vector proves the image
    // data is never written to (and never read back from) the JSON.
    QCOMPARE(decoded.id,      -1);
    QCOMPARE(decoded.version, 0);
    QVERIFY(decoded.imagePaths.isEmpty());
    QVERIFY(decoded.imageLayouts.isEmpty());

    const DetailedSensorySample& d0 = decoded.samples[0];
    QCOMPARE(d0.name,              s1.name);
    QCOMPARE(d0.comments,          s1.comments);
    QCOMPARE(d0.voltage,           s1.voltage);
    QCOMPARE(d0.resistance,        s1.resistance);
    QCOMPARE(d0.power,             s1.power);
    QCOMPARE(d0.heatingTechnology, s1.heatingTechnology);
    QCOMPARE(d0.scores.value("Burn Taste"),              2.0);
    QCOMPARE(d0.scores.value("Flavor Intensity"),        5.0);
    QCOMPARE(d0.scores.value("Throat Irritation"),       3.0);
    QCOMPARE(d0.scores.value("Cough"),                   2.0);
    QCOMPARE(d0.scores.value("Performance Consistency"), 3.0);
    QCOMPARE(d0.scores.value("Vapor Volume"),            4.0);

    const DetailedSensorySample& d1 = decoded.samples[1];
    QCOMPARE(d1.name,              s2.name);
    QCOMPARE(d1.comments,          s2.comments);
    QCOMPARE(d1.heatingTechnology, s2.heatingTechnology);
    QCOMPARE(d1.scores.value("Burn Taste"),         6.0);
    QCOMPARE(d1.scores.value("Nasal Irritation"),   4.0);
    QCOMPARE(d1.scores.value("Vapor vs Oil"),       2.0);
    QCOMPARE(d1.scores.value("Volume Consistency"), 1.0);
}

void TstDetailedSensoryJson::deserializeClampsScoresToMetricRange()
{
    // The deserializer qBound()s every score into [1, metricMax]. Out-of-range
    // blob values (e.g. a corrupt recovery file) must be clamped, and a metric
    // missing from the blob defaults to 1.0.
    QJsonObject obj;
    obj["session_name"] = "Clamp";
    QJsonArray samples;
    QJsonObject s;
    s["name"]       = "D";
    s["Cough"]      = 99.0;   // max 4 -> clamp to 4
    s["Burn Taste"] = 0.0;    // below min -> clamp to 1
    // "Flavor Intensity" omitted -> defaults to 1.0
    samples.append(s);
    obj["samples"] = samples;

    const DetailedSensorySession decoded = detailedSensorySessionFromJson(obj);
    QCOMPARE(decoded.samples.size(), 1);
    QCOMPARE(decoded.samples[0].scores.value("Cough"),            4.0);
    QCOMPARE(decoded.samples[0].scores.value("Burn Taste"),       1.0);
    QCOMPARE(decoded.samples[0].scores.value("Flavor Intensity"), 1.0);
}

void TstDetailedSensoryJson::freshNewSessionIsPlaceholder()
{
    // Mirror DetailedSensoryPanel::newSession() EXACTLY: it leaves sessionName
    // empty (default-constructed, not "New Session"), stamps date + ISO8601
    // timestamp, and appends one default DetailedSensorySample{}. That fresh
    // card must read as a placeholder -- otherwise every brand-new session
    // would be force-written into the recovery snapshot.
    DetailedSensorySession s;
    s.date      = "2026-06-04";                 // QDate::currentDate() in prod
    s.timestamp = "2026-06-04T00:00:00Z";       // currentDateTimeUtc() in prod
    s.samples.append(DetailedSensorySample{});  // newSession() seeds one sample
    QVERIFY(s.sessionName.isEmpty());           // guards against drift from prod
    QVERIFY(isPlaceholderSession(s));
}

void TstDetailedSensoryJson::emptyNameIsPlaceholder()
{
    DetailedSensorySession s;
    s.sessionName = QString();  // empty also counts as placeholder
    QVERIFY(isPlaceholderSession(s));
}

void TstDetailedSensoryJson::typedTitleMakesItNonPlaceholder()
{
    DetailedSensorySession s;
    s.sessionName = "New Session";
    s.testTitle   = "Strawberry vs Mint";  // user typed a header field
    QVERIFY(!isPlaceholderSession(s));
}

void TstDetailedSensoryJson::typedSampleNameMakesItNonPlaceholder()
{
    DetailedSensorySession s;
    s.sessionName = "New Session";
    DetailedSensorySample sample;
    sample.name = "DeviceA";
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstDetailedSensoryJson::typedHeaderFieldMakesItNonPlaceholder()
{
    DetailedSensorySession s;
    s.sessionName = "New Session";
    s.testerName  = "Alice";
    QVERIFY(!isPlaceholderSession(s));

    DetailedSensorySession s2;
    s2.sessionName  = "New Session";
    s2.assessorName = "Reviewer";
    QVERIFY(!isPlaceholderSession(s2));

    DetailedSensorySession s3;
    s3.sessionName = "New Session";
    s3.media       = "Tobacco";
    QVERIFY(!isPlaceholderSession(s3));
}

void TstDetailedSensoryJson::typedSessionLevelFieldMakesItNonPlaceholder()
{
    // Regression guard for the fix that extended the predicate beyond the four
    // header fields. A session whose ONLY content is a session-level S2-1 extra
    // -- here just the facilitator comment, with an empty name and no samples
    // -- must NOT be treated as a placeholder, or the user's note would be
    // silently dropped from recovery.
    DetailedSensorySession s;
    s.facilitatorComment = "needs a re-test next week";
    QVERIFY(!isPlaceholderSession(s));

    // The non-string defaults matter too: toggling only the clog flag away from
    // its false default is real content.
    DetailedSensorySession s2;
    s2.clog = true;
    QVERIFY(!isPlaceholderSession(s2));
}

void TstDetailedSensoryJson::renamedSessionIsNonPlaceholder()
{
    // Once the user renames the session, it is no longer a placeholder even
    // with empty content.
    DetailedSensorySession s;
    s.sessionName = "Detailed run 2026-06-04";
    QVERIFY(!isPlaceholderSession(s));
}

QTEST_MAIN(TstDetailedSensoryJson)
#include "tst_detailedsensoryjson.moc"
