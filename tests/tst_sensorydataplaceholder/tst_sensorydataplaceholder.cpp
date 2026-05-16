#include <QtTest>
#include "pipeline/SensoryData.h"

using namespace DVE;

class TstSensoryDataPlaceholder : public QObject
{
    Q_OBJECT
private slots:
    void defaultNewSessionIsPlaceholder();
    void typedNameMakesItNonPlaceholder();
    void typedTesterMakesItNonPlaceholder();
    void typedSampleNameMakesItNonPlaceholder();
    void renamedSessionIsNonPlaceholder();
    void persistedSessionNeverPlaceholder();
    void typedCommentMakesItNonPlaceholder();
    void voltageMakesItNonPlaceholder();
    void resistanceMakesItNonPlaceholder();
    void heatingTechMakesItNonPlaceholder();
    void adjustedScoreMakesItNonPlaceholder();
    void controlFieldMakesItNonPlaceholder();
    void imageAttachedMakesItNonPlaceholder();
    void primaryDifferencesMakesItNonPlaceholder();
    void multiSampleSecondHasContent();
    void typedAssessorMakesItNonPlaceholder();
    void typedMediaMakesItNonPlaceholder();
    void changedPowerTypeMakesItNonPlaceholder();
    void changedPuffLengthMakesItNonPlaceholder();
};

void TstSensoryDataPlaceholder::defaultNewSessionIsPlaceholder()
{
    // Mirrors what SensoryPanel::newSession() produces: default
    // name, today's date+timestamp, plus one empty SampleCard.
    SensorySession s;
    s.sessionName = "New Session";
    s.date        = "2026-05-15";
    s.timestamp   = "2026-05-15T00:00:00Z";
    s.samples.append(SensorySample{});  // empty default card
    QVERIFY(isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedNameMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.testTitle = "Strawberry vs Mint";  // user typed something
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedTesterMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.testerName = "Alice";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedSampleNameMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.name = "DeviceA";
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::renamedSessionIsNonPlaceholder()
{
    // Once the user renames the session, it's no longer a placeholder
    // regardless of content emptiness.
    SensorySession s;
    s.sessionName = "TPM run 2026-05-15";
    s.samples.append(SensorySample{});
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::persistedSessionNeverPlaceholder()
{
    // Issue 2: once a session has an id, the gate must let it through
    // even if the user blanks out the content -- otherwise NOTIFY would
    // silently roll back the blank state.
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.id = 42;
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedCommentMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.comments = "tasted like burnt rubber";
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::voltageMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.voltage = 3.7;
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::resistanceMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.resistance = 1.2;
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::heatingTechMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.heatingTechnology = "Ceramic Coil";
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::adjustedScoreMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.scores["Burnt Taste"] = 7.0;
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::controlFieldMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.control = "DeviceX baseline";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::imageAttachedMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.imagePaths.append("C:/scans/form.png");
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::primaryDifferencesMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.primaryDifferences = "flavor density";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::multiSampleSecondHasContent()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});      // empty
    SensorySample second;
    second.name = "DeviceB";
    s.samples.append(second);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedAssessorMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.assessorName = "Reviewer 3";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::typedMediaMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    s.samples.append(SensorySample{});
    s.media = "Tobacco";
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::changedPowerTypeMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.powerType = "Variable Power";  // changed from default
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

void TstSensoryDataPlaceholder::changedPuffLengthMakesItNonPlaceholder()
{
    SensorySession s;
    s.sessionName = "New Session";
    SensorySample sample;
    sample.puffLengthSec = 5.0;  // changed from default
    s.samples.append(sample);
    QVERIFY(!isPlaceholderSession(s));
}

QTEST_MAIN(TstSensoryDataPlaceholder)
#include "tst_sensorydataplaceholder.moc"
