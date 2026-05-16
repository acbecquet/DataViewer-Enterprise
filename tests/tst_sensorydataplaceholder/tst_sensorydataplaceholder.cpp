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

QTEST_MAIN(TstSensoryDataPlaceholder)
#include "tst_sensorydataplaceholder.moc"
