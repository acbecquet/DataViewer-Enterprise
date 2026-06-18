#include <QtTest/QtTest>

#include "../../src/database/CompatClassifier.h"

using namespace DVE;

// Pure-function unit tests for the SP4 triage classifier (A3). The DB-browser
// version/health filter buckets rows by classifyEra()/*Health() output, so green
// coverage here pins the era boundaries + health rules the UI depends on.
class TstCompatClassifier : public QObject
{
    Q_OBJECT

private slots:
    void era_stampedVersion_exactBucket();
    void era_stampedWithPrefixAndV();
    void era_unstamped_inferredFromDate_approx();
    void era_unstamped_boundaryOnReleaseDate();
    void era_unstamped_invalidDate_unstampedBucket();
    void era_unstamped_predatesV2_unstampedBucket();
    void health_sensory_legacyString();
    void health_sensory_junkPlaceholderOrZero();
    void health_sensory_healthyIsEmpty();
    void health_file_noSamples();
    void health_file_missingRegimes();
    void health_file_healthyIsEmpty();
};

void TstCompatClassifier::era_stampedVersion_exactBucket()
{
    CompatClass c = classifyEra(QStringLiteral("DataViewer/2.2.5"), QDate(2099, 1, 1));
    QCOMPARE(c.eraLabel, QStringLiteral("v2.2.x"));
    QVERIFY(!c.approx);   // stamped -> exact, never approximate
    QCOMPARE(classifyEra(QStringLiteral("DataViewer/2.0.10"), QDate()).eraLabel,
             QStringLiteral("v2.0.x"));
    QCOMPARE(classifyEra(QStringLiteral("DataViewer/2.4.1"), QDate()).eraLabel,
             QStringLiteral("v2.4.x"));
}

void TstCompatClassifier::era_stampedWithPrefixAndV()
{
    QCOMPARE(classifyEra(QStringLiteral("v2.3.1"), QDate()).eraLabel, QStringLiteral("v2.3.x"));
    QCOMPARE(classifyEra(QStringLiteral("2.1.0"), QDate()).eraLabel,  QStringLiteral("v2.1.x"));
}

void TstCompatClassifier::era_unstamped_inferredFromDate_approx()
{
    // 2026-05-28: after v2.1 start (05-27), before v2.2 start (05-29) -> v2.1.x (approx)
    CompatClass c = classifyEra(QString(), QDate(2026, 5, 28));
    QCOMPARE(c.eraLabel, QStringLiteral("v2.1.x"));
    QVERIFY(c.approx);
    // 2026-06-06: after v2.3 start (06-05) -> v2.3.x (approx)
    CompatClass d = classifyEra(QString(), QDate(2026, 6, 6));
    QCOMPARE(d.eraLabel, QStringLiteral("v2.3.x"));
    QVERIFY(d.approx);
}

void TstCompatClassifier::era_unstamped_boundaryOnReleaseDate()
{
    // creation exactly on the v2.2 start date buckets to v2.2.x
    CompatClass c = classifyEra(QString(), QDate(2026, 5, 29));
    QCOMPARE(c.eraLabel, QStringLiteral("v2.2.x"));
    QVERIFY(c.approx);
}

void TstCompatClassifier::era_unstamped_invalidDate_unstampedBucket()
{
    CompatClass c = classifyEra(QString(), QDate());   // null/invalid date
    QCOMPARE(c.eraLabel, unstampedEraLabel());
    QVERIFY(!c.approx);
}

void TstCompatClassifier::era_unstamped_predatesV2_unstampedBucket()
{
    CompatClass c = classifyEra(QString(), QDate(2026, 1, 1));  // before v2.0.0 (05-13)
    QCOMPARE(c.eraLabel, unstampedEraLabel());
}

void TstCompatClassifier::health_sensory_legacyString()
{
    QStringList h = sensoryHealth(/*legacy*/ true, /*placeholder*/ false, /*samples*/ 1);
    QVERIFY(h.contains(QStringLiteral("Legacy string scores")));
    QVERIFY(!h.contains(QStringLiteral("Junk candidate")));
}

void TstCompatClassifier::health_sensory_junkPlaceholderOrZero()
{
    QVERIFY(sensoryHealth(false, true, 1).contains(QStringLiteral("Junk candidate")));   // placeholder
    QVERIFY(sensoryHealth(false, false, 0).contains(QStringLiteral("Junk candidate")));  // zero samples
}

void TstCompatClassifier::health_sensory_healthyIsEmpty()
{
    QVERIFY(sensoryHealth(false, false, 3).isEmpty());
}

void TstCompatClassifier::health_file_noSamples()
{
    QCOMPARE(fileHealth(0, false), QStringList{ QStringLiteral("No samples") });
}

void TstCompatClassifier::health_file_missingRegimes()
{
    QStringList h = fileHealth(6, true);
    QVERIFY(h.contains(QStringLiteral("Missing puff regimes")));
    QVERIFY(!h.contains(QStringLiteral("No samples")));
}

void TstCompatClassifier::health_file_healthyIsEmpty()
{
    QVERIFY(fileHealth(6, false).isEmpty());
}

QTEST_MAIN(TstCompatClassifier)
#include "tst_compatclassifier.moc"
