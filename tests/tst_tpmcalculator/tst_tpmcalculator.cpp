#include <QtTest>
#include <QVector>
#include <cmath>
#include "TpmCalculator.h"
#include "TestHelpers.h"

class TestTpmCalculator : public QObject
{
    Q_OBJECT

private slots:
    void testCalculate();
    void testCalcIntervals();
    void testAverage();
    void testStddev();
    void testPowerDensity();
    void testVariation();
    void testNormalizedTPM();
    void testCalculateEmptyInputs();
};

void TestTpmCalculator::testCalculate()
{
    QVector<double> puffs     = {10, 20, 30, 40, 50};
    QVector<double> before    = {1.000, 1.000, 1.000, 1.000, 1.000};
    QVector<double> after     = {0.965, 0.965, 0.965, 0.965, 0.965};

    // weight drop = 0.035, interval = 10 each, TPM = 0.035*1000/10 = 3.5
    QVector<double> result = DVE::TpmCalculator::calculate(puffs, before, after);
    QCOMPARE(result.size(), 5);
    for (int i = 0; i < result.size(); ++i) {
        QFUZZY_COMPARE(result[i], 3.5);
    }
}

void TestTpmCalculator::testCalcIntervals()
{
    // Normal case: [10,20,30] -> [10,10,10]
    {
        QVector<double> puffs = {10, 20, 30};
        QVector<double> intervals = DVE::TpmCalculator::calcIntervals(puffs);
        QCOMPARE(intervals.size(), 3);
        QFUZZY_COMPARE(intervals[0], 10.0);
        QFUZZY_COMPARE(intervals[1], 10.0);
        QFUZZY_COMPARE(intervals[2], 10.0);
    }

    // First puff is 0 -> row 0 gets default 10
    {
        QVector<double> puffs = {0, 20, 30};
        QVector<double> intervals = DVE::TpmCalculator::calcIntervals(puffs);
        QCOMPARE(intervals.size(), 3);
        QFUZZY_COMPARE(intervals[0], 10.0);
        QFUZZY_COMPARE(intervals[1], 20.0);
        QFUZZY_COMPARE(intervals[2], 10.0);
    }

    // Single puff
    {
        QVector<double> puffs = {5};
        QVector<double> intervals = DVE::TpmCalculator::calcIntervals(puffs);
        QCOMPARE(intervals.size(), 1);
        QFUZZY_COMPARE(intervals[0], 5.0);
    }
}

void TestTpmCalculator::testAverage()
{
    // Normal case
    {
        QVector<double> vals = {2.0, 4.0, 6.0};
        QFUZZY_COMPARE(DVE::TpmCalculator::average(vals), 4.0);
    }

    // Empty -> 0
    {
        QVector<double> vals;
        QFUZZY_COMPARE(DVE::TpmCalculator::average(vals), 0.0);
    }

    // Single value -> 0 (less than 2 finite values)
    {
        QVector<double> vals = {5.0};
        QFUZZY_COMPARE(DVE::TpmCalculator::average(vals), 0.0);
    }

    // Mixed with NaN/Inf — should ignore them
    {
        QVector<double> vals = {2.0, std::numeric_limits<double>::quiet_NaN(),
                                4.0, std::numeric_limits<double>::infinity()};
        // Two finite values: 2.0 and 4.0, mean = 3.0
        QFUZZY_COMPARE(DVE::TpmCalculator::average(vals), 3.0);
    }
}

void TestTpmCalculator::testStddev()
{
    // Known values: [2, 4, 4, 4, 5, 5, 7, 9] -> population stddev = 2.0
    {
        QVector<double> vals = {2, 4, 4, 4, 5, 5, 7, 9};
        QFUZZY_COMPARE(DVE::TpmCalculator::stddev(vals), 2.0);
    }

    // Empty -> 0
    {
        QVector<double> vals;
        QFUZZY_COMPARE(DVE::TpmCalculator::stddev(vals), 0.0);
    }

    // Uniform -> 0
    {
        QVector<double> vals = {3.0, 3.0, 3.0, 3.0};
        QFUZZY_COMPARE(DVE::TpmCalculator::stddev(vals), 0.0);
    }
}

void TestTpmCalculator::testPowerDensity()
{
    // Normal case
    {
        QVector<double> tpm = {10.0, 20.0, 30.0};
        double power = 5.0;
        QVector<double> result = DVE::TpmCalculator::powerDensity(tpm, power);
        QCOMPARE(result.size(), 3);
        QFUZZY_COMPARE(result[0], 2.0);
        QFUZZY_COMPARE(result[1], 4.0);
        QFUZZY_COMPARE(result[2], 6.0);
    }

    // Power = 0 -> all zeros
    {
        QVector<double> tpm = {10.0, 20.0};
        QVector<double> result = DVE::TpmCalculator::powerDensity(tpm, 0.0);
        QCOMPARE(result.size(), 2);
        QFUZZY_COMPARE(result[0], 0.0);
        QFUZZY_COMPARE(result[1], 0.0);
    }
}

void TestTpmCalculator::testVariation()
{
    // Normal case: [2,3,4] -> [0, 50, 100]
    {
        QVector<double> tpm = {2.0, 3.0, 4.0};
        QVector<double> result = DVE::TpmCalculator::variation(tpm);
        QCOMPARE(result.size(), 3);
        QFUZZY_COMPARE(result[0], 0.0);
        QFUZZY_COMPARE(result[1], 50.0);
        QFUZZY_COMPARE(result[2], 100.0);
    }

    // First element is 0 -> all zeros (division by zero protection)
    {
        QVector<double> tpm = {0.0, 1.0, 2.0};
        QVector<double> result = DVE::TpmCalculator::variation(tpm);
        QCOMPARE(result.size(), 3);
        QFUZZY_COMPARE(result[0], 0.0);
        QFUZZY_COMPARE(result[1], 0.0);
        QFUZZY_COMPARE(result[2], 0.0);
    }

    // Empty
    {
        QVector<double> tpm;
        QVector<double> result = DVE::TpmCalculator::variation(tpm);
        QCOMPARE(result.size(), 0);
    }
}

void TestTpmCalculator::testNormalizedTPM()
{
    // Normal case
    QFUZZY_COMPARE(DVE::TpmCalculator::normalizedTPM(10.0, 5.0), 2.0);

    // Power = 0 -> 0
    QFUZZY_COMPARE(DVE::TpmCalculator::normalizedTPM(10.0, 0.0), 0.0);
}

void TestTpmCalculator::testCalculateEmptyInputs()
{
    QVector<double> empty;
    QVector<double> result = DVE::TpmCalculator::calculate(empty, empty, empty);
    QCOMPARE(result.size(), 0);
}

QTEST_APPLESS_MAIN(TestTpmCalculator)
#include "tst_tpmcalculator.moc"
