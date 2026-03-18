#include "TpmCalculator.h"

#include <cmath>
#include <limits>

namespace DVE {

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
static inline bool isFiniteValue(double v)
{
    return std::isfinite(v);
}

// ---------------------------------------------------------------------------
// calcIntervals
// ---------------------------------------------------------------------------
QVector<double> TpmCalculator::calcIntervals(const QVector<double>& puffs)
{
    const int n = puffs.size();
    QVector<double> intervals(n, 0.0);

    if (n == 0)
        return intervals;

    // Row 0: use the cumulative value directly (default 10 if it is 0)
    double first = puffs[0];
    intervals[0] = (first > 0.0) ? first : 10.0;

    double lastValidInterval = intervals[0];

    for (int i = 1; i < n; ++i) {
        double delta = puffs[i] - puffs[i - 1];
        if (delta > 0.0) {
            intervals[i]       = delta;
            lastValidInterval  = delta;
        } else {
            // Reuse last valid interval; lastValidInterval was seeded from row 0 so
            // it is always at least 10.
            intervals[i] = lastValidInterval;
        }
    }

    return intervals;
}

// ---------------------------------------------------------------------------
// calculate
// ---------------------------------------------------------------------------
QVector<double> TpmCalculator::calculate(
    const QVector<double>& puffs,
    const QVector<double>& beforeWeights,
    const QVector<double>& afterWeights)
{
    const int n = puffs.size();
    QVector<double> result(n, 0.0);

    if (n == 0)
        return result;

    // Ensure all input vectors are the same length; if not, work with what we have.
    const int nBefore = beforeWeights.size();
    const int nAfter  = afterWeights.size();

    QVector<double> intervals = calcIntervals(puffs);

    for (int i = 0; i < n; ++i) {
        double interval = intervals[i];
        if (interval <= 0.0)
            continue;

        double before = (i < nBefore) ? beforeWeights[i] : 0.0;
        double after  = (i < nAfter)  ? afterWeights[i]  : 0.0;

        double tpm = (before - after) * 1000.0 / interval;
        result[i] = tpm;
    }

    return result;
}

// ---------------------------------------------------------------------------
// average
// ---------------------------------------------------------------------------
double TpmCalculator::average(const QVector<double>& tpmValues)
{
    double sum   = 0.0;
    int    count = 0;

    for (double v : tpmValues) {
        if (isFiniteValue(v)) {
            sum += v;
            ++count;
        }
    }

    if (count < 2)
        return 0.0;

    return sum / static_cast<double>(count);
}

// ---------------------------------------------------------------------------
// stddev  (population standard deviation)
// ---------------------------------------------------------------------------
double TpmCalculator::stddev(const QVector<double>& tpmValues)
{
    double sum   = 0.0;
    int    count = 0;

    for (double v : tpmValues) {
        if (isFiniteValue(v)) {
            sum += v;
            ++count;
        }
    }

    if (count < 2)
        return 0.0;

    double mean   = sum / static_cast<double>(count);
    double sqSum  = 0.0;

    for (double v : tpmValues) {
        if (isFiniteValue(v)) {
            double diff = v - mean;
            sqSum += diff * diff;
        }
    }

    return std::sqrt(sqSum / static_cast<double>(count));
}

// ---------------------------------------------------------------------------
// powerDensity
// ---------------------------------------------------------------------------
QVector<double> TpmCalculator::powerDensity(const QVector<double>& tpmValues, double power)
{
    const int n = tpmValues.size();
    QVector<double> result(n, 0.0);

    if (power == 0.0)
        return result;

    for (int i = 0; i < n; ++i)
        result[i] = tpmValues[i] / power;

    return result;
}

// ---------------------------------------------------------------------------
// variation
// ---------------------------------------------------------------------------
QVector<double> TpmCalculator::variation(const QVector<double>& tpmValues)
{
    const int n = tpmValues.size();
    QVector<double> result(n, 0.0);

    if (n == 0)
        return result;

    double base = tpmValues[0];
    if (base == 0.0)
        return result;   // all zeros when base is zero

    result[0] = 0.0;
    for (int i = 1; i < n; ++i)
        result[i] = (tpmValues[i] - base) / base * 100.0;

    return result;
}

// ---------------------------------------------------------------------------
// normalizedTPM
// ---------------------------------------------------------------------------
double TpmCalculator::normalizedTPM(double avgTPM, double power)
{
    if (power == 0.0)
        return 0.0;
    return avgTPM / power;
}

} // namespace DVE
