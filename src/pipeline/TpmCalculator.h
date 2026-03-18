#pragma once
#include "ReportData.h"
#include <QVector>

namespace DVE {

class TpmCalculator {
public:
    // Calculate TPM (mg/puff) for each measurement interval.
    // puffs:          cumulative puff count per row
    // beforeWeights:  gram weight before the puffing interval
    // afterWeights:   gram weight after the puffing interval
    // Returns a vector of TPM values (mg/puff), same length as input.
    static QVector<double> calculate(
        const QVector<double>& puffs,
        const QVector<double>& beforeWeights,
        const QVector<double>& afterWeights);

    // Calculate per-interval puff counts from a cumulative puff-count vector.
    // Row 0: puffs[0] (or 10 if puffs[0] == 0)
    // Row i: puffs[i] - puffs[i-1]; if <= 0, reuse the previous valid interval
    //        (or default to 10 if no valid interval exists yet).
    static QVector<double> calcIntervals(const QVector<double>& puffs);

    // Arithmetic mean of tpmValues.  Ignores NaN / Inf.
    // Returns 0.0 if fewer than 2 finite values are present.
    static double average(const QVector<double>& tpmValues);

    // Population standard deviation of tpmValues.  Ignores NaN / Inf.
    // Returns 0.0 if fewer than 2 finite values are present.
    static double stddev(const QVector<double>& tpmValues);

    // Element-wise tpm[i] / power.  Returns 0 for every element when power == 0.
    static QVector<double> powerDensity(const QVector<double>& tpmValues, double power);

    // Percentage variation relative to the first element:
    //   result[0] = 0
    //   result[i] = (tpm[i] - tpm[0]) / tpm[0] * 100
    // If tpm[0] == 0 all elements are returned as 0.
    static QVector<double> variation(const QVector<double>& tpmValues);

    // Normalised TPM = avgTPM / power.  Returns 0 when power == 0.
    static double normalizedTPM(double avgTPM, double power);
};

} // namespace DVE
