#include "SheetProcessors.h"

#include <QDebug>
#include <cmath>

namespace DVE {

// ===========================================================================
// Internal helpers
// ===========================================================================

// Safe conversion of a QVariant to double (returns 0.0 on failure).
static double varToDouble(const QVariant& v)
{
    bool ok = false;
    double d = v.toDouble(&ok);
    return ok ? d : 0.0;
}

// Safe conversion of a QVariant to QString (trims whitespace).
static QString varToString(const QVariant& v)
{
    return v.toString().trimmed();
}

// Normalise a burn/clog/leak indicator string to "Y", "N", or "".
static QString normaliseBCL(const QString& raw)
{
    QString upper = raw.trimmed().toUpper();
    if (upper == "Y" || upper == "YES")
        return QStringLiteral("Y");
    if (upper == "N" || upper == "NO")
        return QStringLiteral("N");
    return QString();
}

// ---------------------------------------------------------------------------
// Column indices used when reading raw Excel data rows.
// These match the ExcelReader column layout described in the specification.
// ---------------------------------------------------------------------------
namespace ColIdx {
    constexpr int PUFFS         = 0;
    constexpr int BEFORE_WEIGHT = 1;
    constexpr int AFTER_WEIGHT  = 2;
    constexpr int DRAW_PRESSURE = 3;
    constexpr int RESISTANCE    = 4;
    constexpr int SMELL         = 5;
    constexpr int CLOG          = 6;
    constexpr int NOTES         = 7;
    // cols 8-11 are calculated — we may still read them from the raw data
    // but they will be overwritten by calculateMetrics().
    constexpr int BURN_OR_FLAG  = 10;  // position checked for "Burn/Clog/Leak" labels
}

// ===========================================================================
// SheetProcessor — protected base methods
// ===========================================================================

// ---------------------------------------------------------------------------
// buildSampleResult
// ---------------------------------------------------------------------------
SampleResult SheetProcessor::buildSampleResult(const ExcelReader::SampleData& raw,
                                                int sampleIndex,
                                                const QString& templateVersion)
{
    Q_UNUSED(sampleIndex)
    Q_UNUSED(templateVersion)

    SampleResult s;

    // --- copy metadata -------------------------------------------------
    const ExcelReader::SampleMetadata& meta = raw.metadata;
    // sampleID is the unique per-sample identifier; fall back to a generated name
    s.sampleName       = meta.sampleID.isEmpty()
                         ? QString("Sample %1").arg(sampleIndex + 1)
                         : meta.sampleID;
    s.sampleID         = meta.sampleID;
    s.date             = meta.date;
    s.tester           = meta.tester;
    s.media            = meta.media;
    s.viscosity        = meta.viscosity;
    s.resistance       = meta.resistance;
    s.voltage          = meta.voltage;
    s.power            = meta.power;
    s.heatingTechnology = meta.heatingTechnology;
    s.puffingRegime    = meta.puffingRegime;
    s.initialOilMass   = meta.initialOilMass;

    // Write provenance (Phase 2b): the reader's own block origin is ground
    // truth for where this sample's cells physically start.
    s.startColumn      = raw.startColumn;

    // --- build DataRow vector from raw Excel rows ----------------------
    s.rows.reserve(raw.dataRows.size());

    for (const QVector<QVariant>& rowData : raw.dataRows) {
        // Skip completely empty rows (all variants are invalid / null).
        bool hasData = false;
        for (const QVariant& cell : rowData) {
            if (cell.isValid() && !cell.isNull()) {
                hasData = true;
                break;
            }
        }
        if (!hasData)
            continue;

        DataRow dr;

        auto cell = [&](int col) -> QVariant {
            return (col < rowData.size()) ? rowData[col] : QVariant();
        };

        dr.puffs         = varToDouble(cell(ColIdx::PUFFS));
        dr.beforeWeight  = varToDouble(cell(ColIdx::BEFORE_WEIGHT));
        dr.afterWeight   = varToDouble(cell(ColIdx::AFTER_WEIGHT));
        dr.drawPressure  = varToDouble(cell(ColIdx::DRAW_PRESSURE));
        if (m_perRowRegime)
            dr.puffingRegime = varToString(cell(ColIdx::RESISTANCE));   // col 4 = regime string
        else
            dr.resistance    = varToDouble(cell(ColIdx::RESISTANCE));   // col 4 = resistance number
        dr.smell         = varToString(cell(ColIdx::SMELL));
        dr.clog          = varToString(cell(ColIdx::CLOG));
        dr.notes         = varToString(cell(ColIdx::NOTES));

        // Columns 8-11 will be recalculated; read them here for completeness
        // (they will be overwritten in calculateMetrics).
        dr.tpm             = varToDouble(cell(8));
        dr.tpmPowerDensity = varToDouble(cell(9));
        dr.variationTPM    = varToDouble(cell(ColIdx::BURN_OR_FLAG));
        dr.oilConsumed     = varToDouble(cell(11));

        s.rows.append(dr);
    }

    // --- fill in missing Before Weight values --------------------------
    // The Excel template uses formulas: beforeWeight[i] = afterWeight[i-1].
    // When the file was saved by openpyxl (New File, Edit Headers), cached
    // formula values are stripped, so data_only=True reads them as None/0.
    // Repair by applying the known pattern: if beforeWeight is 0 but the
    // previous row has a valid afterWeight, copy it forward.
    for (int i = 1; i < s.rows.size(); ++i) {
        if (s.rows[i].beforeWeight == 0.0 && s.rows[i - 1].afterWeight > 0.0) {
            s.rows[i].beforeWeight = s.rows[i - 1].afterWeight;
        }
    }

    // --- fill in missing Puffs values ----------------------------------
    // The Excel template uses formulas for the puffs column in rows 2+:
    // puffs[i] = puffs[i-1] + <interval>.  When cached values are stripped
    // (openpyxl save), data_only=True reads formula cells as None/0.
    // Repair: find the constant puff interval from the first two valid rows
    // and extrapolate forward for any zero-valued rows.
    //
    // W1 (2026-08-27): NOT when the reader detected that this sheet's formula
    // caches were stripped - then the zeros are destroyed data, the guessed
    // interval is meaningless (it once computed a diff across an unfilled
    // zero row and extrapolated +30 chains into the database), and
    // fabricating values hides the corruption. The sheet is flagged instead
    // (SheetResult::strippedFormulaCells) so the UI warns and blocks the DB
    // save. Healthy sheets (count 0) behave byte-identically to before.
    if (m_strippedFormulaCells == 0
        && !s.rows.isEmpty() && s.rows[0].puffs > 0.0) {
        double increment = s.rows[0].puffs;  // fallback: first row = interval
        for (int i = 1; i < s.rows.size(); ++i) {
            if (s.rows[i].puffs > 0.0) {
                double diff = s.rows[i].puffs - s.rows[i - 1].puffs;
                if (diff > 0.0)
                    increment = diff;
                break;
            }
        }
        for (int i = 1; i < s.rows.size(); ++i) {
            if (s.rows[i].puffs == 0.0)
                s.rows[i].puffs = s.rows[i - 1].puffs + increment;
        }
    }

    // --- burn / clog / leak detection ----------------------------------
    // Strategy 1: look for a dedicated column at index 10 whose header
    // contains "Burn", "Clog", or "Leak" and whose data cell holds "Y"/"N".
    // We check the first data row for these labels (the Excel sheet sometimes
    // places them inline rather than in a dedicated header row).
    //
    // Strategy 2: fall back to scanning the notes column for keywords.

    QString burnVal, clogVal, leakVal;

    for (const QVector<QVariant>& rowData : raw.dataRows) {
        if (rowData.size() <= ColIdx::BURN_OR_FLAG)
            continue;

        // Check columns 8, 9, 10 for inline Burn/Clog/Leak "Y"/"N" flags.
        // In several template versions these appear at fixed offset columns
        // relative to the sample's startColumn, numbered 8-10 within the sample.
        for (int flagCol = 8; flagCol <= 10 && flagCol < rowData.size(); ++flagCol) {
            QString cellStr = varToString(rowData[flagCol]).toUpper();
            if (cellStr == "Y" || cellStr == "N") {
                // Assign in order: first flag → burn, second → clog, third → leak
                if (burnVal.isEmpty())
                    burnVal = cellStr;
                else if (clogVal.isEmpty())
                    clogVal = cellStr;
                else if (leakVal.isEmpty())
                    leakVal = cellStr;
            }
        }
        if (!burnVal.isEmpty())
            break; // found flags in this row — done
    }

    // Strategy 2: scan notes for keywords if strategy 1 found nothing.
    if (burnVal.isEmpty()) {
        for (const DataRow& dr : s.rows) {
            QString notesUpper = dr.notes.toUpper();
            if (notesUpper.contains("BURN")) {
                burnVal = notesUpper.contains("NO BURN") ? "N" : "Y";
            }
            if (notesUpper.contains("CLOG")) {
                clogVal = notesUpper.contains("NO CLOG") ? "N" : "Y";
            }
            if (notesUpper.contains("LEAK")) {
                leakVal = notesUpper.contains("NO LEAK") ? "N" : "Y";
            }
        }
    }

    s.burnStatus = normaliseBCL(burnVal);
    s.clogStatus = normaliseBCL(clogVal);
    s.leakStatus = normaliseBCL(leakVal);

    return s;
}

// ---------------------------------------------------------------------------
// calculateMetrics
// ---------------------------------------------------------------------------
void SheetProcessor::calculateMetrics(SampleResult& s)
{
    const int n = s.rows.size();
    if (n == 0)
        return;

    // Extract raw vectors.
    QVector<double> puffs(n), before(n), after(n);
    for (int i = 0; i < n; ++i) {
        puffs[i]  = s.rows[i].puffs;
        before[i] = s.rows[i].beforeWeight;
        after[i]  = s.rows[i].afterWeight;
    }

    // Calculate per-interval TPM.
    QVector<double> tpmVec = TpmCalculator::calculate(puffs, before, after);

    // Power density and variation.
    QVector<double> pdVec  = TpmCalculator::powerDensity(tpmVec, s.power);
    QVector<double> varVec = TpmCalculator::variation(tpmVec);

    // Oil consumed per interval (g). The before/after weights are already
    // grams, so their difference is grams and needs no scaling.
    //
    // GRAMS IS CANONICAL for every oil quantity in this codebase. The template
    // agrees: its column header is "OilCum(g)" and its formula is
    // =(TPM*puffs)/1000, i.e. it divides milligrams down to grams itself.
    // MetricRegistry declares oil_consumed and initial_oil_mass with unit "g".
    // This function used to be the sole dissenter - it scaled up to milligrams,
    // which forced every display site to divide by 1000 again on the way out.
    QVector<double> oilVec(n);
    for (int i = 0; i < n; ++i)
        oilVec[i] = before[i] - after[i];

    // Write back into DataRows — oil consumed is cumulative (running sum).
    // Only accumulate oil for rows that have actual weight data; empty rows
    // (beforeWeight == 0 || afterWeight == 0) get oilConsumed = 0.
    double cumOil = 0.0;
    for (int i = 0; i < n; ++i) {
        bool hasData = (before[i] > 0.0 && after[i] > 0.0);
        s.rows[i].tpm             = hasData ? ((tpmVec[i] > 100.0) ? 0.0 : tpmVec[i]) : 0.0;
        s.rows[i].tpmPowerDensity = hasData ? pdVec[i]  : 0.0;
        s.rows[i].variationTPM    = hasData ? varVec[i] : 0.0;
        if (hasData) {
            cumOil += oilVec[i];
            // 10 g is the same physical sanity bound the old 10000 mg literal
            // expressed - only the unit changed, not the threshold.
            s.rows[i].oilConsumed = (cumOil > 10.0) ? 0.0 : cumOil;
        } else {
            s.rows[i].oilConsumed = 0.0;
        }
    }

    // Sample-level aggregates — only include rows with actual weight data.
    QVector<double> validTPM;
    for (int i = 0; i < n; ++i) {
        if (before[i] > 0.0 && after[i] > 0.0)
            validTPM.append((tpmVec[i] > 100.0) ? 0.0 : tpmVec[i]);
    }

    s.averageTPM          = TpmCalculator::average(validTPM);
    s.stdDevTPM           = TpmCalculator::stddev(validTPM);
    s.normalizedTPM       = TpmCalculator::normalizedTPM(s.averageTPM, s.power);

    // Average power density (reuse the powerDensity helper on the scalar).
    if (s.power != 0.0)
        s.averagePowerDensity = s.averageTPM / s.power;
    else
        s.averagePowerDensity = 0.0;

    // Total puffs: last row with actual data.
    s.totalPuffs = 0;
    for (int i = n - 1; i >= 0; --i) {
        if (before[i] > 0.0 && after[i] > 0.0 && s.rows[i].puffs > 0.0) {
            s.totalPuffs = static_cast<int>(s.rows[i].puffs);
            break;
        }
    }

    // Total oil consumed (sum of valid intervals only, g).
    double oilSum = 0.0;
    for (int i = 0; i < n; ++i) {
        if (before[i] > 0.0 && after[i] > 0.0)
            oilSum += oilVec[i];
    }
    s.totalOilConsumed = oilSum;

    // Efficiency: (totalOilConsumed / initialOilMass) * 100, clamped 0-100.
    // Both operands are grams, so the ratio is dimensionless and no scale
    // factor belongs here. (The old form divided by initialOilMass * 1000 to
    // cancel this function's own milligram scaling; both sides now agree.)
    if (s.initialOilMass > 0.0) {
        double eff = (s.totalOilConsumed / s.initialOilMass) * 100.0;
        s.efficiencyPercent = std::max(0.0, std::min(100.0, eff));
    } else {
        s.efficiencyPercent = 0.0;
    }
}

// ---------------------------------------------------------------------------
// computeSheetAggregates
// ---------------------------------------------------------------------------
void SheetProcessor::computeSheetAggregates(SheetResult& sheet)
{
    const int nSamples = sheet.samples.size();
    if (nSamples == 0)
        return;

    // Collect per-sample average TPM values.
    QVector<double> avgTPMs;
    avgTPMs.reserve(nSamples);
    for (const SampleResult& sr : sheet.samples)
        avgTPMs.append(sr.averageTPM);

    sheet.overallAvgTPM    = TpmCalculator::average(avgTPMs);
    sheet.overallStdDevTPM = TpmCalculator::stddev(avgTPMs);

    // Build concatenated trend vectors for plotting — skip empty rows.
    sheet.tpmTrend.clear();
    sheet.puffCounts.clear();

    for (const SampleResult& sr : sheet.samples) {
        for (const DataRow& dr : sr.rows) {
            if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
            sheet.tpmTrend.append(dr.tpm);
            sheet.puffCounts.append(dr.puffs);
        }
    }
}

// ===========================================================================
// Generic processor — shared implementation used by all concrete processors
// ===========================================================================

SheetResult GenericSheetProcessor::process(
    const QVector<ExcelReader::SampleData>& rawSamples,
    const QString& sheetName,
    const QString& templateVersion)
{
    SheetResult sheet;
    sheet.sheetName       = sheetName;
    sheet.templateVersion = templateVersion;
    sheet.hasPerRowRegime = m_perRowRegime;
    sheet.samples.reserve(rawSamples.size());

    for (int i = 0; i < rawSamples.size(); ++i) {
        SampleResult sr = buildSampleResult(rawSamples[i], i, templateVersion);
        calculateMetrics(sr);
        sheet.samples.append(sr);
    }

    computeSheetAggregates(sheet);
    return sheet;
}

// ===========================================================================
// Concrete processors — all delegate to the generic implementation.
// They are kept as distinct classes so that future specialisation is
// straightforward (e.g. different column layouts, extra metrics).
// ===========================================================================

SheetResult UserTestSimProcessor::process(
    const QVector<ExcelReader::SampleData>& rawSamples,
    const QString& sheetName,
    const QString& templateVersion)
{
    GenericSheetProcessor generic;
    generic.setPerRowRegime(m_perRowRegime);
    return generic.process(rawSamples, sheetName, templateVersion);
}

SheetResult LongPuffProcessor::process(
    const QVector<ExcelReader::SampleData>& rawSamples,
    const QString& sheetName,
    const QString& templateVersion)
{
    GenericSheetProcessor generic;
    generic.setPerRowRegime(m_perRowRegime);
    return generic.process(rawSamples, sheetName, templateVersion);
}

SheetResult RapidPuffProcessor::process(
    const QVector<ExcelReader::SampleData>& rawSamples,
    const QString& sheetName,
    const QString& templateVersion)
{
    GenericSheetProcessor generic;
    generic.setPerRowRegime(m_perRowRegime);
    return generic.process(rawSamples, sheetName, templateVersion);
}

SheetResult TempCyclingProcessor::process(
    const QVector<ExcelReader::SampleData>& rawSamples,
    const QString& sheetName,
    const QString& templateVersion)
{
    GenericSheetProcessor generic;
    generic.setPerRowRegime(m_perRowRegime);
    return generic.process(rawSamples, sheetName, templateVersion);
}

SheetResult DeviceLifeProcessor::process(
    const QVector<ExcelReader::SampleData>& rawSamples,
    const QString& sheetName,
    const QString& templateVersion)
{
    GenericSheetProcessor generic;
    generic.setPerRowRegime(m_perRowRegime);
    return generic.process(rawSamples, sheetName, templateVersion);
}

SheetResult ExtendedTestProcessor::process(
    const QVector<ExcelReader::SampleData>& rawSamples,
    const QString& sheetName,
    const QString& templateVersion)
{
    GenericSheetProcessor generic;
    generic.setPerRowRegime(m_perRowRegime);
    return generic.process(rawSamples, sheetName, templateVersion);
}

// ===========================================================================
// Factory
// ===========================================================================

SheetProcessor* createProcessor(const QString& sheetName)
{
    // Exact matches first, then substring checks.
    if (sheetName == QStringLiteral("User Test Simulation"))
        return new UserTestSimProcessor();

    if (sheetName == QStringLiteral("Long Puff lifetime Test") ||
        sheetName == QStringLiteral("Long Puff Lifetime Test"))
        return new LongPuffProcessor();

    if (sheetName == QStringLiteral("Rapid Puff Lifetime Test"))
        return new RapidPuffProcessor();

    if (sheetName.contains(QStringLiteral("Temperature Cycling"), Qt::CaseInsensitive))
        return new TempCyclingProcessor();

    if (sheetName.contains(QStringLiteral("Device Life"), Qt::CaseInsensitive))
        return new DeviceLifeProcessor();

    if (sheetName.contains(QStringLiteral("Extended Test"), Qt::CaseInsensitive))
        return new ExtendedTestProcessor();

    return new GenericSheetProcessor();
}

} // namespace DVE
