#pragma once

#include "ReportData.h"
#include "TpmCalculator.h"
#include "../ExcelReader.h"

#include <QString>
#include <QVector>

namespace DVE {

// ---------------------------------------------------------------------------
// SheetProcessor — abstract base class
// ---------------------------------------------------------------------------
class SheetProcessor {
public:
    virtual ~SheetProcessor() = default;

    // Main entry point: process all raw samples for one sheet and return a
    // fully-populated SheetResult.
    virtual SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                                const QString& sheetName,
                                const QString& templateVersion) = 0;

    // Populate all calculated fields on an already-built SampleResult
    // (tpm, oilConsumed, variation, averageTPM, stdDevTPM, etc.).
    void calculateMetrics(SampleResult& s);

    // Populate sheet-level aggregates (overallAvgTPM, tpmTrend, puffCounts …)
    // after all SampleResults have been added to the sheet.
    void computeSheetAggregates(SheetResult& sheet);

    // When true, per-row column index 4 is read as a Puffing Regime string
    // (new template) instead of a Resistance double. Set by DataProcessor
    // from the column-E header before process() is called.
    void setPerRowRegime(bool v) { m_perRowRegime = v; }

protected:
    bool m_perRowRegime = false;

    // Build a SampleResult from one ExcelReader::SampleData entry.
    // sampleIndex is used for diagnostic messages only.
    SampleResult buildSampleResult(const ExcelReader::SampleData& raw,
                                   int sampleIndex,
                                   const QString& templateVersion);
};

// ---------------------------------------------------------------------------
// Concrete processors
// All of these share the same underlying data layout; they delegate to the
// base-class helpers and are distinguished only by their factory key.
// ---------------------------------------------------------------------------

class UserTestSimProcessor : public SheetProcessor {
public:
    SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                        const QString& sheetName,
                        const QString& templateVersion) override;
};

class LongPuffProcessor : public SheetProcessor {
public:
    SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                        const QString& sheetName,
                        const QString& templateVersion) override;
};

class RapidPuffProcessor : public SheetProcessor {
public:
    SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                        const QString& sheetName,
                        const QString& templateVersion) override;
};

class TempCyclingProcessor : public SheetProcessor {
public:
    SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                        const QString& sheetName,
                        const QString& templateVersion) override;
};

class DeviceLifeProcessor : public SheetProcessor {
public:
    SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                        const QString& sheetName,
                        const QString& templateVersion) override;
};

class ExtendedTestProcessor : public SheetProcessor {
public:
    SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                        const QString& sheetName,
                        const QString& templateVersion) override;
};

class GenericSheetProcessor : public SheetProcessor {
public:
    SheetResult process(const QVector<ExcelReader::SampleData>& rawSamples,
                        const QString& sheetName,
                        const QString& templateVersion) override;
};

// ---------------------------------------------------------------------------
// Factory: returns a heap-allocated processor for the given sheet name.
// Caller takes ownership.
// ---------------------------------------------------------------------------
SheetProcessor* createProcessor(const QString& sheetName);

} // namespace DVE
