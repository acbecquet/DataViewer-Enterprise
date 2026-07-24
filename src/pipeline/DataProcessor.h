#pragma once

#include "ReportData.h"
#include "../ExcelReader.h"

#include <QString>
#include <functional>

namespace DVE {

class DataProcessor {
public:
    DataProcessor();
    ~DataProcessor();

    // Load an Excel file, iterate over every sheet, run the appropriate
    // SheetProcessor, and return the aggregated FileResult.
    //
    // Production read path: extraction runs through model::SchemaDrivenReader
    // + model::LegacyAdapter (schema-driven). Byte-identical to the legacy
    // positional path below, proven by the Phase-1 shadow gate (tst_v3shadow).
    //
    // progressCallback(percent, message) is invoked at key milestones.
    // Pass nullptr (or omit the argument) if progress reporting is not needed.
    FileResult processFile(
        const QString& filePath,
        std::function<void(int, const QString&)> progressCallback = nullptr);

    // Process a single sheet that has already been selected on `reader`.
    // Returns a fully-populated SheetResult (or an empty one on error).
    SheetResult processSheet(ExcelReader& reader, const QString& sheetName);

    // Legacy positional read path - shadow-harness referee only. Retained so the
    // Phase-1 shadow gate can diff production against the original positional
    // extraction; deleted in Phase 4 together with ExcelReader's positional
    // getters (getAllSamples / getSample / extractMetadata). Not used in
    // production - call processFile / processSheet instead.
    FileResult  processFileLegacy(
        const QString& filePath,
        std::function<void(int, const QString&)> progressCallback = nullptr);
    SheetResult processSheetLegacy(ExcelReader& reader, const QString& sheetName);

    // Human-readable description of the last error, or empty string on success.
    QString lastError() const { return m_lastError; }

private:
    QString m_lastError;

    // Shared raw-table (SOP / instruction sheet) branch used by both the legacy
    // and v3 read paths - logic lives here once.
    SheetResult processSopSheet(ExcelReader& reader, const QString& sheetName);

    void logDebug(const QString& msg) const;
    void setError(const QString& msg);
};

} // namespace DVE
