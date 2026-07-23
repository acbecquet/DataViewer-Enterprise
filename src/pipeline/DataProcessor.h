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
    // progressCallback(percent, message) is invoked at key milestones.
    // Pass nullptr (or omit the argument) if progress reporting is not needed.
    FileResult processFile(
        const QString& filePath,
        std::function<void(int, const QString&)> progressCallback = nullptr);

    // Process a single sheet that has already been selected on `reader`.
    // Returns a fully-populated SheetResult (or an empty one on error).
    SheetResult processSheet(ExcelReader& reader, const QString& sheetName);

    // v3 (Phase 1): schema-driven read path. Same output contract as
    // processFile/processSheet; extraction runs through model::SchemaDrivenReader
    // + model::LegacyAdapter instead of ExcelReader's positional getters.
    FileResult  processFileV3(
        const QString& filePath,
        std::function<void(int, const QString&)> progressCallback = nullptr);
    SheetResult processSheetV3(ExcelReader& reader, const QString& sheetName);

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
