#pragma once

#include "ReportData.h"
#include "../ExcelReader.h"
#include "../model/Manifest.h"

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
    // `manifest` is the workbook's parsed _dve_schema (the resolver's rung 1);
    // nullptr = no manifest, today's standard/inference ladder.
    SheetResult processSheet(ExcelReader& reader, const QString& sheetName,
                             const model::Manifest::ParseResult* manifest = nullptr);

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

    // True when the most recent processFile() routed ANY sheet through the
    // header-driven schema-inference path (non-standard block layout). Reset at
    // the start of every processFile(); processFileLegacy() never infers, so it
    // always leaves this false. The shadow harness consumes this to skip the
    // legacy-identity assertion on inference-path files (legacy is known-wrong
    // there); correctness is gated by the inference E2E value tests instead.
    bool lastFileUsedInference() const { return m_lastUsedInference; }

    // True when the most recent processFile() found (and parsed) a _dve_schema
    // manifest sheet in the workbook. Reset at the start of every
    // processFile(); processFileLegacy() never parses manifests, so it always
    // leaves this false. Same consumer as lastFileUsedInference(): the shadow
    // harness skips the legacy-identity assertion on manifest workbooks (the
    // legacy referee parses their sheets positionally - known-wrong by
    // design); manifest correctness is gated by the manifest E2E value tests
    // (tst_v3inference) and the round-trip identity harness instead.
    bool lastFileHadManifest() const { return m_lastHadManifest; }

private:
    QString m_lastError;
    bool    m_lastUsedInference = false;
    bool    m_lastHadManifest = false;

    // Shared raw-table (SOP / instruction sheet) branch used by both the legacy
    // and v3 read paths - logic lives here once.
    SheetResult processSopSheet(ExcelReader& reader, const QString& sheetName);

    void logDebug(const QString& msg) const;
    void setError(const QString& msg);
};

} // namespace DVE
