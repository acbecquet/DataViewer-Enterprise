#pragma once
// ── DataViewer Enterprise — DbRepair ────────────────────────────────────────
// Headless utility that backfills data_rows and raw_grid for DB files that
// were saved under the pre-v2.0 migration gap (samples exist with
// averageTPM > 0 but no data_rows). Locates the source .xlsx by file_path
// (or recursive search of opts.sourceDir by file name), re-parses it via
// DataProcessor::processFile, inherits the existing DB id/version, and
// re-saves via DatabaseManager::tryWriteFile so the UPDATE path fires and
// children are re-inserted without creating a new file row.
//
// The core backfill step (backfillFile) is intentionally callable with an
// already-parsed FileResult so tests can exercise it without a bundled
// Python interpreter.

#include "../pipeline/ReportData.h"

#include <QString>
#include <QStringList>

#include <functional>

namespace DVE {

class DatabaseManager;
class DataProcessor;

struct RepairOptions {
    QString sourceDir;          // top-level directory to search for .xlsx by name
    bool    dryRun  = false;    // classify without writing
    QString reportPath;         // write JSON summary to this path when non-empty
    // Audit fix (responsiveness): optional per-record progress callback so a UI
    // caller can drive a QProgressDialog instead of freezing during the
    // python-subprocess re-read loop. (done, total, currentFileName). An empty
    // callback (the default) means headless / no UI -- behavior is unchanged.
    std::function<void(int done, int total, const QString& fileName)> progress;
};

struct RepairSummary {
    int healed          = 0;    // re-parsed .xlsx + saved successfully
    int alreadyComplete = 0;    // had data_rows already; nothing to do
    int skippedNoXlsx   = 0;    // .xlsx could not be located
    int failed          = 0;    // located .xlsx but processFile / saveFile failed
    QStringList details;        // one human-readable line per file
};

// Backfill a single file given a pre-parsed complete FileResult and a
// DatabaseManager that already has the DB record (result.id / result.version
// populated from a prior loadFile). Inherits the DB id/version and calls
// tryWriteFile (UPDATE path). Returns true on success; false on DB write
// error (or dryRun, which also returns true to indicate the file would heal).
// This overload is directly testable without Python / processFile.
bool backfillFile(DatabaseManager& db,
                  const FileResult& liveResult,  // complete, from processFile
                  const FileResult& dbRecord,    // stub, from loadFile (carries id/version)
                  bool dryRun,
                  QString& detail);

// Full repair loop: enumerate DB files, detect incomplete ones, locate .xlsx,
// re-parse via DataProcessor, call backfillFile. Returns the aggregate summary.
RepairSummary runDbRepair(DatabaseManager& db,
                          DataProcessor&   proc,
                          const RepairOptions& opts);

} // namespace DVE
