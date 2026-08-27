#pragma once

#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariant>
#include <QMap>
#include <QByteArray>
#include <QPoint>
#include <QRectF>
#include <QFileInfo>
#include <QRegularExpression>

namespace DVE {

// ─── Per-row measurement ─────────────────────────────────────────────────────
struct DataRow {
    double puffs           = 0.0;
    double beforeWeight    = 0.0;
    double afterWeight     = 0.0;
    double drawPressure    = 0.0;
    double resistance      = 0.0;
    QString puffingRegime;          // per-row regime (new template); empty on old template
    QString smell;
    QString clog;
    QString notes;
    double tpm             = 0.0;  // calculated
    double tpmPowerDensity = 0.0;  // calculated
    double variationTPM    = 0.0;  // calculated
    double oilConsumed     = 0.0;  // calculated

    // Open per-row values from inferred or manifest-declared (non-standard)
    // schemas - metric key -> cell value for columns that have no fixed field
    // above (e.g. PV1..PV5, chronology, a workbook's own `coil_temp`).
    //
    // PERSISTENCE (v3 Phase 3c): these round-trip through the database. Each
    // entry is written to `measurements`, keyed (sample_id, metric_id,
    // sort_order = the row's index within its sample), inside the same
    // transaction as the rest of the save; DatabaseManager::loadFile reads them
    // back; and the offline SQLite snapshot mirrors the same tables (schema v4),
    // so they survive going offline too. Also preserved losslessly through
    // recovery JSON, via the typed envelope in ReportDataJson.cpp.
    //
    // Sparse by rule (index D2): a key with no value is not stored at all,
    // rather than stored as NULL - so a row with nothing extra costs nothing and
    // reads back as the empty map it was constructed with. The metric key is
    // registered in `metric_defs` on first save if the compiled MetricRegistry
    // does not already carry it (MetricDefCache::ensureMetric).
    //
    // The 13 STANDARD fields above are deliberately NOT in here: they still live
    // in the wide `data_rows` columns and move to the long format in Phase 3d.
    // Still not DISPLAYED anywhere - no UI, plots or reports - that is Phase 4.
    QMap<QString, QVariant> extra;

    // Server-assigned row id. -1 means "not yet persisted" (a freshly entered
    // row, or one loaded from an Excel file without a DB round-trip). Plan B
    // Phase 6 uses this to map NOTIFY events on data_rows back to the
    // QTableWidgetItem that represents the row so we can decorate
    // remotely-edited cells (T18) without yanking the user's in-progress edit.
    qint64 id      = -1;
    int    version = 0;   // server-assigned; populated by loadFile, used by C3 OCC
};

// ─── Per-sample result ────────────────────────────────────────────────────────
struct SampleResult {
    // Identification
    QString sampleName;
    QString sampleID;
    QString date;
    QString tester;

    // Device parameters
    QString media;
    double  viscosity          = 0.0;
    double  resistance         = 0.0;
    double  voltage            = 0.0;
    double  power              = 0.0;
    QString heatingTechnology;
    QString puffingRegime;
    double  initialOilMass     = 0.0;

    // Calculated metrics
    double averageTPM          = 0.0;
    double stdDevTPM           = 0.0;
    double averagePowerDensity = 0.0;
    double efficiencyPercent   = 0.0;
    double totalOilConsumed    = 0.0;
    int    totalPuffs          = 0;
    double normalizedTPM       = 0.0;  // TPM per watt

    // Status flags
    QString burnStatus;   // "Y", "N", or ""
    QString clogStatus;
    QString leakStatus;

    // Row-level data
    QVector<DataRow> rows;

    // Open per-sample header values - the DataRow::extra twin, for header-band
    // keys that have no fixed member above. Same contract in every respect
    // except the table: these go to `sample_headers`, keyed (sample_id,
    // field_id) with metric kind 'header'. There is no sort_order dimension - a
    // header field is per-sample, not per-row. See DataRow::extra for the full
    // persistence / sparse-rule / Phase-4 notes.
    QMap<QString, QVariant> extra;

    // User-loaded image file paths (populated via Load Images UI, not from Excel)
    QStringList     imagePaths;
    QVector<QRectF> imageLayouts;  // parallel to imagePaths; position in inches; empty = auto-grid
    QVector<QRectF> imageCrops;    // parallel to imagePaths; normalized [0-1] crop; empty = no crop
    // C3: server-assigned image-row identities, parallel to imagePaths. -1/0
    // means "not yet persisted"; the id-aware upsert in tryWriteFile INSERTs
    // those rows and back-fills these vectors with the RETURNING id/version.
    QVector<qint64> imageIds;
    QVector<int>    imageVersions;

    // Write provenance (Phase 2b): 0-based physical column where this sample's
    // block starts in the source sheet. -1 = unknown (pre-2b recovery
    // snapshots, DB-cache fallbacks) - write-back then uses the legacy
    // sampleIndex*12 math via CellAddress fallback.
    int startColumn = -1;

    // C3: server-assigned identity for this sample row. -1 means "fresh
    // INSERT"; >0 with version > 0 means UPDATE…WHERE id=? AND version=?.
    qint64 id      = -1;
    int    version = 0;
};

// ─── Per-sheet result ─────────────────────────────────────────────────────────
struct SheetResult {
    QString sheetName;
    QString templateVersion;   // "new" | "old"
    bool    hasPerRowRegime = false;   // true => column index 4 is a per-row Puffing Regime string, not Resistance
    // Set when this sheet was parsed by the header-driven schema-inference path
    // (non-standard block layout, e.g. 13-col S26 / 8-col UserSim) rather than
    // the standard 12-wide positional path. Since Phase 2b, write-back on these
    // sheets works through CellAddress + the write provenance below; the
    // MainWindow guards only fire when that provenance is ABSENT (pre-2b
    // recovery snapshots, DB-cache fallbacks). Serialized to recovery JSON and
    // persisted to Postgres/offline snapshot (tests.from_inferred_schema).
    bool    fromInferredSchema = false;
    // W1 poka-yoke (2026-08-27): count of formula cells with NO cached value
    // on this sheet (openpyxl-fallback reads of a cache-stripped workbook -
    // i.e. one destroyed by a pre-surgical app write-back). Nonzero means the
    // parse is missing destroyed data: MainWindow warns and blocks the DB
    // save, and SheetProcessors skips its puff extrapolation. TRANSIENT
    // runtime state - deliberately NOT serialized by ReportDataJson (the
    // frozen legacy referee never emits it and the tst_v3shadow byte-identity
    // gate must hold) and not persisted anywhere.
    int     strippedFormulaCells = 0;
    QVector<SampleResult> samples;
    QStringList columnHeaders;

    // Write provenance (Phase 2b), recorded from the schema the reader
    // resolved. Empty/0 = unknown (see SampleResult::startColumn).
    // headerCells convention (fixed here once): QPoint::x() = 1-based
    // block-relative column of the header VALUE cell, QPoint::y() = 1-based
    // Excel row of that cell.
    int                   blockCols = 0;       // physical block width (12/13/8)
    int                   dataStartRow = 0;    // 1-based Excel row of data row 0 (5 on every known layout)
    QStringList           columnKeys;          // canonical metric key per physical column slot
    QMap<QString, QPoint> headerCells;         // header key -> (x=1-based block-relative col of the VALUE cell, y=1-based Excel row)

    // Aggregate
    double overallAvgTPM    = 0.0;
    double overallStdDevTPM = 0.0;

    // Plot series (parallel vectors, same length)
    QVector<double> tpmTrend;
    QVector<double> puffCounts;

    // Embedded images from Excel (PNG/JPEG bytes)
    QVector<QByteArray> images;

    // Raw table mode — set for sheets like "Test SOP's" that contain
    // plain instructional content rather than sample measurement data.
    // When true, rawHeaders / rawRows hold the full sheet contents and
    // samples is empty; the GUI shows a plain table with no plot.
    bool isRawTable = false;
    QStringList rawHeaders;
    QVector<QStringList> rawRows;

    bool hasSamples() const { return !samples.isEmpty(); }

    // Set by the DB loaders (DatabaseManager::loadFile, OfflineSnapshot::loadFile)
    // when a sheet reloads without its per-row / raw-grid content — indicating
    // a legacy or partially-migrated DB record. Always false on the Excel path.
    bool dbDataIncomplete = false;

    // C3: server-assigned identity for this sheet/test row. Parallel to
    // SampleResult::id/version. Populated by loadFile, consumed by the
    // id-aware upsert path.
    qint64 id      = -1;
    int    version = 0;
};

// ─── Full-file result ─────────────────────────────────────────────────────────
struct FileResult {
    QString filePath;
    QString fileName;
    QString templateVersion;
    QVector<SheetResult> sheets;
    QStringList sheetNames;

    // Optimistic-concurrency anchors. -1 / 0 means "not yet persisted" — a
    // subsequent saveFile / tryWriteFile will INSERT instead of UPDATE.
    // Populated by DatabaseManager::loadFile* on every successful load.
    // C3 widened id to qint64 to match Postgres BIGSERIAL and stop the
    // silent narrowing in QSqlQuery::value(0).toInt() at >2.1B rows.
    qint64 id      = -1;  // server-assigned row id; -1 if not yet persisted.
    int    version = 0;   // server-assigned row version; 0 if unknown.
};

// ─── Report configuration ─────────────────────────────────────────────────────
struct ReportConfig {
    bool includeTPM           = true;
    bool includeEfficiency    = true;
    bool includeBurnClogLeak  = true;
    bool includeStdDev        = true;
    bool includePlots         = true;
    bool includeImages        = true;
    bool singleSheetOnly      = false;
    QString singleSheetName;

    QStringList selectedColumns;   // empty = all
    QString outputPath;
    QString reportTitle;
};

// ─── Test-name fallback ───────────────────────────────────────────────────────
// The sheet name IS the test name in the current standard template. Files from
// the era where the FILENAME named the test carry Excel-default sheet names
// ("Sheet1"), which are meaningless as test names (registry 8.1, owner rule
// 2026-07-29: "always check the filename if the sheet name doesn't fit our
// map"). Wherever a sheet name is DISPLAYED as a test name (navigator, sheet
// combo, report test lists), a generic default falls back to the workbook's
// base filename. Persistence keys and Excel write-back keep the RAW sheet
// name - this is presentation only.
inline bool isGenericSheetName(const QString& name)
{
    static const QRegularExpression re(QStringLiteral("^sheet\\s*\\d*$"),
                                       QRegularExpression::CaseInsensitiveOption);
    return re.match(name.trimmed()).hasMatch();
}

inline QString effectiveTestName(const QString& sheetName, const QString& fileNameOrPath)
{
    if (!sheetName.trimmed().isEmpty() && !isGenericSheetName(sheetName))
        return sheetName;
    const QString base = QFileInfo(fileNameOrPath).completeBaseName();
    return base.isEmpty() ? sheetName : base;
}

// ─── Column name constants ────────────────────────────────────────────────────
namespace Cols {
    constexpr int PUFFS          = 0;
    constexpr int BEFORE_WEIGHT  = 1;
    constexpr int AFTER_WEIGHT   = 2;
    constexpr int DRAW_PRESSURE  = 3;
    constexpr int RESISTANCE     = 4;
    constexpr int SMELL          = 5;
    constexpr int CLOG           = 6;
    constexpr int NOTES          = 7;
    constexpr int TPM            = 8;
    constexpr int TPM_PWR        = 9;
    constexpr int VARIATION      = 10;
    constexpr int OIL_CONSUMED   = 11;
    constexpr int COUNT          = 12;
}

} // namespace DVE
