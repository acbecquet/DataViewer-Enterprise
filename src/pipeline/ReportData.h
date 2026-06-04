#pragma once

#include <QString>
#include <QStringList>
#include <QVector>
#include <QVariant>
#include <QMap>
#include <QByteArray>
#include <QRectF>

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

    // Extra metrics (sheet-type specific)
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
    QVector<SampleResult> samples;
    QStringList columnHeaders;

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
