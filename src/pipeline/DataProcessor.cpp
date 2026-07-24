#include "DataProcessor.h"
#include "SheetProcessors.h"
#include "RegimeUtils.h"

#include "../model/StandardSchema.h"
#include "../model/SchemaDrivenReader.h"
#include "../model/LegacyAdapter.h"

#include <QDebug>
#include <QFileInfo>
#include <memory>

namespace DVE {

// ===========================================================================
// Constructor / Destructor
// ===========================================================================

DataProcessor::DataProcessor() = default;
DataProcessor::~DataProcessor() = default;

// ===========================================================================
// Private helpers
// ===========================================================================

void DataProcessor::logDebug(const QString& msg) const
{
#ifndef QT_NO_DEBUG
    qDebug() << "[DataProcessor]" << msg;
#else
    Q_UNUSED(msg);
#endif
}

void DataProcessor::setError(const QString& msg)
{
    m_lastError = msg;
    qWarning() << "[DataProcessor] ERROR:" << msg;
}

// ===========================================================================
// processFileLegacy — legacy positional read path (shadow-harness referee).
// Not the production path; see processFile below.
// ===========================================================================

FileResult DataProcessor::processFileLegacy(
    const QString& filePath,
    std::function<void(int, const QString&)> progressCallback)
{
    m_lastError.clear();

    FileResult result;
    result.filePath = filePath;
    result.fileName = QFileInfo(filePath).fileName();

    // -----------------------------------------------------------------------
    // 1. Open the file
    // -----------------------------------------------------------------------
    auto notify = [&](int pct, const QString& msg) {
        logDebug(QString("(%1%) %2").arg(pct).arg(msg));
        if (progressCallback)
            progressCallback(pct, msg);
    };

    notify(0, QStringLiteral("Opening file: ") + result.fileName);

    ExcelReader reader;
    if (!reader.loadFile(filePath)) {
        setError(QStringLiteral("Failed to open file: ") + reader.getLastError());
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();  // signal failure to onFileLoadFinished
        return result;
    }

    // -----------------------------------------------------------------------
    // 2. Detect template version
    // -----------------------------------------------------------------------
    QString templateVersion = reader.detectTemplateVersion();
    result.templateVersion  = templateVersion;
    logDebug(QStringLiteral("Template version: ") + templateVersion);

    // -----------------------------------------------------------------------
    // 3. Enumerate sheets
    // -----------------------------------------------------------------------
    QStringList sheetNames = reader.getSheetNames();
    result.sheetNames      = sheetNames;

    if (sheetNames.isEmpty()) {
        setError(QStringLiteral("No sheets found in file: ") + filePath);
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();
        return result;
    }

    notify(5, QString("Found %1 sheet(s)").arg(sheetNames.size()));

    // -----------------------------------------------------------------------
    // 4. Process each sheet
    // -----------------------------------------------------------------------
    const int nSheets = sheetNames.size();

    for (int idx = 0; idx < nSheets; ++idx) {
        const QString& sheetName = sheetNames[idx];

        int progressBase = 5 + static_cast<int>(90.0 * idx / nSheets);
        notify(progressBase, QStringLiteral("Processing sheet: ") + sheetName);

        // Select the sheet in the reader.
        if (!reader.selectSheet(sheetName)) {
            qWarning() << "[DataProcessor] Could not select sheet:" << sheetName
                       << "-" << reader.getLastError();
            // Continue with remaining sheets rather than aborting.
            continue;
        }

        SheetResult sheetResult = processSheetLegacy(reader, sheetName);
        sheetResult.templateVersion = templateVersion;

        // Capture column headers (populated after selectSheet).
        sheetResult.columnHeaders = reader.getColumnHeaders();

        result.sheets.append(sheetResult);
    }

    notify(100, QStringLiteral("Done. Processed ")
               + QString::number(result.sheets.size())
               + QStringLiteral(" sheet(s)."));

    return result;
}

// ===========================================================================
// processSheetLegacy — legacy positional single-sheet read (shadow referee).
// ===========================================================================

SheetResult DataProcessor::processSheetLegacy(ExcelReader& reader, const QString& sheetName)
{
    logDebug(QStringLiteral("processSheetLegacy: ") + sheetName);

    SheetResult empty;
    empty.sheetName = sheetName;

    // Retrieve all samples from the currently-selected sheet.
    QVector<ExcelReader::SampleData> rawSamples = reader.getAllSamples();

    // ── Diagnostic dump (debug builds only) ─────────────────────────────────
#ifndef QT_NO_DEBUG
    qDebug() << "[DVE DIAG] Sheet:" << sheetName
             << " | rawSamples:" << rawSamples.size();
    for (int si = 0; si < rawSamples.size(); ++si) {
        const auto& rs = rawSamples[si];
        qDebug() << "  Sample" << si
                 << "| startCol:" << rs.startColumn
                 << "| dataRows:" << rs.dataRows.size()
                 << "| sampleID:" << rs.metadata.sampleID
                 << "| testName:" << rs.metadata.testName
                 << "| voltage:"  << rs.metadata.voltage
                 << "| resistance:" << rs.metadata.resistance;
        for (int ri = 0; ri < qMin(3, rs.dataRows.size()); ++ri) {
            QString rowStr;
            for (int ci = 0; ci < rs.dataRows[ri].size(); ++ci) {
                const QVariant& v = rs.dataRows[ri][ci];
                rowStr += QString("  [%1] type=%2 val=%3")
                    .arg(ci)
                    .arg(static_cast<int>(v.typeId()))
                    .arg(v.toString());
            }
            qDebug() << "    Row" << ri << ":" << rowStr;
        }
    }
#endif
    // ────────────────────────────────────────────────────────────────────────

    // ── SOP / instruction sheet — show raw table, no plot ───────────────
    if (sheetName.contains(QStringLiteral("SOP"), Qt::CaseInsensitive))
        return processSopSheet(reader, sheetName);

    if (rawSamples.isEmpty()) {
        logDebug(QStringLiteral("No samples found in sheet: ") + sheetName);
        // Return the empty result rather than an error — some sheets may be
        // intentionally blank (e.g. summary / chart sheets).
        return empty;
    }

    logDebug(QString("Sheet '%1': %2 sample(s)").arg(sheetName).arg(rawSamples.size()));

    // Select the appropriate processor and run it.
    std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));

    // Column index 4's header decides per-row Resistance vs Puffing Regime.
    const QStringList hdrs = reader.getColumnHeaders();
    const bool perRowRegime =
        (hdrs.size() > 4) && RegimeUtils::isRegimeHeader(hdrs.at(4));
    processor->setPerRowRegime(perRowRegime);

    // detectTemplateVersion() is a const method on the reader; we call it here
    // so each processSheet call carries the version even when called standalone.
    QString templateVersion = reader.detectTemplateVersion();

    SheetResult result;
    try {
        result = processor->process(rawSamples, sheetName, templateVersion);
    } catch (const std::exception& ex) {
        setError(QString("Exception processing sheet '%1': %2")
                     .arg(sheetName, QString::fromLatin1(ex.what())));
        return empty;
    } catch (...) {
        setError(QString("Unknown exception processing sheet '%1'").arg(sheetName));
        return empty;
    }

    logDebug(QString("Sheet '%1' done: %2 sample(s), overallAvgTPM=%3")
                 .arg(sheetName)
                 .arg(result.samples.size())
                 .arg(result.overallAvgTPM, 0, 'f', 4));

    return result;
}

// ===========================================================================
// processSopSheet — SOP / instruction sheet: raw table, no plot.
// Shared by processSheet (production) and processSheetLegacy; behavior unchanged.
// ===========================================================================

SheetResult DataProcessor::processSopSheet(ExcelReader& reader, const QString& sheetName)
{
    SheetResult sopResult;
    sopResult.sheetName       = sheetName;
    sopResult.templateVersion = reader.detectTemplateVersion();
    sopResult.isRawTable      = true;

    QVector<QStringList> rows = reader.getSheetRows();
    if (!rows.isEmpty()) {
        // First non-empty row becomes the column headers.
        int headerIdx = 0;
        for (int i = 0; i < rows.size(); ++i) {
            bool hasContent = false;
            for (const QString& cell : rows[i]) {
                if (!cell.trimmed().isEmpty()) { hasContent = true; break; }
            }
            if (hasContent) { headerIdx = i; break; }
        }
        sopResult.rawHeaders = rows[headerIdx];
        for (int i = headerIdx + 1; i < rows.size(); ++i)
            sopResult.rawRows.append(rows[i]);
    }
    return sopResult;
}

// ===========================================================================
// processFile — production read path (schema-driven)
//
// Load / sheet-iteration / FileResult assembly identical to processFileLegacy,
// with processSheet doing schema-driven extraction (SchemaDrivenReader +
// LegacyAdapter). The Phase-1 shadow gate (tst_v3shadow) proved this
// byte-identical to processFileLegacy over the whole fixture corpus.
// ===========================================================================

FileResult DataProcessor::processFile(
    const QString& filePath,
    std::function<void(int, const QString&)> progressCallback)
{
    m_lastError.clear();

    FileResult result;
    result.filePath = filePath;
    result.fileName = QFileInfo(filePath).fileName();

    auto notify = [&](int pct, const QString& msg) {
        logDebug(QString("(%1%) %2").arg(pct).arg(msg));
        if (progressCallback)
            progressCallback(pct, msg);
    };

    notify(0, QStringLiteral("Opening file: ") + result.fileName);

    ExcelReader reader;
    if (!reader.loadFile(filePath)) {
        setError(QStringLiteral("Failed to open file: ") + reader.getLastError());
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();  // signal failure to onFileLoadFinished
        return result;
    }

    QString templateVersion = reader.detectTemplateVersion();
    result.templateVersion  = templateVersion;
    logDebug(QStringLiteral("Template version: ") + templateVersion);

    QStringList sheetNames = reader.getSheetNames();
    result.sheetNames      = sheetNames;

    if (sheetNames.isEmpty()) {
        setError(QStringLiteral("No sheets found in file: ") + filePath);
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();
        return result;
    }

    notify(5, QString("Found %1 sheet(s)").arg(sheetNames.size()));

    const int nSheets = sheetNames.size();

    for (int idx = 0; idx < nSheets; ++idx) {
        const QString& sheetName = sheetNames[idx];

        int progressBase = 5 + static_cast<int>(90.0 * idx / nSheets);
        notify(progressBase, QStringLiteral("Processing sheet: ") + sheetName);

        if (!reader.selectSheet(sheetName)) {
            qWarning() << "[DataProcessor] Could not select sheet:" << sheetName
                       << "-" << reader.getLastError();
            continue;
        }

        SheetResult sheetResult = processSheet(reader, sheetName);
        sheetResult.templateVersion = templateVersion;

        // Capture column headers (populated after selectSheet).
        sheetResult.columnHeaders = reader.getColumnHeaders();

        result.sheets.append(sheetResult);
    }

    notify(100, QStringLiteral("Done. Processed ")
               + QString::number(result.sheets.size())
               + QStringLiteral(" sheet(s)."));

    return result;
}

// ===========================================================================
// processSheet — production single-sheet read (schema-driven)
//
// Extraction goes through model::SchemaDrivenReader (positional resolution, to
// mirror ExcelReader::extractRow byte-for-byte) + model::LegacyAdapter, then the
// UNCHANGED SheetProcessor chain. Header/metadata for each block is resolved by
// sniffing the same landmark cells ExcelReader::extractMetadata sniffs and
// applying the matching builtin header layout (Standard / Cart / Project).
// ===========================================================================

namespace {

// Read a header-field set for one block into a raw headers map, exactly like
// SchemaDrivenReader::parseSheet does (raw QVariant cells, block-relative
// 1-based row/col). Used to (re)build a block's headers from a chosen layout.
QMap<QString, QVariant> extractBlockHeaders(const QVector<QVector<QVariant>>& cells,
                                            int off,
                                            const QVector<model::HeaderFieldDef>& fields)
{
    auto cellAt = [&](int r, int c) -> QVariant {
        if (r < 0 || r >= cells.size() || c < 0 || c >= cells[r].size()) return {};
        return cells[r][c];
    };
    QMap<QString, QVariant> headers;
    for (const model::HeaderFieldDef& h : fields)
        headers.insert(h.key, cellAt(h.row - 1, off + h.col - 1));
    return headers;
}

} // namespace

SheetResult DataProcessor::processSheet(ExcelReader& reader, const QString& sheetName)
{
    logDebug(QStringLiteral("processSheet: ") + sheetName);

    SheetResult empty;
    empty.sheetName = sheetName;

    // ── SOP / instruction sheet — identical raw-table branch as legacy ──
    if (sheetName.contains(QStringLiteral("SOP"), Qt::CaseInsensitive))
        return processSopSheet(reader, sheetName);

    const QVector<QVector<QVariant>> cells = reader.currentSheetCells();
    if (cells.isEmpty())
        return empty;

    // Column index 4's header decides per-row Resistance vs Puffing Regime -
    // same rule as processSheetLegacy.
    const QStringList hdrs = reader.getColumnHeaders();
    const bool perRowRegime =
        (hdrs.size() > 4) && RegimeUtils::isRegimeHeader(hdrs.at(4));

    const QString templateVersion = reader.detectTemplateVersion();

    // Data columns: positional resolution reproduces ExcelReader::extractRow's
    // fixed 12-wide-by-position read (name-first matching would shuffle data on
    // the historical Cart / 13-wide layouts). Header fields from the Standard
    // layout are extracted here too; cart/project blocks get theirs replaced
    // below.
    const model::TemplateSchema schema = model::standardV1(perRowRegime);
    model::Sheet mSheet = model::SchemaDrivenReader::parseSheet(
        cells, sheetName, schema, perRowRegime, model::ColumnResolution::Positional);
    if (mSheet.samples.isEmpty())
        return empty;

    // ── Per-block metadata layout: sniff the SAME landmark cells
    //    ExcelReader::extractMetadata sniffs and apply the matching header
    //    layout. Cart / Project fully replace the block's headers; a Standard
    //    "old"-template block drops Heating Technology (extractMetadata's old
    //    branch never reads it). ──
    const model::TemplateSchema cartSchema =
        model::standardV1(perRowRegime, model::HeaderLayout::Cart);
    const model::TemplateSchema projectSchema =
        model::standardV1(perRowRegime, model::HeaderLayout::Project);

    for (int b = 0; b < mSheet.samples.size(); ++b) {
        const int off = b * schema.blockCols;

        auto sniff = [&](int r, int c) -> QString {
            if (r < 0 || r >= cells.size() || (off + c) < 0 || (off + c) >= cells[r].size())
                return QString();
            return cells[r][off + c].toString().trimmed();
        };
        const bool isCart    = sniff(1, 0).contains(QStringLiteral("Cart"), Qt::CaseInsensitive);
        const bool isProject = sniff(0, 5).contains(QStringLiteral("Project:"), Qt::CaseInsensitive);

        model::Sample& sample = mSheet.samples[b];
        if (isCart) {
            sample.headers = extractBlockHeaders(cells, off, cartSchema.headerFields);
        } else if (isProject) {
            sample.headers = extractBlockHeaders(cells, off, projectSchema.headerFields);
            // Assemble sampleID = "<project> <suffix>" (space, not dash), exactly
            // as extractMetadata does; suffix already carries its own hyphen.
            const QString project = sample.headers.value(QStringLiteral("project_name")).toString().trimmed();
            const QString suffix  = sample.headers.value(QStringLiteral("sample_suffix")).toString().trimmed();
            sample.headers.insert(QStringLiteral("sample_id"),
                                  suffix.isEmpty() ? project : project + QStringLiteral(" ") + suffix);
        } else if (templateVersion != QLatin1String("new")) {
            // Mirror extractMetadata's else-structure: only the "new" branch
            // reads Heating Technology. Every other standardized block (the
            // "old" template, and the theoretical "unknown") assigns none, so
            // drop the key here.
            sample.headers.remove(QStringLiteral("heating_technology"));
        }
    }

    // ── Lower each v3 Sample to the legacy raw shape, then run the UNCHANGED
    //    SheetProcessor chain. ──
    QVector<ExcelReader::SampleData> rawSamples;
    rawSamples.reserve(mSheet.samples.size());
    for (int i = 0; i < mSheet.samples.size(); ++i)
        rawSamples.append(model::LegacyAdapter::lowerSample(mSheet.samples[i], schema, i));

    std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));
    processor->setPerRowRegime(perRowRegime);

    SheetResult result;
    try {
        result = processor->process(rawSamples, sheetName, templateVersion);
    } catch (const std::exception& ex) {
        setError(QString("Exception processing sheet '%1': %2")
                     .arg(sheetName, QString::fromLatin1(ex.what())));
        return empty;
    } catch (...) {
        setError(QString("Unknown exception processing sheet '%1'").arg(sheetName));
        return empty;
    }

    logDebug(QString("Sheet '%1' done: %2 sample(s), overallAvgTPM=%3")
                 .arg(sheetName)
                 .arg(result.samples.size())
                 .arg(result.overallAvgTPM, 0, 'f', 4));

    return result;
}

} // namespace DVE
