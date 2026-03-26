#include "DataProcessor.h"
#include "SheetProcessors.h"

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
// processFile
// ===========================================================================

FileResult DataProcessor::processFile(
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
// processSheet
// ===========================================================================

SheetResult DataProcessor::processSheet(ExcelReader& reader, const QString& sheetName)
{
    logDebug(QStringLiteral("processSheet: ") + sheetName);

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
    if (sheetName.contains(QStringLiteral("SOP"), Qt::CaseInsensitive)) {
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

    if (rawSamples.isEmpty()) {
        logDebug(QStringLiteral("No samples found in sheet: ") + sheetName);
        // Return the empty result rather than an error — some sheets may be
        // intentionally blank (e.g. summary / chart sheets).
        return empty;
    }

    logDebug(QString("Sheet '%1': %2 sample(s)").arg(sheetName).arg(rawSamples.size()));

    // Select the appropriate processor and run it.
    std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));

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

    logDebug(QString("Sheet '%1' done: %2 sample(s), overallAvgTPM=%.4f")
                 .arg(sheetName)
                 .arg(result.samples.size())
                 .arg(result.overallAvgTPM));

    return result;
}

} // namespace DVE
