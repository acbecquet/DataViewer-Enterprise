#include "ReportGenerator.h"
#include "../plotting/PlotEngine.h"
#include "../utils/ImageUtils.h"
#include <QDebug>
#include <QFile>
#include <QFileInfo>
#include <QDir>
#include <QDate>
#include <QRegularExpression>
#include <QImage>
#include <QImageReader>
#include <QBuffer>
#include <cmath>
#include <algorithm>

namespace DVE {

// ──────────────────────────────────────────────────────────────────────────────
ReportGenerator::ReportGenerator(QObject* parent)
    : QObject(parent)
{}

void ReportGenerator::setResourcePath(const QString& path)
{
    m_resourcePath = path;
}

// ──────────────────────────────────────────────────────────────────────────────
void ReportGenerator::reportProgress(ProgressFn fn, int pct, const QString& msg)
{
    if (fn) fn(pct, msg);
    emit progressChanged(pct, msg);
}

void ReportGenerator::logDebug(const QString& msg) const
{
#ifndef QT_NO_DEBUG
    qDebug() << "[ReportGenerator]" << msg;
#else
    Q_UNUSED(msg);
#endif
}

// ──────────────────────────────────────────────────────────────────────────────
QVector<QByteArray> ReportGenerator::buildPlots(const SheetResult& sheet, bool includeBarChart)
{
    QVector<QByteArray> plots;

    // Plot 1: TPM Trend — one series per sample, matching the GUI rendering exactly.
    {
        static const QColor kColors[] = {
            QColor(0x00, 0x66, 0xCC), QColor(0xFF, 0x73, 0x00),
            QColor(0x00, 0xAA, 0x44), QColor(0xCC, 0x00, 0x00),
            QColor(0x99, 0x00, 0xCC), QColor(0x00, 0xAA, 0xCC),
        };

        QVector<PlotSeries> series;
        int colorIdx = 0;
        for (const SampleResult& sr : sheet.samples) {
            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
            ps.color     = kColors[colorIdx++ % 6];
            ps.drawLine  = true;
            ps.drawDots  = (sr.rows.size() <= 30);
            ps.lineWidth = 2;
            ps.dotRadius = 3;

            for (const DataRow& row : sr.rows) {
                if (row.beforeWeight == 0.0 || row.afterWeight == 0.0) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.tpm);
            }
            if (!ps.x.isEmpty())
                series.append(std::move(ps));
        }

        if (!series.isEmpty()) {
            PlotConfig cfg;
            cfg.title      = sheet.sheetName + QStringLiteral(" \u2013 TPM Trend");
            cfg.xLabel     = "Cumulative Puffs";
            cfg.yLabel     = "TPM (mg)";
            cfg.width      = 800;
            cfg.height     = 480;
            cfg.autoScale  = true;
            cfg.showGrid   = true;
            cfg.showLegend = (series.size() > 1);

            QPixmap pm  = PlotEngine::renderLinePlot(series, cfg);
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty())
                plots.append(std::move(png));
        }
    }

    // Plot 2: TPM Bar Chart (average TPM per sample)
    if (includeBarChart) {
        QVector<QString> names;
        QVector<double> avgs, sdevs;
        for (const auto& s : sheet.samples) {
            names.append(s.sampleName.isEmpty() ? s.sampleID : s.sampleName);
            avgs.append(s.averageTPM);
            sdevs.append(s.stdDevTPM);
        }
        QPixmap pm = PlotEngine::renderTPMBarChart(names, avgs, sdevs,
                                                    sheet.sheetName + QStringLiteral(" \u2013 Average TPM"));
        QByteArray png = PlotEngine::toPng(pm, 150);
        if (!png.isEmpty()) plots.append(std::move(png));
    }

    // Plot 3: Draw Pressure Trend — one series per sample.
    {
        static const QColor kColors[] = {
            QColor(0x00, 0x66, 0xCC), QColor(0xFF, 0x73, 0x00),
            QColor(0x00, 0xAA, 0x44), QColor(0xCC, 0x00, 0x00),
            QColor(0x99, 0x00, 0xCC), QColor(0x00, 0xAA, 0xCC),
        };

        QVector<PlotSeries> series;
        int colorIdx = 0;
        for (const SampleResult& sr : sheet.samples) {
            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
            ps.color     = kColors[colorIdx++ % 6];
            ps.drawLine  = true;
            ps.drawDots  = (sr.rows.size() <= 30);
            ps.lineWidth = 2;
            ps.dotRadius = 3;

            for (const DataRow& row : sr.rows) {
                if (row.drawPressure == 0.0) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.drawPressure);
            }
            if (!ps.x.isEmpty())
                series.append(std::move(ps));
        }

        if (!series.isEmpty()) {
            PlotConfig cfg;
            cfg.title      = sheet.sheetName + QStringLiteral(" \u2013 Draw Pressure");
            cfg.xLabel     = "Cumulative Puffs";
            cfg.yLabel     = "Draw Pressure (Pa)";
            cfg.width      = 800;
            cfg.height     = 480;
            cfg.autoScale  = true;
            cfg.showGrid   = true;
            cfg.showLegend = (series.size() > 1);

            QPixmap pm  = PlotEngine::renderLinePlot(series, cfg);
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty())
                plots.append(std::move(png));
        }
    }

    return plots;
}

// ──────────────────────────────────────────────────────────────────────────────
// Compute the median of a list of doubles.  Returns 0 if empty.
static double medianOf(QVector<double> vals)
{
    if (vals.isEmpty()) return 0.0;
    std::sort(vals.begin(), vals.end());
    int n = vals.size();
    if (n % 2 == 1) return vals[n / 2];
    return (vals[n / 2 - 1] + vals[n / 2]) / 2.0;
}

SlideTable ReportGenerator::buildTable(const SheetResult& sheet, const ReportConfig& config)
{
    SlideTable tbl;

    // Check whether any sample has non-zero draw pressure data
    bool hasDrawPressure = false;
    for (const SampleResult& s : sheet.samples) {
        for (const DataRow& dr : s.rows) {
            if (dr.drawPressure > 0.0) { hasDrawPressure = true; break; }
        }
        if (hasDrawPressure) break;
    }

    // Default column set — V/R/P combined into one stacked column.
    // Notes is always the last column.
    QStringList defaultCols = {
        "Sample Name", "Media", "Viscosity (cP)",
        "Puffing Regime",
        "Voltage (V)\nResistance (\xCE\xA9)\nPower (W)",   // combined V/R/P
        "Avg TPM\n(mg/puff)", "Std\nDev",
        "Oil\nConsumed\n(g)", "Initial\nOil (g)", "Usage\nEfficiency",
    };
    if (hasDrawPressure)
        defaultCols << "Draw\nPressure";
    defaultCols << "Burn" << "Clog" << "Leak" << "Notes";

    tbl.headers = config.selectedColumns.isEmpty() ? defaultCols : config.selectedColumns;

    for (const auto& s : sheet.samples) {
        QStringList row;
        for (const auto& col : tbl.headers) {
            if      (col.contains("Sample", Qt::CaseInsensitive))    row << (s.sampleName.isEmpty() ? s.sampleID : s.sampleName);
            else if (col.contains("Media",  Qt::CaseInsensitive))    row << s.media;
            else if (col.contains("Visco",  Qt::CaseInsensitive))    row << QString::number(s.viscosity, 'f', 1);
            else if (col.contains("Puffing",Qt::CaseInsensitive))    row << s.puffingRegime;
            else if (col.contains("Voltage",Qt::CaseInsensitive) &&
                     col.contains("Resistance",Qt::CaseInsensitive)) {
                // Stacked V/R/P cell
                row << QString("%1\n%2\n%3")
                       .arg(QString::number(s.voltage, 'f', 2),
                            QString::number(s.resistance, 'f', 3),
                            QString::number(s.power, 'f', 2));
            }
            else if (col.contains("Voltage",Qt::CaseInsensitive))    row << QString::number(s.voltage, 'f', 2);
            else if (col.contains("Resist", Qt::CaseInsensitive))    row << QString::number(s.resistance, 'f', 3);
            else if (col.contains("Power",  Qt::CaseInsensitive))    row << QString::number(s.power, 'f', 2);
            else if (col.contains("Avg TPM",Qt::CaseInsensitive))    row << QString::number(s.averageTPM, 'f', 4);
            else if (col.contains("Std",    Qt::CaseInsensitive) &&
                     col.contains("Dev",    Qt::CaseInsensitive))    row << QString::number(s.stdDevTPM, 'f', 4);
            else if (col.contains("Usage",  Qt::CaseInsensitive) &&
                     col.contains("Efficien",Qt::CaseInsensitive)) {
                if (s.initialOilMass > 0.0) {
                    double oilConsumedG = s.totalOilConsumed / 1000.0;
                    double eff = (oilConsumedG / s.initialOilMass) * 100.0;
                    row << QString::number(eff, 'f', 1) + "%";
                } else {
                    row << "-";
                }
            }
            else if (col.contains("Burn",   Qt::CaseInsensitive))    row << (s.burnStatus.isEmpty() ? "N" : s.burnStatus);
            else if (col.contains("Clog",   Qt::CaseInsensitive))    row << (s.clogStatus.isEmpty() ? "N" : s.clogStatus);
            else if (col.contains("Leak",   Qt::CaseInsensitive))    row << (s.leakStatus.isEmpty() ? "N" : s.leakStatus);
            else if (col.contains("Tester", Qt::CaseInsensitive))    row << s.tester;
            else if (col.contains("Date",   Qt::CaseInsensitive))    row << s.date;
            else if (col.contains("Oil",    Qt::CaseInsensitive) &&
                     col.contains("Consumed",Qt::CaseInsensitive))   row << QString::number(s.totalOilConsumed / 1000.0, 'f', 3);
            else if (col.contains("Initial",Qt::CaseInsensitive) &&
                     col.contains("Oil",    Qt::CaseInsensitive))    row << QString::number(s.initialOilMass, 'f', 3);
            else if (col.contains("Draw",   Qt::CaseInsensitive) &&
                     col.contains("Pressure",Qt::CaseInsensitive)) {
                // Median draw pressure (excludes zero/empty rows)
                QVector<double> dpVals;
                for (const DataRow& dr : s.rows)
                    if (dr.drawPressure > 0.0)
                        dpVals.append(dr.drawPressure);
                row << (dpVals.isEmpty() ? QString("-") : QString::number(medianOf(dpVals), 'f', 2));
            }
            else if (col.contains("Note",   Qt::CaseInsensitive)) {
                // Aggregate all non-empty per-row notes into a single cell.
                QStringList noteParts;
                for (const DataRow& dr : s.rows) {
                    QString n = dr.notes.trimmed();
                    if (!n.isEmpty())
                        noteParts << n;
                }
                // Append data-cleanup summary if present
                if (s.extra.contains("cleanupNote"))
                    noteParts << s.extra["cleanupNote"].toString();
                row << noteParts.join(QStringLiteral("; "));
            }
            else                                                       row << "";
        }
        tbl.rows.append(row);
    }

    // Explicit column width fractions matching the reference layout.
    // Columns: SampleName, Media, Viscosity, PuffRegime, V/R/P, AvgTPM, StdDev,
    //          OilConsumed, InitialOil, UsageEff, [DrawPressure], Burn, Clog, Leak, Notes
    const int nCols = tbl.headers.size();
    tbl.colWidthFractions.clear();
    if (hasDrawPressure) {
        // 15 columns
        tbl.colWidthFractions = {
            0.068, 0.048, 0.058, 0.068,  // Sample, Media, Viscosity, Puffing
            0.090,                         // V/R/P stacked
            0.068, 0.046,                  // Avg TPM, Std Dev
            0.068, 0.050, 0.065,           // Oil Consumed, Initial Oil, Usage Eff
            0.058,                         // Draw Pressure
            0.040, 0.038, 0.038,           // Burn, Clog, Leak
            0.197                          // Notes
        };
    } else {
        // 14 columns (no Draw Pressure)
        tbl.colWidthFractions = {
            0.070, 0.050, 0.060, 0.070,  // Sample, Media, Viscosity, Puffing
            0.092,                         // V/R/P stacked
            0.070, 0.048,                  // Avg TPM, Std Dev
            0.070, 0.052, 0.067,           // Oil Consumed, Initial Oil, Usage Eff
            0.044, 0.042, 0.042,           // Burn, Clog, Leak
            0.223                          // Notes
        };
    }
    // Ensure vector matches column count; fall back to equal split if mismatch
    if (tbl.colWidthFractions.size() != nCols) {
        tbl.colWidthFractions.clear();
        tbl.colWidthFractions.fill(1.0 / nCols, nCols);
    }

    // Table always fills the full slide width; height is computed from row count.
    tbl.x = 0.46;
    tbl.y = 0.92;
    tbl.w = 12.42;
    // Header row ≈ 0.5", each data row ≈ 1/3"; PPTX will honour these heights.
    tbl.h = 0.5 + static_cast<double>(tbl.rows.size()) * (1.0 / 3.0);

    return tbl;
}

QVector<QByteArray> ReportGenerator::collectImages(const SheetResult& sheet)
{
    QVector<QByteArray> result;
    result.reserve(sheet.images.size());
    for (const QByteArray& raw : sheet.images) {
        QByteArray compressed = compressImageBlob(raw);
        result.append(std::move(compressed));
    }
    return result;
}

// ──────────────────────────────────────────────────────────────────────────────
bool ReportGenerator::generateFullReport(const FileResult& data,
                                          const ReportConfig& config,
                                          ProgressFn progress)
{
    logDebug("Generating full report for: " + data.fileName);

    if (data.sheets.isEmpty()) {
        m_lastError = "No sheet data to report";
        return false;
    }

    PptxWriter writer;
    writer.setResourcePath(m_resourcePath);

    // Determine output path
    QString outPath = config.outputPath;
    if (outPath.isEmpty()) {
        QFileInfo fi(data.filePath);
        outPath = fi.dir().absoluteFilePath(fi.baseName() + "_Report.pptx");
    }

    const int totalSheets = data.sheets.size();

    // Cover slide
    reportProgress(progress, 2, "Adding cover slide...");
    QDate today = QDate::currentDate();
    QString dateStr = today.toString("MMMM d, yyyy");
    // Strip file extension from display name (e.g. ".xlsx")
    QString displayFileName = QFileInfo(data.fileName).completeBaseName();
    QString reportTitle = config.reportTitle.isEmpty()
        ? (displayFileName + " Standard Test Report")
        : config.reportTitle;
    writer.addCoverSlide(reportTitle, dateStr);

    // Content slides
    for (int i = 0; i < totalSheets; ++i) {
        SheetResult sheet = data.sheets[i];
        if (!sheet.hasSamples()) continue;

        // Remove completely empty samples (no data rows at all)
        sheet.samples.erase(
            std::remove_if(sheet.samples.begin(), sheet.samples.end(),
                           [](const SampleResult& s) { return s.rows.isEmpty(); }),
            sheet.samples.end());
        if (!sheet.hasSamples()) continue;

        int pct = 5 + (int)(90.0 * i / totalSheets);
        reportProgress(progress, pct, "Processing sheet: " + sheet.sheetName);

        // Build table
        SlideTable tbl = buildTable(sheet, config);

        // Build plots
        QVector<QByteArray> plotPngs;
        if (config.includePlots) {
            plotPngs = buildPlots(sheet, true);
        }

        // Fixed positions for the three report plots (inches from top-left).
        // [0] TPM Trend  [1] Avg TPM Bar  [2] Draw Pressure
        static const struct { double x, y, w, h; } kPlotLayout[] = {
            { 0.10, 3.25, 4.32, 3.20 },
            { 4.51, 3.25, 4.32, 3.20 },
            { 8.91, 3.25, 4.32, 3.20 },
        };

        QVector<SlideImage> plotImages;
        for (int pi = 0; pi < plotPngs.size() && pi < 3; ++pi) {
            SlideImage img;
            img.pngData = std::move(plotPngs[pi]);
            img.x = kPlotLayout[pi].x;
            img.y = kPlotLayout[pi].y;
            img.w = kPlotLayout[pi].w;
            img.h = kPlotLayout[pi].h;
            plotImages.append(std::move(img));
        }

        writer.addContentSlide(sheet.sheetName, tbl, plotImages);

        // Image slides: per-sample user-loaded images first, then sheet-level photos
        if (config.includeImages) {
            for (const SampleResult& sr : sheet.samples) {
                if (sr.imagePaths.isEmpty()) {
                    logDebug(QString("  Sample '%1' has no images assigned")
                             .arg(sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName));
                    continue;
                }
                QString displayName = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
                logDebug(QString("  Sample '%1' has %2 image(s)")
                         .arg(displayName).arg(sr.imagePaths.size()));

                const bool hasLayouts = (sr.imageLayouts.size() == sr.imagePaths.size());
                if (hasLayouts) {
                    QVector<SlideImage> slideImages;
                    for (int i = 0; i < sr.imagePaths.size(); ++i) {
                        QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                        if (data.isEmpty()) {
                            logDebug(QString("    WARNING: Image file missing or unreadable: %1")
                                     .arg(sr.imagePaths[i]));
                            continue;
                        }
                        SlideImage si;
                        si.pngData = std::move(data);
                        const QRectF& r = sr.imageLayouts[i];
                        si.x = r.x(); si.y = r.y();
                        si.w = r.width(); si.h = r.height();
                        slideImages.append(std::move(si));
                    }
                    if (!slideImages.isEmpty())
                        writer.addImageSlide(displayName + " Images", slideImages);
                } else {
                    QVector<QByteArray> imgBytes;
                    for (int i = 0; i < sr.imagePaths.size(); ++i) {
                        QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                        if (data.isEmpty()) {
                            logDebug(QString("    WARNING: Image file missing or unreadable: %1")
                                     .arg(sr.imagePaths[i]));
                            continue;
                        }
                        imgBytes.append(std::move(data));
                    }
                    if (!imgBytes.isEmpty())
                        writer.addImageSlide(displayName + " Images", imgBytes);
                }
            }

            QVector<QByteArray> sheetImgs = collectImages(sheet);
            if (!sheetImgs.isEmpty())
                writer.addImageSlide(sheet.sheetName + " \u2013 Photos", sheetImgs);
        }
    }

    reportProgress(progress, 95, "Saving report...");
    bool ok = writer.save(outPath);
    if (!ok) {
        m_lastError = writer.lastError();
        emit reportFinished(false, outPath);
        return false;
    }

    reportProgress(progress, 100, "Report saved: " + outPath);
    emit reportFinished(true, outPath);
    logDebug("Full report saved to: " + outPath);
    return true;
}

// ──────────────────────────────────────────────────────────────────────────────
bool ReportGenerator::generateTestReport(const FileResult& data,
                                          const QString& sheetName,
                                          const ReportConfig& config,
                                          ProgressFn progress)
{
    logDebug("Generating test report for sheet: " + sheetName);

    // Find the sheet
    const SheetResult* target = nullptr;
    for (const auto& s : data.sheets) {
        if (s.sheetName == sheetName) { target = &s; break; }
    }

    if (!target) {
        m_lastError = "Sheet not found: " + sheetName;
        return false;
    }

    PptxWriter writer;
    writer.setResourcePath(m_resourcePath);

    QString outPath = config.outputPath;
    if (outPath.isEmpty()) {
        QFileInfo fi(data.filePath);
        outPath = fi.dir().absoluteFilePath(fi.baseName() + "_" + sheetName + "_Report.pptx");
        // Sanitize filename
        outPath.replace(QRegularExpression("[\\\\/:*?\"<>|]"), "_");
    }

    reportProgress(progress, 5, "Adding cover slide...");
    QDate today = QDate::currentDate();
    // Strip file extension from display name (e.g. ".xlsx")
    QString displayFileName = QFileInfo(data.fileName).completeBaseName();
    QString reportTitle = config.reportTitle.isEmpty()
        ? (displayFileName + " \u2013 " + sheetName + " Test Report")
        : config.reportTitle;
    writer.addCoverSlide(reportTitle, today.toString("MMMM d, yyyy"));

    // Remove completely empty samples (no data rows at all)
    SheetResult filtered = *target;
    filtered.samples.erase(
        std::remove_if(filtered.samples.begin(), filtered.samples.end(),
                       [](const SampleResult& s) { return s.rows.isEmpty(); }),
        filtered.samples.end());
    target = &filtered;

    reportProgress(progress, 30, "Building table...");
    SlideTable tbl = buildTable(*target, config);

    QVector<QByteArray> plotPngs;
    if (config.includePlots) {
        reportProgress(progress, 50, "Generating plots...");
        plotPngs = buildPlots(*target, target->samples.size() > 1);
    }

    // Fixed positions for the three report plots (inches from top-left).
    static const struct { double x, y, w, h; } kPlotLayout[] = {
        { 0.10, 3.25, 4.32, 3.20 },
        { 4.51, 3.25, 4.32, 3.20 },
        { 8.91, 3.25, 4.32, 3.20 },
    };

    QVector<SlideImage> plotImages;
    for (int pi = 0; pi < plotPngs.size() && pi < 3; ++pi) {
        SlideImage img;
        img.pngData = plotPngs[pi];
        img.x = kPlotLayout[pi].x;
        img.y = kPlotLayout[pi].y;
        img.w = kPlotLayout[pi].w;
        img.h = kPlotLayout[pi].h;
        plotImages.append(img);
    }

    reportProgress(progress, 70, "Adding content slide...");
    writer.addContentSlide(target->sheetName, tbl, plotImages);

    if (config.includeImages) {
        // Per-sample user-loaded image slides
        for (const SampleResult& sr : target->samples) {
            if (sr.imagePaths.isEmpty()) continue;
            QString displayName = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
            const bool hasLayouts = (sr.imageLayouts.size() == sr.imagePaths.size());
            if (hasLayouts) {
                QVector<SlideImage> slideImages;
                for (int i = 0; i < sr.imagePaths.size(); ++i) {
                    QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                    QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                    if (data.isEmpty()) continue;
                    SlideImage si;
                    si.pngData = data;
                    const QRectF& r = sr.imageLayouts[i];
                    si.x = r.x(); si.y = r.y();
                    si.w = r.width(); si.h = r.height();
                    slideImages.append(si);
                }
                if (!slideImages.isEmpty())
                    writer.addImageSlide(displayName + " Images", slideImages);
            } else {
                QVector<QByteArray> imgBytes;
                for (int i = 0; i < sr.imagePaths.size(); ++i) {
                    QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                    QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                    if (!data.isEmpty()) imgBytes.append(data);
                }
                if (!imgBytes.isEmpty())
                    writer.addImageSlide(displayName + " Images", imgBytes);
            }
        }

        // Sheet-level embedded photos
        QVector<QByteArray> sheetImgs = collectImages(*target);
        if (!sheetImgs.isEmpty())
            writer.addImageSlide(target->sheetName + " \u2013 Photos", sheetImgs);
    }

    reportProgress(progress, 95, "Saving...");
    bool ok = writer.save(outPath);
    if (!ok) {
        m_lastError = writer.lastError();
        emit reportFinished(false, outPath);
        return false;
    }

    reportProgress(progress, 100, "Done: " + outPath);
    emit reportFinished(true, outPath);
    return true;
}

} // namespace DVE
