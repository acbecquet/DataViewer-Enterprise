#include "ReportGenerator.h"
#include "../plotting/PlotEngine.h"
#include "../utils/AppTheme.h"
#include "../utils/ImageUtils.h"
#include "../pipeline/RegimeUtils.h"
#include <QDebug>
#include <QFile>
#include <QFileInfo>
#include <QDir>
#include <QDate>
#include <QRegularExpression>
#include <QImage>
#include <QImageReader>
#include <QBuffer>
#include <QSet>
#include <cmath>
#include <algorithm>

namespace {
// Fixed positions (inches) for the 3 report plots: TPM Trend, Avg TPM Bar, Draw Pressure.
struct PlotSlot { double x, y, w, h; };
const PlotSlot kPlotLayout[3] = {
    { 0.10, 3.25, 4.32, 3.20 },
    { 4.51, 3.25, 4.32, 3.20 },
    { 8.91, 3.25, 4.32, 3.20 },
};
QVector<DVE::SlideImage> layoutPlots(QVector<QByteArray> pngs) {
    QVector<DVE::SlideImage> imgs;
    for (int pi = 0; pi < pngs.size() && pi < 3; ++pi) {
        DVE::SlideImage img; img.pngData = std::move(pngs[pi]);
        img.x = kPlotLayout[pi].x; img.y = kPlotLayout[pi].y;
        img.w = kPlotLayout[pi].w; img.h = kPlotLayout[pi].h;
        imgs.append(std::move(img));
    }
    return imgs;
}
} // namespace

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
QColor ReportGenerator::lifetimeBarColor(const QColor& fileBase, int sampleIdx, int totalSamplesInFile)
{
    return AppTheme::shade(fileBase, sampleIdx, totalSamplesInFile);
}

// ──────────────────────────────────────────────────────────────────────────────
int ReportGenerator::regimeSlideCount(const SheetResult& sheet)
{
    // Mirrors emitSheetContentSlides' branching (both derive from the same
    // RegimeUtils calls): old/no-regime -> 1 slide; else 1 per unique regime.
    if (!DVE::RegimeUtils::sheetHasRegimeData(sheet)) return 1;   // old / no per-row regime
    return DVE::RegimeUtils::uniqueRegimeKeys(sheet).size();
}

// Emit one content slide per unique per-row regime key (filtered + recomputed,
// including an "(unspecified)" slide when some rows are blank so no rows drop),
// or a single unchanged slide for old / no-regime sheets. `sheet` must already
// be empty-sample-filtered by the caller.
void ReportGenerator::emitSheetContentSlides(PptxWriter& writer, const SheetResult& sheet,
                                             const ReportConfig& config, bool includeBarChart)
{
    if (!DVE::RegimeUtils::sheetHasRegimeData(sheet)) {
        SlideTable tbl = buildTable(sheet, config);
        QVector<QByteArray> pngs = config.includePlots ? buildPlots(sheet, includeBarChart)
                                                       : QVector<QByteArray>{};
        writer.addContentSlide(sheet.sheetName, tbl, layoutPlots(std::move(pngs)));
        return;
    }
    for (const QString& key : DVE::RegimeUtils::uniqueRegimeKeys(sheet)) {
        const SheetResult fs = DVE::RegimeUtils::filterByRegime(sheet, key);
        if (!fs.hasSamples()) continue;
        SlideTable tbl = buildTable(fs, config);
        QVector<QByteArray> pngs = config.includePlots ? buildPlots(fs, includeBarChart)
                                                       : QVector<QByteArray>{};
        writer.addContentSlide(sheet.sheetName + QStringLiteral(" – ") + key,
                               tbl, layoutPlots(std::move(pngs)));
    }
}

// ──────────────────────────────────────────────────────────────────────────────
QVector<QByteArray> ReportGenerator::buildPlots(const SheetResult& sheet, bool includeBarChart)
{
    qDebug() << "buildPlots: entry, samples:" << sheet.samples.size() << "includeBar:" << includeBarChart;
    QVector<QByteArray> plots;
    const QVector<QColor> palette = AppTheme::seriesColors(sheet.samples.size());

    // Plot 1: TPM Trend — one series per sample, matching the GUI rendering exactly.
    {
        QVector<PlotSeries> series;
        int colorIdx = 0;
        for (const SampleResult& sr : sheet.samples) {
            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
            ps.color     = palette[colorIdx++];  // advance per sample → color stays stable across this sheet's plots
            ps.drawLine  = true;
            ps.drawDots  = true;                                  // markers always on
            ps.lineWidth = 2;

            for (const DataRow& row : sr.rows) {
                if (row.beforeWeight == 0.0 || row.afterWeight == 0.0) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.tpm);
            }
            ps.dotRadius = adaptiveDotRadius(ps.x.size());        // size from this series
            if (!ps.x.isEmpty()) series.append(std::move(ps));
        }

        if (!series.isEmpty()) {
            PlotConfig cfg = reportPlotConfig();
            cfg.title      = sheet.sheetName + QStringLiteral(" \u2013 TPM Trend");
            cfg.xLabel     = "Cumulative Puffs";
            cfg.yLabel     = "TPM (mg)";
            cfg.width      = 800;
            cfg.height     = 480;
            cfg.showGrid   = true;
            cfg.showLegend = (series.size() > 1);
            cfg.autoScale  = false;
            // Standing axis rules: x from 0 to the last puff, y from 0 to max+1.
            PlotEngine::applyDataXRange(cfg, series);
            PlotEngine::applyAnchoredYRange(cfg, series);

            qDebug() << "buildPlots: rendering TPM trend, series:" << series.size();
            QPixmap pm  = PlotEngine::renderLinePlot(series, cfg);
            qDebug() << "buildPlots: TPM trend rendered, size:" << pm.size();
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty()) plots.append(std::move(png));
        }
    }

    // Plot 2: TPM Bar Chart (average TPM per sample)
    if (includeBarChart) {
        qDebug() << "buildPlots: rendering bar chart";
        QVector<QString> names;
        QVector<double> avgs, sdevs;
        for (const auto& s : sheet.samples) {
            names.append(s.sampleName.isEmpty() ? s.sampleID : s.sampleName);
            avgs.append(s.averageTPM);
            sdevs.append(s.stdDevTPM);
        }
        PlotConfig cfg = reportPlotConfig();
        cfg.title      = sheet.sheetName + QStringLiteral(" \u2013 Average TPM");
        cfg.yLabel     = "Avg TPM (mg)";
        cfg.width      = 800;
        cfg.height     = 480;
        // renderBarChart computes its own y range: 0 to max bar/error-bar top + 1.
        QPixmap pm = PlotEngine::renderBarChart(names, avgs, cfg, /*colors=*/{}, sdevs);
        qDebug() << "buildPlots: bar chart rendered, size:" << pm.size();
        QByteArray png = PlotEngine::toPng(pm, 150);
        if (!png.isEmpty()) plots.append(std::move(png));
    }

    qDebug() << "buildPlots: rendering draw pressure";
    // Plot 3: Draw Pressure Trend — one series per sample.
    {
        QVector<PlotSeries> series;
        int colorIdx = 0;
        for (const SampleResult& sr : sheet.samples) {
            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
            ps.color     = palette[colorIdx++];  // advance per sample → color stays stable across this sheet's plots
            ps.drawLine  = true;
            ps.drawDots  = true;
            ps.lineWidth = 2;
            for (const DataRow& row : sr.rows) {
                if (row.drawPressure == 0.0) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.drawPressure);
            }
            ps.dotRadius = adaptiveDotRadius(ps.x.size());
            if (!ps.x.isEmpty()) series.append(std::move(ps));
        }

        if (!series.isEmpty()) {
            PlotConfig cfg = reportPlotConfig();
            cfg.title      = sheet.sheetName + QStringLiteral(" \u2013 Draw Pressure");
            cfg.xLabel     = "Cumulative Puffs";
            cfg.yLabel     = "Draw Pressure (Pa)";
            cfg.width      = 800;
            cfg.height     = 480;
            cfg.autoScale = false;
            // Standing axis rules: x from 0 to the last puff, y from 0 to max+1.
            PlotEngine::applyDataXRange(cfg, series);
            PlotEngine::applyAnchoredYRange(cfg, series);
            cfg.showGrid   = true;
            cfg.showLegend = (series.size() > 1);
            QPixmap pm = PlotEngine::renderLinePlot(series, cfg);
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty()) plots.append(std::move(png));
        }
    }

    return plots;
}

// ──────────────────────────────────────────────────────────────────────────────
SlideTable ReportGenerator::buildTable(const SheetResult& sheet, const ReportConfig& config)
{
    qDebug() << "buildTable: enter, samples:" << sheet.samples.size();
    SlideTable tbl;

    // Check whether any sample has non-zero draw pressure data
    bool hasDrawPressure = false;
    for (int si = 0; si < sheet.samples.size(); ++si) {
        const SampleResult& s = sheet.samples[si];
        qDebug() << "  sample" << si << "rows:" << s.rows.size()
                 << "name:" << s.sampleName;
        for (const DataRow& dr : s.rows) {
            if (dr.drawPressure > 0.0) { hasDrawPressure = true; break; }
        }
        if (hasDrawPressure) break;
    }
    qDebug() << "buildTable: hasDrawPressure:" << hasDrawPressure;

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
    qDebug() << "buildTable: headers:" << tbl.headers.size() << "building rows...";

    for (int sIdx = 0; sIdx < sheet.samples.size(); ++sIdx) {
        const auto& s = sheet.samples[sIdx];
        QStringList row;
        for (int cIdx = 0; cIdx < tbl.headers.size(); ++cIdx) {
            const auto& col = tbl.headers[cIdx];
            qDebug() << "  row" << sIdx << "col" << cIdx << col.left(20);
            if      (col.contains("Sample", Qt::CaseInsensitive))    row << (s.sampleName.isEmpty() ? s.sampleID : s.sampleName);
            else if (col.contains("Media",  Qt::CaseInsensitive))    row << s.media;
            else if (col.contains("Visco",  Qt::CaseInsensitive))    row << QString::number(s.viscosity, 'f', 1);
            else if (col.contains("Puffing",Qt::CaseInsensitive))    row << s.puffingRegime;
            else if (col.contains("Voltage",Qt::CaseInsensitive) &&
                     col.contains("Resistance",Qt::CaseInsensitive)) {
                // Stacked V/R/P cell
                row << QString("%1\n%2\n%3")
                       .arg(QString::number(s.voltage, 'f', 2),
                            QString::number(s.resistance, 'f', 2),
                            QString::number(s.power, 'f', 2));
            }
            else if (col.contains("Voltage",Qt::CaseInsensitive))    row << QString::number(s.voltage, 'f', 2);
            else if (col.contains("Resist", Qt::CaseInsensitive))    row << QString::number(s.resistance, 'f', 2);
            else if (col.contains("Power",  Qt::CaseInsensitive))    row << QString::number(s.power, 'f', 2);
            else if (col.contains("Avg TPM",Qt::CaseInsensitive))    row << QString::number(s.averageTPM, 'f', 2);
            else if (col.contains("Std",    Qt::CaseInsensitive) &&
                     col.contains("Dev",    Qt::CaseInsensitive))    row << QString::number(s.stdDevTPM, 'f', 2);
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
                     col.contains("Consumed",Qt::CaseInsensitive))   row << QString::number(s.totalOilConsumed / 1000.0, 'f', 2);
            else if (col.contains("Initial",Qt::CaseInsensitive) &&
                     col.contains("Oil",    Qt::CaseInsensitive))    row << QString::number(s.initialOilMass, 'f', 2);
            else if (col.contains("Draw",   Qt::CaseInsensitive) &&
                     col.contains("Pressure",Qt::CaseInsensitive)) {
                // Average draw pressure (excludes zero/empty rows)
                double dpSum = 0.0;
                int    dpCount = 0;
                for (const DataRow& dr : s.rows)
                    if (dr.drawPressure > 0.0) { dpSum += dr.drawPressure; ++dpCount; }
                row << (dpCount == 0 ? QString("-")
                                     : QString::number(dpSum / dpCount, 'f', 2));
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
    qDebug() << "buildTable: rows built:" << tbl.rows.size();

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

    // Test Protocol slide
    {
        QStringList testNames;
        for (const SheetResult& sh : data.sheets)
            if (sh.hasSamples())
                testNames << sh.sheetName;

        const QVector<SopEntry> sopRows = loadSopRows(testNames);
        if (!sopRows.isEmpty() || !testNames.isEmpty()) {
            SlideTable t;
            t.headers = {"Test", "Objective", "Pass Criteria", "Equipment", "Quantity", "Est Duration"};
            QSet<QString> covered;
            for (const SopEntry& e : sopRows) {
                covered.insert(e.test.toLower());
                QString dur = QStringLiteral("1mL: %1 / 2mL: %2")
                                  .arg(e.estDuration1mL.isEmpty() ? "-" : e.estDuration1mL,
                                       e.estDuration2mL.isEmpty() ? "-" : e.estDuration2mL);
                t.rows.append({e.test, e.objective, e.passCriteria, e.equipment, e.quantity, dur});
            }
            for (const QString& n : testNames)
                if (!covered.contains(n.toLower()))
                    t.rows.append({n, "—", "—", "—", "—", "—"});
            if (!t.rows.isEmpty())
                writer.addTestProtocolSlide(t);
        }
    }

    // Test Overview slide
    {
        QStringList testNames;
        for (const SheetResult& sh : data.sheets)
            if (sh.hasSamples())
                testNames << sh.sheetName;
        const QString desc = QStringLiteral("Standard performance evaluation of %1 across %2 tests.")
                                 .arg(displayFileName).arg(testNames.size());
        writer.addTestOverviewSlide(desc, testNames);
    }

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

        emitSheetContentSlides(writer, sheet, config, /*includeBarChart=*/true);

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

    // Conclusions slide (always last)
    writer.addConclusionsSlide();

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

    reportProgress(progress, 50, "Building slides...");
    emitSheetContentSlides(writer, *target, config, /*includeBarChart=*/target->samples.size() > 1);

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

// ──────────────────────────────────────────────────────────────────────────────
int ReportGenerator::adaptiveDotRadius(int pointCount) const
{
    if (pointCount <= 30)  return 5;
    if (pointCount >= 150) return 2;
    const double t = (pointCount - 30) / 120.0;
    return static_cast<int>(std::round(5.0 - 3.0 * t));
}

// ──────────────────────────────────────────────────────────────────────────────
PlotConfig ReportGenerator::reportPlotConfig() const
{
    PlotConfig cfg;
    cfg.titleFont = QFont("Segoe UI", 18, QFont::Bold);
    cfg.axisFont  = QFont("Segoe UI", 18);
    cfg.labelFont = QFont("Segoe UI", 14);
    return cfg;
}

// ──────────────────────────────────────────────────────────────────────────────
QVector<SopEntry> ReportGenerator::loadSopRows(const QStringList& reportTestNames) const
{
    const QString xlsx = m_resourcePath +
        QStringLiteral("/templates/Standardized Test Template - December 2025.xlsx");
    QVector<SopEntry> all = SopLoader::load(xlsx);
    if (reportTestNames.isEmpty()) return all;

    QSet<QString> wantLower;
    for (const QString& n : reportTestNames) wantLower.insert(n.toLower());

    QVector<SopEntry> filtered;
    for (const SopEntry& e : all) {
        if (wantLower.contains(e.test.toLower()))
            filtered.append(e);
    }
    return filtered;
}

// ──────────────────────────────────────────────────────────────────────────────
bool ReportGenerator::generateCombinedFullReport(const QVector<FileResult>& files,
                                                  const ReportConfig& /*config*/,
                                                  const QString& outputPath,
                                                  ProgressFn progress)
{
    if (files.isEmpty()) {
        m_lastError = "No files to combine";
        return false;
    }

    PptxWriter writer;
    writer.setResourcePath(m_resourcePath);

    // 1. Cover
    QDate today = QDate::currentDate();
    QString dateStr = today.toString("MMMM d, yyyy");
    writer.addCoverSlide(QStringLiteral("Combined Standard Test Report"), dateStr);

    // 2. Test Protocol — union across files
    QStringList allTests;
    QSet<QString> seen;
    for (const FileResult& f : files) {
        for (const SheetResult& sh : f.sheets) {
            if (sh.hasSamples() && !seen.contains(sh.sheetName.toLower())) {
                seen.insert(sh.sheetName.toLower());
                allTests << sh.sheetName;
            }
        }
    }
    {
        const QVector<SopEntry> sopRows = loadSopRows(allTests);
        SlideTable t;
        t.headers = {"Test", "Objective", "Pass Criteria", "Equipment", "Quantity", "Est Duration"};
        QSet<QString> covered;
        for (const SopEntry& e : sopRows) {
            covered.insert(e.test.toLower());
            QString dur = QStringLiteral("1mL: %1 / 2mL: %2")
                              .arg(e.estDuration1mL.isEmpty() ? "-" : e.estDuration1mL,
                                   e.estDuration2mL.isEmpty() ? "-" : e.estDuration2mL);
            t.rows.append({e.test, e.objective, e.passCriteria, e.equipment, e.quantity, dur});
        }
        for (const QString& n : allTests)
            if (!covered.contains(n.toLower()))
                t.rows.append({n, "—", "—", "—", "—", "—"});
        if (!t.rows.isEmpty()) writer.addTestProtocolSlide(t);
    }

    // 3. Test Overview (combined)
    {
        const QString desc = QStringLiteral("Combined performance evaluation across %1 files and %2 unique tests.")
                                 .arg(files.size()).arg(allTests.size());
        writer.addTestOverviewSlide(desc, allTests);
    }

    // 4. Lifetime TPM Comparison (cross-file bar chart)
    {
        QVector<QString> labels;
        QVector<double>  values;
        QVector<QColor>  colors;

        // One distinct base hue per file that actually appears (files lacking a
        // Lifetime sheet are skipped). hueIdx advances only for appearing files,
        // so they take consecutive, maximally-separated palette slots → strongest
        // between-file contrast; shade() varies samples within each file.
        const QVector<QColor> fileHues = AppTheme::seriesColors(files.size());
        int hueIdx = 0;

        for (int fi = 0; fi < files.size(); ++fi) {
            const FileResult& f = files[fi];
            const SheetResult* lifetime = nullptr;
            for (const SheetResult& sh : f.sheets) {
                if (sh.sheetName.compare("Lifetime Test", Qt::CaseInsensitive) == 0) {
                    lifetime = &sh; break;
                }
            }
            if (!lifetime || !lifetime->hasSamples()) continue;

            int totalSamples = 0;
            for (const SampleResult& s : lifetime->samples)
                if (!s.rows.isEmpty()) ++totalSamples;

            int sIdx = 0;
            for (const SampleResult& s : lifetime->samples) {
                if (s.rows.isEmpty()) continue;
                QString name = s.sampleName.isEmpty() ? s.sampleID : s.sampleName;
                labels.append(name);
                values.append(s.averageTPM);
                colors.append(lifetimeBarColor(fileHues[hueIdx], sIdx, totalSamples));
                ++sIdx;
            }
            ++hueIdx;   // next appearing file → next consecutive palette slot
        }

        if (!labels.isEmpty()) {
            PlotConfig cfg = reportPlotConfig();
            cfg.title         = "Lifetime TPM Comparison";
            cfg.yLabel        = "Avg TPM (mg)";
            cfg.width         = 1200;
            cfg.height        = 600;
            // renderBarChart computes its own y range: 0 to max bar top + 1.
            cfg.labelFont     = QFont("Segoe UI", 18);  // 18pt for auditorium

            QPixmap pm = PlotEngine::renderBarChart(labels, values, cfg, colors, /*stdDev=*/{});
            QByteArray png = PlotEngine::toPng(pm, 150);
            if (!png.isEmpty()) {
                SlideImage img;
                img.pngData = png;
                img.x = 0.30; img.y = 1.10; img.w = 12.7; img.h = 6.20;
                writer.addImageSlide(QStringLiteral("Lifetime TPM Comparison"),
                                     QVector<SlideImage>{img});
            }
        }
    }

    // 5. Per-file sections
    const int totalFiles = files.size();
    for (int fi = 0; fi < totalFiles; ++fi) {
        const FileResult& f = files[fi];
        QString displayName = QFileInfo(f.fileName).completeBaseName();
        reportProgress(progress, 20 + (60 * fi / totalFiles),
                       "Adding section: " + displayName);

        writer.addSectionDividerSlide(displayName);

        QStringList fileTests;
        for (const SheetResult& sh : f.sheets)
            if (sh.hasSamples()) fileTests << sh.sheetName;
        const QString perFileDesc =
            QStringLiteral("Standard performance evaluation of %1 across %2 tests.")
                .arg(displayName).arg(fileTests.size());
        writer.addTestOverviewSlide(perFileDesc, fileTests);

        for (const SheetResult& origSheet : f.sheets) {
            if (!origSheet.hasSamples()) continue;
            SheetResult sheet = origSheet;
            sheet.samples.erase(
                std::remove_if(sheet.samples.begin(), sheet.samples.end(),
                               [](const SampleResult& s){ return s.rows.isEmpty(); }),
                sheet.samples.end());
            if (!sheet.hasSamples()) continue;

            emitSheetContentSlides(writer, sheet, ReportConfig{}, /*includeBarChart=*/true);

            // Image slides — same pattern as generateFullReport
            for (const SampleResult& sr : sheet.samples) {
                if (sr.imagePaths.isEmpty()) continue;
                QString sampleDisplay = sr.sampleName.isEmpty() ? sr.sampleID : sr.sampleName;
                const bool hasLayouts = (sr.imageLayouts.size() == sr.imagePaths.size());
                if (hasLayouts) {
                    QVector<SlideImage> slideImages;
                    for (int i = 0; i < sr.imagePaths.size(); ++i) {
                        QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                        if (data.isEmpty()) continue;
                        SlideImage si;
                        si.pngData = std::move(data);
                        const QRectF& r = sr.imageLayouts[i];
                        si.x = r.x(); si.y = r.y();
                        si.w = r.width(); si.h = r.height();
                        slideImages.append(std::move(si));
                    }
                    if (!slideImages.isEmpty())
                        writer.addImageSlide(sampleDisplay + " Images", slideImages);
                } else {
                    QVector<QByteArray> imgBytes;
                    for (int i = 0; i < sr.imagePaths.size(); ++i) {
                        QRectF crop = (i < sr.imageCrops.size()) ? sr.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sr.imagePaths[i], crop);
                        if (!data.isEmpty()) imgBytes.append(data);
                    }
                    if (!imgBytes.isEmpty())
                        writer.addImageSlide(sampleDisplay + " Images", imgBytes);
                }
            }
            QVector<QByteArray> sheetImgs = collectImages(sheet);
            if (!sheetImgs.isEmpty())
                writer.addImageSlide(sheet.sheetName + " – Photos", sheetImgs);
        }
    }

    // 6. Conclusions
    writer.addConclusionsSlide();

    reportProgress(progress, 95, "Saving combined report...");
    bool ok = writer.save(outputPath);
    if (!ok) { m_lastError = writer.lastError(); return false; }
    reportProgress(progress, 100, "Saved: " + outputPath);
    return true;
}

} // namespace DVE
