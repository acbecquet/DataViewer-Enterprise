#pragma once

#include "../pipeline/ReportData.h"
#include "../plotting/PlotEngine.h"
#include "../utils/SopLoader.h"
#include "PptxWriter.h"
#include <QString>
#include <QObject>
#include <QColor>
#include <functional>

namespace DVE {

using ProgressFn = std::function<void(int /*percent*/, const QString& /*message*/)>;

class ReportGenerator : public QObject {
    Q_OBJECT

public:
    explicit ReportGenerator(QObject* parent = nullptr);
    ~ReportGenerator() override = default;

    // Set resource directory (for branding images)
    void setResourcePath(const QString& path);

    // Generate a full report: all sheets → single .pptx (+ optional Excel)
    bool generateFullReport(const FileResult& data,
                            const ReportConfig& config,
                            ProgressFn progress = nullptr);

    // Generate a single-sheet test report
    bool generateTestReport(const FileResult& data,
                            const QString& sheetName,
                            const ReportConfig& config,
                            ProgressFn progress = nullptr);

    // Generate a combined multi-file report
    bool generateCombinedFullReport(const QVector<FileResult>& files,
                                    const ReportConfig& config,
                                    const QString& outputPath,
                                    ProgressFn progress = nullptr);

    /// Bar color for the Lifetime Comparison slide. The file's distinct base
    /// hue (from AppTheme::seriesColors) is shaded per sample so every bar in
    /// the chart is a unique, projector-safe color while staying grouped by file.
    static QColor lifetimeBarColor(const QColor& fileBase, int sampleIdx, int totalSamplesInFile);

    // Number of content slides a sheet emits: 1 per unique per-row regime,
    // or 1 for old/no-regime sheets. Pure function; a test seam that mirrors
    // emitSheetContentSlides' branching (used by tst_reportgenerator).
    static int regimeSlideCount(const SheetResult& sheet);

    QString lastError() const { return m_lastError; }

    // Test-only access wrappers — exposed publicly so unit tests can call private
    // helpers without befriending Qt Test classes.
    int adaptiveDotRadiusForTesting(int n) const { return adaptiveDotRadius(n); }
    QVector<SopEntry> loadSopRowsForTesting(const QStringList& reportTests) const
        { return loadSopRows(reportTests); }
    SlideTable buildTableForTesting(const SheetResult& s, const ReportConfig& c)
        { return buildTable(s, c); }

signals:
    void progressChanged(int percent, const QString& message);
    void reportFinished(bool success, const QString& filePath);

private:
    QString m_resourcePath;
    QString m_lastError;

    // Build plots for a sheet → PNG bytes, one per plot type
    QVector<QByteArray> buildPlots(const SheetResult& sheet, bool includeBarChart = true);

    void emitSheetContentSlides(PptxWriter& writer, const SheetResult& sheet,
                                const ReportConfig& config, bool includeBarChart);

    // Build PPTX table data from sheet result
    SlideTable buildTable(const SheetResult& sheet, const ReportConfig& config);

    // Build image list for image slide
    QVector<QByteArray> collectImages(const SheetResult& sheet);

    void reportProgress(ProgressFn fn, int pct, const QString& msg);
    void logDebug(const QString& msg) const;

    int adaptiveDotRadius(int pointCount) const;
    PlotConfig reportPlotConfig() const;
    QVector<SopEntry> loadSopRows(const QStringList& reportTestNames) const;
};

} // namespace DVE
