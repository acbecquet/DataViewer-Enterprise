#pragma once

#include "../pipeline/ReportData.h"
#include "../plotting/PlotEngine.h"
#include "PptxWriter.h"
#include <QString>
#include <QObject>
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

    QString lastError() const { return m_lastError; }

    // Test-only access wrappers — exposed publicly so unit tests can call private
    // helpers without befriending Qt Test classes.
    bool isLongPuffForTesting(const SheetResult& s) const { return isLongPuff(s); }
    double computeTpmYMaxForTesting(const SheetResult& s) const { return computeTpmYMax(s); }
    int adaptiveDotRadiusForTesting(int n) const { return adaptiveDotRadius(n); }

signals:
    void progressChanged(int percent, const QString& message);
    void reportFinished(bool success, const QString& filePath);

private:
    QString m_resourcePath;
    QString m_lastError;

    // Build plots for a sheet → PNG bytes, one per plot type
    QVector<QByteArray> buildPlots(const SheetResult& sheet, bool includeBarChart = true);

    // Build PPTX table data from sheet result
    SlideTable buildTable(const SheetResult& sheet, const ReportConfig& config);

    // Build image list for image slide
    QVector<QByteArray> collectImages(const SheetResult& sheet);

    void reportProgress(ProgressFn fn, int pct, const QString& msg);
    void logDebug(const QString& msg) const;

    bool isLongPuff(const SheetResult& sheet) const;
    double computeTpmYMax(const SheetResult& sheet) const;
    int adaptiveDotRadius(int pointCount) const;
    PlotConfig reportPlotConfig() const;
};

} // namespace DVE
