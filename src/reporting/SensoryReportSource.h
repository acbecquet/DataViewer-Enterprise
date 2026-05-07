#pragma once
#include "IReportSource.h"
#include "pipeline/SensoryData.h"

namespace DVE {

class DatabaseManager;

class SensoryReportSource : public IReportSource {
public:
    SensoryReportSource(QVector<SensorySession> sessions,
                        DatabaseManager* db);

    QString modeId() const override { return QStringLiteral("sensory"); }
    QString sourceLabel() const override;

    int slideCount() const override;
    SlideKind slideKind(int idx) const override;
    ReportSlideSpec buildSlide(int idx, const ReportLayout&,
                                const QSet<QString>&) const override;

    QVector<SampleRef> allSamples() const override;

    ReportLayout loadLayout() const override;
    void saveLayout(const ReportLayout&) override;

    [[nodiscard]] bool writePptx(const QString& outPath, const ReportLayout&,
                                  const QSet<QString>&, QString* errorOut) override;

    // Public so tests can verify the layout matches the legacy fast path.
    static ReportLayout computeDefaultLayout(const QVector<SensorySession>& sessions);

    // Shared implementation used by both the legacy fast path
    // (SensoryPanel::generateCombinedPptx) and the new IReportSource entry
    // point (writePptx). When called with computeDefaultLayout(sessions) and
    // an empty exclusion set, output is bit-for-bit identical to the legacy
    // generateCombinedPptx output. Layout overrides (per-slide table/radar
    // rects) and sample exclusion are honored on the per-session content
    // slides and on the cumulative slide. Image-slide layouts are NOT
    // parameterized in this task (Phase 1B/2 concern); image slides always
    // use legacy positions.
    [[nodiscard]] static bool writeSensoryPptx(const QVector<SensorySession>& sessions,
                                                const ReportLayout& layout,
                                                const QSet<QString>& excludedSamples,
                                                const QString& outPath,
                                                QString* errorOut);

private:
    struct SlideEntry { SlideKind kind; int sessionIdx; QString key; };
    void buildSlideIndex();

    QVector<SensorySession> m_sessions;
    DatabaseManager*        m_db;
    QVector<SlideEntry>     m_slides;
};

} // namespace DVE
