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
    QString   slideKey(int idx) const override;
    ReportSlideSpec buildSlide(int idx, const ReportLayout&,
                                const QSet<QString>&) const override;

    QVector<SampleRef> allSamples() const override;

    ReportLayout loadLayout() const override;
    void saveLayout(const ReportLayout&) override;

    [[nodiscard]] bool writePptx(const QString& outPath, const ReportLayout&,
                                  const QSet<QString>&, QString* errorOut) override;

    // Public so tests can verify the layout matches the legacy fast path.
    static ReportLayout computeDefaultLayout(const QVector<SensorySession>& sessions);

    // Mode-agnostic sensory PPTX renderer. Shared implementation used by both
    // the legacy fast path (SensoryPanel::generateCombinedPptx) and the new
    // IReportSource entry point (writePptx).
    //
    // Layout consultation: when `layout.contentSlides[contentKey]` has non-null
    // rects (or `layout.cumulative` for the cumulative slide), those positions
    // override the legacy in-line wrap-aware math. With a default-constructed
    // (empty) ReportLayout, every rect is null and the legacy in-line
    // positioning is used unchanged — this is the bit-for-bit equivalent of
    // the pre-Task-8 SensoryPanel::generateCombinedPptx output.
    //
    // Sample exclusion: skip every sample whose `"<sessionIdx>#<sampleIdx>"`
    // key (matching SampleRef::sampleId from allSamples()) is in
    // `excludedSamples`. Applied to the table, radar input, Highest/Lowest
    // computation, and cumulative device accumulation.
    //
    // Image-slide overrides (layout.imageSlides) and cover-slide overrides
    // (layout.coverTitle, layout.coverSubtitle) are intentionally not
    // consulted by this method — those are deferred to Phase 1B/1C.
    [[nodiscard]] static bool writeSensoryPptx(const QVector<SensorySession>& sessions,
                                                const ReportLayout& layout,
                                                const QSet<QString>& excludedSamples,
                                                const QString& outPath,
                                                QString* errorOut);

    // Build the "properties" textbox XML for a sensory content slide. When
    // rectOverride is non-null it positions the box (WYSIWYG from the Preview);
    // otherwise the box is anchored bottom-right with a height derived from the
    // line count. fontPt 0 = legacy 16 pt. Returns empty when propLines empty.
    [[nodiscard]] static QString buildPropertiesBoxXml(const QStringList& propLines,
                                                       const QRectF& rectOverride,
                                                       int fontPt,
                                                       double slideW,
                                                       double slideH);

private:
    struct SlideEntry { SlideKind kind; int sessionIdx; QString key; };
    void buildSlideIndex();

    // ── Shared content-builder helpers ───────────────────────────────────
    //
    // Extracted from writeSensoryPptx so buildSlide() (canvas) and
    // writeSensoryPptx (PPTX) produce the same data (with the same
    // exclusion semantics). PPTX-specific positioning math (rawTable.x/y/w/h,
    // chart placement) stays inline in writeSensoryPptx to preserve the
    // bit-for-bit legacy fast-path output (Task 8 invariant).
    //
    // Sample exclusion key format: "<sessionIdx>#<sampleIdx>"
    // (matches SampleRef::sampleId emitted by allSamples()).
    static void buildContentTable(const SensorySession& sess,
                                   int sessionIdx,
                                   const QSet<QString>& excludedSamples,
                                   QStringList& outHeaders,
                                   QVector<QStringList>& outRows);

    static QImage renderRadarImage(const SensorySession& sess,
                                    int sessionIdx,
                                    const QSet<QString>& excludedSamples);

    static QString buildPropertiesText(const SensorySession& sess,
                                        int sessionIdx,
                                        const QSet<QString>& excludedSamples);

    // When filterTestTitle is non-empty, only sessions whose testTitle matches
    // contribute to the aggregation. Empty means "all sessions" (legacy path).
    // sessIdx values used in excludedSamples keys remain the original indices
    // into `sessions`, so exclusion keys stay consistent with allSamples().
    static void buildCumulativeData(const QVector<SensorySession>& sessions,
                                     const QSet<QString>& excludedSamples,
                                     QStringList& outHeaders,
                                     QVector<QStringList>& outRows,
                                     SensorySession& outCumSession,
                                     const QString& filterTestTitle = QString());

    static QImage renderCumulativeRadarImage(const SensorySession& cumSession);

    QVector<SensorySession> m_sessions;
    DatabaseManager*        m_db;
    QVector<SlideEntry>     m_slides;
};

} // namespace DVE
