#include "SensoryReportSource.h"
#include <QHash>

namespace DVE {

SensoryReportSource::SensoryReportSource(QVector<SensorySession> sessions,
                                          DatabaseManager* db)
    : m_sessions(std::move(sessions)), m_db(db)
{
    buildSlideIndex();
}

void SensoryReportSource::buildSlideIndex()
{
    m_slides.clear();
    if (m_sessions.isEmpty()) return;

    // Cover slide
    m_slides.push_back({ SlideKind::Cover, -1, QStringLiteral("cover") });

    // Group sessions by tester (mirrors generateCombinedPptx grouping)
    QHash<QString, QVector<int>> byTester;
    QStringList testerOrder;
    for (int i = 0; i < m_sessions.size(); ++i) {
        const QString& t = m_sessions[i].testerName;
        if (!byTester.contains(t)) testerOrder.append(t);
        byTester[t].append(i);
    }
    for (const QString& tester : testerOrder) {
        const QVector<int>& idxs = byTester[tester];
        m_slides.push_back({ SlideKind::Divider, idxs.first(),
                              QStringLiteral("divider_%1").arg(idxs.first()) });
        for (int sIdx : idxs) {
            m_slides.push_back({ SlideKind::Content, sIdx,
                                  QStringLiteral("content_%1").arg(sIdx) });
            if (!m_sessions[sIdx].imagePaths.isEmpty())
                m_slides.push_back({ SlideKind::Image, sIdx,
                                      QStringLiteral("image_%1").arg(sIdx) });
        }
    }

    // Cumulative summary (only when 2+ sessions)
    if (m_sessions.size() >= 2)
        m_slides.push_back({ SlideKind::Cumulative, -1, QStringLiteral("cumulative") });
}

QString SensoryReportSource::sourceLabel() const
{
    if (m_sessions.size() == 1)
        return m_sessions.first().testTitle;
    return QStringLiteral("%1 sessions").arg(m_sessions.size());
}

int SensoryReportSource::slideCount() const { return m_slides.size(); }
SlideKind SensoryReportSource::slideKind(int idx) const { return m_slides[idx].kind; }

QVector<SampleRef> SensoryReportSource::allSamples() const
{
    QVector<SampleRef> out;
    for (int i = 0; i < m_sessions.size(); ++i) {
        const auto& sess = m_sessions[i];
        const QString slideKey = QStringLiteral("content_%1").arg(i);
        for (int sIdx = 0; sIdx < sess.samples.size(); ++sIdx) {
            const auto& samp = sess.samples[sIdx];
            SampleRef r;
            r.slideKey = slideKey;
            r.sampleId = QStringLiteral("%1#%2").arg(i).arg(sIdx);
            r.displayName = samp.name.isEmpty()
                ? QStringLiteral("Sample %1").arg(sIdx + 1) : samp.name;
            r.sessionLabel = sess.testTitle.isEmpty()
                ? QStringLiteral("Session %1").arg(i + 1) : sess.testTitle;
            out.append(r);
        }
    }
    return out;
}

// ── Stubs (filled in by Tasks 5-8) ───────────────────────────────────────
ReportSlideSpec SensoryReportSource::buildSlide(int, const ReportLayout&,
                                                  const QSet<QString>&) const { return {}; }
ReportLayout SensoryReportSource::loadLayout() const { return {}; }
void SensoryReportSource::saveLayout(const ReportLayout&) {}
bool SensoryReportSource::writePptx(const QString&, const ReportLayout&,
                                     const QSet<QString>&, QString*) { return false; }
ReportLayout SensoryReportSource::computeDefaultLayout(const QVector<SensorySession>& sessions)
{
    constexpr double slideW = 13.33;
    constexpr double slideH = 7.5;
    constexpr double tableX = 0.32;
    constexpr double tableW = 12.7;
    constexpr double tableY = 0.75;
    constexpr double gap    = 0.10;
    constexpr double propsW = 3.17;
    constexpr double propsH = 2.00;

    auto contentForSession = [&](const SensorySession& s) {
        ContentSlideLayout c;

        // Title spans full slide width above the table
        c.title = QRectF(tableX, 0.10, tableW, 0.55);

        // Table height grows with sample count: header (~0.5") + each row ~0.33"
        const double tableH = 0.50 + s.samples.size() * (1.0 / 3.0);
        c.table = QRectF(tableX, tableY, tableW, tableH);

        // Radar centered horizontally below the table, square aspect, fills
        // the remaining vertical space (clamped to leave room for the
        // properties textbox at bottom-right).
        const double tableBottom = tableY + tableH + gap;
        const double availH      = slideH - gap - tableBottom;
        const double radarH      = qMax(0.5, availH);
        const double radarW      = radarH;                   // square (1:1)
        const double radarX      = (slideW - radarW) / 2.0;
        c.radar = QRectF(radarX, tableBottom, radarW, radarH);

        // Properties textbox anchored bottom-right
        c.propertiesBox.rect = QRectF(slideW - propsW - 0.05,
                                       slideH - propsH - 0.05,
                                       propsW, propsH);
        c.propertiesBox.text.clear();   // filled by buildSlide() at render time
        return c;
    };

    ReportLayout layout;

    // Cover slide: centered title, subtitle below
    layout.coverTitle    = QRectF(0.5, 2.5, slideW - 1.0, 1.5);
    layout.coverSubtitle = QRectF(0.5, 4.2, slideW - 1.0, 0.8);

    for (int i = 0; i < sessions.size(); ++i) {
        const QString contentKey = QStringLiteral("content_%1").arg(i);
        const QString dividerKey = QStringLiteral("divider_%1").arg(i);
        layout.contentSlides[contentKey] = contentForSession(sessions[i]);
        layout.dividerTitles[dividerKey] = QRectF(0.5, slideH/2 - 0.75,
                                                   slideW - 1.0, 1.5);

        if (!sessions[i].imagePaths.isEmpty()) {
            ImageSlideLayout img;
            // Default grid: up to 4 images per row, equal cells with 0.25" margin
            const int n = sessions[i].imagePaths.size();
            const int cols = qMin(4, n);
            const int rows = (n + cols - 1) / cols;
            const double cellW = (slideW - 0.5) / cols;
            const double cellH = (slideH - 0.5) / qMax(1, rows);
            for (int k = 0; k < n; ++k) {
                const int r = k / cols;
                const int c = k % cols;
                img.imageLayouts.append(QRectF(0.25 + c * cellW,
                                                0.25 + r * cellH,
                                                cellW - 0.1, cellH - 0.1));
                img.imageCrops.append(QRectF(0, 0, 1, 1));
            }
            layout.imageSlides[QStringLiteral("image_%1").arg(i)] = img;
        }
    }

    // Cumulative slide layout — same shape as content, only when 2+ sessions
    if (sessions.size() >= 2)
        layout.cumulative = contentForSession(sessions.first());

    return layout;
}

} // namespace DVE
