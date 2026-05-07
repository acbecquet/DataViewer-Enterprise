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
ReportLayout SensoryReportSource::computeDefaultLayout(const QVector<SensorySession>&) { return {}; }

} // namespace DVE
