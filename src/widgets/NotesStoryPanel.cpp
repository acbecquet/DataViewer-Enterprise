#include "NotesStoryPanel.h"
#include <QVBoxLayout>
#include <QScrollArea>
#include <QScrollBar>
#include <QFrame>

namespace DVE {

NotesStoryPanel::NotesStoryPanel(QWidget* parent) : QWidget(parent) {
    auto* outer = new QVBoxLayout(this);
    outer->setContentsMargins(0, 0, 0, 0);
    m_scroll = new QScrollArea(this);
    m_scroll->setWidgetResizable(true);
    m_scroll->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_scroll->setVerticalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    m_scroll->setFrameShape(QFrame::NoFrame);
    m_content = new QWidget(m_scroll);
    m_vbox = new QVBoxLayout(m_content);
    m_vbox->setContentsMargins(8, 8, 8, 8);
    m_vbox->setSpacing(8);
    m_vbox->addStretch(1);
    m_scroll->setWidget(m_content);
    outer->addWidget(m_scroll);
    setMinimumWidth(280);
}

void NotesStoryPanel::clear() { m_sample = SampleResult{}; m_excluded.clear(); rebuild(); }

void NotesStoryPanel::setSample(const SampleResult& sample, const QSet<int>& excluded, bool perRowRegime) {
    m_sample = sample; m_excluded = excluded; m_perRowRegime = perRowRegime; rebuild();
}

void NotesStoryPanel::rebuild() {
    m_cardByRow.clear();
    QLayoutItem* item;
    while ((item = m_vbox->takeAt(0)) != nullptr) {
        if (QWidget* w = item->widget()) w->deleteLater();
        delete item;
    }
    const QVector<StorySegment> segs = NotesStory::build(m_sample, m_excluded);
    for (int i = 0; i < segs.size(); ++i) {
        const StorySegment& seg = segs[i];
        QWidget* w = (seg.kind == StorySegment::Note)
            ? buildNoteCard(m_sample, seg.rowIndex, seg.excluded, m_perRowRegime)
            : buildSummaryBar(m_sample, seg.summary, i, m_perRowRegime);
        m_vbox->addWidget(w);
        if (seg.kind == StorySegment::Note) m_cardByRow.insert(seg.rowIndex, w);
    }
    m_vbox->addStretch(1);
}

// ── Stubs (implemented in Tasks 3-5) ─────────────────────────────────────────────────
QWidget* NotesStoryPanel::buildNoteCard(const SampleResult&, int, bool, bool) {
    return new QWidget;
}
QWidget* NotesStoryPanel::buildSummaryBar(const SampleResult&, const StorySummary&, int, bool) {
    return new QWidget;
}
void NotesStoryPanel::highlightRow(int) {}

} // namespace DVE
