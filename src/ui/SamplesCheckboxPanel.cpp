#include "SamplesCheckboxPanel.h"
#include <QVBoxLayout>
#include <QCheckBox>
#include <QScrollArea>
#include <QFrame>
#include <QLabel>

namespace DVE {

SamplesCheckboxPanel::SamplesCheckboxPanel(const QVector<SampleRef>& refs, QWidget* p)
    : QWidget(p) {
    // Prefer a narrow column so long session labels wrap rather than pushing
    // the canvas right; allow growth (the ScrollHost catches residual overflow).
    setMinimumWidth(200);

    auto* outer = new QVBoxLayout(this);
    outer->setContentsMargins(0, 0, 0, 0);
    outer->addWidget(new QLabel("<b>Samples</b>"));

    auto* scroll = new QScrollArea;
    scroll->setWidgetResizable(true);
    auto* host = new QWidget;
    auto* hostL = new QVBoxLayout(host);
    hostL->setContentsMargins(2, 2, 2, 2);

    // groups: sessionLabel -> the QVBoxLayout that holds that group's checkboxes.
    // We use a thin-frame container with a wrappable QLabel header instead of
    // QGroupBox — QGroupBox titles don't word-wrap, so long session labels
    // get clipped in narrow columns. The frame gives the same visual grouping
    // without the title-wrap limitation.
    QHash<QString, QVBoxLayout*> groups;
    for (const SampleRef& r : refs) {
        QVBoxLayout* gbLayout = groups.value(r.sessionLabel, nullptr);
        if (!gbLayout) {
            auto* group = new QFrame;
            group->setFrameShape(QFrame::StyledPanel);
            auto* groupL = new QVBoxLayout(group);
            groupL->setContentsMargins(6, 6, 6, 6);
            groupL->setSpacing(2);
            auto* header = new QLabel(QStringLiteral("<b>%1</b>").arg(r.sessionLabel));
            header->setWordWrap(true);                // long labels wrap to new lines
            header->setTextInteractionFlags(Qt::TextSelectableByMouse);
            groupL->addWidget(header);
            gbLayout = groupL;                         // checkboxes go below the label
            groups.insert(r.sessionLabel, gbLayout);
            hostL->addWidget(group);
        }
        auto* box = new QCheckBox(r.displayName);
        box->setChecked(true);
        const QString id = r.sampleId;
        connect(box, &QCheckBox::toggled, this, [this, id](bool on) {
            emit sampleToggled(id, on);
        });
        m_boxes.insert(id, box);
        gbLayout->addWidget(box);
    }
    hostL->addStretch();
    scroll->setWidget(host);
    outer->addWidget(scroll, 1);
}

QSet<QString> SamplesCheckboxPanel::excludedSampleIds() const {
    QSet<QString> out;
    for (auto it = m_boxes.cbegin(); it != m_boxes.cend(); ++it)
        if (!it.value()->isChecked()) out.insert(it.key());
    return out;
}

} // namespace DVE
