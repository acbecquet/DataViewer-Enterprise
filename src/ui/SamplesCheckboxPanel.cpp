#include "SamplesCheckboxPanel.h"
#include <QVBoxLayout>
#include <QGroupBox>
#include <QCheckBox>
#include <QScrollArea>
#include <QLabel>

namespace DVE {

SamplesCheckboxPanel::SamplesCheckboxPanel(const QVector<SampleRef>& refs, QWidget* p)
    : QWidget(p), m_refs(refs) {
    auto* outer = new QVBoxLayout(this);
    outer->setContentsMargins(0, 0, 0, 0);
    outer->addWidget(new QLabel("<b>Samples</b>"));

    auto* scroll = new QScrollArea;
    scroll->setWidgetResizable(true);
    auto* host = new QWidget;
    auto* hostL = new QVBoxLayout(host);
    hostL->setContentsMargins(2, 2, 2, 2);

    QHash<QString, QGroupBox*> groups;
    QStringList groupOrder;
    for (const SampleRef& r : refs) {
        if (!groups.contains(r.sessionLabel)) {
            auto* gb = new QGroupBox(r.sessionLabel);
            new QVBoxLayout(gb);     // owned by gb
            groups[r.sessionLabel] = gb;
            hostL->addWidget(gb);
            groupOrder << r.sessionLabel;
        }
        auto* box = new QCheckBox(r.displayName);
        box->setChecked(true);
        const QString id = r.sampleId;
        connect(box, &QCheckBox::toggled, this, [this, id](bool on) {
            emit sampleToggled(id, on);
        });
        m_boxes.insert(id, box);
        static_cast<QVBoxLayout*>(groups[r.sessionLabel]->layout())->addWidget(box);
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
