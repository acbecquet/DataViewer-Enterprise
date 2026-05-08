#include "PropertiesPanel.h"
#include <QVBoxLayout>
#include <QGridLayout>
#include <QLabel>
#include <QDoubleSpinBox>
#include <QSpinBox>
#include <QPushButton>

namespace DVE {

namespace {
QDoubleSpinBox* makeInchSpin(double minVal, double maxVal) {
    auto* s = new QDoubleSpinBox;
    s->setRange(minVal, maxVal);
    s->setDecimals(2);
    s->setSingleStep(0.10);
    s->setSuffix(QStringLiteral(" in"));
    return s;
}
} // namespace

PropertiesPanel::PropertiesPanel(QWidget* p) : QWidget(p) {
    auto* outer = new QVBoxLayout(this);
    outer->setContentsMargins(0, 0, 0, 0);

    outer->addWidget(new QLabel(QStringLiteral("<b>Properties</b>")));

    m_label = new QLabel(QStringLiteral("(no selection)"));
    outer->addWidget(m_label);

    auto* grid = new QGridLayout;
    m_x = makeInchSpin(0.0, 13.33);   // 16:9 slide width
    m_y = makeInchSpin(0.0,  7.50);   // 16:9 slide height
    m_w = makeInchSpin(0.10, 13.33);
    m_h = makeInchSpin(0.10,  7.50);
    grid->addWidget(new QLabel(QStringLiteral("X")), 0, 0); grid->addWidget(m_x, 0, 1);
    grid->addWidget(new QLabel(QStringLiteral("Y")), 1, 0); grid->addWidget(m_y, 1, 1);
    grid->addWidget(new QLabel(QStringLiteral("W")), 2, 0); grid->addWidget(m_w, 2, 1);
    grid->addWidget(new QLabel(QStringLiteral("H")), 3, 0); grid->addWidget(m_h, 3, 1);

    // Font-size row: only meaningful for TextItems; disabled when current
    // selection has no font (table/plot/image).
    m_fontLabel = new QLabel(QStringLiteral("Font"));
    m_fontSize = new QSpinBox;
    m_fontSize->setRange(6, 96);
    m_fontSize->setSuffix(QStringLiteral(" pt"));
    grid->addWidget(m_fontLabel, 4, 0); grid->addWidget(m_fontSize, 4, 1);
    outer->addLayout(grid);

    // Z-order buttons stacked vertically with arrow glyphs (▲ / ▼) so the
    // direction is visually obvious. Min height accommodates the larger glyph.
    m_forward  = new QPushButton(QStringLiteral("\xe2\x96\xb2 Bring Forward"));
    m_backward = new QPushButton(QStringLiteral("\xe2\x96\xbc Send Backward"));
    m_forward->setMinimumHeight(28);
    m_backward->setMinimumHeight(28);
    outer->addWidget(m_forward);
    outer->addWidget(m_backward);

    outer->addStretch();

    // Emit rectEdited when any spinbox commits its value.
    for (QDoubleSpinBox* sb : {m_x, m_y, m_w, m_h}) {
        connect(sb, &QDoubleSpinBox::editingFinished,
                this, &PropertiesPanel::emitRectIfReady);
    }
    connect(m_fontSize, &QSpinBox::editingFinished, this, [this]() {
        if (!m_currentId.isEmpty() && m_fontSize->isEnabled())
            emit fontSizeEdited(m_currentId, m_fontSize->value());
    });
    connect(m_forward,  &QPushButton::clicked, this, [this]() {
        if (!m_currentId.isEmpty()) emit bringForwardClicked(m_currentId);
    });
    connect(m_backward, &QPushButton::clicked, this, [this]() {
        if (!m_currentId.isEmpty()) emit sendBackwardClicked(m_currentId);
    });

    setControlsEnabled(false);
}

void PropertiesPanel::setSelectedItem(const QString& elementId, const QRectF& r,
                                        int fontPt) {
    m_currentId = elementId;
    m_label->setText(elementId.isEmpty()
                      ? QStringLiteral("(no selection)")
                      : QStringLiteral("Element: ") + elementId);
    // Block signals while populating to avoid spurious rectEdited emissions.
    for (QDoubleSpinBox* sb : {m_x, m_y, m_w, m_h}) sb->blockSignals(true);
    m_x->setValue(r.x());
    m_y->setValue(r.y());
    m_w->setValue(r.width());
    m_h->setValue(r.height());
    for (QDoubleSpinBox* sb : {m_x, m_y, m_w, m_h}) sb->blockSignals(false);

    // Font size: enabled only when caller passed fontPt > 0 (TextItems do, the
    // table/plot/image classes pass 0 to disable).
    m_fontSize->blockSignals(true);
    if (fontPt > 0) m_fontSize->setValue(fontPt);
    m_fontSize->blockSignals(false);
    const bool hasFont = (fontPt > 0);
    m_fontLabel->setEnabled(hasFont);
    m_fontSize->setEnabled(hasFont && !elementId.isEmpty());

    // If caller passed an empty id, keep the controls disabled — clearSelection()
    // is the canonical no-selection path, but be defensive against misuse.
    setControlsEnabled(!elementId.isEmpty());
}

void PropertiesPanel::clearSelection() {
    m_currentId.clear();
    m_label->setText(QStringLiteral("(no selection)"));
    for (QDoubleSpinBox* sb : {m_x, m_y, m_w, m_h}) {
        sb->blockSignals(true);
        sb->setValue(0.0);
        sb->blockSignals(false);
    }
    m_fontSize->blockSignals(true);
    m_fontSize->setValue(14);
    m_fontSize->blockSignals(false);
    setControlsEnabled(false);
}

void PropertiesPanel::setControlsEnabled(bool on) {
    m_x->setEnabled(on); m_y->setEnabled(on);
    m_w->setEnabled(on); m_h->setEnabled(on);
    // m_fontSize is governed separately by setSelectedItem's fontPt parameter
    // (it's a TextItem-only control). Disabling here keeps it off when nothing
    // is selected; setSelectedItem re-enables it when appropriate.
    if (!on) m_fontSize->setEnabled(false);
    m_forward->setEnabled(on); m_backward->setEnabled(on);
}

void PropertiesPanel::emitRectIfReady() {
    if (m_currentId.isEmpty()) return;
    emit rectEdited(m_currentId,
                    QRectF(m_x->value(), m_y->value(),
                            m_w->value(), m_h->value()));
}

} // namespace DVE
