#include "HeaderEditDialog.h"

#include <QVBoxLayout>
#include <QFormLayout>
#include <QGroupBox>
#include <QDialogButtonBox>
#include <QPushButton>
#include <QStyle>
#include <QGraphicsDropShadowEffect>
#include "../utils/AppTheme.h"
#include "../widgets/ScrollHost.h"

namespace DVE {

HeaderEditDialog::HeaderEditDialog(const SampleResult& s, QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle(QString("Edit Headers \u2013 %1").arg(
        s.sampleName.isEmpty() ? s.sampleID : s.sampleName));
    setMinimumSize(360, 300);

    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(16, 16, 16, 16);

    auto* shadow = new QGraphicsDropShadowEffect(this);
    shadow->setBlurRadius(16);
    shadow->setOffset(0, 4);
    shadow->setColor(QColor(0, 0, 0, 30));
    setGraphicsEffect(shadow);

    // Scrollable content (both-direction as-needed via the shared ScrollHost)
    QWidget* content = new QWidget;
    QVBoxLayout* contentLayout = new QVBoxLayout(content);

    // Identification group
    QGroupBox* idGroup = new QGroupBox("Identification", content);
    QFormLayout* idForm = new QFormLayout(idGroup);
    idForm->setVerticalSpacing(12);
    idForm->setHorizontalSpacing(12);
    idForm->setLabelAlignment(Qt::AlignRight | Qt::AlignVCenter);
    m_sampleID = new QLineEdit(s.sampleID, idGroup);
    m_tester   = new QLineEdit(s.tester,   idGroup);
    idForm->addRow("Sample ID:",  m_sampleID);
    idForm->addRow("Tester:",     m_tester);
    contentLayout->addWidget(idGroup);

    // Device parameters group
    QGroupBox* devGroup = new QGroupBox("Device Parameters", content);
    QFormLayout* devForm = new QFormLayout(devGroup);
    devForm->setVerticalSpacing(12);
    devForm->setHorizontalSpacing(12);
    devForm->setLabelAlignment(Qt::AlignRight | Qt::AlignVCenter);
    m_media        = new QLineEdit(s.media, devGroup);
    m_viscosity    = new QLineEdit(s.viscosity > 0 ? QString::number(s.viscosity, 'f', 1) : QString(), devGroup);
    m_resistance   = new QLineEdit(s.resistance > 0 ? QString::number(s.resistance, 'f', 3) : QString(), devGroup);
    m_voltage      = new QLineEdit(s.voltage > 0 ? QString::number(s.voltage, 'f', 2) : QString(), devGroup);
    devForm->addRow("Media:",          m_media);
    devForm->addRow("Viscosity (cP):", m_viscosity);
    devForm->addRow("Resistance (\u03a9):", m_resistance);
    devForm->addRow("Voltage (V):",    m_voltage);
    contentLayout->addWidget(devGroup);

    // Test parameters group
    QGroupBox* testGroup = new QGroupBox("Test Parameters", content);
    QFormLayout* testForm = new QFormLayout(testGroup);
    testForm->setVerticalSpacing(12);
    testForm->setHorizontalSpacing(12);
    testForm->setLabelAlignment(Qt::AlignRight | Qt::AlignVCenter);
    m_puffingRegime = new QLineEdit(s.puffingRegime, testGroup);
    m_oilMass       = new QLineEdit(s.initialOilMass > 0 ? QString::number(s.initialOilMass, 'f', 3) : QString(), testGroup);
    testForm->addRow("Puffing Regime:",    m_puffingRegime);
    testForm->addRow("Initial Oil (g):",   m_oilMass);
    contentLayout->addWidget(testGroup);

    ScrollHost* scroll = ScrollHost::wrap(content);
    mainLayout->addWidget(scroll);

    QDialogButtonBox* box = new QDialogButtonBox(
        QDialogButtonBox::Save | QDialogButtonBox::Cancel, this);
    connect(box, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(box, &QDialogButtonBox::rejected, this, &QDialog::reject);
    if (auto* saveBtn = box->button(QDialogButtonBox::Save)) {
        saveBtn->setProperty("primary", true);
        saveBtn->style()->unpolish(saveBtn);
        saveBtn->style()->polish(saveBtn);
    }
    mainLayout->addWidget(box);
}

HeaderEditDialog::HeaderData HeaderEditDialog::headerData() const
{
    HeaderData d;
    d.sampleID      = m_sampleID->text().trimmed();
    d.tester        = m_tester->text().trimmed();
    d.media         = m_media->text().trimmed();
    d.viscosity     = m_viscosity->text().trimmed();
    d.resistance    = m_resistance->text().trimmed();
    d.voltage       = m_voltage->text().trimmed();
    d.puffingRegime = m_puffingRegime->text().trimmed();
    d.initialOilMass = m_oilMass->text().trimmed();
    return d;
}

} // namespace DVE
