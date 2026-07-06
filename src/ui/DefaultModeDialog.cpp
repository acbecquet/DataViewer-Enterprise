#include "DefaultModeDialog.h"

#include <QLabel>
#include <QPushButton>
#include <QSettings>
#include <QVBoxLayout>

namespace DVE {

namespace {
const QString kKey = QStringLiteral("ui/defaultMode");

QSettings store()
{
    return QSettings(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
}
} // namespace

DefaultModeDialog::DefaultModeDialog(QWidget* parent)
    : QDialog(parent)
{
    setWindowTitle(tr("Default Mode"));
    setModal(true);

    auto* layout = new QVBoxLayout(this);
    layout->setSpacing(8);

    auto* prompt = new QLabel(tr("Pick a default mode:"), this);
    layout->addWidget(prompt);

    struct Choice { const char* objectName; QString label; ReportMode mode; };
    const Choice choices[] = {
        { "tpmButton",      tr("TPM"),              ReportMode::Tpm },
        { "sensoryButton",  tr("Sensory"),          ReportMode::Sensory },
        { "detailedButton", tr("Detailed Sensory"), ReportMode::DetailedSensory },
    };
    for (const Choice& c : choices) {
        auto* btn = new QPushButton(c.label, this);
        btn->setObjectName(QLatin1String(c.objectName));
        btn->setMinimumHeight(32);
        layout->addWidget(btn);
        const ReportMode mode = c.mode;
        connect(btn, &QPushButton::clicked, this, [this, mode]() {
            m_selected = mode;
            setSavedDefaultMode(mode);
            accept();
        });
    }
}

bool DefaultModeDialog::hasSavedDefaultMode()
{
    return store().contains(kKey);
}

ReportMode DefaultModeDialog::savedDefaultMode()
{
    bool ok = false;
    const int v = store().value(kKey).toInt(&ok);
    if (!ok || v < int(ReportMode::Tpm) || v > int(ReportMode::DetailedSensory))
        return ReportMode::Tpm;
    return ReportMode(v);
}

void DefaultModeDialog::setSavedDefaultMode(ReportMode mode)
{
    store().setValue(kKey, int(mode));
}

} // namespace DVE
