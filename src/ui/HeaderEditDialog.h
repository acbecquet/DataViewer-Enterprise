#pragma once

#include <QDialog>
#include <QLineEdit>
#include <QScrollArea>
#include "../pipeline/ReportData.h"

namespace DVE {

// ---------------------------------------------------------------------------
// HeaderEditDialog
//
// Allows editing the header properties (sample ID, tester, media, etc.)
// for a single sample. Pre-populated from the supplied SampleResult.
// On accept the caller should write the returned data back to the xlsx file.
// ---------------------------------------------------------------------------
class HeaderEditDialog : public QDialog
{
    Q_OBJECT
public:
    struct HeaderData {
        QString sampleID;
        QString tester;
        QString media;
        QString viscosity;
        QString resistance;
        QString voltage;
        QString puffingRegime;
        QString initialOilMass;
    };

    explicit HeaderEditDialog(const SampleResult& sample, QWidget* parent = nullptr);

    HeaderData headerData() const;

private:
    QLineEdit* m_sampleID;
    QLineEdit* m_tester;
    QLineEdit* m_media;
    QLineEdit* m_viscosity;
    QLineEdit* m_resistance;
    QLineEdit* m_voltage;
    QLineEdit* m_puffingRegime;
    QLineEdit* m_oilMass;
};

} // namespace DVE
