#pragma once
#include <QWidget>
#include <QVector>
#include "reporting/IReportSource.h"   // for SampleRef

namespace DVE {

// Stub - full implementation in Task 15.
//
// NOTE: Constructor signature must remain ABI-compatible with
// ReportPreviewDialog::buildUi(). When Task 15 replaces this stub, move the
// inline body into a new SamplesCheckboxPanel.cpp and add it to
// DataViewerEnterprise.pro:SOURCES, then re-run qmake so moc picks up new
// signal/slot declarations.
class SamplesCheckboxPanel : public QWidget {
    Q_OBJECT
public:
    explicit SamplesCheckboxPanel(const QVector<SampleRef>& refs,
                                    QWidget* parent = nullptr)
        : QWidget(parent) { Q_UNUSED(refs); }
};

} // namespace DVE
