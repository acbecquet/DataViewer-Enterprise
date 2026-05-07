#pragma once
#include <QWidget>
#include <QVector>
#include "reporting/IReportSource.h"   // for SampleRef

namespace DVE {

// Stub - full implementation in Task 15.
class SamplesCheckboxPanel : public QWidget {
    Q_OBJECT
public:
    explicit SamplesCheckboxPanel(const QVector<SampleRef>& refs,
                                    QWidget* parent = nullptr)
        : QWidget(parent) { Q_UNUSED(refs); }
};

} // namespace DVE
