#pragma once
#include <QWidget>

namespace DVE {

// Stub - full implementation in Task 16.
//
// NOTE: Constructor signature must remain ABI-compatible with
// ReportPreviewDialog::buildUi(). When Task 16 replaces this stub, move the
// inline body into a new PropertiesPanel.cpp and add it to
// DataViewerEnterprise.pro:SOURCES, then re-run qmake so moc picks up new
// signal/slot declarations.
class PropertiesPanel : public QWidget {
    Q_OBJECT
public:
    explicit PropertiesPanel(QWidget* parent = nullptr) : QWidget(parent) {}
};

} // namespace DVE
