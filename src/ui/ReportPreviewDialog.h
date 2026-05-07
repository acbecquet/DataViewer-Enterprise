#pragma once
#include <QDialog>
#include <QGraphicsView>
#include <QGraphicsScene>
#include <QListWidget>
#include <QSet>
#include "reporting/ReportLayout.h"
#include "reporting/IReportSource.h"

namespace DVE {

class SamplesCheckboxPanel;
class PropertiesPanel;

class ReportPreviewDialog : public QDialog {
    Q_OBJECT
public:
    ReportPreviewDialog(IReportSource* source, QWidget* parent = nullptr);

    QString outputPath() const { return m_outputPath; }      // valid only after exec() == Accepted

private slots:
    void onCreateReport();
    void onCancel();
    void onSlideSelected(int row);

private:
    void buildUi();
    void populateThumbnails();

    IReportSource* m_source;
    ReportLayout   m_layout;
    QSet<QString>  m_excludedSamples;
    int            m_currentSlide = 0;
    QString        m_outputPath;

    QListWidget*    m_thumbList = nullptr;
    QGraphicsView*  m_canvas    = nullptr;
    QGraphicsScene* m_scene     = nullptr;
    SamplesCheckboxPanel* m_samplesPanel = nullptr;
    PropertiesPanel*      m_propsPanel   = nullptr;
};

} // namespace DVE
