#include "ReportPreviewDialog.h"
#include "SamplesCheckboxPanel.h"
#include "PropertiesPanel.h"
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QPushButton>
#include <QFileDialog>
#include <QMessageBox>
#include <QGraphicsRectItem>
#include <QPainter>

namespace DVE {

ReportPreviewDialog::ReportPreviewDialog(IReportSource* src, QWidget* p)
    : QDialog(p), m_source(src) {
    Q_ASSERT(src);                          // hard precondition; caller owns + provides
    setWindowTitle("Report Preview - " + src->sourceLabel());
    resize(1200, 720);
    m_layout = src->loadLayout();
    buildUi();
    populateThumbnails();
    if (m_thumbList->count() > 0) m_thumbList->setCurrentRow(0);
}

void ReportPreviewDialog::buildUi() {
    auto* outer = new QHBoxLayout(this);
    outer->setContentsMargins(8, 8, 8, 8);

    // Left column: thumbs + samples
    auto* left = new QVBoxLayout;
    m_thumbList = new QListWidget;
    m_thumbList->setFixedWidth(160);
    connect(m_thumbList, &QListWidget::currentRowChanged,
            this, &ReportPreviewDialog::onSlideSelected);
    left->addWidget(m_thumbList, 1);
    m_samplesPanel = new SamplesCheckboxPanel(m_source->allSamples());
    left->addWidget(m_samplesPanel, 1);
    outer->addLayout(left);

    // Center: canvas
    m_scene = new QGraphicsScene(this);
    m_scene->setSceneRect(0, 0, 800, 450);
    m_scene->setBackgroundBrush(Qt::white);
    m_scene->addRect(0, 0, 800, 450, QPen(QColor(180,180,180)), Qt::NoBrush);
    m_canvas = new QGraphicsView(m_scene);
    m_canvas->setRenderHint(QPainter::Antialiasing);
    outer->addWidget(m_canvas, 1);

    // Right: properties panel + buttons
    auto* right = new QVBoxLayout;
    m_propsPanel = new PropertiesPanel;
    right->addWidget(m_propsPanel, 1);
    auto* btns = new QHBoxLayout;
    auto* cancel = new QPushButton("Cancel");
    auto* create = new QPushButton("Create Report");
    create->setDefault(true);
    btns->addStretch();
    btns->addWidget(cancel);
    btns->addWidget(create);
    right->addLayout(btns);
    outer->addLayout(right);

    connect(cancel, &QPushButton::clicked, this, &ReportPreviewDialog::onCancel);
    connect(create, &QPushButton::clicked, this, &ReportPreviewDialog::onCreateReport);
}

void ReportPreviewDialog::populateThumbnails() {
    m_thumbList->clear();
    for (int i = 0; i < m_source->slideCount(); ++i) {
        const SlideKind k = m_source->slideKind(i);
        const QString kindLabel =
            k == SlideKind::Cover     ? QStringLiteral("Cover") :
            k == SlideKind::Divider   ? QStringLiteral("Divider") :
            k == SlideKind::Content   ? QStringLiteral("Content") :
            k == SlideKind::Image     ? QStringLiteral("Images") :
                                         QStringLiteral("Cumulative");
        m_thumbList->addItem(QString::number(i + 1) + ". " + kindLabel);
    }
}

void ReportPreviewDialog::onSlideSelected(int row) {
    m_currentSlide = row;
    // Canvas population is wired up in Task 18 (buildSlide).
}

void ReportPreviewDialog::onCancel() { reject(); }

void ReportPreviewDialog::onCreateReport() {
    const QString path = QFileDialog::getSaveFileName(this, "Save Report",
        "report.pptx", "PowerPoint (*.pptx)");
    if (path.isEmpty()) return;
    QString err;
    if (!m_source->writePptx(path, m_layout, m_excludedSamples, &err)) {
        QMessageBox::warning(this, "Report Failed", err);
        return;
    }
    m_outputPath = path;
    accept();
}

} // namespace DVE
