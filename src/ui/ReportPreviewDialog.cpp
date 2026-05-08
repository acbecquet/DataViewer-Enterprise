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
#include <QListWidgetItem>
#include <QPixmap>
#include <QIcon>
#include <QFont>
#include <QFontMetrics>
#include <QPen>

namespace DVE {

namespace {
// Renders a 160x90 placeholder thumbnail showing the slide kind label, slide
// number, and (when buildSlide has been wired up in Task 18) an elided title
// preview. With the current default-constructed ReportSlideSpec the title is
// empty and only the kind label + number are drawn.
QPixmap renderThumbnailPlaceholder(int slideNumber, const QString& kindLabel,
                                    const ReportSlideSpec& spec) {
    constexpr int W = 160, H = 90;
    QPixmap pix(W, H);
    pix.fill(Qt::white);
    QPainter p(&pix);
    p.setRenderHint(QPainter::Antialiasing);
    p.setRenderHint(QPainter::TextAntialiasing);

    // Light gray border + slide-aspect frame
    p.setPen(QPen(QColor(200, 200, 200), 1));
    p.setBrush(QColor(248, 248, 248));
    p.drawRect(QRect(0, 0, W - 1, H - 1));

    // Centered slide kind in dark gray
    p.setPen(QColor(60, 60, 60));
    QFont f = p.font();
    f.setPointSize(10);
    f.setBold(true);
    p.setFont(f);
    p.drawText(pix.rect(), Qt::AlignCenter, kindLabel);

    // Slide number top-left
    f.setBold(false);
    f.setPointSize(8);
    p.setFont(f);
    p.setPen(QColor(120, 120, 120));
    p.drawText(QRect(4, 2, 30, 14), Qt::AlignLeft | Qt::AlignTop,
                QString::number(slideNumber));

    // Title preview (visible once Task 18 populates spec.title)
    if (!spec.title.isEmpty()) {
        f.setPointSize(7);
        p.setFont(f);
        p.setPen(QColor(80, 80, 80));
        QFontMetrics fm(f);
        QString elided = fm.elidedText(spec.title, Qt::ElideRight, W - 8);
        p.drawText(QRect(4, H - 18, W - 8, 14),
                    Qt::AlignLeft | Qt::AlignVCenter, elided);
    }

    return pix;
}
} // anonymous namespace

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
    m_thumbList->setFixedWidth(200);   // icon (160) + label + padding
    connect(m_thumbList, &QListWidget::currentRowChanged,
            this, &ReportPreviewDialog::onSlideSelected);
    left->addWidget(m_thumbList, 1);
    m_samplesPanel = new SamplesCheckboxPanel(m_source->allSamples());
    connect(m_samplesPanel, &SamplesCheckboxPanel::sampleToggled,
            this, [this](const QString& id, bool included) {
        if (included) m_excludedSamples.remove(id);
        else          m_excludedSamples.insert(id);
        // Canvas rebuild happens in Task 18 (buildSlide); for now just track exclusion state.
    });
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
    connect(m_propsPanel, &PropertiesPanel::rectEdited,
            this, [this](const QString& id, const QRectF& r) {
        // Update m_layout's content slide for the current slide. The slide kind
        // determines which slot in ContentSlideLayout to write. Element-id format
        // is "<role>_<key>" where role is one of {title, table, radar, props}.
        // Actual application-to-canvas happens in Task 18 (buildSlide).
        Q_UNUSED(id); Q_UNUSED(r);
        // TODO(task-18): mutate m_layout based on id role + current slide key.
    });
    connect(m_propsPanel, &PropertiesPanel::bringForwardClicked,
            this, [this](const QString& id) {
        // Z-order list mutation; applied to canvas in Task 18.
        Q_UNUSED(id);
        // TODO(task-18): m_layout.zOrder manipulation.
    });
    connect(m_propsPanel, &PropertiesPanel::sendBackwardClicked,
            this, [this](const QString& id) {
        Q_UNUSED(id);
        // TODO(task-18): m_layout.zOrder manipulation.
    });
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
    m_thumbList->setIconSize(QSize(160, 90));
    for (int i = 0; i < m_source->slideCount(); ++i) {
        const SlideKind k = m_source->slideKind(i);
        const QString kindLabel =
            k == SlideKind::Cover     ? QStringLiteral("Cover") :
            k == SlideKind::Divider   ? QStringLiteral("Divider") :
            k == SlideKind::Content   ? QStringLiteral("Content") :
            k == SlideKind::Image     ? QStringLiteral("Images") :
                                         QStringLiteral("Cumulative");
        // buildSlide currently returns an empty default-constructed spec
        // (SensoryReportSource stub); Task 18 will populate it with real
        // titles + content. The placeholder renderer handles both states.
        const ReportSlideSpec spec = m_source->buildSlide(i, m_layout, m_excludedSamples);
        QPixmap pix = renderThumbnailPlaceholder(i + 1, kindLabel, spec);
        auto* item = new QListWidgetItem(QIcon(pix),
                                          QString::number(i + 1) + ". " + kindLabel);
        m_thumbList->addItem(item);
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
