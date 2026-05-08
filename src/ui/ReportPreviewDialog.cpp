#include "ReportPreviewDialog.h"
#include "SamplesCheckboxPanel.h"
#include "PropertiesPanel.h"
#include "SlideCanvasItems.h"
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
        populateCanvas();
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
            this, &ReportPreviewDialog::applyRectEdit);
    connect(m_propsPanel, &PropertiesPanel::bringForwardClicked,
            this, [](const QString& id) {
        Q_UNUSED(id);
        // Z-order is per-slide in m_layout.zOrder; needs canvas item z()-property
        // mapping. Deferred — canvas items don't yet expose a setZ() bridge to
        // ReportLayout::zOrder.
    });
    connect(m_propsPanel, &PropertiesPanel::sendBackwardClicked,
            this, [](const QString& id) {
        Q_UNUSED(id);
        // Same as bringForwardClicked — deferred until z-order mapping lands.
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
    populateCanvas();
}

void ReportPreviewDialog::populateCanvas() {
    m_scene->clear();
    // Slide background frame matching the 800x450 scene (16:9 at 60 px/in
    // would be 800x450 for 13.33"x7.5"; the inch-to-px conversion is the same
    // 60 px/in used by ResizableSlideItem::kPxPerInch).
    m_scene->addRect(0, 0, 800, 450, QPen(QColor(180, 180, 180)), QBrush(Qt::white));
    if (m_currentSlide < 0 || m_currentSlide >= m_source->slideCount()) return;

    const ReportSlideSpec spec =
        m_source->buildSlide(m_currentSlide, m_layout, m_excludedSamples);

    auto place = [this](ResizableSlideItem* item, const QRectF& rectInches) {
        // setRectInches no-ops on null rects, so default-constructed layout
        // slots leave the item at its constructor-default position (0,0) and size.
        item->setRectInches(rectInches);
        m_scene->addItem(item);
        connect(item, &ResizableSlideItem::rectChanged, this,
                [this, item](const QRectF& r) {
            applyRectEdit(item->elementId(), r);
        });
        connect(item, &ResizableSlideItem::itemClicked, this,
                [this](ResizableSlideItem* it) {
            m_propsPanel->setSelectedItem(it->elementId(), it->itemRectInches());
        });
    };

    if (spec.kind == SlideKind::Cover) {
        auto* title = new TextItem(QStringLiteral("cover_title"));
        title->setText(spec.title);
        title->setFontPointSize(28);
        place(title, spec.layout.title);
        auto* subtitle = new TextItem(QStringLiteral("cover_subtitle"));
        subtitle->setText(spec.propertiesText);   // date string
        subtitle->setFontPointSize(16);
        place(subtitle, m_layout.coverSubtitle);
    } else if (spec.kind == SlideKind::Divider) {
        auto* title = new TextItem(QStringLiteral("divider_title"));
        title->setText(spec.title);
        title->setFontPointSize(32);
        place(title, spec.layout.title);
    } else if (spec.kind == SlideKind::Content
               || spec.kind == SlideKind::Cumulative) {
        auto* title = new TextItem(QStringLiteral("title"));
        title->setText(spec.title);
        title->setFontPointSize(18);
        place(title, spec.layout.title);

        auto* table = new TableItem(QStringLiteral("table"));
        table->setHeaders(spec.tableHeaders);
        table->setRows(spec.tableRows);
        table->setSort(m_layout.tableSort.column, m_layout.tableSort.order);
        connect(table, &TableItem::columnHeaderClicked, this,
                [this](const QString& c) { applySortChange(c); });
        place(table, spec.layout.table);

        if (!spec.radarPixmap.isNull()) {
            auto* radar = new PlotItem(QStringLiteral("radar"),
                                        QPixmap::fromImage(spec.radarPixmap));
            place(radar, spec.layout.radar);
        }

        if (!spec.propertiesText.isEmpty()) {
            auto* props = new TextItem(QStringLiteral("propertiesBox"));
            props->setText(spec.propertiesText);
            props->setFontPointSize(12);
            place(props, spec.layout.propertiesBox.rect);
        }
    } else if (spec.kind == SlideKind::Image) {
        for (int i = 0; i < spec.imagePaths.size(); ++i) {
            QPixmap pix(spec.imagePaths[i]);
            auto* item = new PlotItem(QStringLiteral("image_%1").arg(i), pix);
            const QRectF rect = (i < spec.imageLayouts.size())
                ? spec.imageLayouts[i] : QRectF();
            place(item, rect);
        }
    }
}

void ReportPreviewDialog::applyRectEdit(const QString& elementId,
                                         const QRectF& rectInches) {
    if (m_currentSlide < 0 || m_currentSlide >= m_source->slideCount()) return;
    const SlideKind kind = m_source->slideKind(m_currentSlide);
    // slideKey is a cheap accessor; calling buildSlide here would needlessly
    // re-render the radar pixmap (and other content) on every drag-release.
    const QString slideKey = m_source->slideKey(m_currentSlide);

    auto applyToContentLayout = [&](ContentSlideLayout& cs) {
        if (elementId == QStringLiteral("title"))             cs.title = rectInches;
        else if (elementId == QStringLiteral("table"))        cs.table = rectInches;
        else if (elementId == QStringLiteral("radar"))        cs.radar = rectInches;
        else if (elementId == QStringLiteral("propertiesBox")) cs.propertiesBox.rect = rectInches;
    };

    if (kind == SlideKind::Content) {
        ContentSlideLayout cs = m_layout.contentSlides.value(slideKey);
        applyToContentLayout(cs);
        m_layout.contentSlides[slideKey] = cs;
    } else if (kind == SlideKind::Cumulative) {
        applyToContentLayout(m_layout.cumulative);
    } else if (kind == SlideKind::Divider) {
        if (elementId == QStringLiteral("divider_title"))
            m_layout.dividerTitles[slideKey] = rectInches;
    } else if (kind == SlideKind::Cover) {
        if (elementId == QStringLiteral("cover_title"))         m_layout.coverTitle = rectInches;
        else if (elementId == QStringLiteral("cover_subtitle")) m_layout.coverSubtitle = rectInches;
    }
    // Image-slide rect persistence deferred (Phase 2 / image-layout overrides).
}

void ReportPreviewDialog::applySortChange(const QString& column) {
    // 3-state cycle: empty -> Descending -> Ascending -> empty
    if (m_layout.tableSort.column != column) {
        m_layout.tableSort.column = column;
        m_layout.tableSort.order = Qt::DescendingOrder;
    } else if (m_layout.tableSort.order == Qt::DescendingOrder) {
        m_layout.tableSort.order = Qt::AscendingOrder;
    } else {
        m_layout.tableSort.column.clear();
        m_layout.tableSort.order = Qt::DescendingOrder;
    }
    populateCanvas();
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
