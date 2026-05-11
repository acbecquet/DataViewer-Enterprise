#include "ReportPreviewDialog.h"
#include "SamplesCheckboxPanel.h"
#include "PropertiesPanel.h"
#include "SlideCanvasItems.h"
#include "../reporting/LayoutCommand.h"
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
#include <QKeySequence>
#include <QPen>
#include <QResizeEvent>
#include <QSettings>
#include <QShortcut>
#include <QToolBar>
#include <QToolButton>

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

    // Title preview, word-wrapped across up to ~3 lines so long titles like
    // "Rosin Box Double Blind Sensory 5-7-2026 - Initial" stay readable.
    if (!spec.title.isEmpty()) {
        f.setPointSize(7);
        p.setFont(f);
        p.setPen(QColor(80, 80, 80));
        constexpr int titleH = 36;     // bottom band for 2-3 wrapped lines at 7pt
        p.drawText(QRect(4, H - titleH - 2, W - 8, titleH),
                    Qt::AlignLeft | Qt::AlignTop | Qt::TextWordWrap,
                    spec.title);
    }

    return pix;
}
} // anonymous namespace

ReportPreviewDialog::ReportPreviewDialog(IReportSource* src, QWidget* p)
    : QDialog(p), m_source(src) {
    Q_ASSERT(src);                          // hard precondition; caller owns + provides
    setWindowTitle("Report Preview - " + src->sourceLabel());
    // Match the 16:9 slide aspect closely after subtracting fixed-width side
    // panels so the canvas fills vertically with minimal letterbox. Also enable
    // the standard minimize/maximize window controls so the user can full-screen
    // the preview when they want more canvas room.
    setWindowFlags(windowFlags() | Qt::WindowMinMaxButtonsHint);
    resize(1600, 720);
    m_layout = src->loadLayout();
    buildUi();
    populateThumbnails();
    // Single-shot debounce timer: every layout-mutating site calls
    // scheduleAutoSave(), which restarts the 500 ms countdown. flushAutoSave
    // (timeout, or done()-override on close) writes m_layout via the source.
    m_autoSaveTimer = new QTimer(this);
    m_autoSaveTimer->setSingleShot(true);
    m_autoSaveTimer->setInterval(500);
    connect(m_autoSaveTimer, &QTimer::timeout,
            this, &ReportPreviewDialog::flushAutoSave);
    if (m_thumbList->count() > 0) m_thumbList->setCurrentRow(0);
}

void ReportPreviewDialog::buildUi() {
    auto* outer = new QHBoxLayout(this);
    outer->setContentsMargins(8, 8, 8, 8);

    // Left column: thumbs + samples
    auto* left = new QVBoxLayout;
    m_thumbList = new QListWidget;
    m_thumbList->setFixedWidth(180);   // 160 px icon + ~14 px scrollbar + padding
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

    // Center: canvas. Wrapped in a vertical layout so the snap-to-grid toolbar
    // (Phase 3, Task 6) sits as a thin strip above the canvas. The toolbar
    // has zero stretch so the canvas keeps the full remaining vertical space.
    m_scene = new QGraphicsScene(this);
    m_scene->setSceneRect(0, 0, 800, 450);
    m_scene->setBackgroundBrush(Qt::white);
    m_scene->addRect(0, 0, 800, 450, QPen(QColor(180,180,180)), Qt::NoBrush);
    m_canvas = new QGraphicsView(m_scene);
    m_canvas->setRenderHint(QPainter::Antialiasing);
    m_canvas->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_canvas->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    // No frame around the view so the slide-aspect frame inside the scene is
    // the only visible boundary.
    m_canvas->setFrameShape(QFrame::NoFrame);

    auto* centerCol = new QVBoxLayout;
    centerCol->setContentsMargins(0, 0, 0, 0);
    m_toolbar = new QToolBar(this);
    m_snapBtn = new QToolButton(m_toolbar);
    m_snapBtn->setText(QStringLiteral("Snap to grid"));
    m_snapBtn->setCheckable(true);
    m_snapBtn->setToolTip(QStringLiteral(
        "Snap dragged items to a 0.05\" grid (Ctrl+drag overrides)"));
    {
        QSettings s(QStringLiteral("SDR"),
                    QStringLiteral("DataViewerEnterprise"));
        m_snapEnabled = s.value(QStringLiteral("preview/snap"), true).toBool();
    }
    m_snapBtn->setChecked(m_snapEnabled);
    m_toolbar->addWidget(m_snapBtn);
    connect(m_snapBtn, &QToolButton::toggled,
            this, &ReportPreviewDialog::onSnapToggled);
    centerCol->addWidget(m_toolbar, 0);
    centerCol->addWidget(m_canvas, 1);
    outer->addLayout(centerCol, 1);

    // Right: properties panel + buttons (narrow column; buttons stack vertically)
    auto* right = new QVBoxLayout;
    m_propsPanel = new PropertiesPanel;
    m_propsPanel->setMaximumWidth(180);
    connect(m_propsPanel, &PropertiesPanel::rectEdited,
            this, &ReportPreviewDialog::applyRectEdit);
    connect(m_propsPanel, &PropertiesPanel::fontSizeEdited,
            this, &ReportPreviewDialog::applyFontSize);
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
    auto* create = new QPushButton("Create Report");
    create->setDefault(true);
    auto* cancel = new QPushButton("Cancel");
    right->addWidget(create);     // stacked vertically per UX request
    right->addWidget(cancel);
    outer->addLayout(right);

    connect(cancel, &QPushButton::clicked, this, &ReportPreviewDialog::onCancel);
    connect(create, &QPushButton::clicked, this, &ReportPreviewDialog::onCreateReport);

    // Undo / redo. QShortcut is dialog-scoped (parent = this) so it only fires
    // while the preview dialog has focus. Task 4 will plumb the existing
    // applyRectEdit / applyFontSize / applySortChange paths into pushCommand;
    // until then the stack is wired but empty.
    auto* undoSc = new QShortcut(QKeySequence::Undo, this);
    connect(undoSc, &QShortcut::activated, this, &ReportPreviewDialog::doUndo);

    auto* redoSc = new QShortcut(QKeySequence::Redo, this);
    connect(redoSc, &QShortcut::activated, this, &ReportPreviewDialog::doRedo);
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
        // The placeholder renderer paints the slide number + kind + (wrapped)
        // title INSIDE the thumbnail, so the QListWidgetItem text would be
        // redundant — pass an empty string for icon-only display.
        const ReportSlideSpec spec = m_source->buildSlide(i, m_layout, m_excludedSamples);
        QPixmap pix = renderThumbnailPlaceholder(i + 1, kindLabel, spec);
        auto* item = new QListWidgetItem(QIcon(pix), QString());
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
            // De-select all other items so only one set of resize handles shows.
            for (QGraphicsItem* gi : m_scene->items()) {
                if (auto* other = dynamic_cast<ResizableSlideItem*>(gi))
                    other->setSelectedItem(other == it);
            }
            // Pass font size for TextItem and TableItem; 0 disables the font
            // row in the properties panel for non-text items (PlotItem, etc.).
            int fontPt = 0;
            if (auto* textItem = dynamic_cast<TextItem*>(it))
                fontPt = textItem->fontPointSize();
            else if (auto* tableItem = dynamic_cast<TableItem*>(it))
                fontPt = tableItem->fontPointSize();
            m_propsPanel->setSelectedItem(it->elementId(),
                                            it->itemRectInches(), fontPt);
        });
    };

    // For TextItems we call place() FIRST so the rect-derived m_w is in place
    // when setText runs — that lets TextItem::setText auto-grow m_h to fit
    // wrapped content without overwriting the user's layout-set width.

    // Helper: if the persisted layout font is 0 ("not set"), fall back to the
    // canvas-side hardcoded default. This means a fresh layout (where every
    // font field is 0) displays at the small, preview-friendly sizes, while
    // a user-customized layout displays at the user's chosen size.
    auto canvasFont = [](int layoutValue, int canvasDefault) {
        return layoutValue > 0 ? layoutValue : canvasDefault;
    };

    // Font sizes are read from m_layout (which is the resolved layout that
    // buildSlide also reads). Cover and Divider slides pull from top-level
    // ReportLayout fields; Content / Cumulative pull from spec.layout.
    const QString slideKey = m_source->slideKey(m_currentSlide);
    if (spec.kind == SlideKind::Cover) {
        auto* title = new TextItem(QStringLiteral("cover_title"));
        title->setFontPointSize(canvasFont(
            m_layout.coverTitleFontPt, ReportLayout::kCanvasCoverTitlePt));
        place(title, spec.layout.title);
        title->setText(spec.title);
        clampToSlide(title);
        auto* subtitle = new TextItem(QStringLiteral("cover_subtitle"));
        subtitle->setFontPointSize(canvasFont(
            m_layout.coverSubtitleFontPt, ReportLayout::kCanvasCoverSubtitlePt));
        place(subtitle, m_layout.coverSubtitle);
        subtitle->setText(spec.propertiesText);   // date string
        clampToSlide(subtitle);
    } else if (spec.kind == SlideKind::Divider) {
        auto* title = new TextItem(QStringLiteral("divider_title"));
        title->setFontPointSize(canvasFont(
            m_layout.dividerTitleFontPts.value(slideKey, 0),
            ReportLayout::kCanvasDividerTitlePt));
        place(title, spec.layout.title);
        title->setText(spec.title);
        clampToSlide(title);
    } else if (spec.kind == SlideKind::Content
               || spec.kind == SlideKind::Cumulative) {
        auto* title = new TextItem(QStringLiteral("title"));
        title->setFontPointSize(canvasFont(
            spec.layout.titleFontPt, ReportLayout::kCanvasContentTitlePt));
        place(title, spec.layout.title);
        title->setText(spec.title);
        clampToSlide(title);

        auto* table = new TableItem(QStringLiteral("table"));
        table->setHeaders(spec.tableHeaders);
        table->setFontPointSize(canvasFont(
            spec.layout.tableFontPt, ReportLayout::kCanvasTablePt));
        table->setSort(m_layout.tableSort.column, m_layout.tableSort.order);
        place(table, spec.layout.table);
        // setRows AFTER place so auto-grow can extend m_h beyond the layout-set
        // height when the row count requires it (last-row clipping fix).
        table->setRows(spec.tableRows);
        clampToSlide(table);
        connect(table, &TableItem::columnHeaderClicked, this,
                [this](const QString& c) { applySortChange(c); });

        if (!spec.radarPixmap.isNull()) {
            auto* radar = new PlotItem(QStringLiteral("radar"),
                                        QPixmap::fromImage(spec.radarPixmap));
            place(radar, spec.layout.radar);
        }

        if (!spec.propertiesText.isEmpty()) {
            auto* props = new TextItem(QStringLiteral("propertiesBox"));
            props->setFontPointSize(canvasFont(
                spec.layout.propertiesBox.fontPt,
                ReportLayout::kCanvasPropertiesBoxPt));
            place(props, spec.layout.propertiesBox.rect);
            props->setText(spec.propertiesText);   // auto-grows m_h to fit
            clampToSlide(props);
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

    // Push the current snap-to-grid toggle state into every fresh item so
    // their mouseMoveEvent uses the correct behavior from the first drag.
    // The dialog's m_snapEnabled was restored from QSettings in buildUi().
    propagateSnapStateToItems();

    // Scale the 800x450 slide to fill the viewport with letterboxing when the
    // viewport aspect doesn't match. Keeps the entire slide visible no matter
    // how the dialog is resized.
    m_canvas->fitInView(m_scene->sceneRect(), Qt::KeepAspectRatio);
}

void ReportPreviewDialog::resizeEvent(QResizeEvent* e) {
    QDialog::resizeEvent(e);
    if (m_canvas && m_scene)
        m_canvas->fitInView(m_scene->sceneRect(), Qt::KeepAspectRatio);
}

QRectF ReportPreviewDialog::currentRectFor(const QString& slideKey,
                                           const QString& elementId,
                                           const SlideKind kind) const {
    auto contentRect = [&](const ContentSlideLayout& cs) -> QRectF {
        if (elementId == QStringLiteral("title"))         return cs.title;
        if (elementId == QStringLiteral("table"))         return cs.table;
        if (elementId == QStringLiteral("radar"))         return cs.radar;
        if (elementId == QStringLiteral("propertiesBox")) return cs.propertiesBox.rect;
        return QRectF();
    };
    QRectF result;
    if (kind == SlideKind::Content) {
        result = contentRect(m_layout.contentSlides.value(slideKey));
    } else if (kind == SlideKind::Cumulative) {
        result = contentRect(m_layout.cumulative);
    } else if (kind == SlideKind::Divider) {
        if (elementId == QStringLiteral("divider_title"))
            result = m_layout.dividerTitles.value(slideKey);
    } else if (kind == SlideKind::Cover) {
        if (elementId == QStringLiteral("cover_title"))         result = m_layout.coverTitle;
        else if (elementId == QStringLiteral("cover_subtitle")) result = m_layout.coverSubtitle;
    }

    // If m_layout has no rect recorded for this element yet (first edit, or
    // a sibling edit triggered an undo that nulled this field), fall back to
    // the canvas item's current rect. That's what the user actually sees,
    // so undo will restore to a visible position rather than QRectF(0,0,0,0).
    // SensoryReportSource::buildSlide only swaps in computed defaults when
    // the WHOLE layout isEmpty() — once any sibling edit has touched the
    // layout, a null cover/divider rect would otherwise leak through unmitigated.
    if (result.isNull()) {
        if (auto* item = findCanvasItem(elementId))
            return item->itemRectInches();
    }
    return result;
}

int ReportPreviewDialog::currentFontPtFor(const QString& slideKey,
                                          const QString& elementId,
                                          const SlideKind kind) const {
    auto contentFont = [&](const ContentSlideLayout& cs) -> int {
        if (elementId == QStringLiteral("title"))         return cs.titleFontPt;
        if (elementId == QStringLiteral("table"))         return cs.tableFontPt;
        if (elementId == QStringLiteral("propertiesBox")) return cs.propertiesBox.fontPt;
        return 0;
    };
    if (kind == SlideKind::Content) {
        return contentFont(m_layout.contentSlides.value(slideKey));
    } else if (kind == SlideKind::Cumulative) {
        return contentFont(m_layout.cumulative);
    } else if (kind == SlideKind::Divider) {
        if (elementId == QStringLiteral("divider_title"))
            return m_layout.dividerTitleFontPts.value(slideKey, 0);
    } else if (kind == SlideKind::Cover) {
        if (elementId == QStringLiteral("cover_title"))    return m_layout.coverTitleFontPt;
        if (elementId == QStringLiteral("cover_subtitle")) return m_layout.coverSubtitleFontPt;
    }
    return 0;
}

void ReportPreviewDialog::applyRectEdit(const QString& elementId,
                                         const QRectF& rectInches) {
    if (m_currentSlide < 0 || m_currentSlide >= m_source->slideCount()) return;
    const SlideKind kind = m_source->slideKind(m_currentSlide);
    // slideKey is a cheap accessor; calling buildSlide here would needlessly
    // re-render the radar pixmap (and other content) on every drag-release.
    const QString slideKey = m_source->slideKey(m_currentSlide);

    // Image-slide rect persistence is still deferred (Phase 2 / image-layout
    // overrides). For all other kinds, route the mutation through
    // RectCommand so it lands on the undo stack. pushCommand() calls
    // cmd->apply(m_layout) and schedules the auto-save itself, so we do NOT
    // mutate m_layout or call scheduleAutoSave() directly here.
    const QRectF oldRect = currentRectFor(slideKey, elementId, kind);
    if (oldRect != rectInches) {
        pushCommand(QSharedPointer<RectCommand>::create(
            slideKey, elementId, oldRect, rectInches));
    }
    // If oldRect == rectInches we still fall through to the live-update block
    // so a stray canvas item that drifted out of sync can be re-snapped.
    // pushCommand is intentionally skipped to avoid no-op undo entries.

    // Live update: push the new rect onto the matching canvas item so spinbox
    // edits in the properties panel reflect immediately. When the edit ORIGINATED
    // from the canvas (drag-resize), the item is already at this rect and
    // setRectInches is a no-op.
    if (auto* item = findCanvasItem(elementId)) {
        if (item->itemRectInches() != rectInches)
            item->setRectInches(rectInches);
    }
}

void ReportPreviewDialog::applyFontSize(const QString& elementId, int newFontPt) {
    if (m_currentSlide < 0 || m_currentSlide >= m_source->slideCount()) return;
    const SlideKind kind = m_source->slideKind(m_currentSlide);
    const QString slideKey = m_source->slideKey(m_currentSlide);

    // Route the persistent font mutation through FontSizeCommand. The
    // command's apply() updates m_layout; pushCommand schedules autosave.
    // No direct m_layout writes / scheduleAutoSave() here.
    const int oldPt = currentFontPtFor(slideKey, elementId, kind);
    if (oldPt != newFontPt) {
        pushCommand(QSharedPointer<FontSizeCommand>::create(
            slideKey, elementId, oldPt, newFontPt));
    }

    // Live update: push the new font size onto the matching canvas item so
    // the spinbox edit reflects immediately. Same idempotent semantics as
    // the rect-edit case — re-applying the same size is harmless.
    auto* item = findCanvasItem(elementId);
    if (auto* textItem = dynamic_cast<TextItem*>(item)) {
        textItem->setFontPointSize(newFontPt);
        // Re-apply text so auto-grow recomputes for the new font metric.
        const QString t = textItem->text();
        if (!t.isEmpty()) {
            textItem->setText(t);
            clampToSlide(textItem);
        }
    } else if (auto* tableItem = dynamic_cast<TableItem*>(item)) {
        tableItem->setFontPointSize(newFontPt);
        // setRows is invoked internally so row-height auto-grow runs against
        // the new font metric.
        clampToSlide(tableItem);
    }
}

ResizableSlideItem* ReportPreviewDialog::findCanvasItem(const QString& elementId) const {
    if (!m_scene) return nullptr;
    for (QGraphicsItem* gi : m_scene->items()) {
        auto* item = dynamic_cast<ResizableSlideItem*>(gi);
        if (item && item->elementId() == elementId) return item;
    }
    return nullptr;
}

void ReportPreviewDialog::clampToSlide(ResizableSlideItem* item) {
    if (!item) return;
    constexpr double slideW = 13.33;
    constexpr double slideH = 7.50;
    const QRectF r = item->itemRectInches();
    const double maxX = qMax(0.0, slideW - r.width());
    const double maxY = qMax(0.0, slideH - r.height());
    const double newX = qBound(0.0, r.x(), maxX);
    const double newY = qBound(0.0, r.y(), maxY);
    if (newX != r.x() || newY != r.y())
        item->setPos(newX * 60.0, newY * 60.0);
}

void ReportPreviewDialog::applySortChange(const QString& column) {
    // 3-state cycle: empty -> Descending -> Ascending -> empty.
    // Compute the new (column, order) target from the old state, then push a
    // SortCommand. pushCommand applies the mutation and schedules autosave.
    const QString       oldCol = m_layout.tableSort.column;
    const Qt::SortOrder oldOrd = m_layout.tableSort.order;
    QString       newCol;
    Qt::SortOrder newOrd;
    if (oldCol != column) {
        newCol = column;
        newOrd = Qt::DescendingOrder;
    } else if (oldOrd == Qt::DescendingOrder) {
        newCol = column;
        newOrd = Qt::AscendingOrder;
    } else {
        newCol.clear();
        newOrd = Qt::DescendingOrder;
    }

    if (oldCol == newCol && oldOrd == newOrd) return;
    pushCommand(QSharedPointer<SortCommand>::create(
        oldCol, oldOrd, newCol, newOrd));

    // Re-render so the table reflects the new sort. doUndo/doRedo also call
    // populateCanvas(), so this keeps the originating click consistent with
    // the undo/redo paths.
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

void ReportPreviewDialog::scheduleAutoSave() {
    // start() on a single-shot QTimer restarts the countdown if already running
    // — that's the debounce behavior we want (last edit within 500ms wins).
    m_autoSaveTimer->start();
}

void ReportPreviewDialog::flushAutoSave() {
    m_source->saveLayout(m_layout);
}

void ReportPreviewDialog::pushCommand(QSharedPointer<LayoutCommand> cmd) {
    if (!cmd) return;
    cmd->apply(m_layout);
    m_undoStack.push(cmd);
    if (m_undoStack.size() > kMaxUndoDepth) m_undoStack.removeFirst();
    m_redoStack.clear();
    scheduleAutoSave();
    // Caller is responsible for any canvas re-render after apply(). The three
    // existing mutation sites (applyRectEdit / applyFontSize / applySortChange)
    // already re-render — Task 4 will plumb commands through those paths.
}

void ReportPreviewDialog::doUndo() {
    if (m_undoStack.isEmpty()) return;
    auto cmd = m_undoStack.pop();
    cmd->undo(m_layout);
    m_redoStack.push(cmd);
    populateCanvas();
    scheduleAutoSave();
}

void ReportPreviewDialog::doRedo() {
    if (m_redoStack.isEmpty()) return;
    auto cmd = m_redoStack.pop();
    cmd->apply(m_layout);
    m_undoStack.push(cmd);
    populateCanvas();
    scheduleAutoSave();
}

void ReportPreviewDialog::onSnapToggled(bool checked) {
    // Persist the toggle to QSettings so the choice survives across sessions.
    // The same key is read in buildUi() to restore state on dialog open.
    m_snapEnabled = checked;
    QSettings s(QStringLiteral("SDR"),
                QStringLiteral("DataViewerEnterprise"));
    s.setValue(QStringLiteral("preview/snap"), checked);
    propagateSnapStateToItems();
}

void ReportPreviewDialog::propagateSnapStateToItems() {
    // Walks all current scene items and tells each ResizableSlideItem about
    // the current snap state. Called from onSnapToggled when the user flips
    // the toolbar button, and from populateCanvas after fresh items land on
    // the scene (via place()).
    if (!m_canvas || !m_scene) return;
    for (QGraphicsItem* gi : m_scene->items()) {
        if (auto* r = dynamic_cast<ResizableSlideItem*>(gi))
            r->setSnapEnabled(m_snapEnabled);
    }
}

void ReportPreviewDialog::done(int r) {
    // Any close path (Cancel, Create Report, ESC, X) routes through done(),
    // so flushing here guarantees pending edits aren't dropped. Stop the timer
    // first so a pending tick can't fire-then-no-op after we've already saved.
    m_autoSaveTimer->stop();
    flushAutoSave();
    QDialog::done(r);
}

} // namespace DVE
