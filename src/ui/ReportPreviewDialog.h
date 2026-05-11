#pragma once
#include <QDialog>
#include <QGraphicsLineItem>
#include <QGraphicsView>
#include <QGraphicsScene>
#include <QLineF>
#include <QListWidget>
#include <QRectF>
#include <QSet>
#include <QSettings>
#include <QStack>
#include <QSharedPointer>
#include <QTimer>
#include <QVector>
#include "reporting/ReportLayout.h"
#include "reporting/IReportSource.h"

class QShortcut;
class QToolBar;
class QToolButton;

namespace DVE {

class SamplesCheckboxPanel;
class PropertiesPanel;
class LayoutCommand;

class ReportPreviewDialog : public QDialog {
    Q_OBJECT
public:
    ReportPreviewDialog(IReportSource* source, QWidget* parent = nullptr);

    QString outputPath() const { return m_outputPath; }      // valid only after exec() == Accepted

    // Override so any close path (Cancel, Create Report, X button, ESC) flushes
    // pending auto-save edits before the dialog tears down.
    void done(int r) override;

protected:
    // Re-fit the slide canvas whenever the dialog resizes so the full 800x450
    // scene is visible regardless of viewport width.
    void resizeEvent(QResizeEvent*) override;

private slots:
    void onCreateReport();
    void onCancel();
    void onSlideSelected(int row);
    void onSnapToggled(bool checked);

private:
    void buildUi();
    void populateThumbnails();
    void populateCanvas();
    void applyRectEdit(const QString& elementId, const QRectF& rectInches);
    void applyFontSize(const QString& elementId, int newFontPt);
    void applySortChange(const QString& column);
    void scheduleAutoSave();
    void flushAutoSave();

    // Apply a command, push it on the undo stack, clear redo, schedule autosave.
    // Capped at kMaxUndoDepth; oldest entry is dropped when the cap is hit.
    void pushCommand(QSharedPointer<LayoutCommand> cmd);
    void doUndo();
    void doRedo();

    // Look up a canvas item by its elementId. Returns nullptr if no item with
    // that id exists in the current scene (e.g. user is on a different slide).
    class ResizableSlideItem* findCanvasItem(const QString& elementId) const;

    // Helpers for Task 4 — read the current value out of m_layout for a given
    // slide+element. Mirror images of the setters inside RectCommand /
    // FontSizeCommand. Used before pushing commands to capture the "old" state.
    QRectF currentRectFor(const QString& slideKey,
                          const QString& elementId,
                          const SlideKind kind) const;
    int    currentFontPtFor(const QString& slideKey,
                            const QString& elementId,
                            const SlideKind kind) const;

    // Push the item back inside the 13.33 x 7.5 inch slide if its current rect
    // (which may have grown larger than the layout-set rect via TextItem /
    // TableItem auto-grow) extends past the slide edges. Position-only fix —
    // does not shrink the item.
    void clampToSlide(class ResizableSlideItem* item);

    // Walks all current scene items and pushes the current m_snapEnabled state
    // down to each ResizableSlideItem. Task 7 will fill in the body once
    // ResizableSlideItem::setSnapEnabled is added.
    void propagateSnapStateToItems();

    // Alignment guide rendering (Phase 3, Task 9). ResizableSlideItem emits
    // alignmentGuidesRequested while dragging when its edges (or its center)
    // align to a neighbour's edges (or to a slide centerline); these slots
    // render the QLineF list as magenta dashed QGraphicsLineItems at z=1000
    // above all canvas items. alignmentGuidesCleared tears them down on
    // release (and whenever the alignment check finds no matches).
    void showAlignmentGuides(const QVector<QLineF>& guides);
    void clearAlignmentGuides();

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
    QTimer*               m_autoSaveTimer = nullptr;

    // Active magenta dashed alignment-guide line items (Phase 3, Task 9).
    // Lifetime: items are children of m_scene, so m_scene->clear() in
    // populateCanvas() deletes them; clearAlignmentGuides() handles the
    // non-clear teardown path (drag-release / no-match). Always clear this
    // vector after m_scene->clear() to avoid dangling pointers.
    QVector<QGraphicsLineItem*> m_guideItems;

    // Snap-to-grid toolbar (Phase 3, Task 6). The thin toolbar sits above the
    // canvas in the preview dialog with one checkable QToolButton. State is
    // persisted to QSettings("SDR", "DataViewerEnterprise") under "preview/snap"
    // (default true) and propagated to ResizableSlideItems via
    // propagateSnapStateToItems(). Task 7 wires the actual snap math.
    QToolBar*    m_toolbar      = nullptr;
    QToolButton* m_snapBtn      = nullptr;
    bool         m_snapEnabled  = true;          // default ON
    static constexpr double kGridInches = 0.05;  // grid spacing (finer per user)

    // Undo/redo command stacks. Capped at kMaxUndoDepth — oldest entry is
    // dropped from the bottom of the undo stack when the cap is hit.
    QStack<QSharedPointer<LayoutCommand>> m_undoStack;
    QStack<QSharedPointer<LayoutCommand>> m_redoStack;
    static constexpr int kMaxUndoDepth = 100;
};

} // namespace DVE
