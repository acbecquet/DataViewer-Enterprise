#pragma once
#include <QDialog>
#include <QGraphicsView>
#include <QGraphicsScene>
#include <QListWidget>
#include <QRectF>
#include <QSet>
#include <QStack>
#include <QSharedPointer>
#include <QTimer>
#include "reporting/ReportLayout.h"
#include "reporting/IReportSource.h"

class QShortcut;

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

    // Push the item back inside the 13.33 x 7.5 inch slide if its current rect
    // (which may have grown larger than the layout-set rect via TextItem /
    // TableItem auto-grow) extends past the slide edges. Position-only fix —
    // does not shrink the item.
    void clampToSlide(class ResizableSlideItem* item);

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

    // Undo/redo command stacks. Capped at kMaxUndoDepth — oldest entry is
    // dropped from the bottom of the undo stack when the cap is hit.
    QStack<QSharedPointer<LayoutCommand>> m_undoStack;
    QStack<QSharedPointer<LayoutCommand>> m_redoStack;
    static constexpr int kMaxUndoDepth = 100;
};

} // namespace DVE
