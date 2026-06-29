#pragma once
#include <QWidget>
#include <QScrollArea>
#include <QResizeEvent>
#include <QLabel>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QComboBox>
#include <QCheckBox>
#include <QPushButton>
#include <QFrame>
#include <QList>
#include <QVector>

#include "PlotEngine.h"
#include "../pipeline/ReportData.h"

namespace DVE {

// ─── PlotWidget ───────────────────────────────────────────────────────────────
// Composite widget that displays a dynamically updated plot for a SheetResult.
//
// Layout:
//  ┌─────────────────────────────────────────────────────────────────────┐
//  │  [Plot Type ▼]  [+] [-] [⊡] [💾]          (top control bar)       │
//  ├─────────────────────────────────────────────────────────────────────┤
//  │  [☑ Sample A] [☑ Sample B] [☑ Sample C] ...  (horizontal scroll)  │
//  ├─────────────────────────────────────────────────────────────────────┤
//  │                                                                     │
//  │                    Plot image area (scrollable)                     │
//  │                                                                     │
//  └─────────────────────────────────────────────────────────────────────┘

class PlotWidget : public QWidget {
    Q_OBJECT
public:
    explicit PlotWidget(QWidget* parent = nullptr);

    // Load a new sheet result and rebuild the sample list / plot.
    void setSheetData(const SheetResult& sheet);

    // Populate the regime picker with the file's unique per-row regimes.
    // Empty list hides the picker (old-template files). Preserves the current selection.
    void setAvailableRegimes(const QStringList& regimes);

    // Remove current data and show a blank state.
    void clear();

    // Return the currently displayed plot as PNG bytes (for report export).
    // Returns an empty QByteArray if no plot is currently displayed.
    QByteArray currentPlotPng(int dpi = 150) const;

    // Dock a widget at the right end of the top control bar (after the stretch),
    // so it shares the Plot Type / Regime / Save plot row instead of taking its
    // own band above the plot. Used by MainWindow to host the presence avatars.
    void setHeaderTrailingWidget(QWidget* w);

public slots:
    // Note→plot linking (DATAVIEWER-6 v1): emphasise the TPM-trend point at the
    // given cumulative-puff count (the row a clicked note card belongs to).
    // Pass -1 to clear the emphasis. Re-renders the current plot.
    void selectPuff(int puffs);

signals:
    void plotTypeChanged(const QString& plotType);

private slots:
    void onPlotTypeChanged(int index);
    void onRegimeChanged(int index);
    void onSampleToggled(int sampleIndex, bool checked);
    void onOilToggled(int sampleIndex, bool checked);
    void onSaveImage();

protected:
    void resizeEvent(QResizeEvent* event) override;

private:
    // Rebuild and display the current plot.
    void updatePlot();

    // Render the currently selected plot type for the visible samples.
    QPixmap renderCurrentPlot() const;

    // Show or hide the oil consumed overlay checkboxes.
    void updateOilCheckboxVisibility(bool visible);

    // ── Widgets ────────────────────────────────────────────────────────────
    QHBoxLayout* m_topBarLayout = nullptr;   // top control-bar layout (trailing widgets append here)
    QComboBox*   m_plotTypeCombo;    // "TPM Trend" | "TPM Bar Chart" | "Power Density"
    QPushButton* m_saveBtn;
    QComboBox*   m_regimeCombo  = nullptr;   // "All regimes" + each unique per-row regime
    QLabel*      m_regimeLabel  = nullptr;

    QLabel*      m_plotLabel;        // Displays the rendered QPixmap
    QScrollArea* m_plotScrollArea;   // Wraps m_plotLabel so large plots can scroll

    QScrollArea* m_checkboxScrollArea; // Horizontal scroll for sample checkboxes
    QWidget*     m_checkboxPanel;
    QHBoxLayout* m_checkboxLayout;
    QList<QCheckBox*> m_sampleCheckboxes;

    // Oil consumed overlay checkboxes (TPM Trend dual-axis)
    QFrame*           m_oilSeparator = nullptr;
    QLabel*           m_oilLabel     = nullptr;
    QList<QCheckBox*> m_oilCheckboxes;
    QVector<bool>     m_oilVisible;

    // ── Data ───────────────────────────────────────────────────────────────
    SheetResult   m_currentSheet;
    QVector<bool> m_sampleVisible;   // parallel to m_currentSheet.samples
    mutable QPixmap m_currentPixmap; // last rendered pixmap (mutable for const getter)
    double        m_zoomFactor = 1.0;
    int           m_selectedPuff = -1;  // emphasised TPM-trend point; -1 = none
};

} // namespace DVE
