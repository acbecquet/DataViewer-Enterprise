#pragma once

#include <QDialog>
#include <QMap>
#include <QSet>
#include <QVector>

#include "../pipeline/ReportData.h"

class QListWidget;
class QTableWidget;
class QDoubleSpinBox;
class QLabel;
class QCheckBox;
class QRadioButton;

namespace DVE {

// ─── DataCleanupDialog ────────────────────────────────────────────────────────
// Modal dialog for non-destructive outlier exclusion.
//
// Threshold controls auto-apply as you type — no separate "Apply" click needed.
// Scope toggle (Current Sample / All Samples) determines which samples the
// threshold affects.  Individual rows can also be toggled manually.
//
// Click Apply (OK) to confirm and close; the caller retrieves the result via
// exclusions(). Cancel discards all changes.
// ─────────────────────────────────────────────────────────────────────────────
class DataCleanupDialog : public QDialog {
    Q_OBJECT
public:
    // sheet          — current sheet (read-only; dialog does NOT modify it)
    // initExclusions — pre-existing exclusions to start with
    explicit DataCleanupDialog(const SheetResult& sheet,
                               const QMap<int, QSet<int>>& initExclusions,
                               QWidget* parent = nullptr);

    // Call after exec() == QDialog::Accepted.
    QMap<int, QSet<int>> exclusions() const { return m_exclusions; }

private slots:
    void onSampleSelected(int listRow);
    void onRowCheckChanged(int tableRow, int col);
    void onThresholdChanged();   // auto-applies on any threshold control change
    void onResetCurrent();
    void onResetAll();

private:
    void populateSampleList();
    void populateRowTable(int sampleIdx);
    void updateSampleListItem(int sampleIdx);
    void updateStats(int sampleIdx);
    // Replace the exclusions for sampleIdx with the current threshold state.
    // Rows that fall outside [minTPM, maxTPM] become excluded; others are cleared.
    void applyThresholdToSample(int sampleIdx);

    // Maps visible table row → index in SampleResult::rows (full vector)
    QVector<int> m_rowMapping;

    const SheetResult& m_sheet;
    QMap<int, QSet<int>> m_exclusions;   // sampleIdx → set of excluded row indices
    int m_currentSampleIdx = -1;

    // Widgets
    QListWidget*    m_sampleList         = nullptr;
    QTableWidget*   m_rowTable           = nullptr;
    QDoubleSpinBox* m_minTPM             = nullptr;
    QDoubleSpinBox* m_maxTPM             = nullptr;
    QCheckBox*      m_hasMin             = nullptr;
    QCheckBox*      m_hasMax             = nullptr;
    QRadioButton*   m_applyCurrentRadio  = nullptr;
    QRadioButton*   m_applyAllRadio      = nullptr;
    QLabel*         m_statsLabel         = nullptr;
};

} // namespace DVE
