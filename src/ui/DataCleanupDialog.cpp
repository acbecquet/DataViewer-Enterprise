#include "DataCleanupDialog.h"

#include <QApplication>
#include <QButtonGroup>
#include <QCheckBox>
#include <QDialogButtonBox>
#include <QDoubleSpinBox>
#include <QFont>
#include <QGroupBox>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QLabel>
#include <QListWidget>
#include <QPushButton>
#include <QRadioButton>
#include <QSplitter>
#include <QTableWidget>
#include <QVBoxLayout>
#include <cmath>

namespace DVE {

// ─────────────────────────────────────────────────────────────────────────────
// Construction
// ─────────────────────────────────────────────────────────────────────────────

DataCleanupDialog::DataCleanupDialog(const SheetResult& sheet,
                                     const QMap<int, QSet<int>>& initExclusions,
                                     QWidget* parent)
    : QDialog(parent)
    , m_sheet(sheet)
    , m_exclusions(initExclusions)
{
    setWindowTitle("Data Cleanup \u2013 " + sheet.sheetName);
    setMinimumSize(700, 520);
    resize(820, 580);

    // ── Threshold controls ────────────────────────────────────────────────────
    QGroupBox* threshBox = new QGroupBox("Threshold Filter  (auto-applies as you type)");
    QHBoxLayout* threshHL = new QHBoxLayout(threshBox);
    threshHL->setSpacing(8);

    m_hasMin = new QCheckBox("Exclude TPM below:");
    m_hasMin->setChecked(true);
    m_minTPM = new QDoubleSpinBox();
    m_minTPM->setRange(0.0, 99999.0);
    m_minTPM->setDecimals(3);
    m_minTPM->setSuffix(" mg/puff");
    m_minTPM->setValue(1.0);
    m_minTPM->setMinimumWidth(120);

    m_hasMax = new QCheckBox("Exclude TPM above:");
    m_hasMax->setChecked(true);
    m_maxTPM = new QDoubleSpinBox();
    m_maxTPM->setRange(0.0, 99999.0);
    m_maxTPM->setDecimals(3);
    m_maxTPM->setSuffix(" mg/puff");
    m_maxTPM->setValue(50.0);
    m_maxTPM->setMinimumWidth(120);

    // Scope: radio button toggle — Current Sample | All Samples
    m_applyCurrentRadio = new QRadioButton("Current sample");
    m_applyAllRadio     = new QRadioButton("All samples");
    m_applyAllRadio->setChecked(true);   // default: apply to all

    auto* scopeGroup = new QButtonGroup(this);
    scopeGroup->addButton(m_applyCurrentRadio);
    scopeGroup->addButton(m_applyAllRadio);

    QFrame* sep = new QFrame();
    sep->setFrameShape(QFrame::VLine);
    sep->setFrameShadow(QFrame::Sunken);
    sep->setStyleSheet("color: #BCBCBC;");

    threshHL->addWidget(m_hasMin);
    threshHL->addWidget(m_minTPM);
    threshHL->addSpacing(16);
    threshHL->addWidget(m_hasMax);
    threshHL->addWidget(m_maxTPM);
    threshHL->addSpacing(8);
    threshHL->addWidget(sep);
    threshHL->addSpacing(8);
    threshHL->addWidget(new QLabel("Apply to:"));
    threshHL->addWidget(m_applyCurrentRadio);
    threshHL->addWidget(m_applyAllRadio);
    threshHL->addStretch(1);

    // ── Left: sample list ─────────────────────────────────────────────────────
    m_sampleList = new QListWidget();
    m_sampleList->setMinimumWidth(110);
    m_sampleList->setMaximumWidth(160);
    m_sampleList->setAlternatingRowColors(true);

    // ── Right: row table ──────────────────────────────────────────────────────
    m_rowTable = new QTableWidget();
    m_rowTable->setColumnCount(7);
    m_rowTable->setHorizontalHeaderLabels({
        "Include", "#", "Puffs", "Before (g)", "After (g)", "TPM (mg/puff)", "Oil (mg)"
    });
    m_rowTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    m_rowTable->horizontalHeader()->setStretchLastSection(true);
    m_rowTable->setColumnWidth(0, 62);
    m_rowTable->setColumnWidth(1, 36);
    m_rowTable->setColumnWidth(2, 56);
    m_rowTable->setColumnWidth(3, 86);
    m_rowTable->setColumnWidth(4, 86);
    m_rowTable->setColumnWidth(5, 106);
    m_rowTable->verticalHeader()->setVisible(false);
    m_rowTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    m_rowTable->setSelectionMode(QAbstractItemView::NoSelection);
    m_rowTable->setAlternatingRowColors(true);

    m_statsLabel = new QLabel("Select a sample to see cleanup statistics.");
    m_statsLabel->setAlignment(Qt::AlignCenter);
    m_statsLabel->setStyleSheet(
        "font-size: 9pt; padding: 4px;"
        "border-top: 1px solid #BCBCBC; background: #F7F7F7;");

    QWidget* rightPanel = new QWidget();
    QVBoxLayout* rightVL = new QVBoxLayout(rightPanel);
    rightVL->setContentsMargins(0, 0, 0, 0);
    rightVL->setSpacing(0);
    rightVL->addWidget(m_rowTable, 1);
    rightVL->addWidget(m_statsLabel);

    QSplitter* splitter = new QSplitter(Qt::Horizontal);
    splitter->addWidget(m_sampleList);
    splitter->addWidget(rightPanel);
    splitter->setStretchFactor(0, 0);
    splitter->setStretchFactor(1, 1);

    // ── Bottom buttons ────────────────────────────────────────────────────────
    auto* resetCurBtn = new QPushButton("Reset Current Sample");
    auto* resetAllBtn = new QPushButton("Reset All Samples");

    auto* btnBox = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel);
    btnBox->button(QDialogButtonBox::Ok)->setText("Apply");

    QHBoxLayout* bottomHL = new QHBoxLayout();
    bottomHL->addWidget(resetCurBtn);
    bottomHL->addWidget(resetAllBtn);
    bottomHL->addStretch(1);
    bottomHL->addWidget(btnBox);

    // ── Main layout ───────────────────────────────────────────────────────────
    QVBoxLayout* mainVL = new QVBoxLayout(this);
    mainVL->addWidget(threshBox);
    mainVL->addWidget(splitter, 1);
    mainVL->addLayout(bottomHL);

    // ── Connections ───────────────────────────────────────────────────────────
    // Threshold controls → auto-apply
    connect(m_hasMin,  &QCheckBox::toggled,
            this, &DataCleanupDialog::onThresholdChanged);
    connect(m_hasMax,  &QCheckBox::toggled,
            this, &DataCleanupDialog::onThresholdChanged);
    connect(m_minTPM,  QOverload<double>::of(&QDoubleSpinBox::valueChanged),
            this, &DataCleanupDialog::onThresholdChanged);
    connect(m_maxTPM,  QOverload<double>::of(&QDoubleSpinBox::valueChanged),
            this, &DataCleanupDialog::onThresholdChanged);
    // Scope toggle also re-applies
    connect(m_applyCurrentRadio, &QRadioButton::toggled,
            this, &DataCleanupDialog::onThresholdChanged);

    // Enable/disable spinboxes based on checkbox state
    connect(m_hasMin, &QCheckBox::toggled, m_minTPM, &QDoubleSpinBox::setEnabled);
    connect(m_hasMax, &QCheckBox::toggled, m_maxTPM, &QDoubleSpinBox::setEnabled);

    connect(m_sampleList, &QListWidget::currentRowChanged,
            this, &DataCleanupDialog::onSampleSelected);

    connect(m_rowTable, &QTableWidget::itemChanged, this, [this](QTableWidgetItem* item) {
        if (item && item->column() == 0)
            onRowCheckChanged(item->row(), 0);
    });

    connect(resetCurBtn, &QPushButton::clicked, this, &DataCleanupDialog::onResetCurrent);
    connect(resetAllBtn, &QPushButton::clicked, this, &DataCleanupDialog::onResetAll);
    connect(btnBox, &QDialogButtonBox::accepted, this, &QDialog::accept);
    connect(btnBox, &QDialogButtonBox::rejected, this, &QDialog::reject);

    // ── Initial state ─────────────────────────────────────────────────────────
    // Apply the default threshold to all samples before showing the UI so the
    // dialog opens with outliers already excluded (user's expectation).
    if (!m_sheet.samples.isEmpty()) {
        for (int i = 0; i < m_sheet.samples.size(); ++i)
            applyThresholdToSample(i);
    }

    populateSampleList();
    if (!m_sheet.samples.isEmpty())
        m_sampleList->setCurrentRow(0);
}


// ─────────────────────────────────────────────────────────────────────────────
// Private helpers
// ─────────────────────────────────────────────────────────────────────────────

void DataCleanupDialog::populateSampleList()
{
    m_sampleList->blockSignals(true);
    m_sampleList->clear();
    for (int i = 0; i < m_sheet.samples.size(); ++i)
        updateSampleListItem(i);
    m_sampleList->blockSignals(false);
}

void DataCleanupDialog::updateSampleListItem(int si)
{
    const SampleResult& sr = m_sheet.samples[si];

    int validRows = 0;
    for (const DataRow& dr : sr.rows)
        if (dr.beforeWeight != 0.0 && dr.afterWeight != 0.0)
            ++validRows;

    const int excl = m_exclusions.value(si).size();

    QString name = sr.sampleName.isEmpty()
                   ? QString("Sample %1").arg(si + 1)
                   : sr.sampleName;

    QString label = excl > 0
                    ? QString("%1\n(%2 rows, %3 excl.)").arg(name).arg(validRows).arg(excl)
                    : QString("%1\n(%2 rows)").arg(name).arg(validRows);

    QListWidgetItem* item = nullptr;
    if (si < m_sampleList->count()) {
        item = m_sampleList->item(si);
        item->setText(label);
    } else {
        item = new QListWidgetItem(label);
        m_sampleList->addItem(item);
    }

    item->setForeground(excl > 0
                        ? QColor(0xCC, 0x44, 0x00)
                        : QApplication::palette().windowText().color());
}

void DataCleanupDialog::populateRowTable(int si)
{
    m_rowTable->blockSignals(true);
    m_rowTable->setRowCount(0);
    m_rowMapping.clear();

    if (si < 0 || si >= m_sheet.samples.size()) {
        m_rowTable->blockSignals(false);
        return;
    }

    const SampleResult& sr   = m_sheet.samples[si];
    const QSet<int>&    excl = m_exclusions.value(si);

    int rowNum = 0;
    for (int i = 0; i < sr.rows.size(); ++i) {
        const DataRow& dr = sr.rows[i];
        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;

        const bool isExcluded = excl.contains(i);
        const int  tRow       = m_rowTable->rowCount();
        m_rowTable->insertRow(tRow);
        m_rowMapping.append(i);

        auto* chk = new QTableWidgetItem();
        chk->setFlags(Qt::ItemIsEnabled | Qt::ItemIsUserCheckable);
        chk->setCheckState(isExcluded ? Qt::Unchecked : Qt::Checked);
        chk->setTextAlignment(Qt::AlignCenter);
        m_rowTable->setItem(tRow, 0, chk);

        auto* numItm = new QTableWidgetItem(QString::number(++rowNum));
        numItm->setTextAlignment(Qt::AlignCenter);
        m_rowTable->setItem(tRow, 1, numItm);

        auto setCell = [&](int col, double val, int dp) {
            auto* itm = new QTableWidgetItem(QString::number(val, 'f', dp));
            itm->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
            m_rowTable->setItem(tRow, col, itm);
        };
        setCell(2, dr.puffs,        0);
        setCell(3, dr.beforeWeight, 4);
        setCell(4, dr.afterWeight,  4);
        setCell(5, dr.tpm,          4);
        setCell(6, dr.oilConsumed,  2);

        if (isExcluded) {
            for (int c = 0; c < m_rowTable->columnCount(); ++c) {
                auto* itm = m_rowTable->item(tRow, c);
                if (!itm) continue;
                itm->setForeground(QColor(0xAA, 0xAA, 0xAA));
                itm->setBackground(QColor(0xF8, 0xF0, 0xF0));
                if (c >= 2) {
                    QFont f = itm->font();
                    f.setStrikeOut(true);
                    itm->setFont(f);
                }
            }
        }
    }

    m_rowTable->blockSignals(false);
}

void DataCleanupDialog::updateStats(int si)
{
    if (si < 0 || si >= m_sheet.samples.size()) {
        m_statsLabel->setText("Select a sample to see cleanup statistics.");
        m_statsLabel->setStyleSheet(
            "font-size: 9pt; padding: 4px;"
            "border-top: 1px solid #BCBCBC; background: #F7F7F7;");
        return;
    }

    const SampleResult& sr   = m_sheet.samples[si];
    const QSet<int>&    excl = m_exclusions.value(si);

    QVector<double> included;
    int totalValid = 0;
    for (int i = 0; i < sr.rows.size(); ++i) {
        const DataRow& dr = sr.rows[i];
        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
        ++totalValid;
        if (!excl.contains(i))
            included.append(dr.tpm);
    }

    double avg = 0.0, stddev = 0.0;
    if (!included.isEmpty()) {
        for (double v : included) avg += v;
        avg /= included.size();
        double sumSq = 0.0;
        for (double v : included) sumSq += (v - avg) * (v - avg);
        stddev = (included.size() > 1)
                 ? std::sqrt(sumSq / (included.size() - 1))
                 : 0.0;
    }

    const int excCount = excl.size();
    QString msg = QString("Cleaned: Avg TPM = %1 mg/puff,  Std Dev = %2  |  "
                          "%3 / %4 rows active,  %5 excluded")
                  .arg(avg,    0, 'f', 4)
                  .arg(stddev, 0, 'f', 4)
                  .arg(included.size())
                  .arg(totalValid)
                  .arg(excCount);

    m_statsLabel->setText(msg);
    m_statsLabel->setStyleSheet(
        excCount > 0
        ? "font-size: 9pt; color: #CC4400; padding: 4px;"
          "border-top: 1px solid #BCBCBC; background: #FFF8F5;"
        : "font-size: 9pt; color: #1A6600; padding: 4px;"
          "border-top: 1px solid #BCBCBC; background: #F5FFF5;");
}

void DataCleanupDialog::applyThresholdToSample(int si)
{
    if (si < 0 || si >= m_sheet.samples.size()) return;

    // Full replace: clear existing exclusions for this sample, then add rows
    // that fall outside the active threshold bounds.
    m_exclusions.remove(si);

    const bool useMin = m_hasMin->isChecked();
    const bool useMax = m_hasMax->isChecked();
    if (!useMin && !useMax) return;   // no filter active — leave cleared

    const double minVal = useMin ? m_minTPM->value() : -1e300;
    const double maxVal = useMax ? m_maxTPM->value() :  1e300;

    const SampleResult& sr = m_sheet.samples[si];
    for (int i = 0; i < sr.rows.size(); ++i) {
        const DataRow& dr = sr.rows[i];
        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
        if (dr.tpm < minVal || dr.tpm > maxVal)
            m_exclusions[si].insert(i);
    }

    if (m_exclusions.value(si).isEmpty())
        m_exclusions.remove(si);
}


// ─────────────────────────────────────────────────────────────────────────────
// Slots
// ─────────────────────────────────────────────────────────────────────────────

void DataCleanupDialog::onSampleSelected(int listRow)
{
    if (listRow < 0 || listRow >= m_sheet.samples.size()) return;
    m_currentSampleIdx = listRow;
    populateRowTable(listRow);
    updateStats(listRow);
}

void DataCleanupDialog::onRowCheckChanged(int tRow, int /*col*/)
{
    if (m_currentSampleIdx < 0) return;
    if (tRow >= m_rowMapping.size()) return;

    auto* item = m_rowTable->item(tRow, 0);
    if (!item) return;

    const int  fullIdx = m_rowMapping[tRow];
    const bool include = (item->checkState() == Qt::Checked);

    if (include)
        m_exclusions[m_currentSampleIdx].remove(fullIdx);
    else
        m_exclusions[m_currentSampleIdx].insert(fullIdx);

    if (m_exclusions.value(m_currentSampleIdx).isEmpty())
        m_exclusions.remove(m_currentSampleIdx);

    m_rowTable->blockSignals(true);
    for (int c = 0; c < m_rowTable->columnCount(); ++c) {
        auto* itm = m_rowTable->item(tRow, c);
        if (!itm) continue;
        if (!include) {
            itm->setForeground(QColor(0xAA, 0xAA, 0xAA));
            itm->setBackground(QColor(0xF8, 0xF0, 0xF0));
            if (c >= 2) { QFont f = itm->font(); f.setStrikeOut(true); itm->setFont(f); }
        } else {
            itm->setForeground(QApplication::palette().windowText().color());
            itm->setBackground(QApplication::palette().base().color());
            if (c >= 2) { QFont f = itm->font(); f.setStrikeOut(false); itm->setFont(f); }
        }
    }
    m_rowTable->blockSignals(false);

    updateSampleListItem(m_currentSampleIdx);
    updateStats(m_currentSampleIdx);
}

void DataCleanupDialog::onThresholdChanged()
{
    // Determine scope from radio buttons
    const bool applyAll = m_applyAllRadio->isChecked();

    if (applyAll) {
        for (int i = 0; i < m_sheet.samples.size(); ++i) {
            applyThresholdToSample(i);
            updateSampleListItem(i);
        }
    } else {
        // "Current sample" — clear other samples' threshold exclusions
        for (int i = 0; i < m_sheet.samples.size(); ++i) {
            if (i != m_currentSampleIdx) {
                m_exclusions.remove(i);
                updateSampleListItem(i);
            }
        }
        if (m_currentSampleIdx >= 0) {
            applyThresholdToSample(m_currentSampleIdx);
            updateSampleListItem(m_currentSampleIdx);
        }
    }

    // Refresh the row table for the currently visible sample
    if (m_currentSampleIdx >= 0) {
        populateRowTable(m_currentSampleIdx);
        updateStats(m_currentSampleIdx);
    }
}

void DataCleanupDialog::onResetCurrent()
{
    if (m_currentSampleIdx < 0) return;
    m_exclusions.remove(m_currentSampleIdx);
    populateRowTable(m_currentSampleIdx);
    updateStats(m_currentSampleIdx);
    updateSampleListItem(m_currentSampleIdx);
}

void DataCleanupDialog::onResetAll()
{
    m_exclusions.clear();
    populateSampleList();
    if (m_currentSampleIdx >= 0) {
        populateRowTable(m_currentSampleIdx);
        updateStats(m_currentSampleIdx);
    }
}

} // namespace DVE
