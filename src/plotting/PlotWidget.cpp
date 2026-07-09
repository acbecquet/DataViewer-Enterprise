#include "PlotWidget.h"
#include "../utils/AppTheme.h"
#include "../utils/OutputPaths.h"
#include "../pipeline/RegimeUtils.h"
#include "../pipeline/TpmCalculator.h"

#include <algorithm>
#include <QLabel>
#include <QComboBox>
#include <QCheckBox>
#include <QPushButton>
#include <QScrollArea>
#include <QHBoxLayout>
#include <QVBoxLayout>
#include <QFrame>
#include <QSizePolicy>
#include <QFont>
#include <QFileDialog>
#include <QMessageBox>
#include <QScrollBar>
#include <QStyle>
#include <QMouseEvent>
#include <QEvent>

namespace DVE {

// ═══════════════════════════════════════════════════════════════════════════════
// Construction
// ═══════════════════════════════════════════════════════════════════════════════

PlotWidget::PlotWidget(QWidget* parent)
    : QWidget(parent)
{
    // ── Top control bar ───────────────────────────────────────────────────────
    QWidget*     topBar    = new QWidget(this);
    m_topBarLayout = new QHBoxLayout(topBar);
    QHBoxLayout* topLayout = m_topBarLayout;
    topLayout->setContentsMargins(6, 4, 6, 4);
    topLayout->setSpacing(6);

    QLabel* typeLabel = new QLabel("Plot Type:", topBar);
    typeLabel->setStyleSheet("font-weight: 600; font-size: 9pt;");

    m_plotTypeCombo = new QComboBox(topBar);
    m_plotTypeCombo->addItem("TPM Trend");
    m_plotTypeCombo->addItem("TPM Bar Chart");
    m_plotTypeCombo->addItem("Power Density");
    m_plotTypeCombo->addItem("Draw Pressure");
    m_plotTypeCombo->setMinimumWidth(140);
    m_plotTypeCombo->setMaximumWidth(200);

    // Zoom / save buttons
    auto makeBtn = [&](const QString& text, const QString& tip) {
        QPushButton* btn = new QPushButton(text, topBar);
        btn->setToolTip(tip);
        btn->setFixedSize(32, 26);
        // Must explicitly set color and padding — local setStyleSheet overrides
        // the application stylesheet completely, losing color: #1A1A1A.
        btn->setStyleSheet(
            "QPushButton { border: 1px solid #BCBCBC; border-radius: 3px;"
            "  background: #FFFFFF; color: #1A1A1A; font-size: 11pt; padding: 0px; }"
            "QPushButton:hover { background: #E0EEFF; border-color: #0066CC; color: #003388; }"
            "QPushButton:pressed { background: #C0D8FF; color: #003388; }");
        return btn;
    };

    m_saveBtn    = makeBtn("",   "Save image");
    m_saveBtn->setIcon(style()->standardIcon(QStyle::SP_DialogSaveButton));

    // Separator between combo and zoom buttons
    QFrame* sep = new QFrame(topBar);
    sep->setFrameShape(QFrame::VLine);
    sep->setFrameShadow(QFrame::Sunken);
    sep->setStyleSheet("color: #BCBCBC;");

    topLayout->addWidget(typeLabel);
    topLayout->addWidget(m_plotTypeCombo);

    m_regimeLabel = new QLabel("Regime:", topBar);
    m_regimeLabel->setStyleSheet("font-weight: 600; font-size: 9pt;");
    m_regimeCombo = new QComboBox(topBar);
    m_regimeCombo->addItem("All regimes");
    m_regimeCombo->setMinimumWidth(130);
    m_regimeCombo->setMaximumWidth(200);
    m_regimeLabel->setVisible(false);          // hidden until a file has regimes
    m_regimeCombo->setVisible(false);

    topLayout->addWidget(m_regimeLabel);
    topLayout->addWidget(m_regimeCombo);
    topLayout->addWidget(sep);
    topLayout->addWidget(new QLabel(tr("Save plot"), topBar));
    topLayout->addWidget(m_saveBtn);
    topLayout->addStretch(1);

    topBar->setStyleSheet(
        "background-color: #F0F0F0;"
        "border-bottom: 1px solid #BCBCBC;"
    );
    // 40px (not 36) so the docked presence avatar bar (28px circle + 2*2px
    // margin = 32px min) renders uncut alongside the controls; the existing
    // combos/buttons are unaffected.
    topBar->setMinimumHeight(40);

    // ── Horizontal sample checkbox bar ────────────────────────────────────────
    m_checkboxPanel  = new QWidget();
    m_checkboxLayout = new QHBoxLayout(m_checkboxPanel);
    m_checkboxLayout->setContentsMargins(6, 3, 6, 3);
    m_checkboxLayout->setSpacing(12);
    m_checkboxLayout->setAlignment(Qt::AlignVCenter | Qt::AlignLeft);
    m_checkboxLayout->addStretch(1);   // push checkboxes left; stretch at end

    m_checkboxScrollArea = new QScrollArea();
    m_checkboxScrollArea->setWidget(m_checkboxPanel);
    m_checkboxScrollArea->setWidgetResizable(true);
    m_checkboxScrollArea->setVerticalScrollBarPolicy(Qt::ScrollBarAlwaysOff);
    m_checkboxScrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAsNeeded);
    // Floor at the v2.6.0 fixed 32px (room for the row + its AsNeeded
    // horizontal scrollbar); the font-derived term takes over under scaling.
    m_checkboxScrollArea->setMinimumHeight(qMax(32, AppTheme::controlHeight() + 6));
    m_checkboxScrollArea->setStyleSheet(
        "QScrollArea { border: none; border-bottom: 1px solid #BCBCBC;"
        "  background-color: #F7F7F7; }"
    );

    // ── Plot display area ─────────────────────────────────────────────────────
    m_plotLabel = new QLabel();
    m_plotLabel->setAlignment(Qt::AlignCenter);
    m_plotLabel->setMinimumSize(400, 300);
    m_plotLabel->setSizePolicy(QSizePolicy::Expanding, QSizePolicy::Expanding);
    m_plotLabel->setStyleSheet("background-color: #FFFFFF;");
    m_plotLabel->setText("<span style='color:#AAAAAA; font-size:11pt;'>"
                         "No data loaded</span>");
    m_plotLabel->setTextFormat(Qt::RichText);
    // DV-18: a plain QLabel has no click signal — catch presses via an event
    // filter and hit-test them against the last captured TPM-trend transform.
    m_plotLabel->installEventFilter(this);

    m_plotScrollArea = new QScrollArea();
    m_plotScrollArea->setWidget(m_plotLabel);
    m_plotScrollArea->setWidgetResizable(false);   // false so we control size
    m_plotScrollArea->setAlignment(Qt::AlignCenter);
    m_plotScrollArea->setStyleSheet(
        "QScrollArea { border: none; background-color: #FFFFFF; }"
    );

    // ── Main layout ───────────────────────────────────────────────────────────
    QVBoxLayout* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(0, 0, 0, 0);
    mainLayout->setSpacing(0);
    mainLayout->addWidget(topBar);
    mainLayout->addWidget(m_checkboxScrollArea);
    mainLayout->addWidget(m_plotScrollArea, 1);

    // ── Connections ───────────────────────────────────────────────────────────
    connect(m_plotTypeCombo,
            QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &PlotWidget::onPlotTypeChanged);
    connect(m_regimeCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &PlotWidget::onRegimeChanged);
    connect(m_saveBtn,    &QPushButton::clicked, this, &PlotWidget::onSaveImage);
}


// ═══════════════════════════════════════════════════════════════════════════════
// Public interface
// ═══════════════════════════════════════════════════════════════════════════════

void PlotWidget::setHeaderTrailingWidget(QWidget* w)
{
    if (!w || !m_topBarLayout) return;
    // The top bar already ends with addStretch(1); appending after it docks the
    // widget hard against the right edge, on the same row as the plot controls.
    w->setParent(nullptr);                 // detach from any prior layout/parent
    m_topBarLayout->addWidget(w);
}

void PlotWidget::setSheetData(const SheetResult& sheet)
{
    m_currentSheet = sheet;
    m_zoomFactor   = 1.0;
    m_selectedPuff = -1;   // a new sheet drops any note-click emphasis

    // ── Clean up old checkboxes ───────────────────────────────────────────────
    for (QCheckBox* cb : m_sampleCheckboxes) delete cb;
    m_sampleCheckboxes.clear();
    m_sampleVisible.clear();

    for (QCheckBox* cb : m_oilCheckboxes) delete cb;
    m_oilCheckboxes.clear();
    m_oilVisible.clear();

    if (m_oilSeparator) { delete m_oilSeparator; m_oilSeparator = nullptr; }
    if (m_oilLabel)     { delete m_oilLabel;     m_oilLabel     = nullptr; }

    while (m_checkboxLayout->count() > 0) {
        QLayoutItem* item = m_checkboxLayout->takeAt(0);
        delete item;
    }

    // ── Rebuild primary (TPM) sample checkboxes ───────────────────────────────
    // Block signals during rebuild to avoid spurious intermediate re-renders
    // (each setChecked(true) would otherwise trigger updatePlot with partial data).
    for (int i = 0; i < sheet.samples.size(); ++i) {
        const SampleResult& sr  = sheet.samples[i];
        QString             lbl = sr.sampleName.isEmpty()
                                  ? QString("Sample %1").arg(i + 1)
                                  : sr.sampleName;

        QCheckBox* cb = new QCheckBox(lbl, m_checkboxPanel);
        cb->blockSignals(true);
        cb->setChecked(true);
        cb->blockSignals(false);

        QColor c = AppTheme::seriesColor(i);
        cb->setStyleSheet(QString(
            "QCheckBox { font-size: 8pt; padding: 0px 2px; }"
            "QCheckBox::indicator:checked { background-color: %1; border: 1px solid %2; }"
        ).arg(c.name(), c.darker(130).name()));

        const int idx = i;
        connect(cb, &QCheckBox::toggled, this, [this, idx](bool checked) {
            onSampleToggled(idx, checked);
        });

        m_checkboxLayout->addWidget(cb);
        m_sampleCheckboxes.append(cb);
        m_sampleVisible.append(true);
    }

    // ── Oil consumed overlay checkboxes (hidden by default) ───────────────────
    if (!sheet.samples.isEmpty()) {
        m_oilSeparator = new QFrame(m_checkboxPanel);
        m_oilSeparator->setFrameShape(QFrame::VLine);
        m_oilSeparator->setFrameShadow(QFrame::Plain);
        m_oilSeparator->setFixedWidth(1);
        m_oilSeparator->setStyleSheet("color: #BCBCBC;");
        m_oilSeparator->setVisible(false);
        m_checkboxLayout->addWidget(m_oilSeparator);

        m_oilLabel = new QLabel("Oil:", m_checkboxPanel);
        m_oilLabel->setStyleSheet(
            "font-size: 8pt; font-weight: 600; color: #885500; padding: 0px 4px;");
        m_oilLabel->setVisible(false);
        m_checkboxLayout->addWidget(m_oilLabel);
    }

    for (int i = 0; i < sheet.samples.size(); ++i) {
        const SampleResult& sr  = sheet.samples[i];
        QString             lbl = (sr.sampleName.isEmpty()
                                   ? QString("S%1").arg(i + 1)
                                   : sr.sampleName) + " (Oil)";

        QCheckBox* cb = new QCheckBox(lbl, m_checkboxPanel);
        cb->setChecked(false);
        cb->setVisible(false);
        QColor oc = AppTheme::seriesColor(i).darker(120);
        cb->setStyleSheet(QString(
            "QCheckBox { font-size: 8pt; padding: 0px 2px; color: #885500; }"
            "QCheckBox::indicator:checked { background-color: %1; border: 1px solid %2; }"
        ).arg(oc.name(), oc.darker(130).name()));

        const int idx = i;
        connect(cb, &QCheckBox::toggled, this, [this, idx](bool checked) {
            onOilToggled(idx, checked);
        });

        m_checkboxLayout->addWidget(cb);
        m_oilCheckboxes.append(cb);
        m_oilVisible.append(false);
    }

    m_checkboxLayout->addStretch(1);

    // Show oil controls only when TPM Trend is selected
    updateOilCheckboxVisibility(m_plotTypeCombo->currentText() == "TPM Trend");

    updatePlot();
}

void PlotWidget::setAvailableRegimes(const QStringList& regimes)
{
    const QString prev = m_regimeCombo ? m_regimeCombo->currentText() : QString();
    m_regimeCombo->blockSignals(true);
    m_regimeCombo->clear();
    m_regimeCombo->addItem("All regimes");
    m_regimeCombo->addItems(regimes);
    int idx = m_regimeCombo->findText(prev);
    m_regimeCombo->setCurrentIndex(idx >= 0 ? idx : 0);     // preserve selection if still present
    m_regimeCombo->blockSignals(false);

    const bool show = !regimes.isEmpty();
    m_regimeLabel->setVisible(show);
    m_regimeCombo->setVisible(show);
    updatePlot();
}

void PlotWidget::clear()
{
    m_currentSheet = SheetResult{};
    m_zoomFactor   = 1.0;

    for (QCheckBox* cb : m_sampleCheckboxes) delete cb;
    m_sampleCheckboxes.clear();
    m_sampleVisible.clear();

    for (QCheckBox* cb : m_oilCheckboxes) delete cb;
    m_oilCheckboxes.clear();
    m_oilVisible.clear();

    if (m_oilSeparator) { delete m_oilSeparator; m_oilSeparator = nullptr; }
    if (m_oilLabel)     { delete m_oilLabel;     m_oilLabel     = nullptr; }

    while (m_checkboxLayout->count() > 0) {
        QLayoutItem* item = m_checkboxLayout->takeAt(0);
        delete item;
    }
    m_checkboxLayout->addStretch(1);

    m_currentPixmap = QPixmap();
    m_selectedPuff  = -1;
    m_plotLabel->setPixmap(QPixmap());
    m_plotLabel->setText("<span style='color:#AAAAAA; font-size:11pt;'>"
                         "No data loaded</span>");
}

void PlotWidget::selectPuff(int puffs)
{
    if (m_selectedPuff == puffs) return;   // no change, skip the re-render
    m_selectedPuff = puffs;
    updatePlot();
}

QByteArray PlotWidget::currentPlotPng(int dpi) const
{
    if (m_currentPixmap.isNull())
        return {};
    return PlotEngine::toPng(m_currentPixmap, dpi);
}


// ═══════════════════════════════════════════════════════════════════════════════
// Slots
// ═══════════════════════════════════════════════════════════════════════════════

void PlotWidget::onPlotTypeChanged(int /*index*/)
{
    m_zoomFactor = 1.0;
    updateOilCheckboxVisibility(m_plotTypeCombo->currentText() == "TPM Trend");
    updatePlot();
    emit plotTypeChanged(m_plotTypeCombo->currentText());
}

void PlotWidget::onRegimeChanged(int /*index*/) { updatePlot(); }

void PlotWidget::onSampleToggled(int sampleIndex, bool checked)
{
    if (sampleIndex >= 0 && sampleIndex < m_sampleVisible.size()) {
        m_sampleVisible[sampleIndex] = checked;
        updatePlot();
    }
}

void PlotWidget::onOilToggled(int sampleIndex, bool checked)
{
    if (sampleIndex >= 0 && sampleIndex < m_oilVisible.size()) {
        m_oilVisible[sampleIndex] = checked;
        updatePlot();
    }
}

void PlotWidget::updateOilCheckboxVisibility(bool visible)
{
    if (m_oilSeparator) m_oilSeparator->setVisible(visible);
    if (m_oilLabel)     m_oilLabel->setVisible(visible);
    for (QCheckBox* cb : m_oilCheckboxes)
        cb->setVisible(visible);
}

void PlotWidget::resizeEvent(QResizeEvent* event)
{
    QWidget::resizeEvent(event);
    if (!m_currentPixmap.isNull())
        updatePlot();
}

bool PlotWidget::eventFilter(QObject* watched, QEvent* event)
{
    if (watched == m_plotLabel && event->type() == QEvent::MouseButtonPress
        && m_lastTransform.valid && !m_currentPixmap.isNull()) {
        auto* me = static_cast<QMouseEvent*>(event);
        // m_plotLabel shows m_currentPixmap scaled by m_zoomFactor; the label
        // is resized to exactly the scaled pixmap's size (see updatePlot()),
        // and QScrollArea::setAlignment(AlignCenter) only repositions the
        // label within the viewport, it never inflates the label's own
        // width()/height() — so this offset is 0 in current usage, but is
        // kept so a click still lands correctly if that ever changes.
        const QPointF lblPos = me->position();
        const double  z      = (m_zoomFactor > 0 ? m_zoomFactor : 1.0);
        const QSize   scaled = m_currentPixmap.size() * z;
        const int     offX   = qMax(0, (m_plotLabel->width()  - scaled.width())  / 2);
        const int     offY   = qMax(0, (m_plotLabel->height() - scaled.height()) / 2);
        // Map the click back into unzoomed native pixmap pixels — the same
        // pixel space m_lastTransform was captured in.
        const QPoint pmPx(int((lblPos.x() - offX) / z), int((lblPos.y() - offY) / z));

        const QVector<NotePoint> pts = collectVisibleNotePoints();
        int si = -1, row = -1;
        if (nearestNotePoint(pmPx, m_lastTransform, pts, 18, si, row))
            emit plotPointActivated(si, row);
        return true;
    }
    return QWidget::eventFilter(watched, event);
}

QVector<NotePoint> PlotWidget::collectVisibleNotePoints() const
{
    // Same filter renderCurrentPlot()'s TPM Trend primary-series loop uses to
    // build PlotSeries::ringed (visible sample, weighed row, regime match,
    // non-empty row.notes) — so hit-testing always agrees with what's
    // actually drawn on screen (a note-bearing row that was filtered out of
    // the plotted series must not be clickable).
    const QString selRegime = (m_regimeCombo && m_regimeCombo->isVisible())
                              ? m_regimeCombo->currentText() : QString();
    const bool filterRegime = !selRegime.isEmpty() && selRegime != QLatin1String("All regimes");

    QVector<NotePoint> pts;
    for (int si = 0; si < m_currentSheet.samples.size(); ++si) {
        if (si >= m_sampleVisible.size() || !m_sampleVisible[si]) continue;
        const SampleResult& sr = m_currentSheet.samples[si];
        for (int ri = 0; ri < sr.rows.size(); ++ri) {
            const DataRow& row = sr.rows[ri];
            if (row.beforeWeight == 0.0 || row.afterWeight == 0.0) continue;
            if (filterRegime && RegimeUtils::regimeKey(row) != selRegime) continue;
            if (row.notes.trimmed().isEmpty()) continue;
            NotePoint np;
            np.sampleIndex  = si;
            np.dataRowIndex = ri;
            np.puffs        = row.puffs;
            np.tpm          = row.tpm;
            pts.append(np);
        }
    }
    return pts;
}

void PlotWidget::onSaveImage()
{
    if (m_currentPixmap.isNull()) {
        QMessageBox::information(this, "Save Image", "No plot to save.");
        return;
    }

    QString path = QFileDialog::getSaveFileName(
        this, "Save Plot Image", OutputPaths::resolveDir(ReportMode::Tpm, QString()),
        "PNG Image (*.png);;JPEG Image (*.jpg);;BMP Image (*.bmp)");

    if (path.isEmpty())
        return;

    if (!m_currentPixmap.save(path)) {
        QMessageBox::warning(this, "Save Image",
            "Failed to save image to:\n" + path);
    }
}


// ═══════════════════════════════════════════════════════════════════════════════
// Private helpers
// ═══════════════════════════════════════════════════════════════════════════════

void PlotWidget::updatePlot()
{
    m_currentPixmap = renderCurrentPlot();

    if (m_currentPixmap.isNull()) {
        m_plotLabel->setPixmap(QPixmap());
        m_plotLabel->setText("<span style='color:#AAAAAA; font-size:10pt;'>"
                             "No plot data available for the current selection."
                             "</span>");
        m_plotLabel->resize(m_plotScrollArea->viewport()->size());
        return;
    }

    m_plotLabel->setText(QString());

    // Apply zoom factor to the native pixmap
    if (m_zoomFactor == 1.0) {
        m_plotLabel->setPixmap(m_currentPixmap);
        m_plotLabel->resize(m_currentPixmap.size());
    } else {
        QSize scaledSize = m_currentPixmap.size() * m_zoomFactor;
        QPixmap scaled = m_currentPixmap.scaled(
            scaledSize, Qt::KeepAspectRatio, Qt::SmoothTransformation);
        m_plotLabel->setPixmap(scaled);
        m_plotLabel->resize(scaled.size());
    }
}

QPixmap PlotWidget::renderCurrentPlot() const
{
    // DV-18: invalidate the captured hit-test transform up front. It is only
    // re-validated below, inside the TPM Trend branch's renderLinePlot/
    // renderLinePlotDualAxis calls — every other plot type (and the
    // single-sample renderTPMTrend() shortcut, which has no out-param) leaves
    // it false, so click hit-testing correctly no-ops for them.
    m_lastTransform.valid = false;
    m_lastTransformType.clear();

    if (!m_currentSheet.hasSamples() && m_currentSheet.tpmTrend.isEmpty())
        return {};

    const QString plotType = m_plotTypeCombo->currentText();
    const int     W        = qMax(400, m_plotScrollArea->viewport()->width());
    const int     H        = qMax(300, m_plotScrollArea->viewport()->height());

    // ── Collect visible samples ───────────────────────────────────────────────
    QVector<int> visIdx;
    for (int i = 0; i < m_sampleVisible.size(); ++i)
        if (m_sampleVisible[i])
            visIdx.append(i);

    // ── Regime filter ─────────────────────────────────────────────────────────
    const QString selRegime = (m_regimeCombo && m_regimeCombo->isVisible())
                              ? m_regimeCombo->currentText() : QString();
    const bool filterRegime = !selRegime.isEmpty() && selRegime != QLatin1String("All regimes");
    auto rowMatches = [&](const DataRow& r) {
        return !filterRegime || RegimeUtils::regimeKey(r) == selRegime;
    };

    // ── TPM Trend ─────────────────────────────────────────────────────────────
    if (plotType == "TPM Trend") {
        if (!filterRegime &&
            !m_currentSheet.tpmTrend.isEmpty() &&
            !m_currentSheet.puffCounts.isEmpty() &&
            (visIdx.isEmpty() || m_currentSheet.samples.size() <= 1))
        {
            return PlotEngine::renderTPMTrend(
                m_currentSheet.puffCounts,
                m_currentSheet.tpmTrend,
                m_currentSheet.sheetName + " \u2013 TPM Trend");
        }

        // Check if any oil overlay series are enabled
        bool hasVisibleOil = false;
        for (bool v : m_oilVisible) if (v) { hasVisibleOil = true; break; }

        QVector<PlotSeries> primarySeries;
        for (int si : visIdx) {
            if (si >= m_currentSheet.samples.size()) continue;
            const SampleResult& sr = m_currentSheet.samples[si];

            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? QString("S%1").arg(si + 1)
                                                    : sr.sampleName;
            ps.color     = AppTheme::seriesColor(si);  // stable per sample; oil overlay matches
            ps.drawLine  = true;
            ps.drawDots  = (sr.rows.size() <= 30);
            ps.lineWidth = 2;
            ps.dotRadius = 3;

            for (const DataRow& row : sr.rows) {
                if (row.beforeWeight == 0.0 || row.afterWeight == 0.0) continue;
                if (!rowMatches(row)) continue;
                const int ptIdx = ps.x.size();
                ps.x.append(row.puffs);
                ps.y.append(row.tpm);
                // Note→plot v1: ring note-bearing points; emphasise the one the
                // user clicked (matched by cumulative-puff count).
                if (!row.notes.trimmed().isEmpty())
                    ps.ringed.append(ptIdx);
                if (m_selectedPuff >= 0 && int(row.puffs) == m_selectedPuff)
                    ps.emphasized = ptIdx;
            }

            if (!ps.x.isEmpty())
                primarySeries.append(ps);
        }

        if (primarySeries.isEmpty() && !hasVisibleOil) return {};

        PlotConfig cfg;
        cfg.title      = m_currentSheet.sheetName + " \u2013 TPM Trend";
        cfg.xLabel     = "Cumulative Puffs";
        cfg.yLabel     = "TPM (mg/puff)";
        cfg.width      = W;
        cfg.height     = H;
        cfg.autoScale  = true;
        cfg.showGrid   = true;
        cfg.showLegend = true;

        if (!hasVisibleOil) {
            cfg.showLegend = (primarySeries.size() > 1);
            QPixmap pm = PlotEngine::renderLinePlot(primarySeries, cfg, &m_lastTransform);
            m_lastTransformType = plotType;
            return pm;
        }

        // Build oil overlay series (right Y axis)
        QVector<PlotSeries> oilSeries;
        for (int si = 0; si < m_currentSheet.samples.size(); ++si) {
            if (si >= m_oilVisible.size() || !m_oilVisible[si]) continue;
            const SampleResult& sr = m_currentSheet.samples[si];

            PlotSeries ps;
            ps.label     = (sr.sampleName.isEmpty() ? QString("S%1").arg(si + 1)
                                                     : sr.sampleName) + " (Oil)";
            ps.color     = AppTheme::seriesColor(si);  // same color as this sample's TPM line
            ps.dashed    = true;
            ps.drawLine  = true;
            ps.drawDots  = (sr.rows.size() <= 30);
            ps.lineWidth = 2;
            ps.dotRadius = 3;

            for (const DataRow& row : sr.rows) {
                if (row.beforeWeight == 0.0 || row.afterWeight == 0.0) continue;
                if (!rowMatches(row)) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.oilConsumed);
            }

            if (!ps.x.isEmpty())
                oilSeries.append(ps);
        }

        cfg.y2Label = "Oil Consumed (mg)";
        QPixmap pm = PlotEngine::renderLinePlotDualAxis(primarySeries, oilSeries, cfg, &m_lastTransform);
        m_lastTransformType = plotType;
        return pm;
    }

    // ── TPM Bar Chart ─────────────────────────────────────────────────────────
    if (plotType == "TPM Bar Chart") {
        QVector<QString> names;
        QVector<double>  avgTPM;
        QVector<double>  stdDev;

        for (int si : visIdx) {
            if (si >= m_currentSheet.samples.size()) continue;
            const SampleResult& sr = m_currentSheet.samples[si];
            const QString nm = sr.sampleName.isEmpty() ? QString("S%1").arg(si + 1)
                                                       : sr.sampleName;
            if (!filterRegime) {
                names.append(nm); avgTPM.append(sr.averageTPM); stdDev.append(sr.stdDevTPM);
            } else {
                QVector<double> t;
                for (const DataRow& r : sr.rows) {
                    if (r.beforeWeight == 0.0 || r.afterWeight == 0.0) continue;
                    if (!rowMatches(r)) continue;
                    t.append(r.tpm);
                }
                if (t.isEmpty()) continue;   // this sample has no rows for the selected regime
                names.append(nm);
                avgTPM.append(TpmCalculator::average(t));
                stdDev.append(TpmCalculator::stddev(t));
            }
        }

        if (names.isEmpty()) return {};

        return PlotEngine::renderTPMBarChart(
            names, avgTPM, stdDev,
            m_currentSheet.sheetName + " \u2013 Average TPM per Sample");
    }

    // ── Power Density ─────────────────────────────────────────────────────────
    if (plotType == "Power Density") {
        QVector<PlotSeries> series;
        for (int si : visIdx) {
            if (si >= m_currentSheet.samples.size()) continue;
            const SampleResult& sr = m_currentSheet.samples[si];

            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? QString("S%1").arg(si + 1)
                                                    : sr.sampleName;
            ps.color     = AppTheme::seriesColor(si);  // stable per sample
            ps.drawLine  = true;
            ps.drawDots  = (sr.rows.size() <= 30);
            ps.lineWidth = 2;
            ps.dotRadius = 3;

            for (const DataRow& row : sr.rows) {
                if (row.beforeWeight == 0.0 || row.afterWeight == 0.0) continue;
                if (!rowMatches(row)) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.tpmPowerDensity);
            }

            if (!ps.x.isEmpty())
                series.append(ps);
        }

        if (series.isEmpty()) return {};

        PlotConfig cfg;
        cfg.title      = m_currentSheet.sheetName + " \u2013 TPM Power Density";
        cfg.xLabel     = "Cumulative Puffs";
        cfg.yLabel     = "TPM / Power (mg/W)";
        cfg.width      = W;
        cfg.height     = H;
        cfg.autoScale  = true;
        cfg.showGrid   = true;
        cfg.showLegend = (series.size() > 1);

        return PlotEngine::renderLinePlot(series, cfg);
    }

    // ── Draw Pressure ──────────────────────────────────────────────────────────
    if (plotType == "Draw Pressure") {
        QVector<PlotSeries> series;
        for (int si : visIdx) {
            if (si >= m_currentSheet.samples.size()) continue;
            const SampleResult& sr = m_currentSheet.samples[si];

            PlotSeries ps;
            ps.label     = sr.sampleName.isEmpty() ? QString("S%1").arg(si + 1)
                                                    : sr.sampleName;
            ps.color     = AppTheme::seriesColor(si);  // stable per sample
            ps.drawLine  = true;
            ps.drawDots  = (sr.rows.size() <= 30);
            ps.lineWidth = 2;
            ps.dotRadius = 3;

            for (const DataRow& row : sr.rows) {
                if (row.drawPressure == 0.0) continue;
                if (!rowMatches(row)) continue;
                ps.x.append(row.puffs);
                ps.y.append(row.drawPressure);
            }

            if (!ps.x.isEmpty())
                series.append(ps);
        }

        if (series.isEmpty()) return {};

        PlotConfig cfg;
        cfg.title      = m_currentSheet.sheetName + " \u2013 Draw Pressure";
        cfg.xLabel     = "Cumulative Puffs";
        cfg.yLabel     = "Draw Pressure (kPa)";
        cfg.width      = W;
        cfg.height     = H;
        cfg.autoScale = false;
        // Standing axis rules: x from 0 to the last puff, y from 0 to max+1.
        PlotEngine::applyDataXRange(cfg, series);
        PlotEngine::applyAnchoredYRange(cfg, series);
        cfg.showGrid   = true;
        cfg.showLegend = (series.size() > 1);

        return PlotEngine::renderLinePlot(series, cfg);
    }

    return {};
}

} // namespace DVE
