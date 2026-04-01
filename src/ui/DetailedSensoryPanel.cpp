#include "DetailedSensoryPanel.h"

#include <QDate>
#include <QFileDialog>
#include <QMessageBox>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QHeaderView>
#include <QScrollBar>
#include <QFileInfo>
#include <QDir>

#include "utils/AppTheme.h"
#include "reporting/PptxWriter.h"
#include "utils/ImageUtils.h"
#include "xlsxdocument.h"

#include <QWheelEvent>
namespace {
class NoWheelDoubleSpinBox : public QDoubleSpinBox {
public:
    using QDoubleSpinBox::QDoubleSpinBox;
    void wheelEvent(QWheelEvent* e) override { e->ignore(); }
};
class NoWheelComboBox : public QComboBox {
public:
    using QComboBox::QComboBox;
    void wheelEvent(QWheelEvent* e) override { e->ignore(); }
};
}

namespace DVE {

static double calcPower(double voltage, double resistance, const QString& tech)
{
    double rOffset = 0.0;
    if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
        rOffset = 0.78;
    else if (tech == "T51")
        rOffset = 0.25;
    double denom = resistance + rOffset;
    return (voltage > 0 && denom > 0) ? (voltage * voltage) / denom : 0.0;
}

DetailedSensoryPanel::DetailedSensoryPanel(DatabaseManager* db, QWidget* parent)
    : QWidget(parent), m_db(db)
{
    m_refreshTimer = new QTimer(this);
    m_refreshTimer->setSingleShot(true);
    m_refreshTimer->setInterval(150);
    connect(m_refreshTimer, &QTimer::timeout, this, &DetailedSensoryPanel::onRefreshChart);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(6, 2, 6, 2);
    mainLayout->setSpacing(4);

    // Header row
    auto* headerWidget = new QWidget(this);
    buildHeaderRow(headerWidget);
    mainLayout->addWidget(headerWidget);

    // Sample navigation bar
    buildSampleNavBar();
    mainLayout->addWidget(m_prevBtn->parentWidget());

    // Main splitter: top=questions, bottom=dual charts
    auto* splitter = new QSplitter(Qt::Vertical, this);

    // Top: question form in scroll area
    m_questionScroll = new QScrollArea(this);
    m_questionScroll->setWidgetResizable(true);
    m_questionScroll->setFrameShape(QFrame::NoFrame);
    buildQuestionForm();
    m_questionScroll->setWidget(m_questionForm);

    // Stacked widget for normal view vs averaged table
    m_topStack = new QStackedWidget(this);
    m_topStack->addWidget(m_questionScroll);
    m_avgOverlayTable = new QTableWidget(this);
    m_avgOverlayTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    m_avgOverlayTable->setAlternatingRowColors(true);
    m_topStack->addWidget(m_avgOverlayTable);
    splitter->addWidget(m_topStack);

    // Bottom: dual radar charts side by side
    auto* chartContainer = new QWidget(this);
    auto* chartLayout = new QHBoxLayout(chartContainer);
    chartLayout->setContentsMargins(0, 0, 0, 0);
    chartLayout->setSpacing(4);

    m_vaporQualityChart = new RadarChartWidget(this);
    m_vaporQualityChart->setCustomAxes(kDetailedVaporQualityMetrics, kDetailedAxisLabels);

    m_consistencyChart = new RadarChartWidget(this);
    m_consistencyChart->setCustomAxes(kDetailedConsistencyMetrics, kDetailedAxisLabels);

    chartLayout->addWidget(m_vaporQualityChart, 1);
    chartLayout->addWidget(m_consistencyChart, 1);
    splitter->addWidget(chartContainer);

    splitter->setStretchFactor(0, 2);
    splitter->setStretchFactor(1, 3);

    mainLayout->addWidget(splitter, 1);
}

void DetailedSensoryPanel::buildHeaderRow(QWidget* container)
{
    auto* hl = new QHBoxLayout(container);
    hl->setContentsMargins(0, 0, 0, 0);

    auto addField = [&](const QString& label) -> QLineEdit* {
        hl->addWidget(new QLabel(label + ":", container));
        auto* edit = new QLineEdit(container);
        edit->setMinimumWidth(100);
        hl->addWidget(edit);
        return edit;
    };

    m_testTitleEdit = addField("Test Title");
    m_assessorEdit  = addField("Assessor");
    m_testerEdit    = addField("Tester");
    m_mediaEdit     = addField("Media");

    hl->addWidget(new QLabel("Date:", container));
    m_dateLabel = new QLabel(QDate::currentDate().toString("yyyy-MM-dd"), container);
    hl->addWidget(m_dateLabel);
    hl->addStretch();

    for (auto* edit : {m_testTitleEdit, m_assessorEdit, m_testerEdit, m_mediaEdit})
        connect(edit, &QLineEdit::textChanged, this, &DetailedSensoryPanel::scheduleChartRefresh);
}

void DetailedSensoryPanel::buildSampleNavBar()
{
    auto* navBar = new QWidget(this);
    auto* hl = new QHBoxLayout(navBar);
    hl->setContentsMargins(0, 0, 0, 0);

    m_prevBtn = new QPushButton(QStringLiteral("\u25C0"), navBar);
    m_nextBtn = new QPushButton(QStringLiteral("\u25B6"), navBar);
    m_prevBtn->setToolTip("Previous sample (Ctrl+Left)");
    m_nextBtn->setToolTip("Next sample (Ctrl+Right)");
    m_prevBtn->setFixedSize(28, 24);
    m_nextBtn->setFixedSize(28, 24);

    m_sampleCountLabel = new QLabel(QStringLiteral("\u2014"), navBar);
    m_sampleCountLabel->setAlignment(Qt::AlignCenter);
    m_sampleCountLabel->setFont(AppTheme::fontSmall());
    m_sampleCountLabel->setMinimumWidth(70);

    m_addSampleBtn = new QPushButton("+ Add Sample", navBar);
    m_removeSampleBtn = new QPushButton("Remove", navBar);

    hl->addWidget(m_prevBtn);
    hl->addWidget(m_sampleCountLabel);
    hl->addWidget(m_nextBtn);
    hl->addStretch();
    hl->addWidget(m_addSampleBtn);
    hl->addWidget(m_removeSampleBtn);

    connect(m_prevBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onPrevSample);
    connect(m_nextBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onNextSample);
    connect(m_addSampleBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onAddSample);
    connect(m_removeSampleBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onRemoveSample);
}

void DetailedSensoryPanel::buildQuestionForm()
{
    m_questionForm = new QWidget(this);
    auto* grid = new QGridLayout(m_questionForm);
    grid->setContentsMargins(4, 4, 4, 4);
    grid->setSpacing(4);

    int row = 0;

    // Sample name
    grid->addWidget(new QLabel("Sample Name:", m_questionForm), row, 0);
    m_sampleNameEdit = new QLineEdit(m_questionForm);
    grid->addWidget(m_sampleNameEdit, row, 1, 1, 3);
    connect(m_sampleNameEdit, &QLineEdit::textChanged, this, &DetailedSensoryPanel::scheduleChartRefresh);
    ++row;

    // Vapor Quality section
    auto* vqLabel = new QLabel("<b>Vapor Quality</b>", m_questionForm);
    grid->addWidget(vqLabel, row, 0, 1, 4);
    ++row;

    auto addSpinQuestion = [&](const QString& metric, double min, double max, double step, double defaultVal) {
        QString label = kDetailedAxisLabels.value(metric, metric);
        label.replace('\n', ' ');
        grid->addWidget(new QLabel(label + ":", m_questionForm), row, 0, 1, 2);
        auto* spin = new NoWheelDoubleSpinBox(m_questionForm);
        spin->setRange(min, max);
        spin->setSingleStep(step);
        spin->setDecimals(1);
        spin->setValue(defaultVal);
        spin->setFixedWidth(70);
        grid->addWidget(spin, row, 2);
        m_spinBoxes[metric] = spin;
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, &DetailedSensoryPanel::scheduleChartRefresh);
        ++row;
    };

    addSpinQuestion("Burn Taste", 1.0, 9.0, 0.1, 1.0);
    addSpinQuestion("Flavor Intensity", 1.0, 9.0, 0.1, 5.0);
    addSpinQuestion("Throat Irritation", 1.0, 9.0, 0.1, 1.0);
    addSpinQuestion("Nasal Irritation", 1.0, 9.0, 0.1, 1.0);
    addSpinQuestion("Vapor Quality Overall", 1.0, 9.0, 0.1, 5.0);

    auto addComboQuestion = [&](const QString& metric, const QVector<ChoiceOption>& options) {
        QString label = kDetailedAxisLabels.value(metric, metric);
        label.replace('\n', ' ');
        grid->addWidget(new QLabel(label + ":", m_questionForm), row, 0, 1, 2);
        auto* combo = new NoWheelComboBox(m_questionForm);
        for (const auto& opt : options)
            combo->addItem(opt.text, opt.value);
        grid->addWidget(combo, row, 2, 1, 2);
        m_comboBoxes[metric] = combo;
        connect(combo, QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, &DetailedSensoryPanel::scheduleChartRefresh);
        ++row;
    };

    addComboQuestion("Cough", kCoughOptions);

    // Consistency section
    auto* consLabel = new QLabel("<b>Consistency</b>", m_questionForm);
    grid->addWidget(consLabel, row, 0, 1, 4);
    ++row;

    addComboQuestion("Volume Consistency", kVolumeConsistencyOptions);
    addComboQuestion("Performance Consistency", kPerformanceConsistencyOptions);
    addComboQuestion("Vapor Temperature", kVaporTemperatureOptions);
    addComboQuestion("Vapor vs Oil", kVaporVsOilOptions);
    addComboQuestion("Vapor Volume", kVaporVolumeOptions);

    // Comments
    grid->addWidget(new QLabel("Comments:", m_questionForm), row, 0);
    m_commentsEdit = new QTextEdit(m_questionForm);
    m_commentsEdit->setMaximumHeight(60);
    grid->addWidget(m_commentsEdit, row, 1, 1, 3);
    ++row;

    grid->setRowStretch(row, 1);
}

// ── Sample navigation ───────────────────────────────────────────────────────

void DetailedSensoryPanel::onPrevSample()
{
    if (m_currentTesterIdx < 0) return;
    saveCurrentSampleToSession();
    if (m_currentSampleIdx > 0) {
        --m_currentSampleIdx;
        displayCurrentSample();
    }
}

void DetailedSensoryPanel::onNextSample()
{
    if (m_currentTesterIdx < 0) return;
    saveCurrentSampleToSession();
    auto* sess = currentSession();
    if (!sess) return;
    if (m_currentSampleIdx < sess->samples.size() - 1) {
        ++m_currentSampleIdx;
        displayCurrentSample();
    }
}

void DetailedSensoryPanel::onAddSample()
{
    if (m_currentTesterIdx < 0) {
        newSession();
        return;
    }
    saveCurrentSampleToSession();
    auto* sess = currentSession();
    if (!sess) return;
    sess->samples.append(DetailedSensorySample{});
    m_currentSampleIdx = sess->samples.size() - 1;
    displayCurrentSample();
    scheduleChartRefresh();
    emit sessionsChanged();
}

void DetailedSensoryPanel::onRemoveSample()
{
    auto* sess = currentSession();
    if (!sess || sess->samples.isEmpty()) return;
    if (sess->samples.size() == 1) return;

    sess->samples.remove(m_currentSampleIdx);
    if (m_currentSampleIdx >= sess->samples.size())
        m_currentSampleIdx = sess->samples.size() - 1;
    displayCurrentSample();
    scheduleChartRefresh();
    emit sessionsChanged();
}

void DetailedSensoryPanel::updateSampleNav()
{
    auto* sess = currentSession();
    if (!sess || sess->samples.isEmpty()) {
        m_sampleCountLabel->setText(QStringLiteral("\u2014"));
        m_prevBtn->setEnabled(false);
        m_nextBtn->setEnabled(false);
        m_removeSampleBtn->setEnabled(false);
        return;
    }
    m_sampleCountLabel->setText(QString("Sample %1 of %2")
                                    .arg(m_currentSampleIdx + 1)
                                    .arg(sess->samples.size()));
    m_prevBtn->setEnabled(m_currentSampleIdx > 0);
    m_nextBtn->setEnabled(m_currentSampleIdx < sess->samples.size() - 1);
    m_removeSampleBtn->setEnabled(sess->samples.size() > 1);
}

void DetailedSensoryPanel::saveCurrentSampleToSession()
{
    auto* sess = currentSession();
    if (!sess || m_currentSampleIdx < 0 || m_currentSampleIdx >= sess->samples.size()) return;

    auto& sample = sess->samples[m_currentSampleIdx];
    sample.name     = m_sampleNameEdit->text();
    sample.comments = m_commentsEdit->toPlainText();

    for (auto it = m_spinBoxes.begin(); it != m_spinBoxes.end(); ++it)
        sample.scores[it.key()] = it.value()->value();

    for (auto it = m_comboBoxes.begin(); it != m_comboBoxes.end(); ++it)
        sample.scores[it.key()] = it.value()->currentData().toDouble();
}

void DetailedSensoryPanel::displayCurrentSample()
{
    auto* sess = currentSession();
    if (!sess || m_currentSampleIdx < 0 || m_currentSampleIdx >= sess->samples.size()) {
        updateSampleNav();
        return;
    }

    const auto& sample = sess->samples[m_currentSampleIdx];

    for (auto* spin : m_spinBoxes) spin->blockSignals(true);
    for (auto* combo : m_comboBoxes) combo->blockSignals(true);
    m_sampleNameEdit->blockSignals(true);

    m_sampleNameEdit->setText(sample.name);
    m_commentsEdit->setPlainText(sample.comments);

    for (auto it = m_spinBoxes.begin(); it != m_spinBoxes.end(); ++it) {
        double val = sample.scores.value(it.key(), it.value()->minimum());
        it.value()->setValue(val);
    }

    for (auto it = m_comboBoxes.begin(); it != m_comboBoxes.end(); ++it) {
        int rawVal = static_cast<int>(sample.scores.value(it.key(), 1.0));
        int idx = it.value()->findData(rawVal);
        if (idx >= 0) it.value()->setCurrentIndex(idx);
    }

    for (auto* spin : m_spinBoxes) spin->blockSignals(false);
    for (auto* combo : m_comboBoxes) combo->blockSignals(false);
    m_sampleNameEdit->blockSignals(false);

    updateSampleNav();
}

// ── Chart refresh ───────────────────────────────────────────────────────────

void DetailedSensoryPanel::scheduleChartRefresh()
{
    m_refreshTimer->start();
}

void DetailedSensoryPanel::onRefreshChart()
{
    saveCurrentSampleToSession();

    QVector<RadarChartWidget::SampleData> vqData, consData;

    for (const auto& sess : m_sessions) {
        for (const auto& sample : sess.samples) {
            RadarChartWidget::SampleData vqSample, consSample;
            vqSample.name = sample.name;
            consSample.name = sample.name;

            for (const QString& metric : kDetailedVaporQualityMetrics) {
                double raw = sample.scores.value(metric, 1.0);
                vqSample.scores[metric] = normalizeToRadar(metric, raw);
            }
            for (const QString& metric : kDetailedConsistencyMetrics) {
                double raw = sample.scores.value(metric, 1.0);
                consSample.scores[metric] = normalizeToRadar(metric, raw);
            }

            vqData.append(vqSample);
            consData.append(consSample);
        }
    }

    m_vaporQualityChart->setCustomData(vqData);
    m_consistencyChart->setCustomData(consData);
}

// ── Session management ──────────────────────────────────────────────────────

DetailedSensorySession* DetailedSensoryPanel::currentSession()
{
    if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size())
        return nullptr;
    return &m_sessions[m_currentTesterIdx];
}

void DetailedSensoryPanel::newSession()
{
    saveCurrentTester();

    DetailedSensorySession sess;
    sess.date = QDate::currentDate().toString("yyyy-MM-dd");
    sess.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    sess.samples.append(DetailedSensorySample{});

    m_sessions.append(sess);
    m_currentTesterIdx = m_sessions.size() - 1;
    m_currentSampleIdx = 0;
    applySession(sess);
    emit sessionsChanged();
}

void DetailedSensoryPanel::closeSessions(const QVector<int>& indices)
{
    QVector<int> sorted = indices;
    std::sort(sorted.begin(), sorted.end(), std::greater<int>());
    for (int idx : sorted) {
        if (idx >= 0 && idx < m_sessions.size())
            m_sessions.remove(idx);
    }
    if (m_sessions.isEmpty()) {
        m_currentTesterIdx = -1;
        m_currentSampleIdx = 0;
        m_testTitleEdit->clear();
        m_assessorEdit->clear();
        m_testerEdit->clear();
        m_mediaEdit->clear();
        m_sampleNameEdit->clear();
        m_commentsEdit->clear();
        for (auto* spin : m_spinBoxes) spin->setValue(spin->minimum());
        for (auto* combo : m_comboBoxes) combo->setCurrentIndex(0);
    } else {
        m_currentTesterIdx = qMin(m_currentTesterIdx, m_sessions.size() - 1);
        m_currentSampleIdx = 0;
        applySession(m_sessions[m_currentTesterIdx]);
    }
    onRefreshChart();
    emit sessionsChanged();
}

void DetailedSensoryPanel::loadSessions(const QVector<DetailedSensorySession>& sessions)
{
    m_sessions = sessions;
    if (!m_sessions.isEmpty()) {
        m_currentTesterIdx = 0;
        m_currentSampleIdx = 0;
        applySession(m_sessions[0]);
    }
    onRefreshChart();
    emit sessionsChanged();
}

void DetailedSensoryPanel::selectSession(int index)
{
    if (index < 0 || index >= m_sessions.size()) return;
    saveCurrentTester();
    m_currentTesterIdx = index;
    m_currentSampleIdx = 0;
    applySession(m_sessions[index]);
    onRefreshChart();
}

void DetailedSensoryPanel::renameSession(int index, const QString& newLabel)
{
    if (index < 0 || index >= m_sessions.size()) return;
    int sep = newLabel.indexOf(" - ");
    if (sep >= 0) {
        m_sessions[index].testTitle  = newLabel.left(sep).trimmed();
        m_sessions[index].testerName = newLabel.mid(sep + 3).trimmed();
    } else {
        m_sessions[index].sessionName = newLabel;
    }
    if (index == m_currentTesterIdx)
        applySession(m_sessions[index]);
    emit sessionsChanged();
}

void DetailedSensoryPanel::showAveragedChart(const QVector<int>& sessionIndices)
{
    QMap<QString, QMap<QString, QVector<double>>> deviceScores;
    for (int idx : sessionIndices) {
        if (idx < 0 || idx >= m_sessions.size()) continue;
        for (const auto& sample : m_sessions[idx].samples) {
            for (const QString& metric : kDetailedAllMetrics) {
                deviceScores[sample.name][metric].append(sample.scores.value(metric, 0.0));
            }
        }
    }

    QVector<RadarChartWidget::SampleData> vqData, consData;
    for (auto it = deviceScores.begin(); it != deviceScores.end(); ++it) {
        RadarChartWidget::SampleData vqSample, consSample;
        vqSample.name = consSample.name = it.key();
        for (const QString& metric : kDetailedVaporQualityMetrics) {
            const auto& vals = it.value().value(metric);
            double avg = 0;
            for (double v : vals) avg += v;
            if (!vals.isEmpty()) avg /= vals.size();
            vqSample.scores[metric] = normalizeToRadar(metric, avg);
        }
        for (const QString& metric : kDetailedConsistencyMetrics) {
            const auto& vals = it.value().value(metric);
            double avg = 0;
            for (double v : vals) avg += v;
            if (!vals.isEmpty()) avg /= vals.size();
            consSample.scores[metric] = normalizeToRadar(metric, avg);
        }
        vqData.append(vqSample);
        consData.append(consSample);
    }

    m_vaporQualityChart->setCustomData(vqData);
    m_consistencyChart->setCustomData(consData);
}

QVector<DetailedSensorySession> DetailedSensoryPanel::allSessions()
{
    saveCurrentTester();
    return m_sessions;
}

QString DetailedSensoryPanel::sessionLabel(const DetailedSensorySession& s) const
{
    QString title  = s.testTitle.isEmpty()  ? s.sessionName : s.testTitle;
    QString tester = s.testerName;
    if (!title.isEmpty() && !tester.isEmpty())
        return title + " - " + tester;
    if (!title.isEmpty()) return title;
    if (!tester.isEmpty()) return tester;
    return s.sessionName.isEmpty() ? "(untitled)" : s.sessionName;
}

DetailedSensorySession DetailedSensoryPanel::buildSession() const
{
    DetailedSensorySession sess;
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size())
        sess = m_sessions[m_currentTesterIdx];

    sess.testTitle    = m_testTitleEdit->text();
    sess.assessorName = m_assessorEdit->text();
    sess.testerName   = m_testerEdit->text();
    sess.media        = m_mediaEdit->text();
    sess.date         = m_dateLabel->text();
    if (sess.sessionName.isEmpty())
        sess.sessionName = sess.testTitle;
    if (sess.timestamp.isEmpty())
        sess.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    return sess;
}

void DetailedSensoryPanel::applySession(const DetailedSensorySession& session)
{
    m_testTitleEdit->blockSignals(true);
    m_assessorEdit->blockSignals(true);
    m_testerEdit->blockSignals(true);
    m_mediaEdit->blockSignals(true);

    m_testTitleEdit->setText(session.testTitle);
    m_assessorEdit->setText(session.assessorName);
    m_testerEdit->setText(session.testerName);
    m_mediaEdit->setText(session.media);
    m_dateLabel->setText(session.date.isEmpty()
                             ? QDate::currentDate().toString("yyyy-MM-dd")
                             : session.date);

    m_testTitleEdit->blockSignals(false);
    m_assessorEdit->blockSignals(false);
    m_testerEdit->blockSignals(false);
    m_mediaEdit->blockSignals(false);

    m_currentSampleIdx = 0;
    displayCurrentSample();
}

void DetailedSensoryPanel::saveCurrentTester()
{
    if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size()) return;
    saveCurrentSampleToSession();
    m_sessions[m_currentTesterIdx] = buildSession();
}

bool DetailedSensoryPanel::isDefaultState() const
{
    return m_sessions.isEmpty()
        || (m_sessions.size() == 1 && m_sessions[0].samples.isEmpty());
}

void DetailedSensoryPanel::showAveragedTable(const QStringList& deviceNames,
                                              const QVector<QMap<QString, double>>& deviceAvgs)
{
    m_avgOverlayTable->clear();
    m_avgOverlayTable->setColumnCount(kDetailedAllMetrics.size() + 1);

    QStringList headers;
    headers << "Device";
    for (const QString& m : kDetailedAllMetrics)
        headers << m;
    m_avgOverlayTable->setHorizontalHeaderLabels(headers);
    m_avgOverlayTable->setRowCount(deviceNames.size());

    for (int r = 0; r < deviceNames.size(); ++r) {
        m_avgOverlayTable->setItem(r, 0, new QTableWidgetItem(deviceNames[r]));
        for (int c = 0; c < kDetailedAllMetrics.size(); ++c) {
            double val = deviceAvgs[r].value(kDetailedAllMetrics[c], 0.0);
            m_avgOverlayTable->setItem(r, c + 1,
                new QTableWidgetItem(QString::number(val, 'f', 1)));
        }
    }

    m_avgOverlayTable->resizeColumnsToContents();
    m_topStack->setCurrentIndex(1);
}

void DetailedSensoryPanel::showNormalView()
{
    m_topStack->setCurrentIndex(0);
}

// ── File I/O ────────────────────────────────────────────────────────────────

void DetailedSensoryPanel::save()
{
    saveCurrentTester();
    if (m_currentTesterIdx < 0) return;

    const auto& sess = m_sessions[m_currentTesterIdx];

    if (m_savePath.isEmpty()) {
        QString defaultName = sess.testTitle.isEmpty() ? "detailed_sensory" : sess.testTitle;
        defaultName.replace(' ', '_');
        m_savePath = QFileDialog::getSaveFileName(
            this, "Save Detailed Sensory Session",
            lastBrowseDir() + "/" + defaultName + ".xlsx",
            "Excel Files (*.xlsx)");
        if (m_savePath.isEmpty()) return;
        setLastBrowseDir(m_savePath);
    }

    saveToExcel(m_savePath, sess);

    if (m_db && m_db->isOpen()) {
        m_db->saveDetailedSensorySession(sess);
    }
}

void DetailedSensoryPanel::saveToExcel(const QString& path, const DetailedSensorySession& sess)
{
    QXlsx::Document xlsx;

    int col = 1;
    xlsx.write(1, col++, "Sample Name");
    for (const QString& metric : kDetailedAllMetrics)
        xlsx.write(1, col++, metric);
    xlsx.write(1, col++, "V");
    xlsx.write(1, col++, "R");
    xlsx.write(1, col++, "P");
    xlsx.write(1, col++, "HT");
    xlsx.write(1, col++, "Media");
    xlsx.write(1, col++, "Viscosity");
    xlsx.write(1, col++, "Oil Smell Liking");
    xlsx.write(1, col++, "Clog");
    xlsx.write(1, col++, "Comments");

    int row = 2;
    for (const auto& sample : sess.samples) {
        col = 1;
        xlsx.write(row, col++, sample.name);
        for (const QString& metric : kDetailedAllMetrics)
            xlsx.write(row, col++, sample.scores.value(metric, 0.0));
        xlsx.write(row, col++, sample.voltage);
        xlsx.write(row, col++, sample.resistance);
        xlsx.write(row, col++, sample.power);
        xlsx.write(row, col++, sample.heatingTechnology);
        xlsx.write(row, col++, sess.media);
        xlsx.write(row, col++, sess.viscosity);
        xlsx.write(row, col++, sess.oilSmellLiking);
        xlsx.write(row, col++, sess.clog ? "Yes" : "No");
        xlsx.write(row, col++, sample.comments);
        ++row;
    }

    if (sess.samples.size() > 1) {
        col = 1;
        xlsx.write(row, col++, "Average");
        for (const QString& metric : kDetailedAllMetrics) {
            double sum = 0;
            for (const auto& s : sess.samples)
                sum += s.scores.value(metric, 0.0);
            xlsx.write(row, col++, sum / sess.samples.size());
        }
        ++row;
    }

    ++row;
    xlsx.write(row, 1, "Test Title"); xlsx.write(row, 2, sess.testTitle); ++row;
    xlsx.write(row, 1, "Tester");     xlsx.write(row, 2, sess.testerName); ++row;
    xlsx.write(row, 1, "Assessor");   xlsx.write(row, 2, sess.assessorName); ++row;
    xlsx.write(row, 1, "Media");      xlsx.write(row, 2, sess.media); ++row;
    xlsx.write(row, 1, "Date");       xlsx.write(row, 2, sess.date); ++row;
    xlsx.write(row, 1, "Facilitator"); xlsx.write(row, 2, sess.facilitatorName); ++row;
    xlsx.write(row, 1, "Oil Smell Liking"); xlsx.write(row, 2, sess.oilSmellLiking); ++row;
    xlsx.write(row, 1, "Viscosity");  xlsx.write(row, 2, sess.viscosity); ++row;

    xlsx.saveAs(path);
}

void DetailedSensoryPanel::loadFiles()
{
    QStringList files = QFileDialog::getOpenFileNames(
        this, "Load Detailed Sensory Data",
        lastBrowseDir(),
        "Excel Files (*.xlsx);;All Files (*)");
    if (files.isEmpty()) return;
    setLastBrowseDir(files.first());

    saveCurrentTester();

    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    int loaded = 0;
    for (const QString& path : files) {
        QXlsx::Document xlsx(path);
        if (!xlsx.load()) {
            QMessageBox::warning(this, "Load Error",
                                 "Could not open Excel file:\n" + path);
            continue;
        }

        // Expect saved format: row 1 = headers ("Sample Name", metrics..., V, R, P, HT, Comments)
        // rows 2..N = sample data, optional "Average" row, then metadata rows
        DetailedSensorySession sess;
        sess.sessionName = QFileInfo(path).baseName();
        sess.date = QDate::currentDate().toString("yyyy-MM-dd");
        sess.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

        // Read header row to find metric columns
        QMap<int, QString> colMetric;
        int lastCol = 1;
        for (int col = 2; col <= 30; ++col) {
            QString hdr = xlsx.read(1, col).toString().trimmed();
            if (hdr.isEmpty()) break;
            lastCol = col;
            // Map known metric names
            for (const QString& metric : kDetailedAllMetrics) {
                if (hdr.compare(metric, Qt::CaseInsensitive) == 0) {
                    colMetric[col] = metric;
                    break;
                }
            }
        }

        // Read data rows (row 2+)
        for (int row = 2; row <= 200; ++row) {
            QString name = xlsx.read(row, 1).toString().trimmed();
            if (name.isEmpty() || name.compare("Average", Qt::CaseInsensitive) == 0)
                break;

            DetailedSensorySample sample;
            sample.name = name;

            for (auto it = colMetric.constBegin(); it != colMetric.constEnd(); ++it) {
                double val = xlsx.read(row, it.key()).toDouble();
                double maxVal = kDetailedMetricMaxScore.value(it.value(), 9);
                sample.scores[it.value()] = qBound(1.0, val, static_cast<double>(maxVal));
            }

            // Read V, R, P, HT, Comments from columns after metrics
            for (int col = 2; col <= lastCol; ++col) {
                QString hdr = xlsx.read(1, col).toString().trimmed();
                if (hdr == "V") sample.voltage = xlsx.read(row, col).toDouble();
                else if (hdr == "R") sample.resistance = xlsx.read(row, col).toDouble();
                else if (hdr == "P") sample.power = xlsx.read(row, col).toDouble();
                else if (hdr == "HT") sample.heatingTechnology = xlsx.read(row, col).toString().trimmed();
                else if (hdr == "Comments") sample.comments = xlsx.read(row, col).toString().trimmed();
            }

            sess.samples.append(sample);
        }

        // Read metadata rows after data (Test Title, Tester, Assessor, Media, Date, etc.)
        for (int row = sess.samples.size() + 3; row <= sess.samples.size() + 15; ++row) {
            QString key = xlsx.read(row, 1).toString().trimmed();
            QString val = xlsx.read(row, 2).toString().trimmed();
            if (key.isEmpty()) continue;
            if (key == "Test Title") sess.testTitle = val;
            else if (key == "Tester") sess.testerName = val;
            else if (key == "Assessor") sess.assessorName = val;
            else if (key == "Media") sess.media = val;
            else if (key == "Date") sess.date = val;
            else if (key == "Facilitator") sess.facilitatorName = val;
            else if (key == "Oil Smell Liking") sess.oilSmellLiking = val.toInt();
            else if (key == "Viscosity") sess.viscosity = val;
        }

        if (sess.sessionName.isEmpty())
            sess.sessionName = sess.testTitle;

        if (!sess.samples.isEmpty()) {
            m_sessions.append(sess);
            ++loaded;
        }
    }

    if (loaded == 0) {
        QMessageBox::warning(this, "No Data",
                             "No sample data found in the selected file(s).");
        return;
    }

    // Save loaded sessions to database
    if (m_db && m_db->isOpen()) {
        for (const DetailedSensorySession& s : m_sessions) {
            if (!s.samples.isEmpty())
                m_db->saveDetailedSensorySession(s);
        }
    }

    m_currentTesterIdx = m_sessions.size() - 1;
    m_currentSampleIdx = 0;
    applySession(m_sessions[m_currentTesterIdx]);
    emit sessionsChanged();
}

void DetailedSensoryPanel::loadFromDatabase()
{
    if (!m_db || !m_db->isOpen()) return;
    auto sessions = m_db->loadDetailedSensorySessions();
    if (!sessions.isEmpty())
        loadSessions(sessions);
}

// ── Report generation (stub — Task 8 will implement fully) ──────────────────

void DetailedSensoryPanel::generateFullReport()
{
    saveCurrentTester();
    if (m_sessions.isEmpty()) return;

    QDialog dlg(this);
    dlg.setWindowTitle("Select Sessions for Report");
    auto* vl = new QVBoxLayout(&dlg);
    auto* listWidget = new QListWidget(&dlg);
    listWidget->setSelectionMode(QAbstractItemView::ExtendedSelection);
    for (const auto& s : m_sessions)
        listWidget->addItem(sessionLabel(s));
    listWidget->selectAll();
    vl->addWidget(listWidget);

    auto* btnBox = new QDialogButtonBox(QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
    connect(btnBox, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
    connect(btnBox, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);
    vl->addWidget(btnBox);

    if (dlg.exec() != QDialog::Accepted) return;

    QVector<DetailedSensorySession> selected;
    for (auto* item : listWidget->selectedItems()) {
        int idx = listWidget->row(item);
        if (idx >= 0 && idx < m_sessions.size())
            selected.append(m_sessions[idx]);
    }
    if (selected.isEmpty()) return;

    QString defaultName = selected[0].testTitle.isEmpty() ? "detailed_sensory_report" : selected[0].testTitle;
    defaultName.replace(' ', '_');
    QString path = QFileDialog::getSaveFileName(
        this, "Save Report",
        lastBrowseDir() + "/" + defaultName + ".pptx",
        "PowerPoint Files (*.pptx)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    QString errorOut;
    if (generateCombinedPptx(selected, path, errorOut)) {
        QMessageBox::information(this, "Report Saved", "Report saved to:\n" + path);
    } else {
        QMessageBox::warning(this, "Report Error", "Failed to generate report:\n" + errorOut);
    }
}

bool DetailedSensoryPanel::generateCombinedPptx(
    const QVector<DetailedSensorySession>& sessions,
    const QString& filePath,
    QString& errorOut)
{
    if (sessions.isEmpty()) {
        errorOut = "No sessions provided";
        return false;
    }

    PptxWriter pptx;

    QString coverTitle = sessions[0].testTitle.isEmpty()
                             ? "Detailed Sensory Evaluation"
                             : sessions[0].testTitle;
    pptx.addCoverSlide(coverTitle, sessions[0].date);

    // Gather unique sample names
    QStringList sampleNames;
    for (const auto& sess : sessions) {
        for (const auto& sample : sess.samples) {
            if (!sampleNames.contains(sample.name))
                sampleNames.append(sample.name);
        }
    }

    // Per sample: table slide + plot slide
    for (const QString& sampleName : sampleNames) {
        struct TesterRow {
            QString testerName;
            QMap<QString, double> scores;
        };
        QVector<TesterRow> rows;

        for (const auto& sess : sessions) {
            for (const auto& sample : sess.samples) {
                if (sample.name == sampleName) {
                    TesterRow row;
                    row.testerName = sess.testerName.isEmpty() ? sess.assessorName : sess.testerName;
                    row.scores = sample.scores;
                    rows.append(row);
                }
            }
        }

        if (rows.isEmpty()) continue;

        // TABLE SLIDE
        QStringList headers;
        headers << "Tester";
        for (const QString& m : kDetailedAllMetrics)
            headers << m;

        QVector<QStringList> tableRows;
        QMap<QString, double> sums;
        for (const auto& r : rows) {
            QStringList cells;
            cells << r.testerName;
            for (const QString& m : kDetailedAllMetrics) {
                double val = r.scores.value(m, 0.0);
                cells << QString::number(val, 'f', 1);
                sums[m] += val;
            }
            tableRows.append(cells);
        }

        // Average row
        QStringList avgRow;
        avgRow << "Average";
        for (const QString& m : kDetailedAllMetrics)
            avgRow << QString::number(sums[m] / rows.size(), 'f', 1);
        tableRows.append(avgRow);

        QVector<double> colWidths;
        colWidths << 0.12;
        double metricW = (1.0 - 0.12) / kDetailedAllMetrics.size();
        for (int i = 0; i < kDetailedAllMetrics.size(); ++i)
            colWidths << metricW;

        SlideTable table;
        table.headers = headers;
        table.rows = tableRows;
        table.colWidthFractions = colWidths;
        table.x = 0.3;
        table.y = 0.64;
        table.w = 12.7;
        pptx.addContentSlide(sampleName, table, {});

        // PLOT SLIDE — dual radar charts side by side
        QVector<RadarChartWidget::SampleData> vqData, consData;
        for (const auto& r : rows) {
            RadarChartWidget::SampleData vqSample, consSample;
            vqSample.name = consSample.name = r.testerName;
            for (const QString& m : kDetailedVaporQualityMetrics)
                vqSample.scores[m] = normalizeToRadar(m, r.scores.value(m, 1.0));
            for (const QString& m : kDetailedConsistencyMetrics)
                consSample.scores[m] = normalizeToRadar(m, r.scores.value(m, 1.0));
            vqData.append(vqSample);
            consData.append(consSample);
        }

        RadarChartWidget vqChart;
        vqChart.setCustomAxes(kDetailedVaporQualityMetrics, kDetailedAxisLabels);
        vqChart.setCustomData(vqData);
        vqChart.setReportMode(true);
        vqChart.setReportCropTop(70);
        vqChart.resize(580, 858);
        QPixmap vqPix = vqChart.grab();

        RadarChartWidget consChart;
        consChart.setCustomAxes(kDetailedConsistencyMetrics, kDetailedAxisLabels);
        consChart.setCustomData(consData);
        consChart.setReportMode(true);
        consChart.setReportCropTop(70);
        consChart.resize(580, 858);
        QPixmap consPix = consChart.grab();

        QString vqPath  = QDir::temp().filePath("dve_vq_report.png");
        QString consPath = QDir::temp().filePath("dve_cons_report.png");
        vqPix.save(vqPath);
        consPix.save(consPath);

        QByteArray vqBytes, consBytes;
        { QFile f(vqPath); if (f.open(QIODevice::ReadOnly)) vqBytes = f.readAll(); }
        { QFile f(consPath); if (f.open(QIODevice::ReadOnly)) consBytes = f.readAll(); }

        SlideImage vqImg;
        vqImg.pngData = vqBytes;
        vqImg.x = 0.3; vqImg.y = 0.8; vqImg.w = 6.0; vqImg.h = 5.5;

        SlideImage consImg;
        consImg.pngData = consBytes;
        consImg.x = 6.7; consImg.y = 0.8; consImg.w = 6.0; consImg.h = 5.5;

        SlideTable emptyTable;
        pptx.addContentSlide(sampleName + " - Charts", emptyTable, {vqImg, consImg});

        QFile::remove(vqPath);
        QFile::remove(consPath);
    }

    if (!pptx.save(filePath)) {
        errorOut = "Failed to write file: " + filePath;
        return false;
    }
    return true;
}

// ── Browse dir helpers ──────────────────────────────────────────────────────

QString DetailedSensoryPanel::lastBrowseDir() const
{
    if (!m_lastBrowseDir.isEmpty()) return m_lastBrowseDir;
    if (m_db) return m_db->getSetting("last_browse_dir");
    return QDir::homePath();
}

void DetailedSensoryPanel::setLastBrowseDir(const QString& filePath)
{
    m_lastBrowseDir = QFileInfo(filePath).absolutePath();
}

} // namespace DVE
