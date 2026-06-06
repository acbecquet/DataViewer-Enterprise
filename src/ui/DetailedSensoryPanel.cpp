#include "DetailedSensoryPanel.h"

#include <QDate>
#include <QDebug>
#include <QFileDialog>
#include <QMessageBox>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QHeaderView>
#include <QScrollBar>
#include <QFileInfo>
#include <QDir>
#include <QSet>
#include <QAction>
#include <QMenu>
#include <QStyle>
#include <functional>

#include "utils/AppTheme.h"
#include "utils/OutputPaths.h"
#include "utils/ResponsiveLayout.h"
#include "reporting/PptxWriter.h"
#include "utils/ImageUtils.h"
#include "database/LiveSync.h"
#include "xlsxdocument.h"

#include <QWheelEvent>
namespace {

// v2.0.4: identical helper to the one in SensoryPanel.cpp — attaches a
// trailing ▼ action to a QLineEdit that pops up a menu of preset
// values fetched fresh from the DB on every click. Kept TU-local so
// both panels can ship the same UX without an external dependency.
void attachPresetDropdown(QLineEdit* edit,
                          std::function<QStringList()> provider)
{
    if (!edit) return;
    QAction* act = edit->addAction(
        edit->style()->standardIcon(QStyle::SP_ArrowDown),
        QLineEdit::TrailingPosition);
    act->setToolTip(QObject::tr("Pick from saved values"));
    QObject::connect(act, &QAction::triggered, edit,
        [edit, provider = std::move(provider)]() {
            const QStringList values = provider();
            QMenu menu(edit);
            if (values.isEmpty()) {
                QAction* empty = menu.addAction(
                    QObject::tr("(no saved values — type and Save Test Headers)"));
                empty->setEnabled(false);
            } else {
                for (const QString& v : values) {
                    QObject::connect(menu.addAction(v), &QAction::triggered, edit,
                        [edit, v]() { edit->setText(v); });
                }
            }
            const QPoint pos = edit->mapToGlobal(QPoint(0, edit->height()));
            menu.exec(pos);
        });
}

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


DetailedSensoryPanel::DetailedSensoryPanel(DatabaseManager* db, QWidget* parent)
    : QWidget(parent), m_db(db)
{
    m_refreshTimer = new QTimer(this);
    m_refreshTimer->setSingleShot(true);
    m_refreshTimer->setInterval(150);
    connect(m_refreshTimer, &QTimer::timeout, this, &DetailedSensoryPanel::onRefreshChart);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setContentsMargins(6, 4, 6, 2);
    mainLayout->setSpacing(4);

    // Combined header + sample navigation row
    auto* headerWidget = new QWidget(this);
    buildHeaderRow(headerWidget);
    mainLayout->addWidget(headerWidget);

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

    // Responsive: hide right column of 14-question grid below 800 px,
    // and stack the dual radar charts vertically below 1000 px.
    connect(&DVE::ResponsiveLayout::instance(),
            &DVE::ResponsiveLayout::widthChanged,
            this, [this, chartContainer](int w) {
        // 14-question grid: right columns (4=num, 5=label, 6=input) hidden below 800 px.
        // The grid is nested inside outerVBox — walk the layout items to find it.
        QGridLayout* questionGrid = nullptr;
        if (auto* vbox = qobject_cast<QVBoxLayout*>(m_questionForm->layout())) {
            for (int i = 0; i < vbox->count(); ++i) {
                if (auto* item = vbox->itemAt(i)) {
                    if (auto* candidate = qobject_cast<QGridLayout*>(item->layout())) {
                        questionGrid = candidate;
                        break;
                    }
                }
            }
        }
        if (questionGrid) {
            const bool narrow = (w < DVE::ResponsiveLayout::kDetailedNarrowThreshold);
            // Hide/show the right half (columns 4, 5, 6)
            for (int r = 0; r < questionGrid->rowCount(); ++r) {
                for (int c : {4, 5, 6}) {
                    if (auto* item = questionGrid->itemAtPosition(r, c)) {
                        if (auto* widget = item->widget())
                            widget->setVisible(!narrow);
                    }
                }
            }
        }

        // Dual radar charts: stack vertically below 1000 px
        if (auto* chartLayout = qobject_cast<QHBoxLayout*>(chartContainer->layout())) {
            const bool stackCharts = (w < DVE::ResponsiveLayout::kDetailedStackChartsThreshold);
            // QHBoxLayout has no orientation toggle; use a QSplitter-style
            // approach: adjust minimum heights so they wrap sensibly.
            // Pragmatic implementation: set fixed heights when stacking to
            // give each chart adequate vertical space, restore to 0 otherwise.
            const int chartMinH = stackCharts ? 220 : 0;
            m_vaporQualityChart->setMinimumHeight(chartMinH);
            m_consistencyChart->setMinimumHeight(chartMinH);
            (void)chartLayout; // used indirectly via chart widgets above
        }
    });
}

void DetailedSensoryPanel::buildHeaderRow(QWidget* container)
{
    auto* hl = new QHBoxLayout(container);
    hl->setContentsMargins(0, 0, 0, 0);
    hl->setSpacing(8);

    auto addField = [&](const QString& label, int maxW = 90) -> QLineEdit* {
        hl->addWidget(new QLabel(label + ":", container));
        auto* edit = new QLineEdit(container);
        edit->setMaximumWidth(maxW);
        hl->addWidget(edit);
        return edit;
    };

    m_testTitleEdit = addField("Test Title", 110);
    m_assessorEdit  = addField("Assessor");
    m_testerEdit    = addField("Tester");
    m_mediaEdit     = addField("Media");

    // v2.0.4: trailing ▼ dropdowns on the two fields most commonly
    // re-entered verbatim across sessions. Lazy-fetch from the DB on
    // every click so other users' saves appear without a restart.
    // Shared preset pool with SensoryPanel — typing happens once,
    // every panel benefits.
    attachPresetDropdown(m_testTitleEdit, [this]() -> QStringList {
        return m_db ? m_db->loadSensoryHeaderPresets("test_name") : QStringList{};
    });
    attachPresetDropdown(m_mediaEdit, [this]() -> QStringList {
        return m_db ? m_db->loadSensoryHeaderPresets("media") : QStringList{};
    });

    // v2.0.4: "Save Test Headers" button next to Media. Records the
    // current Test Title, Media, and per-sample names from the
    // currently-loaded session into the shared preset pool.
    auto* saveHeadersBtn = new QPushButton(tr("Save Test Headers"), container);
    saveHeadersBtn->setToolTip(
        tr("Add the current Test Title, Media, and sample names to the\n"
           "shared dropdown pool so coworkers can pick them instead of\n"
           "retyping. Safe to click repeatedly — duplicates are ignored."));
    connect(saveHeadersBtn, &QPushButton::clicked, this, [this]() {
        if (!m_db) return;
        // Pull sample names from the currently-loaded session. The
        // detailed form only shows one sample at a time, but the
        // session struct carries every sample, so we capture them all.
        QStringList sampleNames;
        const int idx = m_currentTesterIdx;
        if (idx >= 0 && idx < m_sessions.size()) {
            for (const DetailedSensorySample& s : m_sessions[idx].samples) {
                const QString n = s.name.trimmed();
                if (!n.isEmpty()) sampleNames << n;
            }
        }
        const QString testTitle = m_testTitleEdit->text().trimmed();
        const QString media     = m_mediaEdit->text().trimmed();
        if (m_db->saveSensoryHeaderPresets(testTitle, media, sampleNames)) {
            // v2.0.5: confirmation toast — see SensoryPanel for the
            // matching version. Lists exactly what was attempted so
            // the user can verify sample names hit the pool.
            QStringList parts;
            if (!testTitle.isEmpty())
                parts << tr("Test Title: %1").arg(testTitle);
            if (!media.isEmpty())
                parts << tr("Media: %1").arg(media);
            if (!sampleNames.isEmpty())
                parts << tr("Sample names (%1): %2")
                             .arg(sampleNames.size())
                             .arg(sampleNames.join(QStringLiteral(", ")));
            const QString body = parts.isEmpty()
                ? tr("Nothing to save — all fields are empty.")
                : (tr("Saved to the shared dropdown pool:\n\n") +
                   parts.join(QStringLiteral("\n")) +
                   tr("\n\nDuplicates were skipped automatically. "
                      "Coworkers will see the new values in their "
                      "dropdowns next time they open a session."));
            QMessageBox::information(this, tr("Save Test Headers"), body);
        } else {
            QMessageBox::warning(this, tr("Save Test Headers"),
                tr("Could not save test headers to the database.\n%1")
                    .arg(m_db->lastError()));
        }
    });
    hl->addWidget(saveHeadersBtn);

    hl->addWidget(new QLabel("Date:", container));
    m_dateLabel = new QLabel(QDate::currentDate().toString("yyyy-MM-dd"), container);
    hl->addWidget(m_dateLabel);

    hl->addStretch();

    // Sample navigation (merged into header row)
    m_prevBtn = new QPushButton(QStringLiteral("\u25C0"), container);
    m_nextBtn = new QPushButton(QStringLiteral("\u25B6"), container);
    m_prevBtn->setToolTip("Previous sample (Ctrl+Left)");
    m_nextBtn->setToolTip("Next sample (Ctrl+Right)");
    m_prevBtn->setFixedSize(28, 24);
    m_nextBtn->setFixedSize(28, 24);

    m_sampleCountLabel = new QLabel(QStringLiteral("\u2014"), container);
    m_sampleCountLabel->setAlignment(Qt::AlignCenter);
    m_sampleCountLabel->setFont(AppTheme::fontSmall());
    m_sampleCountLabel->setMinimumWidth(70);

    m_addSampleBtn = new QPushButton("+ Add Sample", container);
    m_removeSampleBtn = new QPushButton("Remove", container);

    m_addSampleBtn->setProperty("primary", true);
    m_addSampleBtn->setIcon(AppTheme::icon("file-plus"));
    m_removeSampleBtn->setProperty("destructive", true);
    m_removeSampleBtn->setIcon(AppTheme::icon("x"));
    for (auto* b : { m_addSampleBtn, m_removeSampleBtn }) {
        b->style()->unpolish(b);
        b->style()->polish(b);
    }

    hl->addWidget(m_prevBtn);
    hl->addWidget(m_sampleCountLabel);
    hl->addWidget(m_nextBtn);
    hl->addWidget(m_addSampleBtn);
    hl->addWidget(m_removeSampleBtn);

    for (auto* edit : {m_testTitleEdit, m_assessorEdit, m_testerEdit, m_mediaEdit}) {
        connect(edit, &QLineEdit::textChanged, this, &DetailedSensoryPanel::scheduleChartRefresh);
        // Plan C (C6 fix): arm the recovery snapshot on every keystroke in the
        // header fields too (not just on editingFinished via commitSessionField),
        // so a crash mid-typing — before the field loses focus — still captures
        // the in-progress header text. buildSession() reads these widgets live at
        // snapshot time, so the data is present even though textChanged does not
        // itself write m_sessions. Matches SensoryPanel's per-keystroke header
        // wiring. Double-arming with commitSessionField on focus-out is harmless
        // (noteDirty() just restarts the debounce).
        connect(edit, &QLineEdit::textChanged, this, &DetailedSensoryPanel::dataEdited);
    }

    // v2.0.1: session-level LiveSync emissions. Use editingFinished so we
    // don't broadcast every keystroke. Field names match the canonical
    // JSON serializer in
    // DetailedSensoryData.cpp::detailedSensorySessionToJson.
    connect(m_testTitleEdit, &QLineEdit::editingFinished, this, [this]() {
        commitSessionField(QStringLiteral("test_title"), m_testTitleEdit->text());
    });
    connect(m_assessorEdit, &QLineEdit::editingFinished, this, [this]() {
        commitSessionField(QStringLiteral("assessor_name"), m_assessorEdit->text());
    });
    connect(m_testerEdit, &QLineEdit::editingFinished, this, [this]() {
        commitSessionField(QStringLiteral("tester_name"), m_testerEdit->text());
    });
    connect(m_mediaEdit, &QLineEdit::editingFinished, this, [this]() {
        commitSessionField(QStringLiteral("media"), m_mediaEdit->text());
    });

    connect(m_prevBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onPrevSample);
    connect(m_nextBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onNextSample);
    connect(m_addSampleBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onAddSample);
    connect(m_removeSampleBtn, &QPushButton::clicked, this, &DetailedSensoryPanel::onRemoveSample);
}

void DetailedSensoryPanel::buildSampleNavBar()
{
    // Kept for API compatibility — navigation is now built inside buildHeaderRow
}

void DetailedSensoryPanel::buildQuestionForm()
{
    m_questionForm = new QWidget(this);
    auto* outerVBox = new QVBoxLayout(m_questionForm);
    outerVBox->setContentsMargins(8, 4, 8, 4);
    outerVBox->setSpacing(4);

    // Sample name — full width row at top
    auto* sampleRow = new QHBoxLayout;
    sampleRow->addWidget(new QLabel("Sample Name:", m_questionForm));
    m_sampleNameEdit = new QLineEdit(m_questionForm);
    m_sampleNameEdit->setMinimumHeight(32);
    m_sampleNameEdit->setFont(AppTheme::fontSection());
    sampleRow->addWidget(m_sampleNameEdit, 1);
    outerVBox->addLayout(sampleRow);
    connect(m_sampleNameEdit, &QLineEdit::textChanged, this, &DetailedSensoryPanel::scheduleChartRefresh);
    // v2.0.4: trailing ▼ dropdown on the sample name field. Same
    // shared preset pool as SensoryPanel — sample names typed once
    // appear in every panel's dropdown.
    attachPresetDropdown(m_sampleNameEdit, [this]() -> QStringList {
        return m_db ? m_db->loadSensoryHeaderPresets("sample_name")
                    : QStringList{};
    });
    // v2.0.1: sample-level name commit on focus-out.
    connect(m_sampleNameEdit, &QLineEdit::editingFinished, this, [this]() {
        commitSampleField(QStringLiteral("name"), m_sampleNameEdit->text());
    });

    // Single grid: cols 0-2 = left (num, label, input), col 3 = spacer,
    //              cols 4-6 = right (num, label, input)
    // Equal column stretch on input cols (2 and 6) to match chart 1:1 split
    auto* grid = new QGridLayout;
    grid->setContentsMargins(8, 4, 8, 4);
    grid->setHorizontalSpacing(12);
    grid->setVerticalSpacing(8);
    grid->setColumnMinimumWidth(0, 20);  // number col left
    grid->setColumnMinimumWidth(1, 200); // label col left — min 200 px so labels align
    grid->setColumnMinimumWidth(3, 24);  // spacer between columns
    grid->setColumnMinimumWidth(4, 20);  // number col right
    grid->setColumnMinimumWidth(5, 200); // label col right — min 200 px so labels align
    grid->setColumnStretch(2, 1);        // left input col stretches
    grid->setColumnStretch(6, 1);        // right input col stretches

    int row = 0;
    int qNum = 1;

    const int kComboMaxW = 280;

    // Helper: add a numbered spin question at a specific grid position
    auto addSpinAt = [&](int r, int numCol, int labelCol, int inputCol,
                          int num, const QString& metric,
                          double min, double max, double step, double defaultVal) {
        QString label = kDetailedAxisLabels.value(metric, metric);
        label.replace('\n', ' ');
        auto* numLabel = new QLabel(QString::number(num) + ".", m_questionForm);
        numLabel->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
        grid->addWidget(numLabel, r, numCol);
        grid->addWidget(new QLabel(label + ":", m_questionForm), r, labelCol);
        auto* spin = new NoWheelDoubleSpinBox(m_questionForm);
        spin->setRange(min, max);
        spin->setSingleStep(step);
        spin->setDecimals(1);
        spin->setValue(defaultVal);
        spin->setFixedWidth(70);
        grid->addWidget(spin, r, inputCol, Qt::AlignLeft);
        m_spinBoxes[metric] = spin;
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, &DetailedSensoryPanel::scheduleChartRefresh);
        // v2.0.1: per-metric LiveSync emission. Metric scores are flat keys
        // on the sample object in the serialized JSON (sObj[metric] = value).
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, [this, metric](double v) {
                    commitSampleField(metric, v);
                });
    };

    auto addComboAt = [&](int r, int numCol, int labelCol, int inputCol,
                           int num, const QString& metric,
                           const QVector<ChoiceOption>& options) {
        QString label = kDetailedAxisLabels.value(metric, metric);
        label.replace('\n', ' ');
        auto* numLabel = new QLabel(QString::number(num) + ".", m_questionForm);
        numLabel->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
        grid->addWidget(numLabel, r, numCol);
        grid->addWidget(new QLabel(label + ":", m_questionForm), r, labelCol);
        auto* combo = new NoWheelComboBox(m_questionForm);
        for (const auto& opt : options)
            combo->addItem(opt.text, opt.value);
        combo->setMaximumWidth(kComboMaxW);
        grid->addWidget(combo, r, inputCol, Qt::AlignLeft);
        m_comboBoxes[metric] = combo;
        connect(combo, QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, &DetailedSensoryPanel::scheduleChartRefresh);
        // v2.0.1: combo metric scores are flat sample-level keys; the
        // numeric value lives in currentData().
        connect(combo, QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, [this, metric, combo](int) {
                    commitSampleField(metric, combo->currentData().toDouble());
                });
    };

    // Left = cols 0,1,2   Right = cols 4,5,6
    // Template order (S2-1), left-to-right then next row:

    // Row 0: 1. Oil Smell Liking (L) | 2. Burn Taste (R)
    {
        auto* numLabel = new QLabel(QString::number(qNum++) + ".", m_questionForm);
        numLabel->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
        grid->addWidget(numLabel, row, 0);
        grid->addWidget(new QLabel("Oil Smell Liking (1-5):", m_questionForm), row, 1);
        m_oilSmellSpin = new NoWheelDoubleSpinBox(m_questionForm);
        m_oilSmellSpin->setRange(1, 5);
        m_oilSmellSpin->setSingleStep(1);
        m_oilSmellSpin->setDecimals(0);
        m_oilSmellSpin->setValue(3);
        m_oilSmellSpin->setFixedWidth(70);
        grid->addWidget(m_oilSmellSpin, row, 2, Qt::AlignLeft);
        // v2.0.1: session-level oil-smell liking commit.
        connect(m_oilSmellSpin,
                QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, [this](double v) {
                    commitSessionField(QStringLiteral("oil_smell_liking"),
                                       static_cast<int>(v));
                });
    }
    addSpinAt(row, 4, 5, 6, qNum++, "Burn Taste", 1.0, 9.0, 0.1, 1.0);
    ++row;

    // Row 1: 3. Cough (L) | 4. Vapor vs Oil (R)
    addComboAt(row, 0, 1, 2, qNum++, "Cough", kCoughOptions);
    addComboAt(row, 4, 5, 6, qNum++, "Vapor vs Oil", kVaporVsOilOptions);
    ++row;

    // Row 2: 5. Flavor Intensity (L) | 6. Throat Irritation (R)
    addSpinAt(row, 0, 1, 2, qNum++, "Flavor Intensity", 1.0, 9.0, 0.1, 5.0);
    addSpinAt(row, 4, 5, 6, qNum++, "Throat Irritation", 1.0, 9.0, 0.1, 1.0);
    ++row;

    // Row 3: 7. Nasal Irritation (L) | 8. Vapor Quality Overall (R)
    addSpinAt(row, 0, 1, 2, qNum++, "Nasal Irritation", 1.0, 9.0, 0.1, 1.0);
    addSpinAt(row, 4, 5, 6, qNum++, "Vapor Quality Overall", 1.0, 9.0, 0.1, 5.0);
    ++row;

    // Row 4: 9. Volume Consistency (L) | 10. Vapor Volume (R)
    addComboAt(row, 0, 1, 2, qNum++, "Volume Consistency", kVolumeConsistencyOptions);
    addComboAt(row, 4, 5, 6, qNum++, "Vapor Volume", kVaporVolumeOptions);
    ++row;

    // Row 5: 11. Vapor Temperature (L) | 12. Performance Consistency (R)
    addComboAt(row, 0, 1, 2, qNum++, "Vapor Temperature", kVaporTemperatureOptions);
    addComboAt(row, 4, 5, 6, qNum++, "Performance Consistency", kPerformanceConsistencyOptions);
    ++row;

    // Row 6: 13. Mouthpiece/Draw Resistance (L) | 14. Clog (R)
    {
        auto* numLabel = new QLabel(QString::number(qNum++) + ".", m_questionForm);
        numLabel->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
        grid->addWidget(numLabel, row, 0);
        grid->addWidget(new QLabel("Mouthpiece/Draw Resistance:", m_questionForm), row, 1);
        m_mouthpieceCombo = new NoWheelComboBox(m_questionForm);
        for (const auto& opt : kMouthpieceOptions)
            m_mouthpieceCombo->addItem(opt.text, opt.value);
        m_mouthpieceCombo->setMaximumWidth(kComboMaxW);
        grid->addWidget(m_mouthpieceCombo, row, 2, Qt::AlignLeft);
        // v2.0.1: session-level mouthpiece notes commit.
        connect(m_mouthpieceCombo,
                QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, [this](int) {
                    commitSessionField(QStringLiteral("mouthpiece_notes"),
                                       m_mouthpieceCombo->currentText());
                });
    }
    {
        auto* numLabel = new QLabel(QString::number(qNum++) + ".", m_questionForm);
        numLabel->setAlignment(Qt::AlignRight | Qt::AlignVCenter);
        grid->addWidget(numLabel, row, 4);
        grid->addWidget(new QLabel("Clog:", m_questionForm), row, 5);
        m_clogCombo = new NoWheelComboBox(m_questionForm);
        m_clogCombo->addItem("No", false);
        m_clogCombo->addItem("Yes", true);
        m_clogCombo->setMaximumWidth(kComboMaxW);
        grid->addWidget(m_clogCombo, row, 6, Qt::AlignLeft);
        // v2.0.1: session-level clog commit.
        connect(m_clogCombo,
                QOverload<int>::of(&QComboBox::currentIndexChanged),
                this, [this](int) {
                    commitSessionField(QStringLiteral("clog"),
                                       m_clogCombo->currentData().toBool());
                });
    }
    ++row;

    grid->setRowStretch(row, 1);

    outerVBox->addLayout(grid, 1);

    // Comments — full width at bottom
    auto* commentsRow = new QHBoxLayout;
    commentsRow->addWidget(new QLabel("Comments:", m_questionForm));
    m_commentsEdit = new QTextEdit(m_questionForm);
    m_commentsEdit->setMinimumHeight(80);
    m_commentsEdit->setPlaceholderText(tr("Notes about this sample…"));
    m_commentsEdit->setFrameShape(QFrame::Box);
    commentsRow->addWidget(m_commentsEdit, 1);
    outerVBox->addLayout(commentsRow);
    // v2.0.1: sample-level comments commit. QTextEdit fires textChanged on
    // every keystroke and LiveSync does NOT coalesce server-side — a 50-char
    // comment would produce 50 separate BEGIN/UPDATE/COMMIT transactions.
    // A 500 ms single-shot timer collapses a typing burst into one commit.
    // displayCurrentSample() flushes the pending commit before navigating to
    // a different sample so the in-flight text lands on the correct
    // samples[i] path.
    m_commentsCommitTimer = new QTimer(this);
    m_commentsCommitTimer->setSingleShot(true);
    m_commentsCommitTimer->setInterval(500);
    connect(m_commentsCommitTimer, &QTimer::timeout, this, [this]() {
        commitSampleField(QStringLiteral("comments"),
                          m_commentsEdit->toPlainText());
    });
    connect(m_commentsEdit, &QTextEdit::textChanged, this, [this]() {
        m_commentsCommitTimer->start();
    });
    // Plan C (C6 follow-up): also route comment keystrokes through the same
    // capture path as every other sample field (sample name/spins/combos) so a
    // comment-only edit arms the recovery snapshot. The LiveSync commit timer
    // above never emits dataEdited() and early-returns offline / for unpersisted
    // sessions, so without this a comment typed just before a crash or the
    // updater's std::_Exit was silently lost. scheduleChartRefresh()'s 150 ms
    // timer fires onRefreshChart() -> saveCurrentSampleToSession(), which writes
    // m_commentsEdit into the live sample AND emits the diff-guarded dataEdited()
    // -> noteDirty(). Diff-guarding means a no-op textChanged (e.g. re-setting the
    // same text on navigation) does not re-arm the flush timer. Mirrors the
    // m_sampleNameEdit wiring above; leaves the 500 ms LiveSync path untouched.
    connect(m_commentsEdit, &QTextEdit::textChanged,
            this, &DetailedSensoryPanel::scheduleChartRefresh);
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

    // Plan C (C6 fix): this is the single chokepoint where per-field sample
    // edits (name/comments/spins/combos) land in the in-memory session — the
    // value-change lambdas only call scheduleChartRefresh(), whose 150 ms timer
    // fires onRefreshChart() → here. So emitting dataEdited() here covers every
    // detailed-sensory sample-field edit for the recovery snapshot.
    //
    // But this function is ALSO called on read-only paths (sample navigation,
    // and allSessions()/saveCurrentTester() — which the recovery snapshot
    // itself invokes via captureRecoveryState). Emitting unconditionally would
    // re-arm noteDirty() on every snapshot capture and pin the 2 s flush timer
    // on forever. So diff against the prior value and emit only on a real
    // change. The three fields below are exactly what this function writes;
    // voltage/resistance/power/heatingTechnology are not edited via this form.
    const QString prevName     = sample.name;
    const QString prevComments = sample.comments;
    const QMap<QString, double> prevScores = sample.scores;

    sample.name     = m_sampleNameEdit->text();
    sample.comments = m_commentsEdit->toPlainText();

    for (auto it = m_spinBoxes.begin(); it != m_spinBoxes.end(); ++it)
        sample.scores[it.key()] = it.value()->value();

    for (auto it = m_comboBoxes.begin(); it != m_comboBoxes.end(); ++it)
        sample.scores[it.key()] = it.value()->currentData().toDouble();

    if (sample.name != prevName
        || sample.comments != prevComments
        || sample.scores != prevScores) {
        emit dataEdited();
    }
}

void DetailedSensoryPanel::displayCurrentSample()
{
    // v2.0.1: flush any pending comments commit on the OLD sample before
    // switching. The debounce timer is keyed on whichever sample is currently
    // displayed; without this flush a fast typist's last few keystrokes would
    // commit against the NEW m_currentSampleIdx after navigation.
    if (m_commentsCommitTimer && m_commentsCommitTimer->isActive()) {
        m_commentsCommitTimer->stop();
        commitSampleField(QStringLiteral("comments"),
                          m_commentsEdit->toPlainText());
    }

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
    // v2.0.6: see SensoryPanel::loadSessions for the matching commentary.
    inheritExistingIdsAndVersions();
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

void DetailedSensoryPanel::inheritExistingIdsAndVersions()
{
    // v2.0.6: bulk variant. See SensoryPanel::inheritExistingIdsAndVersions
    // for the rationale (v2.0.5 issued N round-trips on the UI thread per
    // Ctrl+U / 5-second auto-save tick, which froze the app on slower LAN
    // segments). This version does one round-trip regardless of session
    // count, and is invoked from loadSessions so the save path pays
    // nothing.
    if (!m_db) return;
    QVector<DatabaseManager::NaturalKey> keys;
    keys.reserve(m_sessions.size());
    for (const DetailedSensorySession& s : m_sessions) {
        if (s.id > 0) continue;
        DatabaseManager::NaturalKey k;
        k.sessionName = s.sessionName.trimmed();
        k.testerName  = s.testerName.trimmed();
        k.date        = s.date.trimmed();
        if (k.sessionName.isEmpty() && k.testerName.isEmpty() && k.date.isEmpty())
            continue;
        keys.append(k);
    }
    if (keys.isEmpty()) return;

    const auto matches = m_db->findDetailedSensorySessionsByKeys(keys);
    if (matches.isEmpty()) return;

    QHash<QString, DatabaseManager::SessionKey> byKey;
    byKey.reserve(matches.size());
    for (const auto& m : matches) {
        const QString k = m.sessionName + QChar('\x1f')
                        + m.testerName + QChar('\x1f') + m.date;
        byKey.insert(k, {m.id, m.version});
    }
    for (DetailedSensorySession& s : m_sessions) {
        if (s.id > 0) continue;
        const QString k = s.sessionName.trimmed() + QChar('\x1f')
                        + s.testerName.trimmed() + QChar('\x1f')
                        + s.date.trimmed();
        const auto it = byKey.constFind(k);
        if (it != byKey.constEnd()) {
            s.id      = static_cast<int>(it->id);
            s.version = it->version;
        }
    }
}

void DetailedSensoryPanel::syncSavedSessionState(
    const QVector<DetailedSensorySession>& saved)
{
    // Copy a saved session's persistence anchors back into a panel session.
    // Anchors only — never the user-editable content. id<=0 means the save did
    // not land (placeholder/version-mismatch/error), so leave the panel as-is.
    const auto adopt = [](DetailedSensorySession& dst,
                          const DetailedSensorySession& src) {
        if (src.id > 0) {
            dst.id      = src.id;
            dst.version = src.version;
        }
        // Per-image identity back-fill: tryWriteDetailedSensorySession writes
        // back imageIds + imageVersions (parallel to imagePaths) for any new
        // image rows it inserted, so repeat saves UPDATE instead of re-INSERT.
        if (!src.imageIds.isEmpty())
            dst.imageIds = src.imageIds;
        if (!src.imageVersions.isEmpty())
            dst.imageVersions = src.imageVersions;
    };

    // Fast path: allSessions() returned m_sessions verbatim, so the saved copy
    // is index-aligned (mirrors SensoryPanel::syncSavedSessionState).
    if (saved.size() == m_sessions.size()) {
        for (int i = 0; i < m_sessions.size(); ++i)
            adopt(m_sessions[i], saved[i]);
        return;
    }

    // Counts diverged (e.g. a session was added/removed between the snapshot
    // and now) — fall back to a natural-key match so anchors still land on the
    // right rows. First saved session per key wins.
    QHash<QString, int> byKey;
    byKey.reserve(saved.size());
    for (int i = 0; i < saved.size(); ++i) {
        const DetailedSensorySession& s = saved[i];
        const QString k = s.sessionName.trimmed() + QChar('\x1f')
                        + s.testerName.trimmed() + QChar('\x1f')
                        + s.date.trimmed();
        if (!byKey.contains(k))
            byKey.insert(k, i);
    }
    for (DetailedSensorySession& dst : m_sessions) {
        const QString k = dst.sessionName.trimmed() + QChar('\x1f')
                        + dst.testerName.trimmed() + QChar('\x1f')
                        + dst.date.trimmed();
        const auto it = byKey.constFind(k);
        if (it != byKey.constEnd())
            adopt(dst, saved[it.value()]);
    }
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
    if (m_oilSmellSpin)
        sess.oilSmellLiking = static_cast<int>(m_oilSmellSpin->value());
    if (m_clogCombo)
        sess.clog = m_clogCombo->currentData().toBool();
    if (m_mouthpieceCombo)
        sess.mouthpieceNotes = m_mouthpieceCombo->currentText();
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
    if (m_oilSmellSpin) {
        m_oilSmellSpin->blockSignals(true);
        m_oilSmellSpin->setValue(session.oilSmellLiking);
        m_oilSmellSpin->blockSignals(false);
    }
    if (m_clogCombo) {
        m_clogCombo->blockSignals(true);
        m_clogCombo->setCurrentIndex(session.clog ? 1 : 0);
        m_clogCombo->blockSignals(false);
    }
    if (m_mouthpieceCombo) {
        m_mouthpieceCombo->blockSignals(true);
        int idx = m_mouthpieceCombo->findText(session.mouthpieceNotes);
        m_mouthpieceCombo->setCurrentIndex(idx >= 0 ? idx : 0);
        m_mouthpieceCombo->blockSignals(false);
    }

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
            OutputPaths::resolveDir(ReportMode::DetailedSensory,m_lastBrowseDir) + "/" + defaultName + ".xlsx",
            "Excel Files (*.xlsx)");
        if (m_savePath.isEmpty()) return;
        setLastBrowseDir(m_savePath);
    }

    saveToExcel(m_savePath, sess);

    // v2.0.1: LiveSync owns per-cell DB persistence. Manual save still
    // flushes the current session in bulk as a fallback for fresh imports
    // and offline replays.
    if (m_db && m_db->isOpen()) {
        if (!m_db->saveDetailedSensorySession(sess)) {
            QMessageBox::warning(this, "Database Save Failed",
                "The session was saved to Excel but could not be written to "
                "the database:\n\n" + m_db->lastError() +
                "\n\nThis is logged. Try again, or contact the developer if "
                "it keeps failing.");
        }
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

void DetailedSensoryPanel::loadFile(const QString& path)
{
    saveCurrentTester();

    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    QXlsx::Document xlsx(path);
    if (!xlsx.load()) return;

    DetailedSensorySession sess;
    sess.sessionName = QFileInfo(path).baseName();
    sess.date = QDate::currentDate().toString("yyyy-MM-dd");
    sess.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    QMap<int, QString> colMetric;
    int lastCol = 1;
    for (int col = 2; col <= 30; ++col) {
        QString hdr = xlsx.read(1, col).toString().trimmed();
        if (hdr.isEmpty()) break;
        lastCol = col;
        for (const QString& metric : kDetailedAllMetrics) {
            if (hdr.compare(metric, Qt::CaseInsensitive) == 0) {
                colMetric[col] = metric;
                break;
            }
        }
    }

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

    if (sess.samples.isEmpty()) return;

    // If a session with the same natural key already exists in the database
    // (e.g., re-importing a file that was migrated, or in-memory dup from a
    // prior load), inherit id+version so the save below UPDATEs instead of
    // INSERTing. Prevents UniqueViolationDialog on every reimport.
    if (m_db && m_db->isOpen()) {
        for (const DetailedSensoryRecord& r : m_db->listDetailedSensoryRecords()) {
            if (r.sessionName == sess.sessionName
                && r.testerName  == sess.testerName
                && r.date        == sess.date) {
                const DetailedSensorySession existing = m_db->loadDetailedSensorySession(r.id);
                if (existing.id > 0) {
                    sess.id      = existing.id;
                    sess.version = existing.version;
                }
                break;
            }
        }
    }

    m_sessions.append(sess);

    if (m_db && m_db->isOpen()) {
        m_db->saveDetailedSensorySession(m_sessions.last());
    }

    m_currentTesterIdx = m_sessions.size() - 1;
    m_currentSampleIdx = 0;
    applySession(m_sessions[m_currentTesterIdx]);
    emit sessionsChanged();
}

void DetailedSensoryPanel::loadFiles()
{
    QStringList files = QFileDialog::getOpenFileNames(
        this, "Load Detailed Sensory Data",
        OutputPaths::resolveDir(ReportMode::DetailedSensory, lastBrowseDir()),
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

// ─────────────────────────────────────────────────────────────────────────────
// DATAVIEWER-4: DB-authoritative export source (detailed-sensory twin of
// SensoryPanel::dbAuthoritativeSessions, commit adb3545)
// ─────────────────────────────────────────────────────────────────────────────
//
// The detailed-sensory report reads the in-memory session model, which can hold
// stale per-metric scores — this client may have missed a LiveSync NOTIFY, or
// another user edited a score concurrently. The DB blob is the single source of
// truth for scores. This helper makes the report score-authoritative:
//   1. Flush our own pending per-cell edits to the DB once, so the DB row holds
//      this client's latest scores before we re-fetch.
//   2. For each persisted session, re-fetch the DB row and overlay ONLY the
//      kDetailedAllMetrics score values onto a COPY of the in-memory session.
//
// Scores are taken from the DB; everything else (header/session metadata, sample
// names, comments, V/R/power, AND the non-serialized fields the report renderer
// needs — imagePaths/imageLayouts/imageCrops/imageIds/imageVersions/id/version)
// stays in-memory-authoritative. We deliberately do NOT round-trip the struct
// through detailedSensorySessionFromJson: that JSON contract omits the image
// vectors and persistence anchors, so a round-trip would silently drop every
// image from exported reports. Instead we reuse the pure
// mergeDetailedSensoryPreservingDbScores() at the JSON layer purely to compute
// the authoritative per-metric values, then copy those scalars back into the
// live struct. Unsaved sessions (id <= 0) and rows that have since disappeared
// pass through unchanged.
QVector<DetailedSensorySession> DetailedSensoryPanel::dbAuthoritativeSessions(
        const QVector<DetailedSensorySession>& inMem)
{
    if (m_liveSync) m_liveSync->flushNowAndWait();   // our latest scores -> DB
    if (!m_db) return inMem;

    QVector<DetailedSensorySession> out;
    out.reserve(inMem.size());
    for (const DetailedSensorySession& s : inMem) {
        if (s.id <= 0) { out.append(s); continue; }            // never persisted
        const DetailedSensorySession dbSess = m_db->loadDetailedSensorySession(s.id);
        if (dbSess.id <= 0) { out.append(s); continue; }       // row gone / load failed

        // Compute DB-authoritative scores at the JSON layer, then overlay them
        // back onto a copy of the in-memory struct so all non-JSON fields
        // (images, anchors) survive untouched.
        const QJsonObject merged = mergeDetailedSensoryPreservingDbScores(
            detailedSensorySessionToJson(s), detailedSensorySessionToJson(dbSess));
        const QJsonArray mergedSamples = merged.value("samples").toArray();

        DetailedSensorySession authoritative = s;  // keep every in-memory field
        for (int i = 0; i < authoritative.samples.size()
                        && i < mergedSamples.size(); ++i) {
            const QJsonObject ms = mergedSamples[i].toObject();
            for (const QString& metric : kDetailedAllMetrics) {
                if (ms.contains(metric))
                    authoritative.samples[i].scores[metric] = ms.value(metric).toDouble();
            }
        }
        out.append(authoritative);
    }
    return out;
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

    // DATAVIEWER-4: scores in the report must come from the DB, not stale
    // in-memory state. Everything else (metadata, images) stays in-memory.
    // This is the outermost in-memory-model export entry; the static
    // generateCombinedPptx() it calls is deliberately left unrouted (see note
    // at that function).
    selected = dbAuthoritativeSessions(selected);

    const QString titleBase = selected[0].testTitle.isEmpty()
                                  ? QStringLiteral("detailed_sensory")
                                  : selected[0].testTitle;
    const QString dir      = OutputPaths::resolveDir(ReportMode::DetailedSensory, m_lastBrowseDir);
    const QString fileName = OutputPaths::reportFileName(titleBase);
    QString path = QFileDialog::getSaveFileName(
        this, "Save Report",
        dir + "/" + fileName,
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
    // DATAVIEWER-4 note: this static entry point does NOT route through
    // dbAuthoritativeSessions() — and must not. It has no instance access to
    // m_db / m_liveSync, and its other caller (DatabaseBrowserDialog, ~line 903)
    // builds its `sessions` list by loading rows straight from the DB, so those
    // scores are already authoritative. The in-memory-model export
    // (generateFullReport) is the one that needs the reconciliation, and it
    // applies the helper before calling here.
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
    if (!m_lastBrowseDir.isEmpty() && QDir(m_lastBrowseDir).exists())
        return m_lastBrowseDir;

    return OutputPaths::documentsDir();
}

void DetailedSensoryPanel::setLastBrowseDir(const QString& filePath)
{
    m_lastBrowseDir = QFileInfo(filePath).absolutePath();
}

// ─────────────────────────────────────────────────────────────────────────────
// v2.0.1 LiveSync wiring
// ─────────────────────────────────────────────────────────────────────────────

void DetailedSensoryPanel::setLiveSync(LiveSync* sync)
{
    if (m_liveSync == sync) return;
    if (m_liveSync) disconnect(m_liveSync, nullptr, this, nullptr);
    m_liveSync = sync;
    if (m_liveSync) {
        connect(m_liveSync, &LiveSync::cellChanged, this,
                &DetailedSensoryPanel::onRemoteCellChanged);
    }
}

int DetailedSensoryPanel::activeSessionId() const
{
    if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size())
        return -1;
    const int id = m_sessions[m_currentTesterIdx].id;
    return id > 0 ? id : -1;
}

void DetailedSensoryPanel::commitSessionField(const QString& fieldPath,
                                              const QVariant& value)
{
    // Plan C (C6 fix): every session-level field widget (header
    // title/assessor/tester/media + oil-smell/clog/mouthpiece) routes its
    // change here on editingFinished / currentIndexChanged. These edits land
    // in the in-memory session at buildSession() time (called by allSessions()
    // during recovery capture), so arming noteDirty() here is sufficient for
    // the snapshot to pick them up. Emit BEFORE the LiveSync early-return so
    // recovery fires even offline (when m_liveSync is null) and for brand-new
    // unpersisted sessions (activeSessionId() < 0) the DB doesn't have yet.
    emit dataEdited();

    if (!m_liveSync) return;
    const int sessionId = activeSessionId();
    if (sessionId < 0) return;

    // Some session-level fields exist BOTH as scalar columns on
    // detailed_sensory_sessions AND inside json_data (the serializer
    // writes both). The deserializer reads from json_data, but the
    // UNIQUE index on (session_name, tester_name, date) keys off the
    // scalar columns. A live edit must update both so the scalar
    // columns don't go stale — otherwise duplicate-detection breaks
    // and external queries see the wrong value. json-only fields like
    // test_title / oil_smell_liking / clog / mouthpiece_notes live
    // only inside json_data and don't need the second write.
    static const QSet<QString> kScalarDualFields = {
        QStringLiteral("session_name"),
        QStringLiteral("tester_name"),
        QStringLiteral("assessor_name"),
        QStringLiteral("media"),
        QStringLiteral("date")
    };
    if (kScalarDualFields.contains(fieldPath)) {
        m_liveSync->commitCell(QStringLiteral("detailed_sensory_sessions"),
                               sessionId, fieldPath, value);
    }
    m_liveSync->commitCell(QStringLiteral("detailed_sensory_sessions"),
                           sessionId,
                           QStringLiteral("json_path:") + fieldPath,
                           value);
}

void DetailedSensoryPanel::commitSampleField(const QString& fieldPath,
                                             const QVariant& value)
{
    if (!m_liveSync) return;
    const int sessionId = activeSessionId();
    if (sessionId < 0) return;
    const DetailedSensorySession& s = m_sessions[m_currentTesterIdx];
    if (m_currentSampleIdx < 0 || m_currentSampleIdx >= s.samples.size()) return;
    const QString jsonPath = QStringLiteral("json_path:samples[%1].%2")
                                .arg(m_currentSampleIdx).arg(fieldPath);
    m_liveSync->commitCell(QStringLiteral("detailed_sensory_sessions"),
                           sessionId, jsonPath, value);
}

void DetailedSensoryPanel::onRemoteCellChanged(const QString& table,
                                               qint64 rowId,
                                               const QString& column,
                                               const QVariant& newValue)
{
    if (table != QLatin1String("detailed_sensory_sessions")) return;
    if (!column.startsWith(QLatin1String("json_path:"))) return;

    const QString path = column.mid(QStringLiteral("json_path:").size());

    int sessIdx = -1;
    for (int i = 0; i < m_sessions.size(); ++i) {
        if (static_cast<qint64>(m_sessions[i].id) == rowId) { sessIdx = i; break; }
    }
    if (sessIdx < 0) return;

    if (path.startsWith(QLatin1String("samples["))) {
        const QString rest = path.mid(QStringLiteral("samples[").size());
        const int rbr = rest.indexOf(QLatin1Char(']'));
        if (rbr < 0) return;
        bool okIdx = false;
        const int idx = rest.left(rbr).toInt(&okIdx);
        if (!okIdx) return;
        if (rest.size() <= rbr + 2) return;
        const QString fieldPath = rest.mid(rbr + 2);  // skip "]."
        if (idx < 0 || idx >= m_sessions[sessIdx].samples.size()) return;
        applyRemoteFieldToSample(
            m_sessions[sessIdx].samples[idx], fieldPath, newValue);
        if (sessIdx == m_currentTesterIdx && idx == m_currentSampleIdx) {
            rebindCurrentSampleFromMemory();
        }
    } else {
        applyRemoteSessionField(m_sessions[sessIdx], path, newValue);
        if (sessIdx == m_currentTesterIdx) {
            // Patch the visible session-level widgets with signals blocked.
            QSignalBlocker bTitle(m_testTitleEdit);
            QSignalBlocker bAssess(m_assessorEdit);
            QSignalBlocker bTester(m_testerEdit);
            QSignalBlocker bMedia(m_mediaEdit);
            m_testTitleEdit->setText(m_sessions[sessIdx].testTitle);
            m_assessorEdit->setText(m_sessions[sessIdx].assessorName);
            m_testerEdit->setText(m_sessions[sessIdx].testerName);
            m_mediaEdit->setText(m_sessions[sessIdx].media);
            if (m_oilSmellSpin) {
                QSignalBlocker b(m_oilSmellSpin);
                m_oilSmellSpin->setValue(m_sessions[sessIdx].oilSmellLiking);
            }
            if (m_clogCombo) {
                QSignalBlocker b(m_clogCombo);
                m_clogCombo->setCurrentIndex(m_sessions[sessIdx].clog ? 1 : 0);
            }
            if (m_mouthpieceCombo) {
                QSignalBlocker b(m_mouthpieceCombo);
                const int idx = m_mouthpieceCombo->findText(
                    m_sessions[sessIdx].mouthpieceNotes);
                m_mouthpieceCombo->setCurrentIndex(idx >= 0 ? idx : 0);
            }
        }
    }
}

void DetailedSensoryPanel::applyRemoteFieldToSample(
    DetailedSensorySample& sample,
    const QString& fieldPath,
    const QVariant& value)
{
    if (fieldPath == QLatin1String("name"))              sample.name = value.toString();
    else if (fieldPath == QLatin1String("comments"))     sample.comments = value.toString();
    else if (fieldPath == QLatin1String("voltage"))      sample.voltage = value.toDouble();
    else if (fieldPath == QLatin1String("resistance"))   sample.resistance = value.toDouble();
    else if (fieldPath == QLatin1String("power"))        sample.power = value.toDouble();
    else if (fieldPath == QLatin1String("heating_technology"))
                                                         sample.heatingTechnology = value.toString();
    else if (kDetailedAllMetrics.contains(fieldPath))    sample.scores[fieldPath] = value.toDouble();
}

void DetailedSensoryPanel::applyRemoteSessionField(
    DetailedSensorySession& sess,
    const QString& fieldPath,
    const QVariant& value)
{
    if (fieldPath == QLatin1String("session_name"))      sess.sessionName = value.toString();
    else if (fieldPath == QLatin1String("test_title"))   sess.testTitle = value.toString();
    else if (fieldPath == QLatin1String("assessor_name")) sess.assessorName = value.toString();
    else if (fieldPath == QLatin1String("tester_name"))  sess.testerName = value.toString();
    else if (fieldPath == QLatin1String("facilitator_name"))    sess.facilitatorName = value.toString();
    else if (fieldPath == QLatin1String("facilitator_comment")) sess.facilitatorComment = value.toString();
    else if (fieldPath == QLatin1String("media"))        sess.media = value.toString();
    else if (fieldPath == QLatin1String("date"))         sess.date = value.toString();
    else if (fieldPath == QLatin1String("oil_smell_liking")) sess.oilSmellLiking = value.toInt();
    else if (fieldPath == QLatin1String("clog"))         sess.clog = value.toBool();
    else if (fieldPath == QLatin1String("clog_oil_level")) sess.clogOilLevel = value.toString();
    else if (fieldPath == QLatin1String("mouthpiece_notes")) sess.mouthpieceNotes = value.toString();
    else if (fieldPath == QLatin1String("device_return_date")) sess.deviceReturnDate = value.toString();
    else if (fieldPath == QLatin1String("viscosity"))    sess.viscosity = value.toString();
}

void DetailedSensoryPanel::rebindCurrentSampleFromMemory()
{
    if (m_currentTesterIdx < 0
        || m_currentTesterIdx >= m_sessions.size()) return;
    const DetailedSensorySession& sess = m_sessions[m_currentTesterIdx];
    if (m_currentSampleIdx < 0
        || m_currentSampleIdx >= sess.samples.size()) return;
    const DetailedSensorySample& sample = sess.samples[m_currentSampleIdx];

    for (auto* spin : m_spinBoxes) spin->blockSignals(true);
    for (auto* combo : m_comboBoxes) combo->blockSignals(true);
    QSignalBlocker bName(m_sampleNameEdit);
    QSignalBlocker bComments(m_commentsEdit);

    m_sampleNameEdit->setText(sample.name);
    m_commentsEdit->setPlainText(sample.comments);
    for (auto it = m_spinBoxes.begin(); it != m_spinBoxes.end(); ++it) {
        const double val = sample.scores.value(it.key(), it.value()->minimum());
        it.value()->setValue(val);
    }
    for (auto it = m_comboBoxes.begin(); it != m_comboBoxes.end(); ++it) {
        const int rawVal = static_cast<int>(sample.scores.value(it.key(), 1.0));
        const int idx = it.value()->findData(rawVal);
        if (idx >= 0) it.value()->setCurrentIndex(idx);
    }
    for (auto* spin : m_spinBoxes) spin->blockSignals(false);
    for (auto* combo : m_comboBoxes) combo->blockSignals(false);
}

} // namespace DVE
