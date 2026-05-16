#include "SensoryPanel.h"

#include "database/SaveCoordinator.h"

#include <QHBoxLayout>
#include <QFormLayout>
#include <QPushButton>
#include <QFileDialog>
#include <QMessageBox>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QDateTime>
#include <QFile>
#include <QTextStream>
#include <QStyle>
#include <QLabel>
#include <QScrollBar>
#include <QWheelEvent>
#include <QApplication>
#include <QScreen>
#include <QDir>
#include <QPainter>
#include <QPixmap>
#include <QImage>
#include <QBuffer>
#include <QFrame>
#include <QFileInfo>
#include <QHeaderView>
#include <QTreeWidget>
#include <QListWidget>
#include <cmath>
#include <climits>
#include <utility>

#include "xlsxdocument.h"
#include "reporting/PptxWriter.h"
#include "reporting/SensoryReportSource.h"
#include "ReportPreviewDialog.h"
#include "utils/ImageUtils.h"

namespace DVE {

namespace {
// QTableWidget's default sort uses QTableWidgetItem::operator< which compares
// the DisplayRole as a string — so "10" sorts before "9". This subclass
// stores the raw double in UserRole and compares against UserRole instead,
// giving numeric sort order while preserving the formatted display string.
class NumericTableItem : public QTableWidgetItem {
public:
    explicit NumericTableItem(double value)
        : QTableWidgetItem(QString::number(value, 'f', 1)) {
        setData(Qt::UserRole, value);
        setTextAlignment(Qt::AlignCenter);
    }
    bool operator<(const QTableWidgetItem& other) const override {
        return data(Qt::UserRole).toDouble()
             < other.data(Qt::UserRole).toDouble();
    }
};
} // anonymous namespace

// Ignores wheel events unless the spinbox has keyboard focus.
// Prevents accidental value changes when scrolling past sample cards.
class NoWheelDoubleSpinBox : public QDoubleSpinBox
{
public:
    explicit NoWheelDoubleSpinBox(QWidget* parent = nullptr) : QDoubleSpinBox(parent) {
        setFocusPolicy(Qt::StrongFocus);
    }
protected:
    void wheelEvent(QWheelEvent* e) override {
        e->ignore();  // never consume scroll — always let parent scroll area handle it
    }
};

// ─────────────────────────────────────────────────────────────────────────────
// FlowLayout — cards wrap left-to-right, then down
// (Adapted from Qt's Flow Layout example.)
// ─────────────────────────────────────────────────────────────────────────────

FlowLayout::FlowLayout(QWidget* parent, int margin, int hSpacing, int vSpacing)
    : QLayout(parent), m_hSpace(hSpacing), m_vSpace(vSpacing)
{
    setContentsMargins(margin, margin, margin, margin);
}

FlowLayout::~FlowLayout()
{
    QLayoutItem* item;
    while ((item = takeAt(0)))
        delete item;
}

void FlowLayout::addItem(QLayoutItem* item) { m_items.append(item); }

int FlowLayout::horizontalSpacing() const
{
    return (m_hSpace >= 0) ? m_hSpace : 6;
}

int FlowLayout::verticalSpacing() const
{
    return (m_vSpace >= 0) ? m_vSpace : 6;
}

int FlowLayout::count() const { return m_items.size(); }

QLayoutItem* FlowLayout::itemAt(int index) const
{
    return (index >= 0 && index < m_items.size()) ? m_items.at(index) : nullptr;
}

QLayoutItem* FlowLayout::takeAt(int index)
{
    if (index >= 0 && index < m_items.size())
        return m_items.takeAt(index);
    return nullptr;
}

Qt::Orientations FlowLayout::expandingDirections() const { return {}; }

bool FlowLayout::hasHeightForWidth() const { return true; }

int FlowLayout::heightForWidth(int width) const
{
    return doLayout(QRect(0, 0, width, 0), true);
}

void FlowLayout::setGeometry(const QRect& rect)
{
    QLayout::setGeometry(rect);
    doLayout(rect, false);
}

QSize FlowLayout::sizeHint() const
{
    return minimumSize();
}

QSize FlowLayout::minimumSize() const
{
    QSize size;
    for (const QLayoutItem* item : m_items)
        size = size.expandedTo(item->minimumSize());
    int m = contentsMargins().left() + contentsMargins().right();
    return QSize(size.width() + m, size.height() + m);
}

int FlowLayout::doLayout(const QRect& rect, bool testOnly) const
{
    int left, top, right, bottom;
    getContentsMargins(&left, &top, &right, &bottom);
    QRect effectiveRect = rect.adjusted(left, top, -right, -bottom);
    int availW = effectiveRect.width();
    int spaceX = horizontalSpacing();
    int spaceY = verticalSpacing();

    QVector<QSize> sizeHints;
    sizeHints.reserve(m_items.size());
    for (int i = 0; i < m_items.size(); ++i)
        sizeHints.append(m_items[i]->sizeHint());

    struct Row { int firstIdx; int count; int totalWidth; int maxHeight; };
    QVector<Row> rows;
    {
        int rowWidth = 0;
        int rowStart = 0;
        int rowHeight = 0;
        for (int i = 0; i < m_items.size(); ++i) {
            const QSize& sz = sizeHints[i];
            int needed = (i == rowStart) ? sz.width() : rowWidth + spaceX + sz.width();
            if (needed > availW && i != rowStart) {
                rows.append({rowStart, i - rowStart, rowWidth, rowHeight});
                rowStart = i;
                rowWidth = sz.width();
                rowHeight = sz.height();
            } else {
                rowWidth = needed;
                rowHeight = qMax(rowHeight, sz.height());
            }
        }
        if (rowStart < m_items.size())
            rows.append({rowStart, static_cast<int>(m_items.size() - rowStart), rowWidth, rowHeight});
    }

    int y = effectiveRect.y();
    for (const Row& row : rows) {
        int xOffset = (availW - row.totalWidth) / 2;
        int x = effectiveRect.x() + xOffset;
        for (int i = row.firstIdx; i < row.firstIdx + row.count; ++i) {
            const QSize& sz = sizeHints[i];
            if (!testOnly)
                m_items[i]->setGeometry(QRect(QPoint(x, y), sz));
            x += sz.width() + spaceX;
        }
        y += row.maxHeight + spaceY;
    }

    return y - spaceY - rect.y() + bottom;
}

// ─────────────────────────────────────────────────────────────────────────────
// SampleCard
// ─────────────────────────────────────────────────────────────────────────────

SampleCard::SampleCard(int index, QWidget* parent)
    : QGroupBox(QString("Sample %1").arg(index + 1), parent)
{
    // #7: +25px for the power-type row in devGrid and the puff-length row
    // inserted between scoring and comments.
    setFixedSize(263, 485);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(3);
    mainLayout->setContentsMargins(6, 16, 6, 4);

    auto* nameRow = new QHBoxLayout;
    nameRow->setSpacing(3);
    nameRow->addWidget(new QLabel("Name:"));
    m_nameEdit = new QLineEdit;
    m_nameEdit->setPlaceholderText(QString("Sample %1").arg(index + 1));
    nameRow->addWidget(m_nameEdit);
    mainLayout->addLayout(nameRow);

    // ── Per-sample device properties ──
    auto* devGrid = new QGridLayout;
    devGrid->setSpacing(2);
    devGrid->setContentsMargins(0, 0, 0, 0);

    auto makeSmallLabel = [](const QString& text) {
        auto* lbl = new QLabel(text);
        lbl->setStyleSheet("font-size: 7pt; color: #555;");
        return lbl;
    };

    devGrid->addWidget(makeSmallLabel("V:"), 0, 0);
    m_voltageEdit = new QLineEdit;
    m_voltageEdit->setFixedWidth(52);
    m_voltageEdit->setFixedHeight(20);
    m_voltageEdit->setPlaceholderText("0.00");
    m_voltageEdit->setStyleSheet("font-size: 7pt;");
    devGrid->addWidget(m_voltageEdit, 0, 1);

    devGrid->addWidget(makeSmallLabel("R:"), 0, 2);
    m_resistanceEdit = new QLineEdit;
    m_resistanceEdit->setFixedWidth(52);
    m_resistanceEdit->setFixedHeight(20);
    m_resistanceEdit->setPlaceholderText("0.000");
    m_resistanceEdit->setStyleSheet("font-size: 7pt;");
    devGrid->addWidget(m_resistanceEdit, 0, 3);

    devGrid->addWidget(makeSmallLabel("HT:"), 0, 4);
    m_heatingTechCombo = new QComboBox;
    m_heatingTechCombo->setFixedWidth(72);
    m_heatingTechCombo->setFixedHeight(20);
    m_heatingTechCombo->setStyleSheet("font-size: 7pt;");
    m_heatingTechCombo->addItems({"", "EVO", "EVOMAX", "SE", "CCELL3.0", "T58G", "T51", "Competitor"});
    m_heatingTechCombo->setEditable(true);
    devGrid->addWidget(m_heatingTechCombo, 0, 5);

    m_powerLabel = new QLabel;
    m_powerLabel->setStyleSheet("font-size: 7pt; color: #333;");
    devGrid->addWidget(makeSmallLabel("P:"), 1, 0);
    devGrid->addWidget(m_powerLabel, 1, 1, 1, 5);

    // #7: power-type combo lives in the device grid below V/R/P.
    devGrid->addWidget(makeSmallLabel("PT:"), 2, 0);
    m_powerTypeCombo = new QComboBox;
    m_powerTypeCombo->setFixedHeight(20);
    m_powerTypeCombo->setStyleSheet("font-size: 7pt;");
    m_powerTypeCombo->addItems({
        tr("Constant Voltage"), tr("Constant Power"),
        tr("Variable Voltage"), tr("Variable Power")
    });
    devGrid->addWidget(m_powerTypeCombo, 2, 1, 1, 5);

    mainLayout->addLayout(devGrid);

    // Wire up power recalculation
    connect(m_voltageEdit, &QLineEdit::textChanged, this, [this]() { recalcPower(); emit changed(); });
    connect(m_resistanceEdit, &QLineEdit::textChanged, this, [this]() { recalcPower(); emit changed(); });
    connect(m_heatingTechCombo, &QComboBox::currentTextChanged, this, [this]() { recalcPower(); emit changed(); });
    connect(m_powerTypeCombo, &QComboBox::currentTextChanged,
            this, [this](const QString&) { emit changed(); });

    // ── Sensory score spinboxes (decimal, 0.1 step) ──
    auto* formLayout = new QFormLayout;
    formLayout->setSpacing(4);
    formLayout->setContentsMargins(0, 2, 0, 2);
    // Tooltips describe what each end of the 1-9 scale means. Hovering either
    // the label or the spinbox surfaces the same guidance.
    static const QMap<QString, QString> kMetricTooltips = {
        { "Burnt Taste",    QObject::tr("Rank 1-9. 9 means no burn, 1 means bad burn.") },
        { "Vapor Volume",   QObject::tr("Rank 1-9. 9 means big cloud, 1 means no cloud.") },
        { "Overall Flavor", QObject::tr("Rank 1-9. 9 means best flavor, 1 means worst flavor.") },
        { "Smoothness",     QObject::tr("Rank 1-9. 9 means extremely smooth, 1 means extremely harsh.") },
        { "Overall Liking", QObject::tr("Rank 1-9. 9 means it's the best, 1 means it's the worst.") },
    };
    for (const QString& metric : kSensoryMetrics) {
        auto* spin = new NoWheelDoubleSpinBox;
        spin->setRange(1.0, 9.0);
        spin->setSingleStep(0.1);
        spin->setDecimals(1);
        spin->setValue(5.0);
        spin->setFixedWidth(65);
        spin->setFixedHeight(22);
        const QString tip = kMetricTooltips.value(metric);
        if (!tip.isEmpty()) spin->setToolTip(tip);
        m_spinBoxes[metric] = spin;
        formLayout->addRow(metric + ":", spin);
        if (!tip.isEmpty()) {
            if (auto* lbl = formLayout->labelForField(spin)) lbl->setToolTip(tip);
        }
        connect(spin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
                this, &SampleCard::changed);
    }
    mainLayout->addLayout(formLayout);

    // #7: puff length sits between scoring and comments per spec.
    auto* puffRow = new QHBoxLayout;
    puffRow->setSpacing(4);
    auto* puffLabel = new QLabel(tr("Puff length:"));
    puffLabel->setStyleSheet("font-size: 7pt; color: #555;");
    puffRow->addWidget(puffLabel);
    m_puffLengthSpin = new NoWheelDoubleSpinBox;
    m_puffLengthSpin->setRange(0.1, 60.0);
    m_puffLengthSpin->setSingleStep(0.5);
    m_puffLengthSpin->setDecimals(1);
    m_puffLengthSpin->setSuffix(QStringLiteral(" s"));
    m_puffLengthSpin->setValue(3.0);
    m_puffLengthSpin->setFixedWidth(72);
    m_puffLengthSpin->setFixedHeight(20);
    m_puffLengthSpin->setStyleSheet("font-size: 7pt;");
    puffRow->addWidget(m_puffLengthSpin);
    puffRow->addStretch();
    mainLayout->addLayout(puffRow);
    connect(m_puffLengthSpin, QOverload<double>::of(&QDoubleSpinBox::valueChanged),
            this, [this](double) { emit changed(); });

    mainLayout->addWidget(new QLabel("Comments:"));
    m_commentsEdit = new QTextEdit;
    m_commentsEdit->setMinimumHeight(36);
    m_commentsEdit->setMaximumHeight(50);
    m_commentsEdit->setStyleSheet("QTextEdit { border: 1px solid #999; background: white; }");
    mainLayout->addWidget(m_commentsEdit, 1);
    connect(m_commentsEdit, &QTextEdit::textChanged, this, &SampleCard::changed);
    connect(m_nameEdit, &QLineEdit::textChanged, this, &SampleCard::changed);

    auto* removeBtn = new QPushButton("Remove");
    removeBtn->setFixedWidth(60);
    removeBtn->setFixedHeight(20);
    removeBtn->setStyleSheet("font-size: 7pt;");
    auto* btnRow = new QHBoxLayout;
    btnRow->addStretch();
    btnRow->addWidget(removeBtn);
    mainLayout->addLayout(btnRow);

    connect(removeBtn, &QPushButton::clicked, this, [this]() {
        emit removeRequested(this);
    });
}

void SampleCard::recalcPower()
{
    double v = m_voltageEdit->text().toDouble();
    double r = m_resistanceEdit->text().toDouble();
    QString tech = m_heatingTechCombo->currentText().trimmed().toUpper();

    double rOffset = 0.0;
    if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
        rOffset = 0.78;
    else if (tech == "T51")
        rOffset = 0.25;

    double denom = r + rOffset;
    double power = (v > 0 && denom > 0) ? (v * v) / denom : 0.0;

    if (power > 0)
        m_powerLabel->setText(QString("%1W").arg(power, 0, 'f', 2));
    else
        m_powerLabel->setText("");
}

SensorySample SampleCard::toSample() const
{
    SensorySample s;
    s.name     = m_nameEdit->text().trimmed();
    s.comments = m_commentsEdit->toPlainText().trimmed();
    for (auto it = m_spinBoxes.constBegin(); it != m_spinBoxes.constEnd(); ++it) {
        s.scores[it.key()] = it.value()->value();
    }
    s.voltage           = m_voltageEdit->text().toDouble();
    s.resistance        = m_resistanceEdit->text().toDouble();
    s.heatingTechnology = m_heatingTechCombo->currentText().trimmed();
    s.powerType         = m_powerTypeCombo->currentText();
    s.puffLengthSec     = m_puffLengthSpin->value();

    // Compute power
    double rOffset = 0.0;
    QString tech = s.heatingTechnology.trimmed().toUpper();
    if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
        rOffset = 0.78;
    else if (tech == "T51")
        rOffset = 0.25;
    double denom = s.resistance + rOffset;
    s.power = (s.voltage > 0 && denom > 0) ? (s.voltage * s.voltage) / denom : 0.0;

    return s;
}

void SampleCard::fromSample(const SensorySample& s)
{
    m_nameEdit->setText(s.name);
    m_commentsEdit->setPlainText(s.comments);
    for (auto it = m_spinBoxes.constBegin(); it != m_spinBoxes.constEnd(); ++it) {
        if (s.scores.contains(it.key())) {
            it.value()->blockSignals(true);
            it.value()->setValue(s.scores.value(it.key(), 5.0));
            it.value()->blockSignals(false);
        }
    }

    // Per-sample device properties
    m_voltageEdit->blockSignals(true);
    m_resistanceEdit->blockSignals(true);
    m_heatingTechCombo->blockSignals(true);

    m_voltageEdit->setText(s.voltage > 0 ? QString::number(s.voltage, 'f', 2) : QString());
    m_resistanceEdit->setText(s.resistance > 0 ? QString::number(s.resistance, 'f', 3) : QString());
    int htIdx = m_heatingTechCombo->findText(s.heatingTechnology, Qt::MatchFixedString);
    if (htIdx >= 0)
        m_heatingTechCombo->setCurrentIndex(htIdx);
    else
        m_heatingTechCombo->setCurrentText(s.heatingTechnology);

    m_voltageEdit->blockSignals(false);
    m_resistanceEdit->blockSignals(false);
    m_heatingTechCombo->blockSignals(false);

    // #7: per-sample test conditions. Default-aware: tolerate older JSON that
    // didn't carry these fields (deserializer returns "" / 0.0 in that case).
    m_powerTypeCombo->blockSignals(true);
    m_puffLengthSpin->blockSignals(true);
    int ptIdx = m_powerTypeCombo->findText(s.powerType);
    m_powerTypeCombo->setCurrentIndex(ptIdx >= 0 ? ptIdx : 0);
    m_puffLengthSpin->setValue(s.puffLengthSec > 0 ? s.puffLengthSec : 3.0);
    m_powerTypeCombo->blockSignals(false);
    m_puffLengthSpin->blockSignals(false);

    recalcPower();
}

// ─────────────────────────────────────────────────────────────────────────────
// SensoryPanel
// ─────────────────────────────────────────────────────────────────────────────

SensoryPanel::SensoryPanel(DatabaseManager* db, QWidget* parent)
    : QWidget(parent)
    , m_db(db)
{
    m_refreshTimer = new QTimer(this);
    m_refreshTimer->setSingleShot(true);
    m_refreshTimer->setInterval(150);
    connect(m_refreshTimer, &QTimer::timeout, this, &SensoryPanel::onRefreshChart);

    auto* outerLayout = new QVBoxLayout(this);
    outerLayout->setSpacing(4);
    outerLayout->setContentsMargins(4, 4, 4, 4);

    // ── Header row (test title, assessor, tester, media, date) ──
    auto* headerWidget = new QWidget;
    buildHeaderRow(headerWidget);
    outerLayout->addWidget(headerWidget);

    // ── Horizontal splitter: cards (left) + radar chart (right) ──
    m_splitter = new QSplitter(Qt::Horizontal);

    // Left: scroll area with flow layout for sample cards + Add button
    auto* cardsPanel = new QWidget;
    auto* cardsLayout = new QVBoxLayout(cardsPanel);
    cardsLayout->setContentsMargins(0, 0, 0, 0);
    cardsLayout->setSpacing(4);

    m_scrollArea = new QScrollArea;
    m_scrollArea->setWidgetResizable(true);
    m_scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);

    m_flowContainer = new QWidget;
    m_flowLayout    = new FlowLayout(m_flowContainer, 4, 6, 6);
    m_scrollArea->setWidget(m_flowContainer);
    cardsLayout->addWidget(m_scrollArea, 1);

    auto* addBtn = new QPushButton("+ Add Sample");
    cardsLayout->addWidget(addBtn);
    connect(addBtn, &QPushButton::clicked, this, &SensoryPanel::onAddSample);

    // Averaged table overlay (shown when test avg selected, replaces cards).
    // Clicking a column header sorts by that column: alphabetical for the
    // Device column (plain QTableWidgetItem default), numeric for every
    // metric column (NumericTableItem overrides operator< — see below).
    m_avgOverlayTable = new QTableWidget;
    m_avgOverlayTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    m_avgOverlayTable->setAlternatingRowColors(true);
    m_avgOverlayTable->setSelectionMode(QAbstractItemView::NoSelection);
    m_avgOverlayTable->setSortingEnabled(true);
    m_avgOverlayTable->horizontalHeader()->setSectionsClickable(true);
    m_avgOverlayTable->verticalHeader()->setVisible(false);
    m_avgOverlayTable->verticalHeader()->setDefaultSectionSize(22);
    m_avgOverlayTable->setStyleSheet(
        "QTableWidget { font-size: 9pt; }"
        "QTableWidget::item { padding: 2px 4px; }"
        "QHeaderView::section { background: #1F4E79; color: white; font-weight: 600;"
        "  font-size: 8pt; padding: 2px 4px; border: none; }");

    // Stack: index 0 = cards panel, index 1 = averaged table
    m_leftStack = new QStackedWidget;
    m_leftStack->addWidget(cardsPanel);
    m_leftStack->addWidget(m_avgOverlayTable);
    m_leftStack->setCurrentIndex(0);

    m_splitter->addWidget(m_leftStack);

    // Right: radar chart fills remaining space
    m_chart = new RadarChartWidget;
    m_splitter->addWidget(m_chart);

    m_splitter->setStretchFactor(0, 50);   // cards: 50%
    m_splitter->setStretchFactor(1, 50);  // chart: 50%
    m_splitter->setSizes({10000, 10000}); // force equal initial split
    m_splitter->setChildrenCollapsible(false);

    outerLayout->addWidget(m_splitter, 1);

    // Start with one empty session
    SensorySession empty;
    empty.sessionName = QStringLiteral("New Session");
    empty.date = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    empty.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    m_sessions.append(empty);
    m_currentTesterIdx = 0;
    addSampleCard();
}

void SensoryPanel::buildHeaderRow(QWidget* container)
{
    auto* layout = new QHBoxLayout(container);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setSpacing(10);

    auto addField = [&](const QString& label, QLineEdit*& edit) {
        layout->addWidget(new QLabel(label));
        edit = new QLineEdit;
        edit->setFixedWidth(150);
        layout->addWidget(edit);
    };

    addField("Test Title:", m_testTitleEdit);
    addField("Assessor:",  m_assessorEdit);
    addField("Tester:",    m_testerEdit);
    addField("Media:",     m_mediaEdit);

    layout->addWidget(new QLabel("Date:"));
    m_dateLabel = new QLabel(QDateTime::currentDateTime().toString("yyyy-MM-dd"));
    layout->addWidget(m_dateLabel);

    auto* saveChartBtn = new QPushButton;
    saveChartBtn->setIcon(style()->standardIcon(QStyle::SP_DialogSaveButton));
    saveChartBtn->setToolTip("Save chart as image");
    saveChartBtn->setFixedSize(24, 24);
    saveChartBtn->setFlat(true);
    connect(saveChartBtn, &QPushButton::clicked, this, &SensoryPanel::onSaveChart);
    layout->addWidget(saveChartBtn);

    layout->addStretch();
}

// ─────────────────────────────────────────────────────────────────────────────
// Card management
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::addSampleCard(const SensorySample& sample)
{
    int idx = m_cards.size();
    auto* card = new SampleCard(idx, m_flowContainer);
    if (!sample.name.isEmpty() || !sample.scores.isEmpty()) {
        card->fromSample(sample);
    }
    connect(card, &SampleCard::changed,          this, &SensoryPanel::scheduleChartRefresh);
    connect(card, &SampleCard::removeRequested,  this, &SensoryPanel::onRemoveCard);

    m_flowLayout->addWidget(card);
    m_cards.append(card);

    scheduleChartRefresh();
}

void SensoryPanel::clearAllCards()
{
    for (SampleCard* card : m_cards) {
        m_flowLayout->removeWidget(card);
        card->deleteLater();
    }
    m_cards.clear();
}

// ─────────────────────────────────────────────────────────────────────────────
// Session serialisation
// ─────────────────────────────────────────────────────────────────────────────

SensorySession SensoryPanel::buildSession() const
{
    SensorySession sess;
    // Preserve existing sessionName for DB upsert matching; only generate new if empty
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()
        && !m_sessions[m_currentTesterIdx].sessionName.isEmpty()) {
        sess.sessionName = m_sessions[m_currentTesterIdx].sessionName;
    } else {
        sess.sessionName = QString("Session_%1").arg(
            QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss"));
    }
    sess.testTitle    = m_testTitleEdit->text().trimmed();
    sess.assessorName = m_assessorEdit->text().trimmed();
    sess.testerName   = m_testerEdit->text().trimmed();
    sess.media        = m_mediaEdit->text().trimmed();
    sess.date         = m_dateLabel->text();
    sess.timestamp    = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    sess.samples.reserve(m_cards.size());
    for (const SampleCard* card : m_cards) {
        sess.samples.append(card->toSample());
    }

    // Carry over fields not represented by SensoryPanel UI widgets
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        const SensorySession& stored = m_sessions[m_currentTesterIdx];

        // Preserve DB primary key + version. Without this, navigating between
        // sessions clobbers id back to -1, the next save attempts INSERT, and
        // the (session_name, tester_name, date) UNIQUE constraint trips
        // UniqueViolationDialog on every selection change. DetailedSensoryPanel
        // already does this implicitly via `sess = m_sessions[idx]`.
        sess.id      = stored.id;
        sess.version = stored.version;

        sess.imagePaths         = stored.imagePaths;
        sess.imageLayouts       = stored.imageLayouts;
        sess.imageCrops         = stored.imageCrops;
        sess.sourceImagePath    = stored.sourceImagePath;

        // Session-level test properties (edited via MainWindow props panel)
        sess.control            = stored.control;
        sess.isBlind            = stored.isBlind;
        sess.primaryDifferences = stored.primaryDifferences;

        // Legacy fields (preserved for backward compat)
        sess.puffLength         = stored.puffLength;
        sess.burnStatus         = stored.burnStatus;
        sess.clogStatus         = stored.clogStatus;
        sess.leakStatus         = stored.leakStatus;
        sess.resistance         = stored.resistance;
        sess.voltage            = stored.voltage;
        sess.power              = stored.power;
        sess.heatingTechnology  = stored.heatingTechnology;
    }

    return sess;
}

void SensoryPanel::applySession(const SensorySession& session)
{
    clearAllCards();
    m_testTitleEdit->setText(session.testTitle);
    m_assessorEdit->setText(session.assessorName);
    m_testerEdit->setText(session.testerName);
    m_mediaEdit->setText(session.media);
    if (!session.date.isEmpty()) m_dateLabel->setText(session.date);

    for (const SensorySample& s : session.samples) {
        addSampleCard(s);
    }

    onRefreshChart();
}

void SensoryPanel::saveCurrentTester()
{
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        m_sessions[m_currentTesterIdx] = buildSession();
    }
}

void SensoryPanel::inheritExistingIdsAndVersions()
{
    // Reverted on 2026-05-14: this helper triggered a heap-corruption /
    // access-violation crash when sensory sessions were loaded from the
    // database (root cause not yet pinned down — possibly an interaction
    // with implicitly-shared QString state inside SensoryRecord). The
    // primary buildSession fix on this branch already covers the
    // dominant case (navigation between in-memory sessions). The remaining
    // case — re-importing a JSON/Excel file whose natural key matches an
    // existing DB row — falls back to SaveCoordinator's standard
    // UniqueViolationDialog, which is at least correct (just clicky).
    // Restore this body once the underlying crash is diagnosed.
}

// ─────────────────────────────────────────────────────────────────────────────
// Session management (called by MainWindow)
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::selectSession(int index)
{
    if (index < 0 || index >= m_sessions.size()) return;
    if (index == m_currentTesterIdx) return;

    saveCurrentTester();
    m_currentTesterIdx = index;
    applySession(m_sessions[index]);
}

void SensoryPanel::showAveragedChart(const QVector<int>& sessionIndices)
{
    if (sessionIndices.isEmpty()) return;

    // Build a synthetic session with per-device averaged scores
    // across all selected sessions
    saveCurrentTester();

    struct DeviceAccum { QMap<QString, double> sums; int count = 0; };
    QMap<QString, DeviceAccum> deviceMap;
    QStringList deviceOrder;

    for (int si : sessionIndices) {
        if (si < 0 || si >= m_sessions.size()) continue;
        const SensorySession& sess = m_sessions[si];
        for (const SensorySample& s : sess.samples) {
            QString key = s.name.isEmpty() ? QStringLiteral("Sample") : s.name;
            if (!deviceMap.contains(key)) deviceOrder.append(key);
            DeviceAccum& acc = deviceMap[key];
            for (const QString& m : kSensoryMetrics)
                acc.sums[m] += s.scores.value(m, 5.0);
            acc.count++;
        }
    }

    SensorySession avgSess;
    avgSess.testTitle = "Averaged";
    for (const QString& devName : deviceOrder) {
        const DeviceAccum& acc = deviceMap[devName];
        SensorySample avgSample;
        avgSample.name = devName;
        for (const QString& m : kSensoryMetrics)
            avgSample.scores[m] = acc.sums.value(m, 0) / qMax(1, acc.count);
        avgSess.samples.append(avgSample);
    }

    m_chart->setSessions({avgSess});
}

void SensoryPanel::showAveragedTable(const QStringList& deviceNames,
                                      const QVector<QMap<QString, double>>& deviceAvgs)
{
    if (!m_leftStack || !m_avgOverlayTable) return;

    int nMetrics = kSensoryMetrics.size();
    m_avgOverlayTable->setColumnCount(1 + nMetrics);
    QStringList headers;
    headers << "Device";
    for (const QString& m : kSensoryMetrics)
        headers << m;
    m_avgOverlayTable->setHorizontalHeaderLabels(headers);
    m_avgOverlayTable->horizontalHeader()->setStretchLastSection(false);
    m_avgOverlayTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);

    // Distribute column widths: device gets 1.5x share, metrics share the rest
    int totalWidth = m_avgOverlayTable->viewport()->width();
    if (totalWidth < 100) totalWidth = 600;  // fallback before first layout
    double totalShares = 1.5 + nMetrics;
    int metricW = static_cast<int>(totalWidth / totalShares);
    int deviceW = static_cast<int>(metricW * 1.5);
    m_avgOverlayTable->setColumnWidth(0, deviceW);
    for (int c = 1; c <= nMetrics; ++c)
        m_avgOverlayTable->setColumnWidth(c, metricW);

    // Disable sorting while we mutate row contents so Qt doesn't reshuffle
    // mid-fill (when the user has a sort column active and rows get inserted
    // out of order). Re-enable afterwards so header clicks still sort.
    const bool wasSorting = m_avgOverlayTable->isSortingEnabled();
    m_avgOverlayTable->setSortingEnabled(false);

    m_avgOverlayTable->setRowCount(deviceNames.size());
    for (int i = 0; i < deviceNames.size(); ++i) {
        auto* nameItem = new QTableWidgetItem(deviceNames[i]);
        nameItem->setFlags(Qt::ItemIsEnabled);
        QFont f = nameItem->font(); f.setBold(true); nameItem->setFont(f);
        m_avgOverlayTable->setItem(i, 0, nameItem);

        const QMap<QString, double>& avgs = deviceAvgs[i];
        for (int c = 0; c < nMetrics; ++c) {
            double avg = avgs.value(kSensoryMetrics[c], 0.0);
            auto* valItem = new NumericTableItem(avg);
            valItem->setFlags(Qt::ItemIsEnabled);
            m_avgOverlayTable->setItem(i, 1 + c, valItem);
        }
    }

    m_avgOverlayTable->setSortingEnabled(wasSorting);
    m_leftStack->setCurrentIndex(1);
}

void SensoryPanel::showNormalView()
{
    if (m_leftStack)
        m_leftStack->setCurrentIndex(0);
}

void SensoryPanel::newSession()
{
    saveCurrentTester();

    SensorySession empty;
    empty.sessionName = QStringLiteral("New Session");
    empty.date = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    empty.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    m_sessions.append(empty);

    m_currentTesterIdx = m_sessions.size() - 1;

    clearAllCards();
    m_testTitleEdit->clear();
    m_assessorEdit->clear();
    m_testerEdit->clear();
    m_mediaEdit->clear();
    m_dateLabel->setText(empty.date);
    addSampleCard();

    emit sessionsChanged();
}

void SensoryPanel::closeSessions(const QVector<int>& indices)
{
    if (indices.isEmpty()) return;

    // Sort descending so removal doesn't shift later indices
    QVector<int> sorted = indices;
    std::sort(sorted.begin(), sorted.end(), std::greater<int>());

    for (int idx : sorted) {
        if (idx < 0 || idx >= m_sessions.size()) continue;
        m_sessions.removeAt(idx);
    }

    if (m_sessions.isEmpty()) {
        // No sessions remain — show empty panels
        m_currentTesterIdx = -1;
        clearAllCards();
        m_testTitleEdit->clear();
        m_assessorEdit->clear();
        m_testerEdit->clear();
        m_mediaEdit->clear();
        m_dateLabel->clear();
        m_chart->setSessions({});
    } else {
        // Clamp current index and display that session
        m_currentTesterIdx = qBound(0, m_currentTesterIdx, m_sessions.size() - 1);
        applySession(m_sessions[m_currentTesterIdx]);
    }

    emit sessionsChanged();
}

void SensoryPanel::renameSession(int index, const QString& newLabel)
{
    if (index < 0 || index >= m_sessions.size()) return;
    // sessionLabel() formats as "testTitle - testerName". We update testTitle only.
    // Split at the FIRST " - " to avoid consuming testerName.
    // Known limitation: titles containing " - " will be split incorrectly.
    QString title = newLabel;
    QString tester = m_sessions[index].testerName;
    int sep = newLabel.indexOf(" - ");
    if (sep >= 0 && !tester.isEmpty()) {
        title = newLabel.left(sep).trimmed();
    }
    m_sessions[index].testTitle = title;

    // If renaming the current session, update the header field too
    if (index == m_currentTesterIdx)
        m_testTitleEdit->setText(title);

    emit sessionsChanged();
}

void SensoryPanel::loadSessions(const QVector<SensorySession>& sessions)
{
    saveCurrentTester();

    // Remove the initial empty session if it's still in default state
    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    for (const SensorySession& s : sessions) {
        if (!s.samples.isEmpty())
            m_sessions.append(s);
    }

    if (!m_sessions.isEmpty()) {
        m_currentTesterIdx = m_sessions.size() - 1;
        applySession(m_sessions[m_currentTesterIdx]);
    }

    emit sessionsChanged();
}

QVector<SensorySession> SensoryPanel::allSessions()
{
    saveCurrentTester();
    return m_sessions;
}

SensorySession* SensoryPanel::currentSession()
{
    saveCurrentTester();
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size())
        return &m_sessions[m_currentTesterIdx];
    return nullptr;
}

QString SensoryPanel::sessionLabel(const SensorySession& s) const
{
    QString title  = s.testTitle.isEmpty() ? QString() : s.testTitle;
    QString tester = s.testerName.isEmpty() ? s.assessorName : s.testerName;

    if (!title.isEmpty() && !tester.isEmpty())
        return title + " - " + tester;
    if (!title.isEmpty())
        return title;
    if (!tester.isEmpty())
        return tester;
    return s.sessionName.isEmpty() ? QStringLiteral("(unnamed)") : s.sessionName;
}

// ─────────────────────────────────────────────────────────────────────────────
// Chart
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::scheduleChartRefresh()
{
    m_refreshTimer->start();
}

void SensoryPanel::onRefreshChart()
{
    SensorySession sess = buildSession();
    m_chart->setSessions({sess});
}

void SensoryPanel::onAddSample()
{
    addSampleCard();
}

void SensoryPanel::onRemoveCard(SampleCard* card)
{
    m_flowLayout->removeWidget(card);
    m_cards.removeOne(card);
    card->deleteLater();
    scheduleChartRefresh();
}

void SensoryPanel::onSaveChart()
{
    if (!m_chart) return;

    QString path = QFileDialog::getSaveFileName(
        this, "Save Chart Image", lastBrowseDir(),
        "PNG Image (*.png);;JPEG Image (*.jpg);;BMP Image (*.bmp)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    QPixmap pixmap(m_chart->size());
    pixmap.fill(Qt::white);
    m_chart->render(&pixmap);

    if (!pixmap.save(path)) {
        QMessageBox::warning(this, "Save Chart",
            "Failed to save chart image to:\n" + path);
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Last browse directory
// ─────────────────────────────────────────────────────────────────────────────

QString SensoryPanel::lastBrowseDir() const
{
    if (!m_lastBrowseDir.isEmpty() && QDir(m_lastBrowseDir).exists())
        return m_lastBrowseDir;

    const QString weeklyReports =
        "C:/Users/S1134987/OneDrive - Shenzhen Smoore Technology Limited"
        "/Shared Files Between Computers/Weekly_Reports_Transfer";
    if (QDir(weeklyReports).exists())
        return weeklyReports;

    return QDir::homePath() + "/Documents";
}

void SensoryPanel::setLastBrowseDir(const QString& filePath)
{
    m_lastBrowseDir = QFileInfo(filePath).absolutePath();
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

bool SensoryPanel::isDefaultState() const
{
    if (m_cards.size() != 1) return false;
    if (!m_testTitleEdit->text().trimmed().isEmpty()) return false;
    if (!m_assessorEdit->text().trimmed().isEmpty()) return false;
    if (!m_testerEdit->text().trimmed().isEmpty()) return false;
    if (!m_mediaEdit->text().trimmed().isEmpty()) return false;

    SensorySample s = m_cards.first()->toSample();
    if (!s.name.isEmpty()) return false;
    if (!s.comments.isEmpty()) return false;
    for (auto it = s.scores.constBegin(); it != s.scores.constEnd(); ++it) {
        if (qAbs(it.value() - 5.0) > 0.01) return false;
    }
    return true;
}

QString SensoryPanel::resolveTestName()
{
    QString testName = m_testTitleEdit->text().trimmed();
    if (!testName.isEmpty()) return testName;

    QString previousName;
    for (const SensorySession& s : m_sessions) {
        if (!s.testTitle.isEmpty()) {
            previousName = s.testTitle;
            break;
        }
    }

    if (!previousName.isEmpty()) {
        auto answer = QMessageBox::question(this, "Test Name Blank",
            "Test name is blank — use same test name as previous?\n\n\"" + previousName + "\"",
            QMessageBox::Yes | QMessageBox::No);
        if (answer == QMessageBox::Yes) {
            m_testTitleEdit->setText(previousName);
            return previousName;
        }
        return QString();
    }

    QString defaultName = m_db ? m_db->nextDefaultTestName() : QStringLiteral("test_0001");
    auto answer = QMessageBox::question(this, "No Test Name",
        "No test name — set default test name?\n\n\"" + defaultName + "\"",
        QMessageBox::Yes | QMessageBox::No);
    if (answer == QMessageBox::Yes) {
        m_testTitleEdit->setText(defaultName);
        return defaultName;
    }
    return QString();
}

// ─────────────────────────────────────────────────────────────────────────────
// Save
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::save()
{
    if (m_cards.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Add at least one sample before saving.");
        return;
    }

    QString testName = resolveTestName();
    if (testName.isEmpty() && m_testTitleEdit->text().trimmed().isEmpty())
        return;

    // Flush current UI state into m_sessions
    saveCurrentTester();

    // Show Save As dialog if no prior save, or if the session has changed
    QString expectedBase = sessionLabel(buildSession());
    bool needsSaveAs = m_savePath.isEmpty()
        || QFileInfo(m_savePath).fileName() != expectedBase;

    if (needsSaveAs) {
        QString suggested = lastBrowseDir() + "/" + expectedBase;
        QString path = QFileDialog::getSaveFileName(
            this, "Save Sensory Session", suggested,
            "Excel files (*.xlsx);;All files (*)");
        if (path.isEmpty()) return;
        setLastBrowseDir(path);
        if (path.endsWith(".xlsx", Qt::CaseInsensitive))
            path.chop(5);
        else if (path.endsWith(".json", Qt::CaseInsensitive))
            path.chop(5);
        m_savePath = path;
    }

    // Save current session's JSON/Excel files
    SensorySession sess = buildSession();
    saveToJson(m_savePath + ".json", sess);
    saveToExcel(m_savePath + ".xlsx", sess);

    // Save ALL sessions to the database (not just the current one).
    // Iterate by index so the coordinator can mutate version/id in-place.
    int dbSaved = 0;
    if (m_saveCoord) {
        for (int i = 0; i < m_sessions.size(); ++i) {
            if (m_sessions[i].samples.isEmpty()) continue;
            if (m_saveCoord->saveSensorySession(m_sessions[i], this)
                == SaveCoordinator::Saved)
                ++dbSaved;
        }
    } else if (m_db) {
        for (const SensorySession& s : m_sessions) {
            if (s.samples.isEmpty()) continue;
            if (m_db->saveSensorySession(s))
                ++dbSaved;
        }
    }

    emit sessionsChanged();

    {
        QMessageBox msg(QMessageBox::Information, "Saved",
            QString("Session saved to:\n%1.json\n%2.xlsx\nDatabase (%3 session%4)")
                .arg(QFileInfo(m_savePath).fileName(),
                     QFileInfo(m_savePath).fileName(),
                     QString::number(dbSaved),
                     dbSaved == 1 ? "" : "s"),
            QMessageBox::Ok, this);
        // Remove the extra spacing Qt adds between icon and text
        if (auto* grid = qobject_cast<QGridLayout*>(msg.layout())) {
            grid->setHorizontalSpacing(8);
            grid->setContentsMargins(10, 10, 10, 10);
        }
        msg.exec();
    }
}

void SensoryPanel::saveToJson(const QString& path, const SensorySession& sess)
{
    // Routes through the canonical pipeline-layer encoder so the on-disk
    // .json wire format stays byte-identical to the Postgres JSONB blob
    // and the offline-snapshot copy. Any new SensorySession field added
    // there is automatically included here.
    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Text)) return;
    f.write(QJsonDocument(sensorySessionToJson(sess)).toJson(QJsonDocument::Indented));
}

void SensoryPanel::saveToExcel(const QString& path, const SensorySession& sess)
{
    QXlsx::Document xlsx;

    // Write header row once with bold formatting (avoids read-back cycle).
    QXlsx::Format hdrFmt;
    hdrFmt.setFontBold(true);
    xlsx.write(1, 1, QStringLiteral("Sample"), hdrFmt);
    for (int i = 0; i < kSensoryMetrics.size(); ++i)
        xlsx.write(1, i + 2, kSensoryMetrics[i], hdrFmt);
    xlsx.write(1, kSensoryMetrics.size() + 2, QStringLiteral("Comments"), hdrFmt);

    int row = 2;
    for (const SensorySample& s : sess.samples) {
        xlsx.write(row, 1, s.name.isEmpty() ? QString("Sample %1").arg(row - 1) : s.name);
        for (int i = 0; i < kSensoryMetrics.size(); ++i)
            xlsx.write(row, i + 2, s.scores.value(kSensoryMetrics[i], 5.0));
        xlsx.write(row, kSensoryMetrics.size() + 2, s.comments);
        ++row;
    }

    row += 2;
    xlsx.write(row, 1, "Test Title");  xlsx.write(row, 2, sess.testTitle);   ++row;
    xlsx.write(row, 1, "Tester");      xlsx.write(row, 2, sess.testerName);  ++row;
    xlsx.write(row, 1, "Assessor");    xlsx.write(row, 2, sess.assessorName);++row;
    xlsx.write(row, 1, "Media");       xlsx.write(row, 2, sess.media);       ++row;
    xlsx.write(row, 1, "Date");        xlsx.write(row, 2, sess.date);        ++row;
    xlsx.write(row, 1, "Control");     xlsx.write(row, 2, sess.control);     ++row;
    xlsx.write(row, 1, "Blind?");      xlsx.write(row, 2, sess.isBlind ? "Y" : "N"); ++row;
    xlsx.write(row, 1, "Primary Difference(s)"); xlsx.write(row, 2, sess.primaryDifferences); ++row;

    xlsx.saveAs(path);
}

// ─────────────────────────────────────────────────────────────────────────────
// Load Excel
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::loadFile(const QString& path)
{
    saveCurrentTester();

    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    int loaded = 0;
    QString ext = QFileInfo(path).suffix().toLower();

    if (ext == "json") {
        QFile f(path);
        if (!f.open(QIODevice::ReadOnly | QIODevice::Text)) return;
        QJsonDocument doc = QJsonDocument::fromJson(f.readAll());
        if (doc.isNull() || !doc.isObject()) return;

        // Route through the canonical decoder so this user-facing import
        // accepts the same wire format as the DB JSONB column and the
        // offline-snapshot copy. Backward-compat defaults for new keys
        // (power_type / puff_length_sec) are handled there.
        SensorySession sess = sensorySessionFromJson(doc.object());

        if (sess.sessionName.isEmpty())
            sess.sessionName = QFileInfo(path).baseName();

        if (!sess.samples.isEmpty()) {
            m_sessions.append(sess);
            ++loaded;
        }
    } else {
        QXlsx::Document xlsx(path);
        if (!xlsx.load()) return;

        QString testTitle = QFileInfo(path).baseName();
        QStringList excelMetrics;
        for (int col = 2; col <= 6; ++col) {
            QString hdr = xlsx.read(1, col).toString().trimmed();
            if (!hdr.isEmpty()) excelMetrics << hdr;
        }
        if (excelMetrics.size() != 5)
            excelMetrics = QStringList(kSensoryMetrics.begin(), kSensoryMetrics.end());

        QString a1 = xlsx.read(1, 1).toString().trimmed();
        bool isSavedFormat = (a1.toLower() == "sample");

        if (isSavedFormat)
            loaded += loadExcelSavedFormat(xlsx, excelMetrics, testTitle, path);
        else
            loaded += loadExcelStandardFormat(xlsx, excelMetrics, testTitle);
    }

    if (loaded == 0) return;

    if (m_saveCoord) {
        for (int i = 0; i < m_sessions.size(); ++i) {
            if (!m_sessions[i].samples.isEmpty())
                m_saveCoord->saveSensorySession(m_sessions[i], this);
        }
    } else if (m_db) {
        for (const SensorySession& s : m_sessions) {
            if (!s.samples.isEmpty())
                m_db->saveSensorySession(s);
        }
    }

    m_currentTesterIdx = m_sessions.size() - 1;
    applySession(m_sessions[m_currentTesterIdx]);
    emit sessionsChanged();
}

void SensoryPanel::loadFiles()
{
    QStringList paths = QFileDialog::getOpenFileNames(
        this, "Load Sensory Files", lastBrowseDir(),
        "Sensory files (*.xlsx *.json);;Excel files (*.xlsx);;JSON files (*.json);;All files (*)");
    if (paths.isEmpty()) return;
    setLastBrowseDir(paths.first());

    saveCurrentTester();

    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    int loaded = 0;
    for (const QString& path : paths) {
        QString ext = QFileInfo(path).suffix().toLower();

        if (ext == "json") {
            // ── Load JSON file ──────────────────────────────────────────
            QFile f(path);
            if (!f.open(QIODevice::ReadOnly | QIODevice::Text)) {
                QMessageBox::warning(this, "Load Error",
                                     "Could not open JSON file:\n" + path);
                continue;
            }
            QJsonDocument doc = QJsonDocument::fromJson(f.readAll());
            if (doc.isNull() || !doc.isObject()) {
                QMessageBox::warning(this, "Parse Error",
                                     "Invalid JSON format:\n" + path);
                continue;
            }

            // Route through the canonical decoder so this user-facing import
            // accepts the same wire format as the DB JSONB column and the
            // offline-snapshot copy.
            SensorySession sess = sensorySessionFromJson(doc.object());

            if (sess.sessionName.isEmpty())
                sess.sessionName = QFileInfo(path).baseName();

            if (!sess.samples.isEmpty()) {
                m_sessions.append(sess);
                ++loaded;
            }
        } else {
            // ── Load Excel file ─────────────────────────────────────────
            QXlsx::Document xlsx(path);
            if (!xlsx.load()) {
                QMessageBox::warning(this, "Load Error",
                                     "Could not open Excel file:\n" + path);
                continue;
            }

            QString testTitle = QFileInfo(path).baseName();
            QString dateNow   = QDateTime::currentDateTime().toString("yyyy-MM-dd");
            QString tsNow     = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

            // Read metric headers from columns B-F (row 1)
            QStringList excelMetrics;
            for (int col = 2; col <= 6; ++col) {
                QString hdr = xlsx.read(1, col).toString().trimmed();
                if (!hdr.isEmpty()) excelMetrics << hdr;
            }
            if (excelMetrics.size() != 5)
                excelMetrics = QStringList(kSensoryMetrics.begin(), kSensoryMetrics.end());

            // ── Format detection ────────────────────────────────────────
            // Saved format: A1 == "Sample", flat table, metadata below
            // Standardized template: A1 == device name, multi-tester blocks
            QString a1 = xlsx.read(1, 1).toString().trimmed();
            bool isSavedFormat = (a1.toLower() == "sample");

            if (isSavedFormat) {
                // ── Parse saved format (single-tester, flat table) ──────
                loaded += loadExcelSavedFormat(xlsx, excelMetrics, testTitle, path);
            } else {
                // ── Parse standardized template (multi-tester blocks) ───
                loaded += loadExcelStandardFormat(xlsx, excelMetrics, testTitle);
            }
        }
    }

    if (loaded == 0) {
        QMessageBox::warning(this, "No Data",
                             "No sample data found in the selected file(s).");
        return;
    }

    inheritExistingIdsAndVersions();

    // Immediately save all loaded sessions to the database
    if (m_saveCoord) {
        for (int i = 0; i < m_sessions.size(); ++i) {
            if (!m_sessions[i].samples.isEmpty())
                m_saveCoord->saveSensorySession(m_sessions[i], this);
        }
    } else if (m_db) {
        for (const SensorySession& s : m_sessions) {
            if (!s.samples.isEmpty())
                m_db->saveSensorySession(s);
        }
    }

    m_currentTesterIdx = m_sessions.size() - 1;
    applySession(m_sessions[m_currentTesterIdx]);
    emit sessionsChanged();
}

// ─────────────────────────────────────────────────────────────────────────────
// Load Excel — saved format (single tester, flat "Sample | Metrics | Comments")
// ─────────────────────────────────────────────────────────────────────────────

int SensoryPanel::loadExcelSavedFormat(QXlsx::Document& xlsx,
                                        const QStringList& excelMetrics,
                                        const QString& testTitle,
                                        const QString& filePath)
{
    QString dateNow = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    QString tsNow   = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    int nMetrics    = excelMetrics.size();

    // Read sample rows (starting at row 2, until empty row)
    QVector<SensorySample> samples;
    for (int row = 2; row <= 100; ++row) {
        QString sampleName = xlsx.read(row, 1).toString().trimmed();
        if (sampleName.isEmpty()) break;

        SensorySample sample;
        sample.name = sampleName;
        bool hasData = false;
        for (int col = 2; col <= nMetrics + 1; ++col) {
            QVariant val = xlsx.read(row, col);
            if (val.isValid() && !val.toString().trimmed().isEmpty()) {
                double score = qBound(1.0, val.toDouble(), 9.0);
                if (score == 0.0) score = 5.0;
                int metricIdx = col - 2;
                QString metricName = (metricIdx < excelMetrics.size())
                    ? excelMetrics[metricIdx] : kSensoryMetrics[metricIdx];
                for (const QString& km : kSensoryMetrics) {
                    if (km.toLower() == metricName.toLower() ||
                        metricName.toLower().contains(km.toLower().left(6))) {
                        sample.scores[km] = score;
                        hasData = true;
                        break;
                    }
                }
            }
        }
        sample.comments = xlsx.read(row, nMetrics + 2).toString().trimmed();
        for (const QString& metric : kSensoryMetrics) {
            if (!sample.scores.contains(metric))
                sample.scores[metric] = 5.0;
        }
        if (hasData) samples.append(sample);
    }

    if (samples.isEmpty()) return 0;

    // Read metadata section (below data, after a gap)
    QString testerName, assessorName, media, date;
    QString control, primaryDifferences;
    bool isBlind = false;
    for (int row = samples.size() + 3; row <= samples.size() + 16; ++row) {
        QString label = xlsx.read(row, 1).toString().trimmed().toLower();
        QString value = xlsx.read(row, 2).toString().trimmed();
        if (label == "test title")              { /* testTitle already from filename */ }
        else if (label == "tester")               testerName         = value;
        else if (label == "assessor")             assessorName       = value;
        else if (label == "media")                media              = value;
        else if (label == "date")                 date               = value;
        else if (label == "control")              control            = value;
        else if (label == "blind?")               isBlind            = (value.toUpper() == "Y");
        else if (label == "primary difference(s)") primaryDifferences = value;
    }

    SensorySession sess;
    sess.sessionName       = testTitle + (testerName.isEmpty() ? "" : " - " + testerName);
    sess.testTitle         = testTitle;
    sess.testerName        = testerName;
    sess.assessorName      = assessorName;
    sess.media             = media;
    sess.date              = date.isEmpty() ? dateNow : date;
    sess.timestamp         = tsNow;
    sess.control           = control;
    sess.isBlind           = isBlind;
    sess.primaryDifferences = primaryDifferences;
    sess.samples           = samples;

    m_sessions.append(sess);
    m_savePath = filePath;
    return 1;
}

// ─────────────────────────────────────────────────────────────────────────────
// Load Excel — standardized template (multi-tester, device-block structure)
// ─────────────────────────────────────────────────────────────────────────────

int SensoryPanel::loadExcelStandardFormat(QXlsx::Document& xlsx,
                                           const QStringList& excelMetrics,
                                           const QString& testTitle)
{
    QString dateNow = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    QString tsNow   = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    struct PanelistEntry {
        QString testerName;
        QString deviceName;
        QMap<QString, double> scores;
        QString comments;
    };

    QVector<PanelistEntry> allEntries;
    QString currentDevice;
    bool inBlock = false;

    QString a1 = xlsx.read(1, 1).toString().trimmed();
    if (!a1.isEmpty() && !a1.toLower().contains("averages")) {
        currentDevice = a1;
        inBlock = true;
    }

    for (int row = 2; row <= 29; ++row) {
        QString cellA = xlsx.read(row, 1).toString().trimmed();

        bool rowEmpty = cellA.isEmpty();
        if (rowEmpty) {
            bool anyData = false;
            for (int c = 2; c <= 7; ++c) {
                if (!xlsx.read(row, c).toString().trimmed().isEmpty()) {
                    anyData = true;
                    break;
                }
            }
            if (!anyData) {
                inBlock = false;
                continue;
            }
        }

        if (cellA.toLower().contains("average"))
            break;

        if (!inBlock && !cellA.isEmpty()) {
            bool hasScores = false;
            for (int c = 2; c <= 6; ++c) {
                QVariant v = xlsx.read(row, c);
                if (v.isValid() && !v.toString().trimmed().isEmpty()) {
                    bool ok;
                    double val = v.toDouble(&ok);
                    if (ok && val >= 1.0 && val <= 9.0) { hasScores = true; break; }
                }
            }

            if (!hasScores) {
                currentDevice = cellA;
                inBlock = true;
                continue;
            } else {
                inBlock = true;
            }
        }

        if (!inBlock && cellA.isEmpty())
            continue;

        if (!cellA.isEmpty()) {
            PanelistEntry entry;
            entry.testerName = cellA;
            entry.deviceName = currentDevice;

            bool hasData = false;
            for (int col = 2; col <= 6; ++col) {
                QVariant val = xlsx.read(row, col);
                if (val.isValid() && !val.toString().trimmed().isEmpty()) {
                    double score = qBound(1.0, val.toDouble(), 9.0);
                    if (score == 0.0) score = 5.0;
                    int metricIdx = col - 2;
                    QString metricName = (metricIdx < excelMetrics.size())
                        ? excelMetrics[metricIdx] : kSensoryMetrics[metricIdx];
                    for (const QString& km : kSensoryMetrics) {
                        if (km.toLower() == metricName.toLower() ||
                            metricName.toLower().contains(km.toLower().left(6))) {
                            entry.scores[km] = score;
                            hasData = true;
                            break;
                        }
                    }
                }
            }
            entry.comments = xlsx.read(row, 7).toString().trimmed();

            if (hasData)
                allEntries.append(entry);
        }
    }

    QMap<QString, QVector<int>> testerIndices;
    QStringList testerOrder;
    for (int i = 0; i < allEntries.size(); ++i) {
        const QString& tn = allEntries[i].testerName;
        if (!testerIndices.contains(tn)) testerOrder.append(tn);
        testerIndices[tn].append(i);
    }

    int loaded = 0;
    for (const QString& testerName : testerOrder) {
        SensorySession sess;
        sess.sessionName = testTitle + " - " + testerName;
        sess.testTitle   = testTitle;
        sess.testerName  = testerName;
        sess.date        = dateNow;
        sess.timestamp   = tsNow;

        for (int idx : testerIndices[testerName]) {
            const PanelistEntry& e = allEntries[idx];
            SensorySample sample;
            sample.name     = e.deviceName;
            sample.scores   = e.scores;
            sample.comments = e.comments;
            for (const QString& metric : kSensoryMetrics) {
                if (!sample.scores.contains(metric))
                    sample.scores[metric] = 5.0;
            }
            sess.samples.append(sample);
        }

        if (!sess.samples.isEmpty()) {
            m_sessions.append(sess);
            ++loaded;
        }
    }
    return loaded;
}

// ─────────────────────────────────────────────────────────────────────────────
// Load from Database
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::loadFromDatabase()
{
    if (!m_db) {
        QMessageBox::warning(this, "Database", "No database connection available.");
        return;
    }

    QVector<SensoryRecord> records = m_db->listSensoryRecords();
    if (records.isEmpty()) {
        QMessageBox::information(this, "Database", "No sensory sessions found in the database.");
        return;
    }

    QDialog picker(this);
    picker.setWindowTitle("Load from Database");
    picker.setMinimumSize(700, 400);
    picker.resize(750, 500);

    auto* layout = new QVBoxLayout(&picker);

    auto* tree = new QTreeWidget;
    tree->setHeaderLabels({"Test Title / Tester", "Assessor", "Media", "Date", "Samples"});
    tree->setSelectionMode(QAbstractItemView::ExtendedSelection);
    tree->setRootIsDecorated(true);

    QMap<QString, QVector<const SensoryRecord*>> groups;
    QStringList order;
    for (const SensoryRecord& r : records) {
        QString title = r.testTitle.isEmpty() ? QStringLiteral("(untitled)") : r.testTitle;
        if (!groups.contains(title)) order.append(title);
        groups[title].append(&r);
    }

    for (const QString& title : order) {
        const auto& recs = groups[title];
        auto* parent = new QTreeWidgetItem(tree);
        parent->setText(0, title);
        parent->setText(4, QString("%1 tester(s)").arg(recs.size()));
        parent->setData(0, Qt::UserRole, -1);
        QFont f = parent->font(0);
        f.setBold(true);
        parent->setFont(0, f);

        for (const SensoryRecord* r : recs) {
            auto* child = new QTreeWidgetItem(parent);
            child->setText(0, r->testerName.isEmpty() ? r->sessionName : r->testerName);
            child->setText(1, r->assessorName);
            child->setText(2, r->media);
            child->setText(3, r->date);
            child->setText(4, QString::number(r->sampleCount));
            child->setData(0, Qt::UserRole, r->id);
        }

        parent->setExpanded(true);
    }

    tree->header()->setSectionResizeMode(0, QHeaderView::Stretch);
    for (int i = 1; i < 5; ++i)
        tree->header()->setSectionResizeMode(i, QHeaderView::ResizeToContents);

    layout->addWidget(tree, 1);

    auto* btnRow = new QHBoxLayout;
    btnRow->addStretch();
    auto* loadBtn = new QPushButton("Load Selected");
    auto* cancelBtn = new QPushButton("Cancel");
    btnRow->addWidget(loadBtn);
    btnRow->addWidget(cancelBtn);
    layout->addLayout(btnRow);

    connect(cancelBtn, &QPushButton::clicked, &picker, &QDialog::reject);
    connect(loadBtn, &QPushButton::clicked, &picker, &QDialog::accept);
    connect(tree, &QTreeWidget::itemDoubleClicked, &picker, &QDialog::accept);

    if (picker.exec() != QDialog::Accepted) return;

    QVector<int> selectedIds;
    const auto selected = tree->selectedItems();
    for (QTreeWidgetItem* item : selected) {
        int id = item->data(0, Qt::UserRole).toInt();
        if (id > 0) {
            selectedIds.append(id);
        } else if (id == -1) {
            for (int c = 0; c < item->childCount(); ++c) {
                int childId = item->child(c)->data(0, Qt::UserRole).toInt();
                if (childId > 0 && !selectedIds.contains(childId))
                    selectedIds.append(childId);
            }
        }
    }

    if (selectedIds.isEmpty()) return;

    QVector<SensorySession> sessions;
    for (int id : selectedIds) {
        SensorySession sess = m_db->loadSensorySession(id);
        if (!sess.samples.isEmpty())
            sessions.append(sess);
    }

    if (sessions.isEmpty()) {
        QMessageBox::warning(this, "Database", "Could not load the selected session(s).");
        return;
    }

    loadSessions(sessions);
}

// ─────────────────────────────────────────────────────────────────────────────
// Generate Sensory Report
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::generateFullReport()
{
    saveCurrentTester();

    if (m_sessions.isEmpty()) {
        QMessageBox::warning(this, "No Data", "No sessions loaded.");
        return;
    }

    // Show a selection dialog — user can Ctrl+Click or Shift+Click to select
    QDialog picker(this);
    picker.setWindowTitle("Select Sessions for Combined Report");
    picker.setMinimumSize(500, 400);
    picker.resize(550, 450);

    auto* layout = new QVBoxLayout(&picker);
    layout->addWidget(new QLabel("Select sessions to include (Ctrl+Click or Shift+Click):"));

    auto* list = new QListWidget;
    list->setSelectionMode(QAbstractItemView::ExtendedSelection);
    for (int i = 0; i < m_sessions.size(); ++i) {
        auto* item = new QListWidgetItem(sessionLabel(m_sessions[i]));
        item->setData(Qt::UserRole, i);
        list->addItem(item);
        item->setSelected(true);  // select all by default
    }
    layout->addWidget(list, 1);

    auto* btnRow = new QHBoxLayout;
    btnRow->addStretch();
    auto* genBtn = new QPushButton("Generate Report");
    auto* cancelBtn = new QPushButton("Cancel");
    btnRow->addWidget(genBtn);
    btnRow->addWidget(cancelBtn);
    layout->addLayout(btnRow);

    connect(cancelBtn, &QPushButton::clicked, &picker, &QDialog::reject);
    connect(genBtn, &QPushButton::clicked, &picker, &QDialog::accept);

    if (picker.exec() != QDialog::Accepted) return;

    QVector<SensorySession> selected;
    for (QListWidgetItem* item : list->selectedItems()) {
        int idx = item->data(Qt::UserRole).toInt();
        if (idx >= 0 && idx < m_sessions.size())
            selected.append(m_sessions[idx]);
    }

    if (selected.isEmpty()) {
        QMessageBox::warning(this, "No Selection", "No sessions selected.");
        return;
    }

    auto* src = new SensoryReportSource(selected, m_db);
    ReportPreviewDialog dlg(src, this);
    const int rc = dlg.exec();
    delete src;

    // Dialog handles its own QFileDialog + writePptx + success confirmation.
    // Nothing more to do here; just return.
    Q_UNUSED(rc);
}

// ─────────────────────────────────────────────────────────────────────────────
// Generate Statistics Report (.csv)
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::generateStats()
{
    if (m_cards.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Add at least one sample before generating stats.");
        return;
    }

    QString path = QFileDialog::getSaveFileName(
        this, "Save Statistics Report", lastBrowseDir(),
        "CSV files (*.csv);;All files (*)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    writeStatsCsv(path);
}

void SensoryPanel::writeStatsCsv(const QString& path)
{
    SensorySession sess = buildSession();
    int n = sess.samples.size();

    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Text)) {
        QMessageBox::warning(this, "Save Error",
                             "Could not save statistics report:\n" + path);
        return;
    }
    QTextStream out(&f);

    out << "Sensory Evaluation Statistics Report\n";
    if (!sess.testTitle.isEmpty())
        out << "Test Title," << sess.testTitle << "\n";
    out << "Date," << sess.date << "\n";
    out << "Assessor," << sess.assessorName << "\n";
    if (!sess.testerName.isEmpty())
        out << "Tester," << sess.testerName << "\n";
    out << "Media," << sess.media << "\n";
    out << "Sample Count," << n << "\n\n";

    out << "Metric,Sample Size,Mean,Std Dev,Std Error,Min,Max,Range,Outliers,Chi-Square\n";

    for (const QString& metric : kSensoryMetrics) {
        QVector<double> vals;
        double sum = 0;
        for (const SensorySample& s : sess.samples) {
            double v = s.scores.value(metric, 5.0);
            vals.append(v);
            sum += v;
        }

        double mean = (n > 0) ? sum / n : 0;
        double sumSq = 0;
        for (double v : vals) sumSq += (v - mean) * (v - mean);
        double variance = (n > 1) ? sumSq / (n - 1) : 0;
        double stddev = std::sqrt(qMax(0.0, variance));
        double stderr_ = (n > 0) ? stddev / std::sqrt(static_cast<double>(n)) : 0;

        double minVal = 9.0, maxVal = 1.0;
        for (double v : vals) { minVal = qMin(minVal, v); maxVal = qMax(maxVal, v); }

        QStringList outliers;
        for (int i = 0; i < vals.size(); ++i) {
            if (stddev > 0 && std::abs(vals[i] - mean) > 2.0 * stddev) {
                QString sName = (i < sess.samples.size() && !sess.samples[i].name.isEmpty())
                    ? sess.samples[i].name : QString("Sample %1").arg(i + 1);
                outliers << QString("%1(%2)").arg(sName).arg(vals[i], 0, 'f', 1);
            }
        }

        double chiSq = 0;
        double expected = static_cast<double>(n) / 9.0;
        if (expected > 0) {
            // Use rounded values for chi-square (integer bins)
            QMap<int, int> freq;
            for (double v : vals) freq[qRound(v)]++;
            for (int score = 1; score <= 9; ++score) {
                double obs = freq.value(score, 0);
                chiSq += (obs - expected) * (obs - expected) / expected;
            }
        }

        out << "\"" << metric << "\""
            << "," << n
            << "," << QString::number(mean, 'f', 2)
            << "," << QString::number(stddev, 'f', 2)
            << "," << QString::number(stderr_, 'f', 3)
            << "," << QString::number(minVal, 'f', 1)
            << "," << QString::number(maxVal, 'f', 1)
            << "," << QString::number(maxVal - minVal, 'f', 1)
            << "," << (outliers.isEmpty() ? "None" : outliers.join("; "))
            << "," << QString::number(chiSq, 'f', 2)
            << "\n";
    }

    out << "\nRaw Data\n";
    out << "Sample";
    for (const QString& m : kSensoryMetrics) out << "," << "\"" << m << "\"";
    out << ",Comments\n";

    for (const SensorySample& s : sess.samples) {
        out << "\"" << (s.name.isEmpty() ? QString("Sample") : s.name) << "\"";
        for (const QString& m : kSensoryMetrics)
            out << "," << QString::number(s.scores.value(m, 5.0), 'f', 1);
        out << ",\"" << QString(s.comments).replace("\"", "\"\"") << "\"\n";
    }

    f.close();

    QMessageBox::information(this, "Statistics Report Saved",
        "Statistics report saved:\n" + QFileInfo(path).fileName());
}

// ─────────────────────────────────────────────────────────────────────────────
// Combined report: one slide per session (static)
// ─────────────────────────────────────────────────────────────────────────────

bool SensoryPanel::generateCombinedPptx(const QVector<SensorySession>& sessions,
                                         const QString& filePath,
                                         QString& errorOut)
{
    // Thin caller for the legacy fast path. The whole body now lives in
    // SensoryReportSource::writeSensoryPptx so the new IReportSource entry
    // point and this legacy entry point share one renderer.
    //
    // Pass an empty ReportLayout so the 5-arg addContentSlide falls through to
    // the legacy in-line positions (wrap-aware table height, aspect-matched
    // radar). computeDefaultLayout is intentionally NOT used here — it's a
    // simpler shape designed for the preview dialog (square radar; flat
    // 0.50 + N/3 table height; cumulative sized from the FIRST session's
    // sample count) and would override the legacy in-line wrap-aware math via
    // the 5-arg PptxWriter::addContentSlide overload, breaking visual
    // equivalence with the pre-Task-8 output.
    const ReportLayout emptyLayout;
    QString err;
    const bool ok = SensoryReportSource::writeSensoryPptx(
        sessions, emptyLayout, /*excludedSamples=*/{}, filePath, &err);
    if (!ok) errorOut = err;
    return ok;
}

} // namespace DVE
