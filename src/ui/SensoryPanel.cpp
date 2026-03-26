#include "SensoryPanel.h"

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
#include <utility>

#include "xlsxdocument.h"
#include "reporting/PptxWriter.h"
#include "utils/ImageUtils.h"

namespace DVE {

// Ignores wheel events unless the spinbox has keyboard focus.
// Prevents accidental value changes when scrolling past sample cards.
class NoWheelSpinBox : public QSpinBox
{
public:
    explicit NoWheelSpinBox(QWidget* parent = nullptr) : QSpinBox(parent) {
        setFocusPolicy(Qt::StrongFocus);
    }
protected:
    void wheelEvent(QWheelEvent* e) override {
        if (hasFocus()) QSpinBox::wheelEvent(e);
        else e->ignore();
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
            rows.append({rowStart, m_items.size() - rowStart, rowWidth, rowHeight});
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
    setFixedSize(175, 340);

    auto* mainLayout = new QVBoxLayout(this);
    mainLayout->setSpacing(4);
    mainLayout->setContentsMargins(6, 16, 6, 4);

    auto* nameRow = new QHBoxLayout;
    nameRow->setSpacing(3);
    nameRow->addWidget(new QLabel("Name:"));
    m_nameEdit = new QLineEdit;
    m_nameEdit->setPlaceholderText(QString("Sample %1").arg(index + 1));
    nameRow->addWidget(m_nameEdit);
    mainLayout->addLayout(nameRow);

    auto* formLayout = new QFormLayout;
    formLayout->setSpacing(6);
    formLayout->setContentsMargins(0, 2, 0, 2);
    for (const QString& metric : kSensoryMetrics) {
        auto* spin = new NoWheelSpinBox;
        spin->setRange(1, 9);
        spin->setValue(5);
        spin->setFixedWidth(55);
        spin->setFixedHeight(24);
        m_spinBoxes[metric] = spin;
        formLayout->addRow(metric + ":", spin);
        connect(spin, QOverload<int>::of(&QSpinBox::valueChanged),
                this, &SampleCard::changed);
    }
    mainLayout->addLayout(formLayout);

    mainLayout->addWidget(new QLabel("Comments:"));
    m_commentsEdit = new QTextEdit;
    m_commentsEdit->setMinimumHeight(40);
    m_commentsEdit->setMaximumHeight(60);
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

SensorySample SampleCard::toSample() const
{
    SensorySample s;
    s.name     = m_nameEdit->text().trimmed();
    s.comments = m_commentsEdit->toPlainText().trimmed();
    for (auto it = m_spinBoxes.constBegin(); it != m_spinBoxes.constEnd(); ++it) {
        s.scores[it.key()] = it.value()->value();
    }
    return s;
}

void SampleCard::fromSample(const SensorySample& s)
{
    m_nameEdit->setText(s.name);
    m_commentsEdit->setPlainText(s.comments);
    for (auto it = m_spinBoxes.constBegin(); it != m_spinBoxes.constEnd(); ++it) {
        if (s.scores.contains(it.key())) {
            it.value()->blockSignals(true);
            it.value()->setValue(s.scores.value(it.key()));
            it.value()->blockSignals(false);
        }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Resource path helper
// ─────────────────────────────────────────────────────────────────────────────

static const QString& findResourcePath()
{
    // Cache the result -- filesystem probing is not free and the answer
    // never changes during a single process lifetime.
    static const QString cached = []() -> QString {
        const QStringList candidates = {
            QApplication::applicationDirPath() + "/resources",
            QApplication::applicationDirPath() + "/../resources",
            "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer/resources",
            "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/resources"
        };
        for (const QString& p : candidates) {
            if (QDir(p).exists()) return p;
        }
        return candidates.first();
    }();
    return cached;
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

    // Averaged table overlay (shown when test avg selected, replaces cards)
    m_avgOverlayTable = new QTableWidget;
    m_avgOverlayTable->setEditTriggers(QAbstractItemView::NoEditTriggers);
    m_avgOverlayTable->setAlternatingRowColors(true);
    m_avgOverlayTable->setSelectionMode(QAbstractItemView::NoSelection);
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

    // Carry over image data from the stored session
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        const SensorySession& stored = m_sessions[m_currentTesterIdx];
        sess.imagePaths   = stored.imagePaths;
        sess.imageLayouts = stored.imageLayouts;
        sess.imageCrops   = stored.imageCrops;
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
                acc.sums[m] += s.scores.value(m, 5);
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
            avgSample.scores[m] = qRound(acc.sums.value(m, 0) / qMax(1, acc.count));
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

    m_avgOverlayTable->setRowCount(deviceNames.size());
    for (int i = 0; i < deviceNames.size(); ++i) {
        auto* nameItem = new QTableWidgetItem(deviceNames[i]);
        nameItem->setFlags(Qt::ItemIsEnabled);
        QFont f = nameItem->font(); f.setBold(true); nameItem->setFont(f);
        m_avgOverlayTable->setItem(i, 0, nameItem);

        const QMap<QString, double>& avgs = deviceAvgs[i];
        for (int c = 0; c < nMetrics; ++c) {
            double avg = avgs.value(kSensoryMetrics[c], 0.0);
            auto* valItem = new QTableWidgetItem(QString::number(avg, 'f', 1));
            valItem->setFlags(Qt::ItemIsEnabled);
            valItem->setTextAlignment(Qt::AlignCenter);
            m_avgOverlayTable->setItem(i, 1 + c, valItem);
        }
    }

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
        if (it.value() != 5) return false;
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

    SensorySession sess = buildSession();

    saveToJson(m_savePath + ".json", sess);
    saveToExcel(m_savePath + ".xlsx", sess);

    if (m_db) {
        m_db->saveSensorySession(sess);
    }

    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        m_sessions[m_currentTesterIdx] = sess;
        emit sessionsChanged();
    }

    QMessageBox::information(this, "Saved",
        QString("Session saved to:\n  %1.json\n  %2.xlsx\n  Database")
            .arg(QFileInfo(m_savePath).fileName(),
                 QFileInfo(m_savePath).fileName()));
}

void SensoryPanel::saveToJson(const QString& path, const SensorySession& sess)
{
    QJsonObject root;
    root["session_name"]  = sess.sessionName;
    root["test_title"]    = sess.testTitle;
    root["assessor_name"] = sess.assessorName;
    root["tester_name"]   = sess.testerName;
    root["media"]         = sess.media;
    root["date"]          = sess.date;
    root["timestamp"]     = sess.timestamp;

    QJsonArray samplesArr;
    for (const SensorySample& s : sess.samples) {
        QJsonObject sObj;
        sObj["name"]     = s.name;
        sObj["comments"] = s.comments;
        for (const QString& metric : kSensoryMetrics)
            sObj[metric] = s.scores.value(metric, 5);
        samplesArr.append(sObj);
    }
    root["samples"] = samplesArr;

    QFile f(path);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Text)) return;
    f.write(QJsonDocument(root).toJson(QJsonDocument::Indented));
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
            xlsx.write(row, i + 2, s.scores.value(kSensoryMetrics[i], 5));
        xlsx.write(row, kSensoryMetrics.size() + 2, s.comments);
        ++row;
    }

    row += 2;
    xlsx.write(row, 1, "Test Title");  xlsx.write(row, 2, sess.testTitle);   ++row;
    xlsx.write(row, 1, "Tester");      xlsx.write(row, 2, sess.testerName);  ++row;
    xlsx.write(row, 1, "Assessor");    xlsx.write(row, 2, sess.assessorName);++row;
    xlsx.write(row, 1, "Media");       xlsx.write(row, 2, sess.media);       ++row;
    xlsx.write(row, 1, "Date");        xlsx.write(row, 2, sess.date);

    xlsx.saveAs(path);
}

// ─────────────────────────────────────────────────────────────────────────────
// Load Excel
// ─────────────────────────────────────────────────────────────────────────────

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
            QJsonObject root = doc.object();

            SensorySession sess;
            sess.sessionName  = root["session_name"].toString();
            sess.testTitle    = root["test_title"].toString();
            sess.assessorName = root["assessor_name"].toString();
            sess.testerName   = root["tester_name"].toString();
            sess.media        = root["media"].toString();
            sess.date         = root["date"].toString();
            sess.timestamp    = root["timestamp"].toString();

            if (sess.sessionName.isEmpty())
                sess.sessionName = QFileInfo(path).baseName();

            for (const QJsonValue& sv : root["samples"].toArray()) {
                QJsonObject sObj = sv.toObject();
                SensorySample sample;
                sample.name     = sObj["name"].toString();
                sample.comments = sObj["comments"].toString();
                for (const QString& metric : kSensoryMetrics)
                    sample.scores[metric] = qBound(1, sObj[metric].toInt(5), 9);
                sess.samples.append(sample);
            }

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
                int score = qBound(1, val.toInt(), 9);
                if (score == 0) score = 5;
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
                sample.scores[metric] = 5;
        }
        if (hasData) samples.append(sample);
    }

    if (samples.isEmpty()) return 0;

    // Read metadata section (below data, after a gap)
    QString testerName, assessorName, media, date;
    for (int row = samples.size() + 3; row <= samples.size() + 12; ++row) {
        QString label = xlsx.read(row, 1).toString().trimmed().toLower();
        QString value = xlsx.read(row, 2).toString().trimmed();
        if (label == "test title")  { /* testTitle already from filename */ }
        else if (label == "tester")   testerName   = value;
        else if (label == "assessor") assessorName = value;
        else if (label == "media")    media        = value;
        else if (label == "date")     date         = value;
    }

    SensorySession sess;
    sess.sessionName  = testTitle + (testerName.isEmpty() ? "" : " - " + testerName);
    sess.testTitle    = testTitle;
    sess.testerName   = testerName;
    sess.assessorName = assessorName;
    sess.media        = media;
    sess.date         = date.isEmpty() ? dateNow : date;
    sess.timestamp    = tsNow;
    sess.samples      = samples;

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
        QMap<QString, int> scores;
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
                    int val = v.toInt(&ok);
                    if (ok && val >= 1 && val <= 9) { hasScores = true; break; }
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
                    int score = val.toInt();
                    score = qBound(1, score, 9);
                    if (score == 0) score = 5;
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
                    sample.scores[metric] = 5;
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
// Chart → image helper
// ─────────────────────────────────────────────────────────────────────────────

QByteArray SensoryPanel::renderChartToImage(int width, int height) const
{
    RadarChartWidget tempChart;
    tempChart.resize(width, height);
    SensorySession sess = buildSession();
    tempChart.setSessions({sess});

    QPixmap pixmap(width, height);
    pixmap.fill(QColor(248, 248, 248));
    tempChart.render(&pixmap);

    QByteArray pngBytes;
    QBuffer buf(&pngBytes);
    buf.open(QIODevice::WriteOnly);
    pixmap.save(&buf, "PNG");
    return pngBytes;
}

// ─────────────────────────────────────────────────────────────────────────────
// Generate Single PowerPoint Report
// ─────────────────────────────────────────────────────────────────────────────

void SensoryPanel::generateSingleReport()
{
    if (m_cards.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Add at least one sample before generating a report.");
        return;
    }

    QString path = QFileDialog::getSaveFileName(
        this, "Save Sensory PowerPoint Report", lastBrowseDir(),
        "PowerPoint files (*.pptx);;All files (*)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    SensorySession sess = buildSession();

    PptxWriter pptx;
    pptx.setResourcePath(findResourcePath());

    SlideTable rawTable;
    rawTable.headers << "Sample";
    for (const QString& m : kSensoryMetrics)
        rawTable.headers << m;
    rawTable.headers << "Comments";

    for (const SensorySample& s : sess.samples) {
        QStringList row;
        row << (s.name.isEmpty() ? QString("Sample") : s.name);
        for (const QString& m : kSensoryMetrics)
            row << QString::number(s.scores.value(m, 5));
        row << s.comments;
        rawTable.rows.append(row);
    }

    rawTable.x = 0.3;
    rawTable.y = 0.7;
    rawTable.w = 7.2;
    rawTable.colWidthFractions = {0.10, 0.11, 0.10, 0.11, 0.11, 0.13, 0.34};

    QByteArray chartPng = renderChartToImage(720, 600);

    double chartW = 5.0;
    double chartH = 4.2;
    double chartX = 8.0;
    double chartY = 0.7 + (6.8 - chartH) / 2.0;

    QVector<SlideImage> plots;
    SlideImage chartImg;
    chartImg.pngData = std::move(chartPng);
    chartImg.x = chartX;
    chartImg.y = chartY;
    chartImg.w = chartW;
    chartImg.h = chartH;
    plots.append(std::move(chartImg));

    QString title = "Sensory Evaluation";
    if (!sess.testTitle.isEmpty()) title = sess.testTitle;
    if (!sess.testerName.isEmpty()) title += " - " + sess.testerName;

    QString coverTitle = sess.testTitle.isEmpty() ? QStringLiteral("Sensory Evaluation") : sess.testTitle;
    pptx.addCoverSlide(coverTitle, sess.date.isEmpty()
        ? QDate::currentDate().toString("MMMM d, yyyy") : sess.date);

    pptx.addContentSlide(title, rawTable, plots);

    // ── Image slide (if session has images) ────────────────────────────────
    if (!sess.imagePaths.isEmpty()) {
        const bool hasLayouts = (sess.imageLayouts.size() == sess.imagePaths.size());
        if (hasLayouts) {
            QVector<SlideImage> slideImages;
            for (int i = 0; i < sess.imagePaths.size(); ++i) {
                QRectF crop = (i < sess.imageCrops.size())
                    ? sess.imageCrops[i] : QRectF(0,0,1,1);
                QByteArray data = loadAndCropImage(sess.imagePaths[i], crop);
                if (data.isEmpty()) continue;
                SlideImage si;
                si.pngData = std::move(data);
                const QRectF& r = sess.imageLayouts[i];
                si.x = r.x(); si.y = r.y();
                si.w = r.width(); si.h = r.height();
                slideImages.append(std::move(si));
            }
            if (!slideImages.isEmpty())
                pptx.addImageSlide(title + " - Images", slideImages);
        } else {
            QVector<QByteArray> imgBytes;
            for (int i = 0; i < sess.imagePaths.size(); ++i) {
                QRectF crop = (i < sess.imageCrops.size())
                    ? sess.imageCrops[i] : QRectF(0,0,1,1);
                QByteArray data = loadAndCropImage(sess.imagePaths[i], crop);
                if (!data.isEmpty()) imgBytes.append(std::move(data));
            }
            if (!imgBytes.isEmpty())
                pptx.addImageSlide(title + " - Images", imgBytes);
        }
    }

    if (!pptx.save(path)) {
        QMessageBox::warning(this, "Save Error",
                             "Could not save PowerPoint:\n" + pptx.lastError());
        return;
    }

    QMessageBox msg(this);
    msg.setWindowTitle("Report Saved");
    msg.setText("PowerPoint report saved successfully.");
    msg.setInformativeText(QFileInfo(path).fileName());
    msg.setIcon(QMessageBox::Information);
    msg.setStandardButtons(QMessageBox::Ok);
    msg.exec();
}

// ─────────────────────────────────────────────────────────────────────────────
// Generate Full (combined) Sensory Report
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

    QString path = QFileDialog::getSaveFileName(
        this, "Save Combined Sensory Report", lastBrowseDir(),
        "PowerPoint files (*.pptx);;All files (*)");
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    QString error;
    if (!generateCombinedPptx(selected, path, error)) {
        QMessageBox::warning(this, "Save Error",
                             "Could not save report:\n" + error);
        return;
    }

    QMessageBox msg(this);
    msg.setWindowTitle("Report Saved");
    msg.setText("Combined sensory report saved successfully.");
    msg.setInformativeText(QFileInfo(path).fileName());
    msg.setIcon(QMessageBox::Information);
    msg.setStandardButtons(QMessageBox::Ok);
    msg.exec();
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
        QVector<int> vals;
        double sum = 0;
        for (const SensorySample& s : sess.samples) {
            int v = s.scores.value(metric, 5);
            vals.append(v);
            sum += v;
        }

        double mean = (n > 0) ? sum / n : 0;
        double sumSq = 0;
        for (int v : vals) sumSq += (v - mean) * (v - mean);
        double variance = (n > 1) ? sumSq / (n - 1) : 0;
        double stddev = std::sqrt(qMax(0.0, variance));
        double stderr_ = (n > 0) ? stddev / std::sqrt(static_cast<double>(n)) : 0;

        int minVal = 9, maxVal = 1;
        for (int v : vals) { minVal = qMin(minVal, v); maxVal = qMax(maxVal, v); }

        QStringList outliers;
        for (int i = 0; i < vals.size(); ++i) {
            if (stddev > 0 && std::abs(vals[i] - mean) > 2.0 * stddev) {
                QString sName = (i < sess.samples.size() && !sess.samples[i].name.isEmpty())
                    ? sess.samples[i].name : QString("Sample %1").arg(i + 1);
                outliers << QString("%1(%2)").arg(sName).arg(vals[i]);
            }
        }

        double chiSq = 0;
        double expected = static_cast<double>(n) / 9.0;
        if (expected > 0) {
            QMap<int, int> freq;
            for (int v : vals) freq[v]++;
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
            << "," << minVal
            << "," << maxVal
            << "," << (maxVal - minVal)
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
            out << "," << s.scores.value(m, 5);
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
    if (sessions.isEmpty()) {
        errorOut = "No sessions to include in the report.";
        return false;
    }

    PptxWriter pptx;
    pptx.setResourcePath(findResourcePath());

    QString coverTitle = sessions.first().testTitle;
    if (coverTitle.isEmpty()) coverTitle = QStringLiteral("Sensory Evaluation");
    QString coverDate = sessions.first().date.isEmpty()
        ? QDate::currentDate().toString("MMMM d, yyyy")
        : sessions.first().date;
    pptx.addCoverSlide(coverTitle, coverDate);

    QMap<QString, QVector<int>> groups;
    QStringList groupOrder;
    for (int i = 0; i < sessions.size(); ++i) {
        QString title = sessions[i].testTitle.isEmpty()
            ? QStringLiteral("Sensory Evaluation") : sessions[i].testTitle;
        if (!groups.contains(title)) groupOrder.append(title);
        groups[title].append(i);
    }

    bool multiGroup = groups.size() > 1;

    for (const QString& groupTitle : groupOrder) {
        if (multiGroup) {
            QString groupDate;
            int firstIdx = groups[groupTitle].first();
            groupDate = sessions[firstIdx].date.isEmpty()
                ? QDate::currentDate().toString("MMMM d, yyyy")
                : sessions[firstIdx].date;
            pptx.addCoverSlide(groupTitle, groupDate);
        }

        for (int si : groups[groupTitle]) {
            const SensorySession& sess = sessions[si];

            SlideTable rawTable;
            rawTable.headers << "Sample";
            for (const QString& m : kSensoryMetrics)
                rawTable.headers << m;
            rawTable.headers << "Comments";

            for (const SensorySample& s : sess.samples) {
                QStringList row;
                row << (s.name.isEmpty() ? QString("Sample") : s.name);
                for (const QString& m : kSensoryMetrics)
                    row << QString::number(s.scores.value(m, 5));
                row << s.comments;
                rawTable.rows.append(row);
            }

            rawTable.x = 0.3;
            rawTable.y = 0.7;
            rawTable.w = 7.2;
            rawTable.colWidthFractions = {0.10, 0.11, 0.10, 0.11, 0.11, 0.13, 0.34};

            RadarChartWidget tempChart;
            tempChart.setSessions({sess});
            tempChart.resize(720, 600);

            QPixmap pixmap(720, 600);
            pixmap.fill(Qt::white);
            QPainter painter(&pixmap);
            tempChart.render(&painter);
            painter.end();

            QByteArray chartPng;
            {
                QBuffer buf(&chartPng);
                buf.open(QIODevice::WriteOnly);
                pixmap.save(&buf, "PNG");
            }

            double chartW = 5.0;
            double chartH = 4.2;
            double chartX = 8.0;
            double chartY = 0.7 + (6.8 - chartH) / 2.0;

            QVector<SlideImage> plots;
            SlideImage chartImg;
            chartImg.pngData = std::move(chartPng);
            chartImg.x = chartX;
            chartImg.y = chartY;
            chartImg.w = chartW;
            chartImg.h = chartH;
            plots.append(std::move(chartImg));

            QString title = "Sensory Evaluation";
            if (!sess.testTitle.isEmpty()) title = sess.testTitle;
            if (!sess.testerName.isEmpty()) title += " - " + sess.testerName;

            pptx.addContentSlide(title, rawTable, plots);

            // ── Image slide for this session ────────────────────────────────
            if (!sess.imagePaths.isEmpty()) {
                const bool hasLayouts = (sess.imageLayouts.size() == sess.imagePaths.size());
                if (hasLayouts) {
                    QVector<SlideImage> slideImages;
                    for (int i = 0; i < sess.imagePaths.size(); ++i) {
                        QRectF crop = (i < sess.imageCrops.size())
                            ? sess.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sess.imagePaths[i], crop);
                        if (data.isEmpty()) continue;
                        SlideImage si;
                        si.pngData = data;
                        const QRectF& r = sess.imageLayouts[i];
                        si.x = r.x(); si.y = r.y();
                        si.w = r.width(); si.h = r.height();
                        slideImages.append(si);
                    }
                    if (!slideImages.isEmpty())
                        pptx.addImageSlide(title + " - Images", slideImages);
                } else {
                    QVector<QByteArray> imgBytes;
                    for (int i = 0; i < sess.imagePaths.size(); ++i) {
                        QRectF crop = (i < sess.imageCrops.size())
                            ? sess.imageCrops[i] : QRectF(0,0,1,1);
                        QByteArray data = loadAndCropImage(sess.imagePaths[i], crop);
                        if (!data.isEmpty()) imgBytes.append(std::move(data));
                    }
                    if (!imgBytes.isEmpty())
                        pptx.addImageSlide(title + " - Images", imgBytes);
                }
            }
        }
    }

    // ── Cumulative slide: averages across all sessions ──────────────────────
    {
        // Build a single session whose samples are per-device averages
        // across all users. Each unique device name becomes one sample
        // with averaged scores.
        struct DeviceAccum { QMap<QString, double> sums; int count = 0; };
        QMap<QString, DeviceAccum> deviceMap;
        QStringList deviceOrder;

        for (const SensorySession& sess : sessions) {
            for (const SensorySample& s : sess.samples) {
                QString key = s.name.isEmpty() ? QStringLiteral("Sample") : s.name;
                if (!deviceMap.contains(key)) deviceOrder.append(key);
                DeviceAccum& acc = deviceMap[key];
                for (const QString& m : kSensoryMetrics)
                    acc.sums[m] += s.scores.value(m, 5);
                acc.count++;
            }
        }

        SensorySession cumSess;
        cumSess.testTitle = coverTitle;

        SlideTable cumTable;
        cumTable.headers << "Device";
        for (const QString& m : kSensoryMetrics)
            cumTable.headers << m;

        for (const QString& devName : deviceOrder) {
            const DeviceAccum& acc = deviceMap[devName];
            SensorySample avgSample;
            avgSample.name = devName;
            QStringList row;
            row << devName;
            for (const QString& m : kSensoryMetrics) {
                double avg = acc.sums.value(m, 0) / qMax(1, acc.count);
                avgSample.scores[m] = qRound(avg);
                row << QString::number(avg, 'f', 1);
            }
            cumSess.samples.append(avgSample);
            cumTable.rows.append(row);
        }

        cumTable.x = 0.3;
        cumTable.y = 0.7;
        cumTable.w = 7.2;
        cumTable.colWidthFractions = {0.22, 0.156, 0.156, 0.156, 0.156, 0.156};

        // Render cumulative radar chart
        RadarChartWidget cumChart;
        cumChart.setSessions({cumSess});
        cumChart.resize(720, 600);

        QPixmap cumPix(720, 600);
        cumPix.fill(Qt::white);
        QPainter cumPainter(&cumPix);
        cumChart.render(&cumPainter);
        cumPainter.end();

        QByteArray cumPng;
        {
            QBuffer buf(&cumPng);
            buf.open(QIODevice::WriteOnly);
            cumPix.save(&buf, "PNG");
        }

        QVector<SlideImage> cumPlots;
        SlideImage cumImg;
        cumImg.pngData = std::move(cumPng);
        cumImg.x = 8.0;
        cumImg.y = 0.7 + (6.8 - 4.2) / 2.0;
        cumImg.w = 5.0;
        cumImg.h = 4.2;
        cumPlots.append(std::move(cumImg));

        QString cumTitle = "Sensory - " + coverTitle + " - Cumulative";
        pptx.addContentSlide(cumTitle, cumTable, cumPlots);
    }

    if (!pptx.save(filePath)) {
        errorOut = pptx.lastError();
        return false;
    }
    return true;
}

} // namespace DVE
