#include "SensoryDialog.h"

#include <QHBoxLayout>
#include <QFormLayout>
#include <QPushButton>
#include <QToolButton>
#include <QFileDialog>
#include <QMessageBox>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QDateTime>
#include <QFile>
#include <QTextStream>
#include <QShortcut>
#include <QStyle>
#include <QLabel>
#include <QScrollBar>
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
#include <cmath>

#include "xlsxdocument.h"
#include "reporting/PptxWriter.h"

namespace DVE {

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

    // Cache sizeHint for all items (avoids calling it twice per item)
    QVector<QSize> sizeHints;
    sizeHints.reserve(m_items.size());
    for (int i = 0; i < m_items.size(); ++i)
        sizeHints.append(m_items[i]->sizeHint());

    // First pass: group items into rows to compute per-row widths
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

    // Second pass: position items, centering each row horizontally
    int y = effectiveRect.y();
    for (const Row& row : rows) {
        int xOffset = (availW - row.totalWidth) / 2;  // center
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

    // Name row
    auto* nameRow = new QHBoxLayout;
    nameRow->setSpacing(3);
    nameRow->addWidget(new QLabel("Name:"));
    m_nameEdit = new QLineEdit;
    m_nameEdit->setPlaceholderText(QString("Sample %1").arg(index + 1));
    nameRow->addWidget(m_nameEdit);
    mainLayout->addLayout(nameRow);

    // Score rows — enough vertical spacing for spinbox arrows to render
    auto* formLayout = new QFormLayout;
    formLayout->setSpacing(6);
    formLayout->setContentsMargins(0, 2, 0, 2);
    for (const QString& metric : kSensoryMetrics) {
        auto* spin = new QSpinBox;
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

    // Comments — scrollable with visible border
    mainLayout->addWidget(new QLabel("Comments:"));
    m_commentsEdit = new QTextEdit;
    m_commentsEdit->setMinimumHeight(40);
    m_commentsEdit->setMaximumHeight(60);
    m_commentsEdit->setStyleSheet("QTextEdit { border: 1px solid #999; background: white; }");
    mainLayout->addWidget(m_commentsEdit, 1);
    connect(m_commentsEdit, &QTextEdit::textChanged, this, &SampleCard::changed);
    connect(m_nameEdit, &QLineEdit::textChanged, this, &SampleCard::changed);

    // Remove button
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
    for (auto it = m_spinBoxes.begin(); it != m_spinBoxes.end(); ++it) {
        if (s.scores.contains(it.key())) {
            it.value()->blockSignals(true);
            it.value()->setValue(s.scores[it.key()]);
            it.value()->blockSignals(false);
        }
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Resource path helper (static, used for PPTX image resources)
// ─────────────────────────────────────────────────────────────────────────────

static QString findResourcePath()
{
    QStringList candidates = {
        QApplication::applicationDirPath() + "/resources",
        QApplication::applicationDirPath() + "/../resources",
        "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer/resources",
        "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/resources"
    };
    for (const QString& p : candidates) {
        if (QDir(p).exists()) return p;
    }
    return candidates.first();
}

// ─────────────────────────────────────────────────────────────────────────────
// SensoryDialog
// ─────────────────────────────────────────────────────────────────────────────

SensoryDialog::SensoryDialog(DatabaseManager* db, QWidget* parent)
    : QDialog(parent)
    , m_db(db)
{
    init();
    // Start with one empty session in the navigator
    SensorySession empty;
    empty.sessionName = QStringLiteral("New Session");
    empty.date = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    empty.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    m_sessions.append(empty);
    m_currentTesterIdx = 0;
    addSampleCard();
    refreshNavigator();
}

SensoryDialog::SensoryDialog(DatabaseManager* db, QWidget* parent,
                             const SensorySession& initialSession)
    : QDialog(parent)
    , m_db(db)
{
    init();
    m_sessions.append(initialSession);
    populateTesterCombo();
    // Combo signal will call applySession for index 0
}

SensoryDialog::SensoryDialog(DatabaseManager* db, QWidget* parent,
                             const QVector<SensorySession>& sessions)
    : QDialog(parent)
    , m_db(db)
{
    init();
    m_sessions = sessions;
    populateTesterCombo();
    // Combo signal will call applySession for index 0
}

void SensoryDialog::init()
{
    setWindowTitle("Sensory Evaluation");

    setMinimumSize(1200, 600);
    resize(1500, 800);

    m_refreshTimer = new QTimer(this);
    m_refreshTimer->setSingleShot(true);
    m_refreshTimer->setInterval(150);
    connect(m_refreshTimer, &QTimer::timeout, this, &SensoryDialog::onRefreshChart);

    auto* outerLayout = new QVBoxLayout(this);
    outerLayout->setSpacing(4);
    outerLayout->setContentsMargins(6, 6, 6, 6);

    // ── Toolbar ──
    buildToolBar();
    outerLayout->addWidget(m_toolbar);

    // ── Header row (assessor, media, date) ──
    auto* headerWidget = new QWidget;
    buildHeaderRow(headerWidget);
    outerLayout->addWidget(headerWidget);

    // ── Horizontal splitter: navigator | cards | radar chart ──
    m_splitter = new QSplitter(Qt::Horizontal);

    // Navigator panel (far left) — shows loaded sessions
    m_navPanel = new QWidget;
    auto* navLayout = new QVBoxLayout(m_navPanel);
    navLayout->setContentsMargins(0, 0, 0, 0);
    navLayout->setSpacing(4);
    navLayout->addWidget(new QLabel("Sessions"));
    m_navigator = new QListWidget;
    m_navigator->setMaximumWidth(200);
    m_navigator->setMinimumWidth(120);
    navLayout->addWidget(m_navigator, 1);
    connect(m_navigator, &QListWidget::currentRowChanged,
            this, &SensoryDialog::onNavigatorClicked);
    m_navPanel->setVisible(true);  // always visible
    m_splitter->addWidget(m_navPanel);

    // Middle: scroll area with flow layout for sample cards + Add button
    auto* leftPanel = new QWidget;
    auto* leftLayout = new QVBoxLayout(leftPanel);
    leftLayout->setContentsMargins(0, 0, 0, 0);
    leftLayout->setSpacing(4);

    m_scrollArea = new QScrollArea;
    m_scrollArea->setWidgetResizable(true);
    m_scrollArea->setHorizontalScrollBarPolicy(Qt::ScrollBarAlwaysOff);

    m_flowContainer = new QWidget;
    m_flowLayout    = new FlowLayout(m_flowContainer, 4, 6, 6);
    m_scrollArea->setWidget(m_flowContainer);
    leftLayout->addWidget(m_scrollArea, 1);

    auto* addBtn = new QPushButton("+ Add Sample");
    leftLayout->addWidget(addBtn);
    connect(addBtn, &QPushButton::clicked, this, &SensoryDialog::onAddSample);

    m_splitter->addWidget(leftPanel);

    // Right side: radar chart fills all available space
    m_chart = new RadarChartWidget;
    m_splitter->addWidget(m_chart);

    // Navigator fixed, cards panel ~565px, chart gets the rest
    m_splitter->setStretchFactor(0, 0);  // navigator
    m_splitter->setStretchFactor(1, 0);  // cards
    m_splitter->setStretchFactor(2, 1);  // chart
    m_splitter->setSizes({160, 565, 935});

    outerLayout->addWidget(m_splitter, 1);

    // ── Close button ──
    auto* btnRow = new QHBoxLayout;
    btnRow->addStretch();
    auto* closeBtn = new QPushButton("Close");
    closeBtn->setDefault(true);
    connect(closeBtn, &QPushButton::clicked, this, &QDialog::accept);
    btnRow->addWidget(closeBtn);
    outerLayout->addLayout(btnRow);
}

void SensoryDialog::buildToolBar()
{
    m_toolbar = new QToolBar;
    m_toolbar->setMovable(false);
    m_toolbar->setIconSize(QSize(16, 16));

    // ── File cascade menu ──
    auto* fileMenu = new QMenu("File", this);
    fileMenu->addAction(style()->standardIcon(QStyle::SP_FileIcon),
                        "New Session",       this, &SensoryDialog::onNewSession);
    fileMenu->addSeparator();
    fileMenu->addAction(style()->standardIcon(QStyle::SP_DialogOpenButton),
                        "Load Excel...",     this, &SensoryDialog::onLoadExcel);
    fileMenu->addAction(style()->standardIcon(QStyle::SP_DriveNetIcon),
                        "Load from Database...", this, &SensoryDialog::onLoadFromDatabase);
    fileMenu->addSeparator();
    fileMenu->addAction(style()->standardIcon(QStyle::SP_DialogSaveButton),
                        "Save  (Ctrl+S)",    this, &SensoryDialog::onSave);

    auto* fileBtn = new QToolButton;
    fileBtn->setText("File");
    fileBtn->setMenu(fileMenu);
    fileBtn->setPopupMode(QToolButton::InstantPopup);
    fileBtn->setIcon(style()->standardIcon(QStyle::SP_DirIcon));
    fileBtn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    m_toolbar->addWidget(fileBtn);

    m_toolbar->addSeparator();

    // ── Report cascade menu ──
    auto* reportMenu = new QMenu("Report", this);
    reportMenu->addAction("Generate PowerPoint Report...",
                          this, &SensoryDialog::onGeneratePptx);
    reportMenu->addAction("Generate Statistics Report...",
                          this, &SensoryDialog::onGenerateStats);

    auto* reportBtn = new QToolButton;
    reportBtn->setText("Report");
    reportBtn->setMenu(reportMenu);
    reportBtn->setPopupMode(QToolButton::InstantPopup);
    reportBtn->setIcon(style()->standardIcon(QStyle::SP_FileDialogDetailedView));
    reportBtn->setToolButtonStyle(Qt::ToolButtonTextBesideIcon);
    m_toolbar->addWidget(reportBtn);

    // ── Ctrl+S shortcut ──
    auto* saveShortcut = new QShortcut(QKeySequence::Save, this);
    connect(saveShortcut, &QShortcut::activated, this, &SensoryDialog::onSave);
}

void SensoryDialog::buildHeaderRow(QWidget* container)
{
    auto* layout = new QHBoxLayout(container);
    layout->setContentsMargins(0, 0, 0, 0);
    layout->setSpacing(10);

    // Tester selector combo (hidden when only one tester)
    layout->addWidget(new QLabel("Tester:"));
    m_testerCombo = new QComboBox;
    m_testerCombo->setFixedWidth(180);
    m_testerCombo->setVisible(false);
    layout->addWidget(m_testerCombo);
    connect(m_testerCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &SensoryDialog::onTesterChanged);

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

void SensoryDialog::addSampleCard(const SensorySample& sample)
{
    int idx = m_cards.size();
    auto* card = new SampleCard(idx, m_flowContainer);
    if (!sample.name.isEmpty() || !sample.scores.isEmpty()) {
        card->fromSample(sample);
    }
    connect(card, &SampleCard::changed,          this, &SensoryDialog::scheduleChartRefresh);
    connect(card, &SampleCard::removeRequested,  this, &SensoryDialog::onRemoveCard);

    m_flowLayout->addWidget(card);
    m_cards.append(card);

    scheduleChartRefresh();
}

void SensoryDialog::clearAllCards()
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

SensorySession SensoryDialog::buildSession() const
{
    SensorySession sess;
    sess.sessionName  = QString("Session_%1").arg(
        QDateTime::currentDateTime().toString("yyyyMMdd_HHmmss"));
    sess.testTitle    = m_testTitleEdit->text().trimmed();
    sess.assessorName = m_assessorEdit->text().trimmed();
    sess.testerName   = m_testerEdit->text().trimmed();
    sess.media        = m_mediaEdit->text().trimmed();
    sess.date         = m_dateLabel->text();
    sess.timestamp    = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

    for (const SampleCard* card : m_cards) {
        sess.samples.append(card->toSample());
    }
    return sess;
}

void SensoryDialog::applySession(const SensorySession& session)
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

void SensoryDialog::saveCurrentTester()
{
    // Save current UI state back into m_sessions at current index
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        m_sessions[m_currentTesterIdx] = buildSession();
    }
}

void SensoryDialog::onTesterChanged(int index)
{
    if (index < 0 || index >= m_sessions.size()) return;

    // Save current tester's edits before switching
    saveCurrentTester();

    m_currentTesterIdx = index;
    applySession(m_sessions[index]);

    // Sync navigator selection
    m_navigator->blockSignals(true);
    for (int i = 0; i < m_navigator->count(); ++i) {
        QListWidgetItem* item = m_navigator->item(i);
        if (item->data(Qt::UserRole).toInt() == index
            && (item->flags() & Qt::ItemIsSelectable)) {
            m_navigator->setCurrentItem(item);
            break;
        }
    }
    m_navigator->blockSignals(false);
}

QString SensoryDialog::sessionLabel(const SensorySession& s) const
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

void SensoryDialog::populateTesterCombo()
{
    m_testerCombo->blockSignals(true);
    m_testerCombo->clear();

    for (const SensorySession& s : m_sessions)
        m_testerCombo->addItem(sessionLabel(s));

    m_testerCombo->setVisible(m_sessions.size() > 1);
    m_testerCombo->blockSignals(false);

    // Apply the session at m_currentTesterIdx (or last if out of range)
    if (!m_sessions.isEmpty()) {
        if (m_currentTesterIdx < 0 || m_currentTesterIdx >= m_sessions.size())
            m_currentTesterIdx = m_sessions.size() - 1;
        m_testerCombo->setCurrentIndex(m_currentTesterIdx);
        applySession(m_sessions[m_currentTesterIdx]);
    }

    refreshNavigator();
}

void SensoryDialog::refreshNavigator()
{
    m_navigator->blockSignals(true);
    m_navigator->clear();

    for (int i = 0; i < m_sessions.size(); ++i) {
        auto* item = new QListWidgetItem(sessionLabel(m_sessions[i]));
        item->setData(Qt::UserRole, i);
        m_navigator->addItem(item);
    }

    // Select current
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_navigator->count())
        m_navigator->setCurrentRow(m_currentTesterIdx);

    m_navigator->blockSignals(false);

    m_splitter->setSizes({160, 565, 935});
}

void SensoryDialog::onNavigatorClicked(int row)
{
    if (row < 0 || row >= m_navigator->count()) return;
    QListWidgetItem* item = m_navigator->item(row);
    if (!(item->flags() & Qt::ItemIsSelectable)) return;

    int sessionIdx = item->data(Qt::UserRole).toInt();
    if (sessionIdx < 0 || sessionIdx >= m_sessions.size()) return;
    if (sessionIdx == m_currentTesterIdx) return;

    saveCurrentTester();
    m_currentTesterIdx = sessionIdx;
    applySession(m_sessions[sessionIdx]);

    // Sync the tester combo
    m_testerCombo->blockSignals(true);
    m_testerCombo->setCurrentIndex(sessionIdx);
    m_testerCombo->blockSignals(false);
}

void SensoryDialog::onLoadFromDatabase()
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

    // Build a selection dialog with a tree widget
    QDialog picker(this);
    picker.setWindowTitle("Load from Database");
    picker.setMinimumSize(700, 400);
    picker.resize(750, 500);

    auto* layout = new QVBoxLayout(&picker);

    auto* tree = new QTreeWidget;
    tree->setHeaderLabels({"Test Title / Tester", "Assessor", "Media", "Date", "Samples"});
    tree->setSelectionMode(QAbstractItemView::ExtendedSelection);
    tree->setRootIsDecorated(true);

    // Group by test title
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

    // Collect selected IDs
    QVector<int> selectedIds;
    const auto selected = tree->selectedItems();
    for (QTreeWidgetItem* item : selected) {
        int id = item->data(0, Qt::UserRole).toInt();
        if (id > 0) {
            selectedIds.append(id);
        } else if (id == -1) {
            // Parent selected — add all children
            for (int c = 0; c < item->childCount(); ++c) {
                int childId = item->child(c)->data(0, Qt::UserRole).toInt();
                if (childId > 0 && !selectedIds.contains(childId))
                    selectedIds.append(childId);
            }
        }
    }

    if (selectedIds.isEmpty()) return;

    // Load sessions
    saveCurrentTester();

    // Remove the initial empty session if it's still in default state
    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    for (int id : selectedIds) {
        SensorySession sess = m_db->loadSensorySession(id);
        if (!sess.samples.isEmpty())
            m_sessions.append(sess);
    }

    if (m_sessions.isEmpty()) {
        QMessageBox::warning(this, "Database", "Could not load the selected session(s).");
        return;
    }

    populateTesterCombo();
}

QString SensoryDialog::resolveTestName()
{
    QString testName = m_testTitleEdit->text().trimmed();
    if (!testName.isEmpty()) return testName;

    // Check if any loaded session has a test name
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

    // No previous name exists
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
// Chart
// ─────────────────────────────────────────────────────────────────────────────

void SensoryDialog::scheduleChartRefresh()
{
    m_refreshTimer->start(); // restarts the 150ms timer
}

void SensoryDialog::onRefreshChart()
{
    SensorySession sess = buildSession();
    m_chart->setSessions({sess});
}

// ─────────────────────────────────────────────────────────────────────────────
// Last browse directory
// ─────────────────────────────────────────────────────────────────────────────

QString SensoryDialog::lastBrowseDir() const
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

void SensoryDialog::setLastBrowseDir(const QString& filePath)
{
    m_lastBrowseDir = QFileInfo(filePath).absolutePath();
}

// ─────────────────────────────────────────────────────────────────────────────
// Toolbar actions
// ─────────────────────────────────────────────────────────────────────────────

bool SensoryDialog::isDefaultState() const
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

void SensoryDialog::onNewSession()
{
    // Save current session state before adding a new one
    saveCurrentTester();

    // Create a new empty session
    SensorySession empty;
    empty.sessionName = QStringLiteral("New Session");
    empty.date = QDateTime::currentDateTime().toString("yyyy-MM-dd");
    empty.timestamp = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);
    m_sessions.append(empty);

    // Switch to the new session
    m_currentTesterIdx = m_sessions.size() - 1;

    clearAllCards();
    m_testTitleEdit->clear();
    m_assessorEdit->clear();
    m_testerEdit->clear();
    m_mediaEdit->clear();
    m_dateLabel->setText(empty.date);
    addSampleCard();

    populateTesterCombo();
}

void SensoryDialog::onSave()
{
    if (m_cards.isEmpty()) {
        QMessageBox::warning(this, "No Data", "Add at least one sample before saving.");
        return;
    }

    // Resolve test name if blank
    QString testName = resolveTestName();
    if (testName.isEmpty() && m_testTitleEdit->text().trimmed().isEmpty())
        return;  // User declined to set a test name

    // If no save path yet, ask the user
    if (m_savePath.isEmpty()) {
        QString suggested = lastBrowseDir() + "/" + sessionLabel(buildSession());
        QString path = QFileDialog::getSaveFileName(
            this, "Save Sensory Session", suggested,
            "Excel files (*.xlsx);;All files (*)");
        if (path.isEmpty()) return;
        setLastBrowseDir(path);
        // Strip extension — we'll add .json / .xlsx ourselves
        if (path.endsWith(".xlsx", Qt::CaseInsensitive))
            path.chop(5);
        else if (path.endsWith(".json", Qt::CaseInsensitive))
            path.chop(5);
        m_savePath = path;
    }

    SensorySession sess = buildSession();

    // 1. Save to JSON
    saveToJson(m_savePath + ".json", sess);

    // 2. Save to Excel
    saveToExcel(m_savePath + ".xlsx", sess);

    // 3. Save to database
    if (m_db) {
        m_db->saveSensorySession(sess);
    }

    // Update the session in m_sessions if applicable
    if (m_currentTesterIdx >= 0 && m_currentTesterIdx < m_sessions.size()) {
        m_sessions[m_currentTesterIdx] = sess;
        refreshNavigator();
    }

    QMessageBox::information(this, "Saved",
        QString("Session saved to:\n  %1.json\n  %2.xlsx\n  Database")
            .arg(QFileInfo(m_savePath).fileName(),
                 QFileInfo(m_savePath).fileName()));
}

void SensoryDialog::saveToJson(const QString& path, const SensorySession& sess)
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

void SensoryDialog::saveToExcel(const QString& path, const SensorySession& sess)
{
    QXlsx::Document xlsx;

    // Header row
    xlsx.write(1, 1, "Sample");
    for (int i = 0; i < kSensoryMetrics.size(); ++i)
        xlsx.write(1, i + 2, kSensoryMetrics[i]);
    xlsx.write(1, kSensoryMetrics.size() + 2, "Comments");

    // Bold header format
    QXlsx::Format hdrFmt;
    hdrFmt.setFontBold(true);
    for (int c = 1; c <= kSensoryMetrics.size() + 2; ++c)
        xlsx.write(1, c, xlsx.read(1, c), hdrFmt);

    // Data rows
    int row = 2;
    for (const SensorySample& s : sess.samples) {
        xlsx.write(row, 1, s.name.isEmpty() ? QString("Sample %1").arg(row - 1) : s.name);
        for (int i = 0; i < kSensoryMetrics.size(); ++i)
            xlsx.write(row, i + 2, s.scores.value(kSensoryMetrics[i], 5));
        xlsx.write(row, kSensoryMetrics.size() + 2, s.comments);
        ++row;
    }

    // Metadata in a second row block
    row += 2;
    xlsx.write(row, 1, "Test Title");  xlsx.write(row, 2, sess.testTitle);   ++row;
    xlsx.write(row, 1, "Tester");      xlsx.write(row, 2, sess.testerName);  ++row;
    xlsx.write(row, 1, "Assessor");    xlsx.write(row, 2, sess.assessorName);++row;
    xlsx.write(row, 1, "Media");       xlsx.write(row, 2, sess.media);       ++row;
    xlsx.write(row, 1, "Date");        xlsx.write(row, 2, sess.date);

    xlsx.saveAs(path);
}

void SensoryDialog::onLoadExcel()
{
    QStringList paths = QFileDialog::getOpenFileNames(
        this, "Load Sensory Excel Template(s)", lastBrowseDir(),
        "Excel files (*.xlsx);;All files (*)");
    if (paths.isEmpty()) return;
    setLastBrowseDir(paths.first());

    // Save current state before loading new files
    saveCurrentTester();

    // Remove the initial empty session if it's still in default state
    if (m_sessions.size() == 1 && isDefaultState())
        m_sessions.clear();

    int loaded = 0;
    for (const QString& path : paths) {
        QXlsx::Document xlsx(path);
        if (!xlsx.load()) {
            QMessageBox::warning(this, "Load Error",
                                 "Could not open Excel file:\n" + path);
            continue;
        }

        QString testTitle = QFileInfo(path).baseName();
        QString dateNow   = QDateTime::currentDateTime().toString("yyyy-MM-dd");
        QString tsNow     = QDateTime::currentDateTimeUtc().toString(Qt::ISODate);

        // Read the metric headers from row 1, columns B-F
        QStringList excelMetrics;
        for (int col = 2; col <= 6; ++col) {
            QString hdr = xlsx.read(1, col).toString().trimmed();
            if (!hdr.isEmpty()) excelMetrics << hdr;
        }
        if (excelMetrics.size() != 5)
            excelMetrics = QStringList(kSensoryMetrics.begin(), kSensoryMetrics.end());

        // ── Parse device blocks ──────────────────────────────────────────────
        // Layout: device blocks separated by blank rows. Each block starts with
        // the device name in column A (first non-empty row). Subsequent rows are
        // panelist data: A=panelist name, B-F=scores, G=comments.
        // Rows 30+ are averages — stop there.

        struct PanelistEntry {
            QString testerName;
            QString deviceName;
            QMap<QString, int> scores;
            QString comments;
        };

        QVector<PanelistEntry> allEntries;
        QString currentDevice;
        bool inBlock = false;

        // Row 1 is special: A1 is the first device name, B1-F1 are headers
        QString a1 = xlsx.read(1, 1).toString().trimmed();
        if (!a1.isEmpty() && !a1.toLower().contains("averages")) {
            currentDevice = a1;
            inBlock = true;
        }

        for (int row = 2; row <= 29; ++row) {
            QString cellA = xlsx.read(row, 1).toString().trimmed();

            // Check if row is empty (no data in A-G)
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

            // Stop at averages section
            if (cellA.toLower().contains("average"))
                break;

            // If we're not in a block and column A has text, this is a new device
            if (!inBlock && !cellA.isEmpty()) {
                // Check if this row has scores — if not, it's a device header
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
                    // This is a device name header row
                    currentDevice = cellA;
                    inBlock = true;
                    continue;
                } else {
                    // Row has both a name and scores — treat name as tester
                    // (device was set from previous header or A1)
                    inBlock = true;
                }
            }

            if (!inBlock && cellA.isEmpty())
                continue;

            // If column A has text and we're in a block, this is a panelist row
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

        // ── Pivot: group entries by tester name → one session per tester ─────
        QMap<QString, QVector<int>> testerIndices; // testerName → entry indices
        QStringList testerOrder;
        for (int i = 0; i < allEntries.size(); ++i) {
            const QString& tn = allEntries[i].testerName;
            if (!testerIndices.contains(tn)) testerOrder.append(tn);
            testerIndices[tn].append(i);
        }

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
                // Fill missing metrics with default 5
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
    }

    if (loaded == 0) {
        QMessageBox::warning(this, "No Data",
                             "No sample data found in the selected file(s).");
        return;
    }

    populateTesterCombo();
}

void SensoryDialog::onAddSample()
{
    addSampleCard();
}

void SensoryDialog::onRemoveCard(SampleCard* card)
{
    m_flowLayout->removeWidget(card);
    m_cards.removeOne(card);
    card->deleteLater();
    scheduleChartRefresh();
}

// ─────────────────────────────────────────────────────────────────────────────
// Chart → image helper
// ─────────────────────────────────────────────────────────────────────────────

QByteArray SensoryDialog::renderChartToImage(int width, int height) const
{
    // Create a temporary RadarChartWidget and render it to a pixmap
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
// Generate PowerPoint Report
// ─────────────────────────────────────────────────────────────────────────────

void SensoryDialog::onGeneratePptx()
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
    int n = sess.samples.size();

    PptxWriter pptx;

    pptx.setResourcePath(findResourcePath());

    // ── Single slide: raw data table (top-left), stats table (bottom-left),
    //    radar chart (right, centered vertically) ──

    // Top-left: Raw data table organized by sample with Comments column
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

    // Compute actual rendered heights (must match PptxWriter row constants)
    const double kHdrH = 0.30;   // header row height in inches
    const double kRowH = 0.22;   // data row height in inches
    const double kGap  = 0.3;    // gap between tables

    double rawActualH = kHdrH + qMax(n, 1) * kRowH;
    double avgActualH = kHdrH + 5 * kRowH;   // 5 metrics always
    double totalBlockH = rawActualH + kGap + avgActualH;

    // Center the averages table vertically, raw table sits above it
    double usableTop = 0.7;
    double usableH   = 6.6;
    double blockTop   = usableTop + (usableH - totalBlockH) / 2.0;
    // Shift raw table up 1" from centered position
    double rawTop     = qMax(usableTop, blockTop - 1.0);

    rawTable.x = 0.3;
    rawTable.y = rawTop;
    rawTable.w = 7.2;
    rawTable.h = rawActualH;

    // Column width fractions: Sample, OvLiking, BurntTaste, VaporVol, OvFlavor, Smoothness, Comments
    rawTable.colWidthFractions = {0.10, 0.11, 0.10, 0.11, 0.11, 0.13, 0.34};

    // Bottom-left: Averages table (Metric, Average, Std Dev only)
    SlideTable avgTable;
    avgTable.headers << "Metric" << "Average" << "Std Dev";

    for (const QString& metric : kSensoryMetrics) {
        double sum = 0, sumSq = 0;

        for (const SensorySample& s : sess.samples) {
            int v = s.scores.value(metric, 5);
            sum += v;
            sumSq += v * v;
        }

        double avg = (n > 0) ? sum / n : 0;
        double variance = (n > 1) ? (sumSq - sum * sum / n) / (n - 1) : 0;
        double stddev = std::sqrt(qMax(0.0, variance));

        QStringList row;
        row << metric
            << QString::number(avg, 'f', 2)
            << QString::number(stddev, 'f', 2);
        avgTable.rows.append(row);
    }

    // Position below raw data table
    avgTable.x = 0.3;
    avgTable.y = rawTop + rawActualH + kGap;
    avgTable.w = 3.5;
    avgTable.h = avgActualH;
    avgTable.colWidthFractions = {0.46, 0.27, 0.27};

    // Right side: Radar chart image, centered vertically on the slide
    QByteArray chartPng = renderChartToImage(720, 600);

    double chartW = 5.0;
    double chartH = 4.2;
    double chartX = 8.0;
    // Center vertically: slide is 7.5" tall, title takes ~0.7"
    double chartY = 0.7 + (6.8 - chartH) / 2.0;

    QVector<SlideImage> plots;
    SlideImage chartImg;
    chartImg.pngData = chartPng;
    chartImg.x = chartX;
    chartImg.y = chartY;
    chartImg.w = chartW;
    chartImg.h = chartH;
    plots.append(chartImg);

    QString title = "Sensory Evaluation";
    if (!sess.testTitle.isEmpty()) title = sess.testTitle;
    if (!sess.testerName.isEmpty()) title += " - " + sess.testerName;

    // Cover slide
    QString coverTitle = sess.testTitle.isEmpty() ? QStringLiteral("Sensory Evaluation") : sess.testTitle;
    pptx.addCoverSlide(coverTitle, sess.date.isEmpty()
        ? QDate::currentDate().toString("MMMM d, yyyy") : sess.date);

    pptx.addDualTableSlide(title, rawTable, avgTable, plots);

    if (!pptx.save(path)) {
        QMessageBox::warning(this, "Save Error",
                             "Could not save PowerPoint:\n" + pptx.lastError());
        return;
    }

    {
        QMessageBox msg(this);
        msg.setWindowTitle("Report Saved");
        msg.setText("PowerPoint report saved successfully.");
        msg.setInformativeText(QFileInfo(path).fileName());
        msg.setIcon(QMessageBox::Information);
        msg.setStandardButtons(QMessageBox::Ok);
        msg.exec();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// Generate Statistics Report (.csv)
// ─────────────────────────────────────────────────────────────────────────────

void SensoryDialog::onGenerateStats()
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

void SensoryDialog::writeStatsCsv(const QString& path)
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
// Combined report: one slide per session
// ─────────────────────────────────────────────────────────────────────────────

bool SensoryDialog::generateCombinedReport(const QVector<SensorySession>& sessions,
                                            const QString& filePath,
                                            QString& errorOut)
{
    if (sessions.isEmpty()) {
        errorOut = "No sessions to include in the report.";
        return false;
    }

    PptxWriter pptx;
    pptx.setResourcePath(findResourcePath());

    // Cover slide: use the test title from the first session
    QString coverTitle = sessions.first().testTitle;
    if (coverTitle.isEmpty()) coverTitle = QStringLiteral("Sensory Evaluation");
    QString coverDate = sessions.first().date.isEmpty()
        ? QDate::currentDate().toString("MMMM d, yyyy")
        : sessions.first().date;
    pptx.addCoverSlide(coverTitle, coverDate);

    // Group sessions by test title for section title slides
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
        // Add a section title slide if there are multiple test titles
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
        int n = sess.samples.size();

        // Raw data table (top-left)
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

        // Compute actual rendered heights (must match PptxWriter row constants)
        const double kHdrH = 0.30;
        const double kRowH = 0.22;
        const double kGap  = 0.3;

        double rawActualH = kHdrH + qMax(n, 1) * kRowH;
        double avgActualH = kHdrH + 5 * kRowH;
        double totalBlockH = rawActualH + kGap + avgActualH;

        double usableTop = 0.7;
        double usableH   = 6.6;
        double blockTop   = usableTop + (usableH - totalBlockH) / 2.0;
        double rawTop     = qMax(usableTop, blockTop - 1.0);

        rawTable.x = 0.3;
        rawTable.y = rawTop;
        rawTable.w = 7.2;
        rawTable.h = rawActualH;
        rawTable.colWidthFractions = {0.10, 0.11, 0.10, 0.11, 0.11, 0.13, 0.34};

        // Averages table (bottom-left)
        SlideTable avgTable;
        avgTable.headers << "Metric" << "Average" << "Std Dev";

        for (const QString& metric : kSensoryMetrics) {
            double sum = 0, sumSq = 0;
            for (const SensorySample& s : sess.samples) {
                int v = s.scores.value(metric, 5);
                sum += v;
                sumSq += v * v;
            }
            double avg = (n > 0) ? sum / n : 0;
            double variance = (n > 1) ? (sumSq - sum * sum / n) / (n - 1) : 0;
            double stddev = std::sqrt(qMax(0.0, variance));

            QStringList row;
            row << metric
                << QString::number(avg, 'f', 2)
                << QString::number(stddev, 'f', 2);
            avgTable.rows.append(row);
        }

        avgTable.x = 0.3;
        avgTable.y = rawTop + rawActualH + kGap;
        avgTable.w = 3.5;
        avgTable.h = avgActualH;
        avgTable.colWidthFractions = {0.46, 0.27, 0.27};

        // Radar chart (right, centered vertically)
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
        chartImg.pngData = chartPng;
        chartImg.x = chartX;
        chartImg.y = chartY;
        chartImg.w = chartW;
        chartImg.h = chartH;
        plots.append(chartImg);

        QString title = "Sensory Evaluation";
        if (!sess.testTitle.isEmpty()) title = sess.testTitle;
        if (!sess.testerName.isEmpty()) title += " - " + sess.testerName;

        pptx.addDualTableSlide(title, rawTable, avgTable, plots);
    } // end session loop
    } // end group loop

    if (!pptx.save(filePath)) {
        errorOut = pptx.lastError();
        return false;
    }
    return true;
}

} // namespace DVE
