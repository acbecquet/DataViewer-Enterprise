#include "MainWindow.h"
#include "utils/AppTheme.h"
#include "pipeline/SheetProcessors.h"
#include "ui/SensoryPanel.h"

#include <QApplication>
#include <QFileDialog>
#include <QMessageBox>
#include <QInputDialog>
#include <QSettings>
#include <QCloseEvent>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QMimeData>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QDesktopServices>
#include <QUrl>
#include <QDir>
#include <QStandardPaths>
#include <QTreeWidgetItem>
#include <QtConcurrent/QtConcurrent>
#include <QDebug>
#include <QRegularExpression>
#include <QDate>
#include <QStyle>
#include <QProcess>
#include <QProcessEnvironment>
#include <QFile>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QTemporaryFile>
#include <QFileSystemWatcher>

namespace DVE {

// ─── Construction / Destruction ───────────────────────────────────────────────
MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
    , m_processor(new DataProcessor())
    , m_reportGen(new ReportGenerator(this))
    , m_db(new DatabaseManager(this))
    , m_loadWatcher(new QFutureWatcher<FileResult>(this))
{
    setWindowTitle("DataViewer Enterprise");
    setMinimumSize(1280, 800);
    resize(1600, 900);
    setAcceptDrops(true);

    m_db->open(defaultDbPath());

    setupUI();
    setupConnections();
    restoreSettings();

    // Debounce timer for batching Excel cell writes
    m_excelWriteTimer = new QTimer(this);
    m_excelWriteTimer->setSingleShot(true);
    m_excelWriteTimer->setInterval(500);
    connect(m_excelWriteTimer, &QTimer::timeout, this, &MainWindow::flushExcelWrites);

    // Auto-save to database 5 seconds after last modification
    m_dbSaveTimer = new QTimer(this);
    m_dbSaveTimer->setSingleShot(true);
    m_dbSaveTimer->setInterval(5000);
    connect(m_dbSaveTimer, &QTimer::timeout, this, [this]() { onUpdateDatabase(); });

    updateStatusBar("Ready");
}

MainWindow::~MainWindow()
{
    flushExcelWrites();
    saveSettings();
    delete m_processor;
}

// ─── Setup ────────────────────────────────────────────────────────────────────
void MainWindow::setupUI()
{
    setupRibbon();
    setupCentralWidget();
    setupDockPanels();
    setupStatusBar();
}

void MainWindow::setupRibbon()
{
    m_ribbon = new RibbonWidget(this);

    buildHomeTab(m_ribbon->addTab("Home"));
    buildReportsTab(m_ribbon->addTab("Reports"));
    buildToolsTab(m_ribbon->addTab("Tools"));

    QWidget* ribbonWrapper = new QWidget(this);
    QVBoxLayout* vl = new QVBoxLayout(ribbonWrapper);
    vl->setContentsMargins(0, 0, 0, 0);
    vl->setSpacing(0);
    vl->addWidget(m_ribbon);
    setMenuWidget(ribbonWrapper);
}

void MainWindow::buildHomeTab(RibbonTab* tab)
{
    auto* fileGrp  = tab->addGroup("File");

    // TPM buttons
    m_homeNewBtn   = fileGrp->addLargeButton("New File",
        style()->standardIcon(QStyle::SP_FileIcon), "Create a new test file from template");
    m_homeLoadBtn  = fileGrp->addLargeButton("Load File",
        style()->standardIcon(QStyle::SP_DialogOpenButton), "Open an Excel file (Ctrl+O)");
    m_homeCloseBtn = fileGrp->addLargeButton("Close",
        style()->standardIcon(QStyle::SP_DialogCloseButton), "Close current file");

    connect(m_homeNewBtn,   &QToolButton::clicked, this, &MainWindow::onNewFile);
    connect(m_homeLoadBtn,  &QToolButton::clicked, this, &MainWindow::onLoadFile);
    connect(m_homeCloseBtn, &QToolButton::clicked, this, &MainWindow::onCloseFile);

    // Sensory-mode buttons (initially hidden)
    m_homeSensNewBtn   = fileGrp->addLargeButton("New\nSession",
        style()->standardIcon(QStyle::SP_FileIcon), "Create a new sensory session");
    m_homeSensSaveBtn  = fileGrp->addLargeButton("Save",
        style()->standardIcon(QStyle::SP_DialogSaveButton), "Save session (Ctrl+S)");
    m_homeSensLoadXlBtn = fileGrp->addLargeButton("Load\nExcel",
        style()->standardIcon(QStyle::SP_DialogOpenButton), "Load sensory data from Excel");
    m_homeSensCloseBtn  = fileGrp->addLargeButton("Close",
        style()->standardIcon(QStyle::SP_DialogCloseButton), "Close selected session(s)");

    m_homeSensNewBtn->setVisible(false);
    m_homeSensSaveBtn->setVisible(false);
    m_homeSensLoadXlBtn->setVisible(false);
    m_homeSensCloseBtn->setVisible(false);

    connect(m_homeSensNewBtn,   &QToolButton::clicked, this, [this]() {
        if (m_sensoryPanel) m_sensoryPanel->newSession();
    });
    connect(m_homeSensSaveBtn,  &QToolButton::clicked, this, [this]() {
        if (m_sensoryPanel) m_sensoryPanel->save();
    });
    connect(m_homeSensLoadXlBtn, &QToolButton::clicked, this, [this]() {
        if (m_sensoryPanel) m_sensoryPanel->loadFiles();
    });
    connect(m_homeSensCloseBtn, &QToolButton::clicked, this, [this]() {
        if (!m_sensoryPanel) return;
        QVector<int> indices;
        for (auto* item : m_sensoryNav->selectedItems())
            indices.append(m_sensoryNav->row(item));
        if (indices.isEmpty() && m_sensoryPanel->currentSessionIndex() >= 0)
            indices.append(m_sensoryPanel->currentSessionIndex());
        if (indices.isEmpty()) return;
        m_sensoryPanel->closeSessions(indices);
        updateImageButton();
    });
}

void MainWindow::buildReportsTab(RibbonTab* tab)
{
    auto* rptGrp  = tab->addGroup("Generate");
    m_reportBtn1 = rptGrp->addLargeButton("Test Report",
        style()->standardIcon(QStyle::SP_FileDialogDetailedView),
        "Generate a PPTX report for the current sheet");
    m_reportBtn2 = rptGrp->addLargeButton("Full Report",
        style()->standardIcon(QStyle::SP_FileDialogListView),
        "Generate a PPTX report for all sheets");

    connect(m_reportBtn1, &QToolButton::clicked, this, &MainWindow::onGenerateTestReport);
    connect(m_reportBtn2, &QToolButton::clicked, this, &MainWindow::onGenerateFullReport);

    // ── Data Cleanup group ────────────────────────────────────────────────────
    m_cleanupGroup = tab->addGroup("Data Cleanup");
    auto* cleanBtn   = m_cleanupGroup->addLargeButton("Clean Data",
        style()->standardIcon(QStyle::SP_DialogResetButton),
        "Open the data cleanup dialog to exclude outliers from plots and reports");
    m_resetCleanupBtn = m_cleanupGroup->addLargeButton("Reset Cleanup",
        style()->standardIcon(QStyle::SP_BrowserReload),
        "Remove all data exclusions for the current sheet");
    m_resetCleanupBtn->setEnabled(false);

    connect(cleanBtn,         &QToolButton::clicked, this, &MainWindow::onCleanData);
    connect(m_resetCleanupBtn, &QToolButton::clicked, this, &MainWindow::onResetCleanup);
}

void MainWindow::buildViewTab(RibbonTab* tab)
{
    auto* layoutGrp = tab->addGroup("Layout");
    auto* tableBtn  = layoutGrp->addLargeButton("Table Only",
        style()->standardIcon(QStyle::SP_FileDialogDetailedView));
    auto* plotBtn   = layoutGrp->addLargeButton("Plots Only",
        style()->standardIcon(QStyle::SP_DesktopIcon));
    auto* bothBtn   = layoutGrp->addLargeButton("Both",
        style()->standardIcon(QStyle::SP_TitleBarMaxButton));

    connect(tableBtn, &QToolButton::clicked, this, &MainWindow::onViewDataTable);
    connect(plotBtn,  &QToolButton::clicked, this, &MainWindow::onViewPlots);
    connect(bothBtn,  &QToolButton::clicked, this, &MainWindow::onViewBoth);
}

void MainWindow::buildToolsTab(RibbonTab* tab)
{
    auto* grp     = tab->addGroup("Utilities");
    m_sensoryBtn = grp->addLargeButton("Sensory",
        QIcon(resourcePath() + "/images/ccell_icon.png"), "Toggle sensory evaluation mode");
    m_sensoryBtn->setCheckable(true);
    connect(m_sensoryBtn, &QToolButton::toggled, this, &MainWindow::toggleSensoryMode);

    auto* dbBtn   = grp->addLargeButton("Database",
        style()->standardIcon(QStyle::SP_DriveHDIcon), "Browse file database");
    connect(dbBtn,   &QToolButton::clicked, this, &MainWindow::onOpenDatabaseBrowser);

    auto* imgGrp = tab->addGroup("Images");
    m_inboxBtn = imgGrp->addLargeButton("Images",
        style()->standardIcon(QStyle::SP_DirOpenIcon),
        "Open Image Inbox to assign photos to samples");
    connect(m_inboxBtn, &QToolButton::clicked, this, &MainWindow::onOpenImageInbox);
}

void MainWindow::setupCentralWidget()
{
    // Vertical splitter: data table on top, plot panel on bottom
    m_centralSplitter = new QSplitter(Qt::Vertical, this);

    // ── Table panel (top) ──────────────────────────────────────────────────────
    m_tablePanel = new QWidget(this);
    m_tablePanel->setObjectName("tablePanel");
    QVBoxLayout* tableVL = new QVBoxLayout(m_tablePanel);
    tableVL->setContentsMargins(4, 4, 4, 2);
    tableVL->setSpacing(4);

    // Nav bar: search | Prev | count | Next  (compact, sits above the table)
    m_sampleNavBar = new QWidget(this);
    QHBoxLayout* navHL = new QHBoxLayout(m_sampleNavBar);
    navHL->setContentsMargins(2, 2, 2, 2);
    navHL->setSpacing(4);

    m_prevBtn = new QPushButton(QStringLiteral("\u25C0"), this);
    m_nextBtn = new QPushButton(QStringLiteral("\u25B6"), this);
    m_prevBtn->setToolTip("Previous sample (Ctrl+Left)");
    m_nextBtn->setToolTip("Next sample (Ctrl+Right)");
    m_prevBtn->setFixedSize(28, 24);
    m_nextBtn->setFixedSize(28, 24);
    // Must set padding:0 so global "padding: 5px 14px" doesn't crush text
    // inside these small buttons.  Also set color explicitly.
    const QString navBtnStyle =
        "QPushButton { border:1px solid #BCBCBC; border-radius:3px;"
        "  background:#E0E0E0; color:#1A1A1A; font-size:11pt; padding:0px; }"
        "QPushButton:hover { background:#CCE4FF; border-color:#0066CC; color:#003388; }"
        "QPushButton:pressed { background:#99CAFF; }"
        "QPushButton:disabled { background:#EBEBEB; color:#AAAAAA; }";
    m_prevBtn->setStyleSheet(navBtnStyle);
    m_nextBtn->setStyleSheet(navBtnStyle);
    m_prevBtn->setEnabled(false);
    m_nextBtn->setEnabled(false);

    m_sampleCountLabel = new QLabel("—", this);
    m_sampleCountLabel->setAlignment(Qt::AlignCenter);
    m_sampleCountLabel->setFont(AppTheme::fontSmall());
    m_sampleCountLabel->setMinimumWidth(70);

    m_addRowBtn    = new QPushButton("+  Add Row",    this);
    m_removeRowBtn = new QPushButton("\u2212  Remove Row", this);
    for (auto* btn : {m_addRowBtn, m_removeRowBtn}) {
        btn->setFixedHeight(24);
        btn->setStyleSheet(
            "QPushButton { border:1px solid #BCBCBC; border-radius:3px;"
            "  background:#FFFFFF; color:#1A1A1A; font-size:8pt; padding:0px 8px; }"
            "QPushButton:hover   { background:#E0EEFF; border-color:#0066CC; }"
            "QPushButton:pressed { background:#C0D8FF; }"
            "QPushButton:disabled{ color:#AAAAAA; }");
    }
    m_removeRowBtn->setEnabled(false);

    navHL->addWidget(m_prevBtn);
    navHL->addWidget(m_sampleCountLabel);
    navHL->addWidget(m_nextBtn);
    navHL->addStretch();
    navHL->addWidget(m_addRowBtn);
    navHL->addWidget(m_removeRowBtn);

    connect(m_prevBtn,      &QPushButton::clicked, this, &MainWindow::onPrevSample);
    connect(m_nextBtn,      &QPushButton::clicked, this, &MainWindow::onNextSample);
    connect(m_addRowBtn,    &QPushButton::clicked, this, &MainWindow::onAddRow);
    connect(m_removeRowBtn, &QPushButton::clicked, this, &MainWindow::onRemoveRow);

    tableVL->addWidget(m_sampleNavBar);

    m_dataTable = new QTableWidget(0, dataTableHeaders().size(), this);
    m_dataTable->setHorizontalHeaderLabels(dataTableHeaders());
    m_dataTable->horizontalHeader()->setStretchLastSection(false);
    m_dataTable->verticalHeader()->setDefaultSectionSize(22);
    m_dataTable->setAlternatingRowColors(true);
    m_dataTable->setSelectionBehavior(QAbstractItemView::SelectRows);
    m_dataTable->setEditTriggers(QAbstractItemView::DoubleClicked | QAbstractItemView::AnyKeyPressed);
    m_dataTable->setWordWrap(false);
    connect(m_dataTable, &QTableWidget::cellChanged,
            this, &MainWindow::onTableCellChanged);
    connect(m_dataTable, &QTableWidget::itemSelectionChanged, this, [this]() {
        if (m_removeRowBtn)
            m_removeRowBtn->setEnabled(m_dataTable->currentRow() >= 0);
    });

    // Notes column (7) stretches; all others are interactive with explicit widths
    m_dataTable->horizontalHeader()->setStretchLastSection(false);
    m_dataTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
    m_dataTable->horizontalHeader()->setSectionResizeMode(7, QHeaderView::Stretch);
    QList<int> colWidths = {50, 80, 80, 75, 75, 50, 50, 0/*stretch*/, 115, 120, 95, 135};
    for (int i = 0; i < colWidths.size() && i < m_dataTable->columnCount(); ++i)
        if (colWidths[i] > 0)
            m_dataTable->setColumnWidth(i, colWidths[i]);

    tableVL->addWidget(m_dataTable, 1);
    m_tablePanel->setMinimumHeight(120);

    // ── Plot panel (bottom) ────────────────────────────────────────────────────
    m_plotWidget = new PlotWidget(this);
    m_plotWidget->setMinimumHeight(200);

    m_centralSplitter->addWidget(m_tablePanel);
    m_centralSplitter->addWidget(m_plotWidget);
    m_centralSplitter->setStretchFactor(0, 50);
    m_centralSplitter->setStretchFactor(1, 50);

    // Wrap in a stacked widget (index 0 = TPM, index 1 = sensory, added lazily)
    m_centralStack = new QStackedWidget(this);
    m_centralStack->addWidget(m_centralSplitter);   // index 0
    setCentralWidget(m_centralStack);
}

void MainWindow::setupDockPanels()
{
    // ── Single left dock with vertical splitter ───────────────────────────────
    // Top half:    file / sheet browser
    // Bottom half: sample properties
    m_fileDock = new QDockWidget("Navigator", this);
    m_fileDock->setAllowedAreas(Qt::LeftDockWidgetArea | Qt::RightDockWidgetArea);
    m_fileDock->setFeatures(QDockWidget::DockWidgetMovable | QDockWidget::DockWidgetFloatable);

    QSplitter* leftSplitter = new QSplitter(Qt::Vertical, this);

    // ── File / sheet browser (tree only — no combo boxes) ──────────────────
    QWidget*     filePanel = new QWidget();
    QVBoxLayout* filePL    = new QVBoxLayout(filePanel);
    filePL->setContentsMargins(6, 4, 6, 4);
    filePL->setSpacing(2);

    // Hidden combo boxes (kept for internal navigation logic)
    m_fileCombo  = new QComboBox(filePanel);
    m_fileCombo->setVisible(false);
    m_sheetCombo = new QComboBox(filePanel);
    m_sheetCombo->setVisible(false);

    m_navLabel = new QLabel("Loaded Files:", filePanel);
    m_navLabel->setTextFormat(Qt::RichText);
    m_navLabel->setFont(AppTheme::fontSmall());

    // Stacked widget: index 0 = file tree (TPM), index 1 = session list (Sensory)
    m_navStack = new QStackedWidget(filePanel);

    m_fileTree = new QTreeWidget(filePanel);
    m_fileTree->setHeaderHidden(true);
    m_fileTree->setRootIsDecorated(true);
    m_fileTree->setIndentation(14);
    m_fileTree->setAlternatingRowColors(true);
    m_navStack->addWidget(m_fileTree);   // index 0

    m_sensoryNav = new QListWidget(filePanel);
    m_sensoryNav->setAlternatingRowColors(true);
    m_sensoryNav->setSelectionMode(QAbstractItemView::ExtendedSelection);
    m_navStack->addWidget(m_sensoryNav); // index 1

    filePL->addWidget(m_navLabel);
    filePL->addWidget(m_navStack, 1);

    leftSplitter->addWidget(filePanel);

    // ── Test Averages panel (sensory mode only) ─────────────────────────────
    m_testAvgPanel = new QWidget();
    QVBoxLayout* avgVL = new QVBoxLayout(m_testAvgPanel);
    avgVL->setContentsMargins(0, 0, 0, 0);
    avgVL->setSpacing(2);

    QLabel* avgHeader = new QLabel("  Test Averages", m_testAvgPanel);
    avgHeader->setFixedHeight(22);
    avgHeader->setStyleSheet(
        "background:#1F4E79; color:white; font-weight:600; font-size:8pt;");

    m_testAvgList = new QListWidget(m_testAvgPanel);
    m_testAvgList->setAlternatingRowColors(true);
    m_testAvgList->setStyleSheet(
        "QListWidget { font-size: 8pt; }"
        "QListWidget::item { padding: 1px 2px; }");
    m_testAvgList->setSpacing(0);

    // Hidden table — data shown in the main panel overlay instead
    m_testAvgTable = new QTableWidget(0, 2, m_testAvgPanel);
    m_testAvgTable->setVisible(false);

    m_testAvgAssessors = new QLabel("Assessors: —", m_testAvgPanel);
    m_testAvgAssessors->setWordWrap(true);
    m_testAvgAssessors->setStyleSheet("font-size: 7pt; padding: 1px 4px; color: #444;");
    m_testAvgTesters = new QLabel("Testers: —", m_testAvgPanel);
    m_testAvgTesters->setWordWrap(true);
    m_testAvgTesters->setStyleSheet("font-size: 7pt; padding: 1px 4px; color: #444;");
    m_testAvgCount = new QLabel("Sessions: 0", m_testAvgPanel);
    m_testAvgCount->setStyleSheet("font-size: 7pt; padding: 1px 4px; color: #444;");

    avgVL->addWidget(avgHeader);
    avgVL->addWidget(m_testAvgList, 1);
    avgVL->addWidget(m_testAvgAssessors);
    avgVL->addWidget(m_testAvgTesters);
    avgVL->addWidget(m_testAvgCount);

    m_testAvgPanel->setVisible(false);   // hidden until sensory mode

    connect(m_testAvgList, &QListWidget::itemClicked,
            this, [this]() { onTestAvgSelectionChanged(); });

    leftSplitter->addWidget(m_testAvgPanel);

    // ── Sample properties ────────────────────────────────────────────────────
    m_propPanel = new QWidget();
    QVBoxLayout* propVL = new QVBoxLayout(m_propPanel);
    propVL->setContentsMargins(0, 0, 0, 4);
    propVL->setSpacing(0);

    QLabel* propHeader = new QLabel("  Sample Properties", m_propPanel);
    propHeader->setFixedHeight(22);
    propHeader->setStyleSheet(
        "background:#1F4E79; color:white; font-weight:600; font-size:8pt;");

    m_propTable = new QTableWidget(0, 2, m_propPanel);
    m_propTable->setHorizontalHeaderLabels({"Property", "Value"});
    m_propTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::ResizeToContents);
    m_propTable->horizontalHeader()->setSectionResizeMode(1, QHeaderView::Stretch);
    m_propTable->verticalHeader()->setVisible(false);
    m_propTable->verticalHeader()->setDefaultSectionSize(20);
    m_propTable->setShowGrid(false);
    m_propTable->setAlternatingRowColors(false);
    m_propTable->setSelectionMode(QAbstractItemView::SingleSelection);
    m_propTable->setEditTriggers(QAbstractItemView::DoubleClicked | QAbstractItemView::AnyKeyPressed);
    m_propTable->setStyleSheet(
        "QTableWidget { border: none; font-size: 8pt; }"
        "QTableWidget::item { padding: 1px 4px; }"
        "QHeaderView::section { background: #E8E8E8; color: #1A1A1A; font-size: 8pt; padding: 2px 4px; border: none; border-bottom: 1px solid #BCBCBC; }");
    connect(m_propTable, &QTableWidget::cellChanged, this, &MainWindow::onPropCellChanged);

    propVL->addWidget(propHeader);
    propVL->addWidget(m_propTable, 1);

    // ── Image buttons below the properties table ──────────────────────────────
    QWidget*     imgBar    = new QWidget(m_propPanel);
    QHBoxLayout* imgLayout = new QHBoxLayout(imgBar);
    imgLayout->setContentsMargins(4, 3, 4, 3);
    imgLayout->setSpacing(4);

    m_loadImagesBtn = new QPushButton("Load Images", imgBar);
    m_viewImagesBtn = new QPushButton("View Images (0)", imgBar);
    m_viewImagesBtn->setEnabled(false);

    const QString imgBtnSS =
        "QPushButton { border: 1px solid #BCBCBC; border-radius: 3px;"
        "  background: #FFFFFF; color: #1A1A1A; font-size: 8pt; padding: 2px 6px; }"
        "QPushButton:hover  { background: #E0EEFF; border-color: #0066CC; }"
        "QPushButton:pressed{ background: #C0D8FF; }"
        "QPushButton:disabled{ color: #AAAAAA; }";
    m_loadImagesBtn->setStyleSheet(imgBtnSS);
    m_viewImagesBtn->setStyleSheet(imgBtnSS);

    imgLayout->addWidget(m_loadImagesBtn);
    imgLayout->addWidget(m_viewImagesBtn);
    imgBar->setFixedHeight(32);

    propVL->addWidget(imgBar);

    connect(m_loadImagesBtn, &QPushButton::clicked, this, &MainWindow::onLoadImages);
    connect(m_viewImagesBtn, &QPushButton::clicked, this, &MainWindow::onViewImages);

    leftSplitter->addWidget(m_propPanel);

    // Navigator ~35%, test averages ~15%, properties ~50%
    leftSplitter->setStretchFactor(0, 35);
    leftSplitter->setStretchFactor(1, 15);
    leftSplitter->setStretchFactor(2, 50);
    leftSplitter->setSizes({ 210, 80, 310 });

    m_propDock = nullptr;   // no separate right dock

    m_fileDock->setWidget(leftSplitter);
    addDockWidget(Qt::LeftDockWidgetArea, m_fileDock);
    m_fileDock->setMinimumWidth(220);
    m_fileDock->setMaximumWidth(320);
}

void MainWindow::setupStatusBar()
{
    m_statusLabel = new QLabel("Ready", this);
    m_progressBar = new QProgressBar(this);
    m_progressBar->setRange(0, 100);
    m_progressBar->setVisible(false);
    m_progressBar->setMaximumWidth(200);
    m_progressBar->setMaximumHeight(16);
    m_fileInfoLabel = new QLabel("", this);
    m_fileInfoLabel->setFont(AppTheme::fontSmall());
    m_fileInfoLabel->setStyleSheet("color: #555;");

    m_dbSyncLabel = new QLabel(this);
    m_dbSyncLabel->setFont(AppTheme::fontSmall());

    statusBar()->addWidget(m_statusLabel, 1);
    statusBar()->addPermanentWidget(m_dbSyncLabel);
    statusBar()->addPermanentWidget(m_progressBar);
    statusBar()->addPermanentWidget(m_fileInfoLabel);

    updateDbSyncIndicator();
}

void MainWindow::setupConnections()
{
    connect(m_fileCombo,  QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &MainWindow::onFileSelected);
    connect(m_sheetCombo, QOverload<int>::of(&QComboBox::currentIndexChanged),
            this, &MainWindow::onSheetSelected);

    connect(m_fileTree, &QTreeWidget::itemActivated,
            this, [this](QTreeWidgetItem* item, int) {
                if (!item->parent()) return;
                QString fileName  = item->parent()->text(0);
                QString sheetName = item->text(0);
                for (int fi = 0; fi < m_loadedFiles.size(); ++fi) {
                    if (m_loadedFiles[fi].fileName != fileName) continue;
                    m_fileCombo->setCurrentIndex(fi);
                    const auto& sheets = m_loadedFiles[fi].sheets;
                    for (int si = 0; si < sheets.size(); ++si) {
                        if (sheets[si].sheetName == sheetName) {
                            m_sheetCombo->setCurrentIndex(si);
                            break;
                        }
                    }
                    break;
                }
            });

    connect(m_loadWatcher, &QFutureWatcher<FileResult>::finished,
            this, &MainWindow::onFileLoadFinished);
    connect(m_reportGen, &ReportGenerator::reportFinished,
            this, &MainWindow::onReportFinished);
    connect(m_reportGen, &ReportGenerator::progressChanged,
            this, [this](int pct, const QString& msg) { setProgress(pct, msg); });

    // Keyboard shortcuts
    auto* prevAct = new QAction(this);
    prevAct->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_Left));
    connect(prevAct, &QAction::triggered, this, &MainWindow::onPrevSample);
    addAction(prevAct);

    auto* nextAct = new QAction(this);
    nextAct->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_Right));
    connect(nextAct, &QAction::triggered, this, &MainWindow::onNextSample);
    addAction(nextAct);

    auto* loadAct = new QAction(this);
    loadAct->setShortcut(QKeySequence::Open);
    connect(loadAct, &QAction::triggered, this, &MainWindow::onLoadFile);
    addAction(loadAct);

    // Ctrl+U: update database with modified files
    auto* dbUpdateAct = new QAction(this);
    dbUpdateAct->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_U));
    connect(dbUpdateAct, &QAction::triggered, this, &MainWindow::onUpdateDatabase);
    addAction(dbUpdateAct);

    // Sensory navigator selection → switch session or show averaged chart
    connect(m_sensoryNav, &QListWidget::itemSelectionChanged, this, [this]() {
        if (!m_sensoryPanel) return;
        auto selected = m_sensoryNav->selectedItems();
        if (selected.size() == 1) {
            // Single selection: switch to that session and show its data
            int idx = m_sensoryNav->row(selected.first());
            m_sensoryPanel->selectSession(idx);
            // Force chart refresh even if session didn't change (e.g. deselecting from multi)
            m_sensoryPanel->showAveragedChart({idx});
            m_sensoryPanel->showNormalView();
            updateSensoryProperties();
            if (m_testAvgList) {
                m_testAvgList->clearSelection();
                m_testAvgList->setCurrentRow(-1);
            }
        } else if (selected.size() > 1) {
            // Multi-selection (Ctrl+Click): show averaged radar chart
            QVector<int> indices;
            for (auto* item : selected)
                indices.append(m_sensoryNav->row(item));
            m_sensoryPanel->showAveragedChart(indices);
        } else if (selected.isEmpty()) {
            // All deselected — clear chart
            m_sensoryPanel->showAveragedChart({});
        }
        updateImageButton();
    });

    connect(m_sensoryNav, &QListWidget::itemChanged, this, [this](QListWidgetItem* item) {
        if (!m_sensoryPanel) return;
        int row = m_sensoryNav->row(item);
        if (row < 0) return;
        QSignalBlocker blocker(m_sensoryNav);
        m_sensoryPanel->renameSession(row, item->text());
    });

    // Ctrl+S shortcut — routes to sensory save when in sensory mode
    auto* saveAct = new QAction(this);
    saveAct->setShortcut(QKeySequence::Save);
    connect(saveAct, &QAction::triggered, this, [this]() {
        if (m_sensoryMode && m_sensoryPanel)
            m_sensoryPanel->save();
    });
    addAction(saveAct);

    // Inbox file watcher
    m_inboxWatcher = new QFileSystemWatcher(this);
    connect(m_inboxWatcher, &QFileSystemWatcher::directoryChanged,
            this, &MainWindow::onInboxFolderChanged);
}

// ─── File operations ──────────────────────────────────────────────────────────
void MainWindow::onNewFile()
{
    // --- locate template ---------------------------------------------------
    const QString tmpl = templatePath();
    if (tmpl.isEmpty()) {
        showError("Template Not Found",
                  "Could not locate the test template file.\n"
                  "Expected: resources/templates/Standardized Test Template - December 2025.xlsx");
        return;
    }

    // --- get available sheet names from template via Python ----------------
    const QString python = findPython();
    if (python.isEmpty()) {
        showError("Python Not Found",
                  "Python 3 with openpyxl is required to create new files.");
        return;
    }

    static const char* kListSheets = R"PY(
import sys, json
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1], read_only=True, data_only=True)
names = [n for n in wb.sheetnames if n != "Sheet1"]
print(json.dumps(names))
)PY";

    QString err;
    const QString out = runPython(python, kListSheets, { tmpl }, err);
    if (out.isEmpty()) {
        showError("Could Not Read Template", err.isEmpty() ? "No output from Python." : err);
        return;
    }

    QJsonDocument doc = QJsonDocument::fromJson(out.toUtf8());
    QStringList availableTests;
    for (const auto& v : doc.array())
        availableTests << v.toString();

    if (availableTests.isEmpty()) {
        showError("Template Empty", "No test sheets found in the template.");
        return;
    }

    // --- show test selection dialog ----------------------------------------
    NewFileDialog dlg(availableTests, this);
    if (dlg.exec() != QDialog::Accepted) return;

    const QStringList selected = dlg.selectedTests();
    if (selected.isEmpty()) {
        showInfo("No Tests Selected", "Please select at least one test to include.");
        return;
    }

    // --- get save path -----------------------------------------------------
    const QString savePath = QFileDialog::getSaveFileName(
        this, "Save New Test File",
        lastBrowseDir() + "/New Test File.xlsx",
        "Excel Files (*.xlsx)");
    if (savePath.isEmpty()) return;
    setLastBrowseDir(savePath);

    // --- Python: copy template + remove unselected sheets ------------------
    static const char* kCreateFile = R"PY(
import sys, shutil, json
from openpyxl import load_workbook

template_path = sys.argv[1]
new_path      = sys.argv[2]
keep          = json.loads(sys.argv[3])

shutil.copy(template_path, new_path)
wb = load_workbook(new_path)
for sheet in list(wb.sheetnames):
    if sheet not in keep and sheet != "Sheet1":
        del wb[sheet]
wb.save(new_path)
print("OK")
)PY";

    QJsonDocument keepDoc(QJsonArray::fromStringList(selected));
    QString createErr;
    const QString result = runPython(
        python, kCreateFile,
        { tmpl, savePath, QString::fromUtf8(keepDoc.toJson(QJsonDocument::Compact)) },
        createErr);

    if (result.trimmed() != "OK") {
        showError("Create File Failed",
                  createErr.isEmpty() ? "Python returned unexpected output." : createErr);
        return;
    }

    // --- load the new file -------------------------------------------------
    loadFile(savePath);
    updateStatusBar("New file created: " + QFileInfo(savePath).fileName());
}

void MainWindow::onEditHeaders()
{
    const FileResult*  file  = currentFile();
    const SheetResult* sheet = currentSheet();
    if (!file || !sheet || sheet->samples.isEmpty()) {
        showInfo("No Sample", "Load a file and select a sample first.");
        return;
    }

    const SampleResult& sample = sheet->samples[m_currentSampleIndex];

    HeaderEditDialog dlg(sample, this);
    if (dlg.exec() != QDialog::Accepted) return;

    const HeaderEditDialog::HeaderData hd = dlg.headerData();

    // --- write back to xlsx via Python ------------------------------------
    const QString python = findPython();
    if (python.isEmpty()) {
        showError("Python Not Found",
                  "Python 3 with openpyxl is required to save header changes.");
        return;
    }

    // Build JSON payload for the Python script.
    // Cell layout (0-based): offset = sampleIndex * 12
    //   row 0 col offset+5 : sampleID
    //   row 1 col offset+1 : media
    //   row 1 col offset+3 : resistance
    //   row 1 col offset+7 : puffingRegime
    //   row 2 col offset+1 : viscosity
    //   row 2 col offset+3 : tester
    //   row 2 col offset+5 : voltage
    //   row 2 col offset+7 : initialOilMass
    QJsonObject payload;
    payload["file_path"]   = file->filePath;
    payload["sheet_name"]  = sheet->sheetName;
    payload["sample_index"] = m_currentSampleIndex;
    payload["sampleID"]    = hd.sampleID;
    payload["tester"]      = hd.tester;
    payload["media"]       = hd.media;
    payload["viscosity"]   = hd.viscosity;
    payload["resistance"]  = hd.resistance;
    payload["voltage"]     = hd.voltage;
    payload["puffingRegime"] = hd.puffingRegime;
    payload["oilMass"]     = hd.initialOilMass;

    static const char* kWriteHeaders = R"PY(
import sys, json
from openpyxl import load_workbook

d    = json.loads(sys.argv[1])
wb   = load_workbook(d['file_path'])
ws   = wb[d['sheet_name']]
off  = d['sample_index'] * 12   # 0-based column offset

def num(v):
    try:    return float(v) if v else None
    except: return v if v else None

# Write cells (openpyxl uses 1-based row/col)
ws.cell(row=1, column=off+6,  value=d['sampleID'])
ws.cell(row=2, column=off+2,  value=d['media'])
ws.cell(row=2, column=off+4,  value=num(d['resistance']))
ws.cell(row=2, column=off+8,  value=d['puffingRegime'])
ws.cell(row=3, column=off+2,  value=num(d['viscosity']))
ws.cell(row=3, column=off+4,  value=d['tester'])
ws.cell(row=3, column=off+6,  value=num(d['voltage']))
ws.cell(row=3, column=off+8,  value=num(d['oilMass']))
wb.save(d['file_path'])
print("OK")
)PY";

    const QString jsonArg = QString::fromUtf8(
        QJsonDocument(payload).toJson(QJsonDocument::Compact));
    QString writeErr;
    const QString result = runPython(python, kWriteHeaders, { jsonArg }, writeErr);

    if (result.trimmed() != "OK") {
        showError("Save Failed",
                  writeErr.isEmpty() ? "Python returned unexpected output." : writeErr);
        return;
    }

    // --- reload the file to reflect changes --------------------------------
    // loadFile's "Replace if already loaded" branch preserves m_currentSheetIndex
    // and m_currentSampleIndex automatically, so we just reload.
    loadFile(file->filePath);
    updateStatusBar("Header data saved.");
}

void MainWindow::onTableCellChanged(int row, int col)
{
    // Columns 8-11 are calculated — ignore changes to those
    if (col < 0 || col > 7) return;

    SheetResult* sheet = currentSheet();
    FileResult*  file  = currentFile();
    if (!sheet || !file || sheet->samples.isEmpty()) return;
    if (m_currentSampleIndex >= sheet->samples.size()) return;

    SampleResult& sample = sheet->samples[m_currentSampleIndex];

    // Table skips rows where both weights are zero — map table row → data row
    int dataRow = -1;
    int tRow = 0;
    for (int i = 0; i < sample.rows.size(); ++i) {
        if (sample.rows[i].beforeWeight == 0.0 || sample.rows[i].afterWeight == 0.0)
            continue;
        if (tRow == row) { dataRow = i; break; }
        ++tRow;
    }
    if (dataRow < 0) return;

    QTableWidgetItem* item = m_dataTable->item(row, col);
    if (!item) return;
    const QString text = item->text().trimmed();

    DataRow& dr = sample.rows[dataRow];
    switch (col) {
        case 0: dr.puffs        = text.toDouble(); break;
        case 1: dr.beforeWeight = text.toDouble(); break;
        case 2: dr.afterWeight  = text.toDouble(); break;
        case 3: dr.drawPressure = text.toDouble(); break;
        case 4: dr.resistance   = text.toDouble(); break;
        case 5: dr.smell        = text; break;
        case 6: dr.clog         = text; break;
        case 7: dr.notes        = text; break;
        default: return;
    }

    // Recalculate metrics in-place and refresh display
    recalculateSampleMetrics(*sheet);

    // Update only the calculated columns (8-11), skipping filtered rows
    m_dataTable->blockSignals(true);
    int tIdx = 0;
    for (int i = 0; i < sample.rows.size(); ++i) {
        if (sample.rows[i].beforeWeight == 0.0 || sample.rows[i].afterWeight == 0.0)
            continue;
        const DataRow& dr2 = sample.rows[i];
        auto refreshNum = [&](int c, double v, int dp) {
            QTableWidgetItem* it = m_dataTable->item(tIdx, c);
            if (!it) {
                it = new QTableWidgetItem();
                it->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
                it->setFlags(it->flags() & ~Qt::ItemIsEditable);
                m_dataTable->setItem(tIdx, c, it);
            }
            it->setText(QString::number(v, 'f', dp));
        };
        refreshNum(8,  dr2.tpm,             4);
        refreshNum(9,  dr2.tpmPowerDensity, 4);
        refreshNum(10, dr2.variationTPM,    2);
        refreshNum(11, dr2.oilConsumed,     2);
        ++tIdx;
    }
    m_dataTable->blockSignals(false);

    if (currentSheetHasCleanup()) {
        const SheetResult cleaned = buildCleanedSheet(*sheet, m_currentFileIndex, m_currentSheetIndex);
        m_plotWidget->setSheetData(cleaned);
    } else {
        m_plotWidget->setSheetData(*sheet);
    }
    updateProperties(sample);
    markFileModified();

    // Queue cell write to Excel (debounced — batches rapid edits into one Python call)
    int excelRow = dataRow + 5;
    int excelCol = m_currentSampleIndex * 12 + col + 1;
    queueExcelWrite(file->filePath, sheet->sheetName, excelRow, excelCol, text);
}

void MainWindow::onPropCellChanged(int row, int col)
{
    if (col != 1) return;  // only value column triggers edits

    if (m_sensoryMode && m_sensoryPanel) {
        m_propTable->blockSignals(true);
        SensorySession* sess = m_sensoryPanel->currentSession();
        if (sess) {
            QTableWidgetItem* labelItem = m_propTable->item(row, 0);
            QTableWidgetItem* dataItem  = m_propTable->item(row, 1);
            if (labelItem && dataItem) {
                QString label = labelItem->text();
                QString value = dataItem->text().trimmed();
                if      (label == "Burn")                     sess->burnStatus = value;
                else if (label == "Clog")                     sess->clogStatus = value;
                else if (label == "Leak")                     sess->leakStatus = value;
                else if (label == "Puff Time (est., s)")      sess->puffLength = value;
                if (m_db->isOpen()) m_db->saveSensorySession(*sess);
            }
        }
        m_propTable->blockSignals(false);
        return;
    }

    SheetResult* sheet = currentSheet();
    FileResult*  file  = currentFile();
    if (!sheet || !file || sheet->samples.isEmpty()) return;
    if (m_currentSampleIndex >= sheet->samples.size()) return;

    SampleResult& s = sheet->samples[m_currentSampleIndex];
    QTableWidgetItem* item = m_propTable->item(row, col);
    if (!item) return;
    const QString text = item->text().trimmed();

    // Excel column offset (0-based): sampleIndex * 12
    int off = m_currentSampleIndex * 12;

    // Map prop table row to field + Excel cell (1-based row, 1-based col)
    int excelRow = -1, excelCol = -1;
    bool affectsPower = false;

    switch (row) {
        case 1:  // Sample ID
            s.sampleName = text; s.sampleID = text;
            excelRow = 1; excelCol = off + 6; break;
        case 3:  // Tester
            s.tester = text;
            excelRow = 3; excelCol = off + 4; break;
        case 5:  // Media
            s.media = text;
            excelRow = 2; excelCol = off + 2; break;
        case 6:  // Viscosity
            s.viscosity = text.toDouble();
            excelRow = 3; excelCol = off + 2; break;
        case 7:  // Resistance
            s.resistance = text.toDouble(); affectsPower = true;
            excelRow = 2; excelCol = off + 4; break;
        case 8:  // Voltage
            s.voltage = text.toDouble(); affectsPower = true;
            excelRow = 3; excelCol = off + 6; break;
        case 10: // Heating Tech
            s.heatingTechnology = text; affectsPower = true;
            excelRow = 1; excelCol = off + 8; break;
        case 12: // Puffing Regime
            s.puffingRegime = text;
            excelRow = 2; excelCol = off + 8; break;
        case 13: // Initial Oil
            s.initialOilMass = text.toDouble();
            excelRow = 3; excelCol = off + 8; break;
        default:
            return;  // read-only row, ignore
    }

    if (affectsPower) {
        // Recalculate power: P = V^2 / (R + Roffset)
        // Offset depends on heating technology (matches Excel SWITCH formula)
        double rOffset = 0.0;
        QString tech = s.heatingTechnology.trimmed().toUpper();
        if (tech == "CCELL3.0" || tech == "CCELL 3.0" || tech == "T58G")
            rOffset = 0.78;
        else if (tech == "T51")
            rOffset = 0.25;
        double denom = s.resistance + rOffset;
        s.power = (s.voltage > 0 && denom > 0) ? (s.voltage * s.voltage) / denom : 0.0;
    }

    // Recalculate all row-level metrics (TPM, power density, etc.)
    recalculateSampleMetrics(*sheet);

    // Refresh the properties panel (shows updated calculated values)
    updateProperties(s);

    // Refresh the data table calculated columns
    m_dataTable->blockSignals(true);
    for (int r = 0; r < s.rows.size(); ++r) {
        const DataRow& dr = s.rows[r];
        auto refreshNum = [&](int c, double v, int dp) {
            QTableWidgetItem* it = m_dataTable->item(r, c);
            if (!it) {
                it = new QTableWidgetItem();
                it->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
                it->setFlags(it->flags() & ~Qt::ItemIsEditable);
                m_dataTable->setItem(r, c, it);
            }
            it->setText(QString::number(v, 'f', dp));
        };
        refreshNum(8,  dr.tpm,             4);
        refreshNum(9,  dr.tpmPowerDensity, 4);
        refreshNum(10, dr.variationTPM,    2);
        refreshNum(11, dr.oilConsumed,     2);
    }
    m_dataTable->blockSignals(false);

    if (currentSheetHasCleanup()) {
        const SheetResult cleaned = buildCleanedSheet(*sheet, m_currentFileIndex, m_currentSheetIndex);
        m_plotWidget->setSheetData(cleaned);
    } else {
        m_plotWidget->setSheetData(*sheet);
    }
    markFileModified();

    // Queue cell write to Excel (debounced)
    if (excelRow > 0 && excelCol > 0)
        queueExcelWrite(file->filePath, sheet->sheetName, excelRow, excelCol, text);
}

void MainWindow::onLoadFile()
{
    QString path = QFileDialog::getOpenFileName(
        this, "Open Excel File", lastBrowseDir(),
        "Excel Files (*.xlsx *.xls);;All Files (*)"
    );
    if (path.isEmpty()) return;
    setLastBrowseDir(path);
    loadFile(path);
}


void MainWindow::onCloseFile()
{
    if (m_currentFileIndex < 0 || m_currentFileIndex >= m_loadedFiles.size()) return;

    // If this file has unsaved DB changes, prompt the user
    const QString& fp = m_loadedFiles[m_currentFileIndex].filePath;
    if (m_modifiedFilePaths.contains(fp)) {
        auto result = QMessageBox::question(
            this, "Unsaved Database Changes",
            "This file has unsaved database changes.\n"
            "Would you like to update the database before closing?",
            QMessageBox::Yes | QMessageBox::No | QMessageBox::Cancel);
        if (result == QMessageBox::Cancel) return;
        if (result == QMessageBox::Yes) {
            if (m_db->saveFile(m_loadedFiles[m_currentFileIndex]))
                m_modifiedFilePaths.remove(fp);
        }
    }

    m_modifiedFilePaths.remove(m_loadedFiles[m_currentFileIndex].filePath);

    // Remove all cleanup exclusions that belonged to this file
    const QString prefix = QString("%1:").arg(m_currentFileIndex);
    const QStringList keys = m_excludedRows.keys();
    for (const QString& k : keys)
        if (k.startsWith(prefix))
            m_excludedRows.remove(k);

    m_loadedFiles.removeAt(m_currentFileIndex);
    updateDbSyncIndicator();

    if (m_loadedFiles.isEmpty()) {
        // No files remain — clear everything
        m_currentFileIndex   = -1;
        m_currentSheetIndex  = -1;
        m_currentSampleIndex = 0;
        m_fileCombo->blockSignals(true);
        m_sheetCombo->blockSignals(true);
        populateFileTree();
        m_fileCombo->clear();
        m_sheetCombo->clear();
        m_fileCombo->blockSignals(false);
        m_sheetCombo->blockSignals(false);
        m_dataTable->setRowCount(0);
        m_plotWidget->clear();
        m_propTable->setRowCount(0);
        m_sampleCountLabel->setText("No file loaded");
        updateStatusBar("File closed.");
        updateCleanupButtons();
    } else {
        // Select a remaining file (clamp index to valid range)
        m_currentFileIndex   = qMin(m_currentFileIndex, m_loadedFiles.size() - 1);
        m_currentSheetIndex  = 0;
        m_currentSampleIndex = 0;
        populateFileTree();
        populateSheetCombo();
        updateStatusBar("File closed. Showing: " + m_loadedFiles[m_currentFileIndex].fileName);
    }
}

void MainWindow::onRecentFileTriggered(const QString& path) { loadFile(path); }

void MainWindow::loadFile(const QString& path)
{
    if (m_loading) { showInfo("Busy", "Please wait for the current load to finish."); return; }

    m_loading = true;
    updateStatusBar("Loading: " + QFileInfo(path).fileName() + "...");
    m_progressBar->setVisible(true);
    m_progressBar->setValue(10);

    DataProcessor* proc = m_processor;
    QString capturedPath = path;
    QFuture<FileResult> future = QtConcurrent::run([capturedPath, proc]() {
        return proc->processFile(capturedPath);
    });
    m_loadWatcher->setFuture(future);
}

void MainWindow::onFileLoadFinished()
{
    m_loading = false;
    m_progressBar->setVisible(false);

    FileResult result = m_loadWatcher->result();
    if (result.filePath.isEmpty()) {
        showError("Load Error", "Failed to load file.\n" + m_processor->lastError());
        updateStatusBar("Load failed.");
        return;
    }

    // Replace if already loaded — preserve images/layouts/crops from the in-memory version
    for (int i = 0; i < m_loadedFiles.size(); ++i) {
        if (m_loadedFiles[i].filePath == result.filePath) {
            // Copy per-sample image data from the existing in-memory file into the fresh result
            const FileResult& existing = m_loadedFiles[i];
            for (int si = 0; si < result.sheets.size() && si < existing.sheets.size(); ++si) {
                SheetResult& newSheet = result.sheets[si];
                const SheetResult& oldSheet = existing.sheets[si];
                for (int sj = 0; sj < newSheet.samples.size() && sj < oldSheet.samples.size(); ++sj) {
                    SampleResult& ns = newSheet.samples[sj];
                    const SampleResult& os = oldSheet.samples[sj];
                    if (!os.imagePaths.isEmpty()) {
                        ns.imagePaths   = os.imagePaths;
                        ns.imageLayouts = os.imageLayouts;
                        ns.imageCrops   = os.imageCrops;
                    }
                }
            }
            m_loadedFiles[i] = result;
            m_currentFileIndex = i;
            populateFileTree();
            populateSheetCombo();
            displayCurrentSample();
            m_db->saveFile(result);
            m_modifiedFilePaths.remove(result.filePath);
            updateDbSyncIndicator();
            updateStatusBar("Refreshed: " + result.fileName);
            return;
        }
    }

    m_loadedFiles.append(result);
    m_currentFileIndex   = m_loadedFiles.size() - 1;
    m_currentSheetIndex  = 0;
    m_currentSampleIndex = 0;

    m_db->saveFile(result);
    m_modifiedFilePaths.remove(result.filePath);
    updateDbSyncIndicator();
    populateFileTree();
    populateSheetCombo();
    displayCurrentSample();

    int totalSamples = 0;
    QString diagMsg = "DIAGNOSTIC — " + result.fileName + "\n\n";
    for (const auto& s : result.sheets) {
        totalSamples += s.samples.size();
        diagMsg += QString("  Sheet '%1': %2 samples").arg(s.sheetName).arg(s.samples.size());
        if (!s.samples.isEmpty()) {
            const auto& s0 = s.samples[0];
            diagMsg += QString(", sample[0].rows=%1, voltage=%2, resistance=%3")
                .arg(s0.rows.size()).arg(s0.voltage).arg(s0.resistance);
        }
        diagMsg += "\n";
    }
#ifndef QT_NO_DEBUG
    qDebug() << "[DVE DIAG]" << diagMsg;
#endif
    updateStatusBar("Loaded: " + result.fileName
                    + "  |  " + QString::number(result.sheets.size()) + " sheets"
                    + "  |  " + QString::number(totalSamples) + " samples");
}

// ─── Navigation ───────────────────────────────────────────────────────────────
void MainWindow::onFileSelected(int index)
{
    if (index < 0 || index >= m_loadedFiles.size()) return;
    m_currentFileIndex   = index;
    m_currentSheetIndex  = 0;
    m_currentSampleIndex = 0;
    populateSheetCombo();
}

void MainWindow::onSheetSelected(int index)
{
    if (m_currentFileIndex < 0) return;
    const auto* f = currentFile();
    if (!f || index < 0 || index >= f->sheets.size()) return;
    m_currentSheetIndex  = index;
    m_currentSampleIndex = 0;
    displayCurrentSample();
}

void MainWindow::onPrevSample()
{
    if (m_currentSampleIndex > 0) { --m_currentSampleIndex; displayCurrentSample(); }
}

void MainWindow::onNextSample()
{
    const SheetResult* s = currentSheet();
    if (!s) return;
    if (m_currentSampleIndex < (int)s->samples.size() - 1) {
        ++m_currentSampleIndex; displayCurrentSample();
    }
}

// ─── Display ─────────────────────────────────────────────────────────────────
void MainWindow::populateFileTree()
{
    m_fileTree->clear();
    m_fileCombo->blockSignals(true);
    m_fileCombo->clear();

    // Cache icons outside the loop to avoid repeated lookups
    const QIcon fileIcon  = style()->standardIcon(QStyle::SP_FileIcon);
    const QIcon sheetIcon = style()->standardIcon(QStyle::SP_FileDialogDetailedView);

    for (const auto& f : m_loadedFiles) {
        m_fileCombo->addItem(f.fileName);
        auto* fi = new QTreeWidgetItem(m_fileTree, {f.fileName});
        fi->setIcon(0, fileIcon);
        for (const auto& sheet : f.sheets) {
            auto* si = new QTreeWidgetItem(fi, {sheet.sheetName});
            si->setIcon(0, sheetIcon);
        }
        fi->setExpanded(true);
    }
    if (m_currentFileIndex >= 0 && m_currentFileIndex < m_fileCombo->count())
        m_fileCombo->setCurrentIndex(m_currentFileIndex);
    m_fileCombo->blockSignals(false);
}

void MainWindow::populateSheetCombo()
{
    m_sheetCombo->blockSignals(true);
    m_sheetCombo->clear();
    const auto* f = currentFile();
    if (f) {
        for (const auto& sheet : f->sheets)
            m_sheetCombo->addItem(sheet.sheetName);
    }
    if (m_currentSheetIndex >= 0 && m_currentSheetIndex < m_sheetCombo->count())
        m_sheetCombo->setCurrentIndex(m_currentSheetIndex);
    m_sheetCombo->blockSignals(false);
    displayCurrentSample();
}

void MainWindow::displayCurrentSample()
{
    const SheetResult* sheet = currentSheet();

    // ── Raw table (SOP / instruction sheets) ──────────────────────────────────
    if (sheet && sheet->isRawTable) {
        m_plotWidget->hide();
        m_sampleCountLabel->setText("SOP View");
        m_prevBtn->setEnabled(false);
        m_nextBtn->setEnabled(false);
        m_propTable->setRowCount(0);

        // Rebuild column headers for raw table
        const int nCols = sheet->rawHeaders.size();
        m_dataTable->blockSignals(true);
        m_dataTable->setWordWrap(true);
        m_dataTable->setColumnCount(nCols > 0 ? nCols : 1);
        m_dataTable->setHorizontalHeaderLabels(
            nCols > 0 ? sheet->rawHeaders : QStringList{"Content"});
        m_dataTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
        m_dataTable->horizontalHeader()->setStretchLastSection(true);
        // Fixed column width forces word-wrap to trigger (ResizeToContents prevents it)
        for (int c = 0; c < nCols; ++c)
            m_dataTable->setColumnWidth(c, 500);

        m_dataTable->setRowCount(sheet->rawRows.size());
        for (int r = 0; r < sheet->rawRows.size(); ++r) {
            const QStringList& rowData = sheet->rawRows[r];
            for (int c = 0; c < nCols; ++c) {
                const QString& text = (c < rowData.size()) ? rowData[c] : QString();
                auto* item = new QTableWidgetItem(text);
                item->setTextAlignment(Qt::AlignLeft | Qt::AlignTop);
                m_dataTable->setItem(r, c, item);
            }
        }
        m_dataTable->resizeRowsToContents();
        m_dataTable->blockSignals(false);

        const auto* f = currentFile();
        m_fileInfoLabel->setText(
            QString("%1  |  %2").arg(f ? f->fileName : "").arg(sheet->sheetName));
        return;
    }

    // Ensure plot is visible when showing normal sheets
    m_plotWidget->show();

    if (!sheet || sheet->samples.isEmpty()) {
        // Restore standard column headers if coming from a raw-table sheet
        m_dataTable->setColumnCount(dataTableHeaders().size());
        m_dataTable->setHorizontalHeaderLabels(dataTableHeaders());
        m_dataTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
        m_dataTable->horizontalHeader()->setStretchLastSection(true);
        m_dataTable->setRowCount(0);
        m_plotWidget->clear();
        m_propTable->setRowCount(0);
        m_sampleCountLabel->setText("0 / 0");
        m_prevBtn->setEnabled(false);
        m_nextBtn->setEnabled(false);
        return;
    }
    // Restore standard column layout if we came from a raw-table sheet
    if (m_dataTable->columnCount() != dataTableHeaders().size()) {
        m_dataTable->setWordWrap(false);
        m_dataTable->setColumnCount(dataTableHeaders().size());
        m_dataTable->setHorizontalHeaderLabels(dataTableHeaders());
        m_dataTable->horizontalHeader()->setStretchLastSection(false);
        m_dataTable->horizontalHeader()->setSectionResizeMode(QHeaderView::Interactive);
        m_dataTable->horizontalHeader()->setSectionResizeMode(7, QHeaderView::Stretch);
        QList<int> colWidths = {50, 80, 80, 75, 75, 50, 50, 0/*stretch*/, 115, 120, 95, 135};
        for (int i = 0; i < colWidths.size() && i < m_dataTable->columnCount(); ++i)
            if (colWidths[i] > 0)
                m_dataTable->setColumnWidth(i, colWidths[i]);
    }

    m_currentSampleIndex = qBound(0, m_currentSampleIndex, (int)sheet->samples.size() - 1);
    const SampleResult& sample = sheet->samples[m_currentSampleIndex];

    // ── Data table ────────────────────────────────────────────────────────────
    m_dataTable->blockSignals(true);

    // Count visible rows (skip rows where both weights are zero — empty template rows)
    int visibleRows = 0;
    for (const DataRow& dr : sample.rows)
        if (dr.beforeWeight != 0.0 && dr.afterWeight != 0.0)
            ++visibleRows;
    m_dataTable->setRowCount(visibleRows);

    // Exclusions for the current sample (may be empty — zero cost to look up)
    const QSet<int> curExcluded =
        exclusionsFor(m_currentFileIndex, m_currentSheetIndex, m_currentSampleIndex);

    // Lambdas that reuse existing QTableWidgetItems instead of allocating new ones
    auto getItem = [&](int r, int c) -> QTableWidgetItem* {
        QTableWidgetItem* it = m_dataTable->item(r, c);
        if (!it) {
            it = new QTableWidgetItem();
            m_dataTable->setItem(r, c, it);
        }
        return it;
    };

    int tRow = 0;
    for (int rowIdx = 0; rowIdx < sample.rows.size(); ++rowIdx) {
        const DataRow& dr = sample.rows[rowIdx];
        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;

        int col = 0;
        auto setNum = [&](double v, int dp = 4) {
            auto* item = getItem(tRow, col++);
            item->setText(QString::number(v, 'f', dp));
            item->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
            item->setFlags(item->flags() | Qt::ItemIsEditable);
            item->setForeground(QColor(Qt::black));
            item->setBackground(QColor(Qt::white));
            item->setFont(QFont());
        };
        auto setEmpty = [&]() {
            auto* item = getItem(tRow, col++);
            item->setText(QString());
            item->setTextAlignment(Qt::AlignCenter);
            item->setFlags(item->flags() | Qt::ItemIsEditable);
            item->setForeground(QColor(Qt::black));
            item->setBackground(QColor(Qt::white));
            item->setFont(QFont());
        };
        auto setStr = [&](const QString& v) {
            auto* item = getItem(tRow, col++);
            item->setText(v);
            item->setTextAlignment(Qt::AlignCenter);
            item->setFlags(item->flags() | Qt::ItemIsEditable);
            item->setForeground(QColor(Qt::black));
            item->setBackground(QColor(Qt::white));
            item->setFont(QFont());
        };

        setNum(dr.puffs, 0);
        setNum(dr.beforeWeight);
        setNum(dr.afterWeight);
        (dr.drawPressure == 0.0) ? setEmpty() : setNum(dr.drawPressure, 2);
        (dr.resistance   == 0.0) ? setEmpty() : setNum(dr.resistance, 3);
        setStr(dr.smell);
        setStr(dr.clog);
        { auto* item = getItem(tRow, col++); item->setText(dr.notes);
          item->setFlags(item->flags() | Qt::ItemIsEditable);
          item->setForeground(QColor(Qt::black)); item->setBackground(QColor(Qt::white)); item->setFont(QFont()); }
        auto setCalc = [&](double v, int dp) {
            auto* item = getItem(tRow, col++);
            item->setText(QString::number(v, 'f', dp));
            item->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
            item->setFlags(item->flags() & ~Qt::ItemIsEditable);
            item->setForeground(QColor(0x44, 0x44, 0x88));
            item->setBackground(QColor(Qt::white));
            item->setFont(QFont());
        };
        setCalc(dr.tpm, 4);
        setCalc(dr.tpmPowerDensity, 4);
        setCalc(dr.variationTPM, 2);
        setCalc(dr.oilConsumed, 2);

        // Visual indicator for excluded rows (strikethrough + light red background)
        if (curExcluded.contains(rowIdx)) {
            for (int c = 0; c < m_dataTable->columnCount(); ++c) {
                auto* itm = m_dataTable->item(tRow, c);
                if (!itm) continue;
                itm->setForeground(QColor(0xAA, 0xAA, 0xAA));
                itm->setBackground(QColor(0xF8, 0xF0, 0xF0));
                QFont f = itm->font();
                f.setStrikeOut(true);
                itm->setFont(f);
            }
        }

        ++tRow;
    }
    m_dataTable->blockSignals(false);

    // ── Plots & Properties ────────────────────────────────────────────────────
    // When cleanup is active, pass cleaned data to the plot and property panel
    // so the stats reflect only the included rows. The raw table above is
    // unchanged — excluded rows are just visually marked, not removed.
    if (currentSheetHasCleanup()) {
        const SheetResult cleaned = buildCleanedSheet(*sheet, m_currentFileIndex, m_currentSheetIndex);
        m_plotWidget->setSheetData(cleaned);
        updateProperties(cleaned.samples[m_currentSampleIndex]);
    } else {
        m_plotWidget->setSheetData(*sheet);
        updateProperties(sample);
    }

    updateSampleNav();
    updateImageButton();
    updateCleanupButtons();

    // ── File info bar ─────────────────────────────────────────────────────────
    const auto* f = currentFile();
    m_fileInfoLabel->setText(
        QString("%1  |  %2  |  Sample %3 / %4")
            .arg(f ? f->fileName : "")
            .arg(sheet->sheetName)
            .arg(m_currentSampleIndex + 1)
            .arg(sheet->samples.size())
    );
}

void MainWindow::updateSampleNav()
{
    const SheetResult* s = currentSheet();
    int total = s ? (int)s->samples.size() : 0;
    m_sampleCountLabel->setText(
        total > 0
        ? QString("Sample %1 / %2").arg(m_currentSampleIndex + 1).arg(total)
        : "No samples");
    m_prevBtn->setEnabled(m_currentSampleIndex > 0);
    m_nextBtn->setEnabled(s && m_currentSampleIndex < total - 1);
}

void MainWindow::updateProperties(const SampleResult& s)
{
    // Row layout:
    //  0 ─── Identification (section header, span)
    //  1  Sample ID      editable  Excel row=1 col=off+6
    //  2  Date           read-only
    //  3  Tester         editable  Excel row=3 col=off+4
    //  4 ─── Device Parameters (section header)
    //  5  Media          editable  Excel row=2 col=off+2
    //  6  Viscosity (cP) editable  Excel row=3 col=off+2
    //  7  Resistance (Ω) editable  Excel row=2 col=off+4
    //  8  Voltage (V)    editable  Excel row=3 col=off+6
    //  9  Power (W)      read-only (calculated)
    // 10  Heating Tech   editable  Excel row=1 col=off+8
    // 11 ─── Test Parameters (section header)
    // 12  Puffing Regime editable  Excel row=2 col=off+8
    // 13  Initial Oil (g) editable Excel row=3 col=off+8
    // 14  Oil Consumed(g) read-only
    // 15  Total Puffs    read-only
    // 16 ─── Calculated Metrics (section header)
    // 17  Avg TPM (mg/puff) read-only
    // 18  Std Dev TPM    read-only
    // 19  Burn/Clog/Leak read-only

    m_propTable->blockSignals(true);
    m_propTable->setRowCount(20);
    m_propTable->setColumnCount(2);

    // Helper lambdas
    auto makeHeader = [&](int row, const QString& title) {
        QTableWidgetItem* it = new QTableWidgetItem(title);
        it->setFlags(Qt::ItemIsEnabled);
        it->setBackground(QColor(0x1F, 0x4E, 0x79));
        it->setForeground(Qt::white);
        QFont f = it->font(); f.setBold(true); f.setPointSize(8); it->setFont(f);
        m_propTable->setItem(row, 0, it);
        m_propTable->setItem(row, 1, new QTableWidgetItem());
        m_propTable->item(row, 1)->setFlags(Qt::ItemIsEnabled);
        m_propTable->item(row, 1)->setBackground(QColor(0x1F, 0x4E, 0x79));
        m_propTable->setSpan(row, 0, 1, 2);
        m_propTable->setRowHeight(row, 18);
    };

    auto makeEditable = [&](int row, const QString& label, const QString& value) {
        QTableWidgetItem* lbl = new QTableWidgetItem(label);
        lbl->setFlags(Qt::ItemIsEnabled);
        lbl->setForeground(QColor(0x55, 0x55, 0x55));
        QFont f = lbl->font(); f.setBold(true); f.setPointSize(8); lbl->setFont(f);
        m_propTable->setItem(row, 0, lbl);

        QTableWidgetItem* val = new QTableWidgetItem(value);
        val->setFlags(Qt::ItemIsEnabled | Qt::ItemIsEditable | Qt::ItemIsSelectable);
        m_propTable->setItem(row, 1, val);
        m_propTable->setRowHeight(row, 20);
    };

    auto makeReadOnly = [&](int row, const QString& label, const QString& value) {
        QTableWidgetItem* lbl = new QTableWidgetItem(label);
        lbl->setFlags(Qt::ItemIsEnabled);
        lbl->setForeground(QColor(0x55, 0x55, 0x55));
        QFont f = lbl->font(); f.setBold(true); f.setPointSize(8); lbl->setFont(f);
        m_propTable->setItem(row, 0, lbl);

        QTableWidgetItem* val = new QTableWidgetItem(value);
        val->setFlags(Qt::ItemIsEnabled);
        val->setForeground(QColor(0x11, 0x11, 0x11));
        m_propTable->setItem(row, 1, val);
        m_propTable->setRowHeight(row, 20);
    };

    makeHeader(0, "  Identification");
    makeEditable(1, "Sample ID",       s.sampleName);
    makeReadOnly(2, "Date",            s.date);
    makeEditable(3, "Tester",          s.tester);

    makeHeader(4, "  Device Parameters");
    makeEditable(5, "Media",           s.media);
    makeEditable(6, "Viscosity (cP)",  QString::number(s.viscosity, 'f', 1));
    makeEditable(7, "Resistance (\u03A9)", QString::number(s.resistance, 'f', 3));
    makeEditable(8, "Voltage (V)",     QString::number(s.voltage, 'f', 2));
    makeReadOnly(9, "Power (W)",       QString::number(s.power, 'f', 2));
    makeEditable(10,"Heating Tech",    s.heatingTechnology.isEmpty() ? "" : s.heatingTechnology);

    makeHeader(11, "  Test Parameters");
    makeEditable(12,"Puffing Regime",  s.puffingRegime);
    makeEditable(13,"Initial Oil (g)", QString::number(s.initialOilMass, 'f', 3));
    makeReadOnly(14,"Oil Consumed (g)",QString::number(s.totalOilConsumed / 1000.0, 'f', 3));
    makeReadOnly(15,"Total Puffs",     QString::number(s.totalPuffs));

    makeHeader(16, "  Calculated Metrics");
    makeReadOnly(17,"Avg TPM (mg/puff)",QString::number(s.averageTPM, 'f', 4));
    makeReadOnly(18,"Std Dev TPM",      QString::number(s.stdDevTPM, 'f', 4));
    makeReadOnly(19,"Burn / Clog / Leak",
                 QString("%1 / %2 / %3")
                 .arg(s.burnStatus.isEmpty() ? "N" : s.burnStatus)
                 .arg(s.clogStatus.isEmpty()  ? "N" : s.clogStatus)
                 .arg(s.leakStatus.isEmpty()  ? "N" : s.leakStatus));

    m_propTable->blockSignals(false);
}

// ─── Reports ─────────────────────────────────────────────────────────────────
void MainWindow::onGenerateTestReport()
{
    const SheetResult* sheet = currentSheet();
    const FileResult*  file  = currentFile();
    if (!sheet || !file) { showInfo("No Data", "Load a file and select a sheet first."); return; }

    // Sanitize sheet name for filename
    QString safeName = sheet->sheetName;
    safeName.remove(QRegularExpression("[\\\\/:*?\"<>|]"));

    QString path = QFileDialog::getSaveFileName(
        this, "Save Test Report",
        lastBrowseDir() + "/" + file->fileName.chopped(5) + "_" + safeName + "_Report.pptx",
        "PowerPoint (*.pptx)"
    );
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    ReportConfig cfg;
    cfg.outputPath = path;
    m_reportGen->setResourcePath(resourcePath());
    // Build cleaned file so the report reflects any active data exclusions
    const FileResult reportFile = buildCleanedFile(*file);
    m_reportGen->generateTestReport(reportFile, sheet->sheetName, cfg);
}

void MainWindow::onGenerateFullReport()
{
    const FileResult* file = currentFile();
    if (!file) { showInfo("No Data", "Load a file first."); return; }

    QString path = QFileDialog::getSaveFileName(
        this, "Save Full Report",
        lastBrowseDir() + "/" + file->fileName.chopped(5) + "_Report.pptx",
        "PowerPoint (*.pptx)"
    );
    if (path.isEmpty()) return;
    setLastBrowseDir(path);

    ReportConfig cfg;
    cfg.outputPath = path;
    m_reportGen->setResourcePath(resourcePath());
    // Build cleaned file so the report reflects any active data exclusions
    const FileResult reportFile = buildCleanedFile(*file);
    m_reportGen->generateFullReport(reportFile, cfg);
}


void MainWindow::onReportFinished(bool success, const QString& path)
{
    m_progressBar->setVisible(false);
    if (success) {
        QMessageBox msgBox(this);
        msgBox.setWindowTitle("Report Saved");
        msgBox.setText("Report saved:\n" + path);
        msgBox.setIcon(QMessageBox::Information);
        auto* openBtn = msgBox.addButton("Open Folder", QMessageBox::ActionRole);
        msgBox.addButton(QMessageBox::Ok);
        msgBox.exec();
        if (msgBox.clickedButton() == openBtn)
            QDesktopServices::openUrl(QUrl::fromLocalFile(QFileInfo(path).dir().absolutePath()));
    } else {
        showError("Report Failed", "Could not generate report:\n" + m_reportGen->lastError());
    }
}

// ─── View ─────────────────────────────────────────────────────────────────────
void MainWindow::onViewDataTable() { if (!m_sensoryMode) { m_tablePanel->show(); m_plotWidget->hide(); } }
void MainWindow::onViewPlots()     { if (!m_sensoryMode) { m_tablePanel->hide(); m_plotWidget->show(); } }
void MainWindow::onViewBoth()      { if (!m_sensoryMode) { m_tablePanel->show(); m_plotWidget->show(); } }
void MainWindow::onZoomIn()  {}
void MainWindow::onZoomOut() {}
void MainWindow::onFitToWindow() {}

// ─── Sensory mode ─────────────────────────────────────────────────────────────

void MainWindow::toggleSensoryMode(bool checked)
{
    m_sensoryMode = checked;

    if (checked) {
        if (!m_sensoryPanel) {
            initSensoryPanel();
        }
        m_centralStack->setCurrentIndex(1);   // sensory panel
        m_navStack->setCurrentIndex(1);        // sensory navigator
        m_navLabel->setText("Sessions:  <span style='color:gray; font-size:11px;'>select multiple to show average sensory score</span>");
        refreshSensoryNavigator();
        if (m_testAvgPanel) m_testAvgPanel->setVisible(true);
        refreshSensoryAverages();
        updateSensoryProperties();
    } else {
        m_centralStack->setCurrentIndex(0);   // TPM splitter
        m_navStack->setCurrentIndex(0);        // file tree
        m_navLabel->setText("Loaded Files:");
        if (m_testAvgPanel) m_testAvgPanel->setVisible(false);
        // Restore TPM properties or clear table
        if (currentSheet() && m_currentSampleIndex >= 0
            && m_currentSampleIndex < currentSheet()->samples.size()) {
            updateProperties(currentSheet()->samples[m_currentSampleIndex]);
        } else {
            m_propTable->setRowCount(0);
        }
    }

    updateRibbonForMode();
    updateImageButton();
}

void MainWindow::initSensoryPanel()
{
    m_sensoryPanel = new SensoryPanel(m_db, this);
    m_centralStack->addWidget(m_sensoryPanel);   // index 1

    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::refreshSensoryNavigator);
    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::refreshSensoryAverages);
    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::updateImageButton);
    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::updateSensoryProperties);
}

void MainWindow::updateRibbonForMode()
{
    // Home tab: show/hide TPM vs sensory buttons
    m_homeNewBtn->setVisible(!m_sensoryMode);
    m_homeLoadBtn->setVisible(!m_sensoryMode);
    m_homeCloseBtn->setVisible(!m_sensoryMode);

    m_homeSensNewBtn->setVisible(m_sensoryMode);
    m_homeSensSaveBtn->setVisible(m_sensoryMode);
    m_homeSensLoadXlBtn->setVisible(m_sensoryMode);
    m_homeSensCloseBtn->setVisible(m_sensoryMode);

    // Reports tab: swap labels and connections
    // Always disconnect ALL clicked connections first to prevent lambda accumulation
    // when toggling sensory mode on/off repeatedly.
    disconnect(m_reportBtn1, &QToolButton::clicked, nullptr, nullptr);
    disconnect(m_reportBtn2, &QToolButton::clicked, nullptr, nullptr);

    if (m_sensoryMode) {
        m_reportBtn1->setText("Single Sensory\nReport");
        m_reportBtn1->setIcon(QIcon(resourcePath() + "/images/ccell_icon.png"));
        m_reportBtn1->setToolTip("Generate PPTX report for the current sensory session");
        m_reportBtn2->setText("Full Sensory\nReport");
        m_reportBtn2->setIcon(QIcon(resourcePath() + "/images/ccell_icon.png"));
        m_reportBtn2->setToolTip("Generate combined PPTX report for selected sessions");

        connect(m_reportBtn1, &QToolButton::clicked, this, [this]() {
            if (m_sensoryPanel) m_sensoryPanel->generateSingleReport();
        });
        connect(m_reportBtn2, &QToolButton::clicked, this, [this]() {
            if (m_sensoryPanel) m_sensoryPanel->generateFullReport();
        });

        // Hide cleanup group (not applicable to sensory)
        if (m_cleanupGroup) m_cleanupGroup->setVisible(false);
    } else {
        m_reportBtn1->setText("Test Report");
        m_reportBtn1->setIcon(style()->standardIcon(QStyle::SP_FileDialogDetailedView));
        m_reportBtn1->setToolTip("Generate a PPTX report for the current sheet");
        m_reportBtn2->setText("Full Report");
        m_reportBtn2->setIcon(style()->standardIcon(QStyle::SP_FileDialogListView));
        m_reportBtn2->setToolTip("Generate a PPTX report for all sheets");

        connect(m_reportBtn1, &QToolButton::clicked, this, &MainWindow::onGenerateTestReport);
        connect(m_reportBtn2, &QToolButton::clicked, this, &MainWindow::onGenerateFullReport);

        if (m_cleanupGroup) m_cleanupGroup->setVisible(true);
    }
}

void MainWindow::refreshSensoryNavigator()
{
    if (!m_sensoryPanel) return;

    {
        QSignalBlocker blocker(m_sensoryNav);
        m_sensoryNav->clear();

        auto sessions = m_sensoryPanel->allSessions();
        for (int i = 0; i < sessions.size(); ++i) {
            auto* navItem = new QListWidgetItem(m_sensoryPanel->sessionLabel(sessions[i]));
            navItem->setFlags(navItem->flags() | Qt::ItemIsEditable);
            m_sensoryNav->addItem(navItem);
        }

        int cur = m_sensoryPanel->currentSessionIndex();
        if (cur >= 0 && cur < m_sensoryNav->count())
            m_sensoryNav->setCurrentRow(cur);
    }   // QSignalBlocker restores prior blocked state here
}

// ─── Sensory Properties ──────────────────────────────────────────────────────
void MainWindow::updateSensoryProperties()
{
    if (!m_sensoryPanel) {
        m_propTable->setRowCount(0);
        return;
    }

    SensorySession* sess = m_sensoryPanel->currentSession();
    if (!sess) {
        m_propTable->setRowCount(0);
        return;
    }

    m_propTable->blockSignals(true);
    m_propTable->setRowCount(12);
    m_propTable->setColumnCount(2);

    // ── Helper lambdas (same style as updateProperties) ──
    auto makeHeader = [&](int row, const QString& title) {
        QTableWidgetItem* it = new QTableWidgetItem(title);
        it->setFlags(Qt::ItemIsEnabled);
        it->setBackground(QColor(0x1F, 0x4E, 0x79));
        it->setForeground(Qt::white);
        QFont f = it->font(); f.setBold(true); f.setPointSize(8); it->setFont(f);
        m_propTable->setItem(row, 0, it);
        m_propTable->setItem(row, 1, new QTableWidgetItem());
        m_propTable->item(row, 1)->setFlags(Qt::ItemIsEnabled);
        m_propTable->item(row, 1)->setBackground(QColor(0x1F, 0x4E, 0x79));
        m_propTable->setSpan(row, 0, 1, 2);
        m_propTable->setRowHeight(row, 18);
    };

    auto makeReadOnly = [&](int row, const QString& label, const QString& value) {
        QTableWidgetItem* lbl = new QTableWidgetItem(label);
        lbl->setFlags(Qt::ItemIsEnabled);
        lbl->setForeground(QColor(0x55, 0x55, 0x55));
        QFont f = lbl->font(); f.setBold(true); f.setPointSize(8); lbl->setFont(f);
        m_propTable->setItem(row, 0, lbl);

        QTableWidgetItem* val = new QTableWidgetItem(value);
        val->setFlags(Qt::ItemIsEnabled);
        val->setForeground(QColor(0x11, 0x11, 0x11));
        m_propTable->setItem(row, 1, val);
        m_propTable->setRowHeight(row, 20);
    };

    // ── Section: Session Info ──
    makeHeader(0, "  Session Info");
    makeReadOnly(1, "Test Title",  sess->testTitle);
    makeReadOnly(2, "Assessor",    sess->assessorName);
    makeReadOnly(3, "Tester",      sess->testerName);
    makeReadOnly(4, "Media",       sess->media);
    makeReadOnly(5, "Date",        sess->date);
    makeReadOnly(6, "Samples",     QString::number(sess->samples.size()));

    auto makeEditable = [&](int row, const QString& label, const QString& value) {
        QTableWidgetItem* lbl = new QTableWidgetItem(label);
        lbl->setFlags(Qt::ItemIsEnabled);
        lbl->setForeground(QColor(0x55, 0x55, 0x55));
        QFont f = lbl->font(); f.setBold(true); f.setPointSize(8); lbl->setFont(f);
        m_propTable->setItem(row, 0, lbl);
        QTableWidgetItem* val = new QTableWidgetItem(value);
        val->setFlags(Qt::ItemIsEnabled | Qt::ItemIsEditable | Qt::ItemIsSelectable);
        m_propTable->setItem(row, 1, val);
        m_propTable->setRowHeight(row, 20);
    };

    // ── Section: Device Properties ──
    makeHeader(7, "  Device Properties");
    makeEditable(8,  "Burn",                sess->burnStatus);
    makeEditable(9,  "Clog",               sess->clogStatus);
    makeEditable(10, "Leak",               sess->leakStatus);
    makeEditable(11, "Puff Time (est., s)", sess->puffLength);

    // ── Extend table for computed rows ──
    // Compute highest/lowest rated by "Overall Liking"
    QString highestRated, lowestRated;
    if (!sess->samples.isEmpty()) {
        int maxScore = INT_MIN;
        int minScore = INT_MAX;
        QStringList maxNames, minNames;

        for (const SensorySample& samp : sess->samples) {
            int score = samp.scores.value("Overall Liking", -1);
            if (score < 0) continue;
            if (score > maxScore) {
                maxScore = score;
                maxNames.clear();
                maxNames.append(samp.name);
            } else if (score == maxScore) {
                maxNames.append(samp.name);
            }
            if (score < minScore) {
                minScore = score;
                minNames.clear();
                minNames.append(samp.name);
            } else if (score == minScore) {
                minNames.append(samp.name);
            }
        }
        highestRated = maxNames.join(", ") + QString(" (%1)").arg(maxScore);
        lowestRated  = minNames.join(", ") + QString(" (%1)").arg(minScore);
    }

    m_propTable->setRowCount(15);
    makeHeader(12, "  Computed");
    makeReadOnly(13, "Highest Rated Device", highestRated);
    makeReadOnly(14, "Lowest Rated Device",  lowestRated);

    m_propTable->blockSignals(false);
}

// ─── Test Averages panel ─────────────────────────────────────────────────────
void MainWindow::refreshSensoryAverages()
{
    if (!m_testAvgList || !m_sensoryPanel) return;

    // Remember current selection
    QString prevSelection;
    if (m_testAvgList->currentItem())
        prevSelection = m_testAvgList->currentItem()->text();

    m_testAvgList->blockSignals(true);
    m_testAvgList->clear();

    auto sessions = m_sensoryPanel->allSessions();
    // Collect unique test titles in order
    QStringList titles;
    QSet<QString> seen;
    for (const auto& s : sessions) {
        if (!seen.contains(s.testTitle)) {
            seen.insert(s.testTitle);
            titles.append(s.testTitle);
        }
    }

    for (const auto& t : titles)
        m_testAvgList->addItem(t);

    // Restore previous selection if still present
    int restoreRow = -1;
    if (!prevSelection.isEmpty()) {
        for (int i = 0; i < m_testAvgList->count(); ++i) {
            if (m_testAvgList->item(i)->text() == prevSelection) {
                restoreRow = i;
                break;
            }
        }
    }

    // Set the row while signals are still blocked to avoid a redundant
    // allSessions() call inside onTestAvgSelectionChanged().
    if (restoreRow >= 0)
        m_testAvgList->setCurrentRow(restoreRow);
    // Don't auto-select the first item — only show averages when user clicks

    m_testAvgList->blockSignals(false);

    // If a row was restored, update the details
    if (restoreRow >= 0)
        onTestAvgSelectionChanged();
}

void MainWindow::onTestAvgSelectionChanged()
{
    if (!m_testAvgTable || !m_sensoryPanel) return;

    m_testAvgTable->setRowCount(0);
    m_testAvgAssessors->setText("Assessors: \u2014");
    m_testAvgTesters->setText("Testers: \u2014");
    m_testAvgCount->setText("Sessions: 0");

    if (!m_testAvgList->currentItem()) {
        // No test avg selected — restore normal cards view
        m_sensoryPanel->showNormalView();
        return;
    }

    QString selectedTitle = m_testAvgList->currentItem()->text();
    auto sessions = m_sensoryPanel->allSessions();

    // Filter sessions matching the selected test title
    QVector<SensorySession> matching;
    for (const auto& s : sessions) {
        if (s.testTitle == selectedTitle)
            matching.append(s);
    }

    if (matching.isEmpty()) return;

    // Group by DEVICE (sample name), average across USERS (sessions)
    struct DeviceAccum { QMap<QString, double> sums; int count = 0; };
    QMap<QString, DeviceAccum> deviceMap;
    QStringList deviceOrder;
    QSet<QString> assessors, testers;

    for (const auto& sess : matching) {
        if (!sess.assessorName.isEmpty())
            assessors.insert(sess.assessorName);
        if (!sess.testerName.isEmpty())
            testers.insert(sess.testerName);

        for (const auto& sample : sess.samples) {
            QString key = sample.name.isEmpty() ? QStringLiteral("Sample") : sample.name;
            if (!deviceMap.contains(key)) deviceOrder.append(key);
            DeviceAccum& acc = deviceMap[key];
            for (const QString& m : kSensoryMetrics)
                acc.sums[m] += sample.scores.value(m, 5);
            acc.count++;
        }
    }

    // Build table: columns = Device + each metric
    int nMetrics = kSensoryMetrics.size();
    m_testAvgTable->setColumnCount(1 + nMetrics);
    QStringList headers;
    headers << "Device";
    for (const QString& m : kSensoryMetrics)
        headers << m;
    m_testAvgTable->setHorizontalHeaderLabels(headers);
    m_testAvgTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::Stretch);
    for (int c = 1; c <= nMetrics; ++c)
        m_testAvgTable->horizontalHeader()->setSectionResizeMode(c, QHeaderView::ResizeToContents);

    m_testAvgTable->setRowCount(deviceOrder.size());
    for (int i = 0; i < deviceOrder.size(); ++i) {
        const QString& devName = deviceOrder[i];
        const DeviceAccum& acc = deviceMap[devName];

        auto* nameItem = new QTableWidgetItem(devName);
        nameItem->setFlags(Qt::ItemIsEnabled);
        QFont f = nameItem->font(); f.setBold(true); nameItem->setFont(f);
        m_testAvgTable->setItem(i, 0, nameItem);

        for (int c = 0; c < nMetrics; ++c) {
            double avg = acc.sums.value(kSensoryMetrics[c], 0) / qMax(1, acc.count);
            auto* valItem = new QTableWidgetItem(QString::number(avg, 'f', 1));
            valItem->setFlags(Qt::ItemIsEnabled);
            valItem->setTextAlignment(Qt::AlignCenter);
            m_testAvgTable->setItem(i, 1 + c, valItem);
        }
    }

    // Assessors / Testers / Session count
    QStringList assessorList = assessors.values();
    assessorList.sort(Qt::CaseInsensitive);
    QStringList testerList = testers.values();
    testerList.sort(Qt::CaseInsensitive);

    m_testAvgAssessors->setText("Assessors: " +
        (assessorList.isEmpty() ? QString("\u2014") : assessorList.join(", ")));
    m_testAvgTesters->setText("Testers: " +
        (testerList.isEmpty() ? QString("\u2014") : testerList.join(", ")));
    m_testAvgCount->setText("Sessions: " + QString::number(matching.size()));

    // Show averaged radar chart in the main sensory panel
    // Build a synthetic session with per-device averaged scores
    SensorySession avgSess;
    avgSess.testTitle = selectedTitle + " (Average)";
    for (const QString& devName : deviceOrder) {
        const DeviceAccum& acc = deviceMap[devName];
        SensorySample avgSample;
        avgSample.name = devName;
        for (const QString& m : kSensoryMetrics)
            avgSample.scores[m] = qRound(acc.sums.value(m, 0) / qMax(1, acc.count));
        avgSess.samples.append(avgSample);
    }

    // Update the radar chart and left panel to show averaged data
    if (m_sensoryPanel) {
        // Find all session indices matching this test title to use showAveragedChart
        QVector<int> matchingIndices;
        auto allSess = m_sensoryPanel->allSessions();
        for (int i = 0; i < allSess.size(); ++i) {
            if (allSess[i].testTitle == selectedTitle)
                matchingIndices.append(i);
        }
        if (!matchingIndices.isEmpty())
            m_sensoryPanel->showAveragedChart(matchingIndices);

        // Show averaged table in the left panel (replaces sample cards)
        QStringList devNames;
        QVector<QMap<QString, double>> devAvgs;
        for (const QString& devName : deviceOrder) {
            const DeviceAccum& acc = deviceMap[devName];
            devNames << devName;
            QMap<QString, double> avgs;
            for (const QString& m : kSensoryMetrics)
                avgs[m] = acc.sums.value(m, 0) / qMax(1, acc.count);
            devAvgs.append(avgs);
        }
        m_sensoryPanel->showAveragedTable(devNames, devAvgs);
    }
}

// ─── Tools ────────────────────────────────────────────────────────────────────
void MainWindow::onOpenDatabaseBrowser()
{
    DatabaseBrowserDialog dlg(m_db, this);
    if (dlg.exec() != QDialog::Accepted) return;

    // ── Sensory selection: switch to sensory mode and load sessions ──
    if (dlg.isSensorySelection()) {
        const QVector<int> sensoryIds = dlg.selectedSensoryIds();
        if (sensoryIds.isEmpty()) return;

        QVector<SensorySession> sessions;
        for (int id : sensoryIds) {
            SensorySession sess = m_db->loadSensorySession(id);
            if (!sess.samples.isEmpty())
                sessions.append(sess);
        }

        if (sessions.isEmpty()) {
            showError("Database Load", "Could not load sensory session(s) from the database.");
            return;
        }

        // Switch to sensory mode if not already
        if (!m_sensoryMode) {
            m_sensoryBtn->setChecked(true);  // triggers toggleSensoryMode
        }
        m_sensoryPanel->loadSessions(sessions);
        return;
    }

    // ── TPM file selection ──
    const QVector<int> ids = dlg.selectedFileIds();
    if (ids.isEmpty()) return;

    int loaded = 0;
    for (int id : ids) {
        FileResult result = m_db->loadFile(id);
        if (result.filePath.isEmpty()) continue;

        bool alreadyLoaded = false;
        for (int i = 0; i < m_loadedFiles.size(); ++i) {
            if (m_loadedFiles[i].filePath == result.filePath) {
                m_loadedFiles[i] = result;
                m_currentFileIndex = i;
                alreadyLoaded = true;
                break;
            }
        }

        if (!alreadyLoaded) {
            m_loadedFiles.append(result);
            m_currentFileIndex = m_loadedFiles.size() - 1;
        }
        ++loaded;
    }

    if (loaded > 0) {
        m_currentSheetIndex  = 0;
        m_currentSampleIndex = 0;
        populateFileTree();
        populateSheetCombo();
        displayCurrentSample();
        updateStatusBar(QString("Loaded %1 file(s) from database.").arg(loaded));
    } else {
        showError("Database Load", "No files could be loaded from the database.");
    }
}

// ─── Database ──────────────────────────────────────────────────────────────────
void MainWindow::onUpdateDatabase()
{
    int saved = 0, failed = 0;

    // ── Save TPM files ──
    for (const FileResult& fr : m_loadedFiles) {
        if (!m_modifiedFilePaths.contains(fr.filePath)) continue;
        if (m_db->saveFile(fr)) {
            m_modifiedFilePaths.remove(fr.filePath);
            ++saved;
        } else {
            ++failed;
        }
    }

    // ── Save sensory sessions ──
    int sensSaved = 0;
    if (m_sensoryMode && m_sensoryPanel) {
        auto sessions = m_sensoryPanel->allSessions();
        for (const SensorySession& sess : sessions) {
            if (sess.samples.isEmpty()) continue;
            if (m_db->saveSensorySession(sess))
                ++sensSaved;
            else
                ++failed;
        }
    }

    updateDbSyncIndicator();

    int total = saved + sensSaved;
    if (total == 0 && failed == 0) {
        updateStatusBar("Database already up to date.");
        return;
    }

    if (failed == 0) {
        QString msg = QString("Database updated (%1 file%2")
                          .arg(total).arg(total > 1 ? "s" : "");
        if (sensSaved > 0)
            msg += QString(", %1 sensory session%2")
                       .arg(sensSaved).arg(sensSaved > 1 ? "s" : "");
        msg += " saved).";
        updateStatusBar(msg);
    } else {
        showError("Database Error",
                  QString("%1 item(s) failed to save: %2").arg(failed).arg(m_db->lastError()));
    }
}

void MainWindow::markFileModified()
{
    FileResult* f = currentFile();
    if (!f) return;
    m_modifiedFilePaths.insert(f->filePath);
    updateDbSyncIndicator();
    // Restart debounce timer — saves 5 s after last change
    m_dbSaveTimer->start();
}

void MainWindow::updateDbSyncIndicator()
{
    if (!m_dbSyncLabel) return;
    bool isNas = m_db->currentPath().startsWith("//") ||
                 m_db->currentPath().startsWith("\\\\");
    QString prefix = isNas ? " NAS DB: " : " Local DB: ";
    if (m_modifiedFilePaths.isEmpty()) {
        m_dbSyncLabel->setText(prefix + "Synced ");
        m_dbSyncLabel->setStyleSheet("color: #2e7d32; font-weight: bold;");
    } else {
        int n = m_modifiedFilePaths.size();
        m_dbSyncLabel->setText(
            prefix + QString("%1 modified (Ctrl+U) ").arg(n));
        m_dbSyncLabel->setStyleSheet("color: #e65100; font-weight: bold;");
    }
}

bool MainWindow::promptSaveDatabase()
{
    if (m_modifiedFilePaths.isEmpty()) return true;

    int n = m_modifiedFilePaths.size();
    auto result = QMessageBox::question(
        this, "Unsaved Database Changes",
        QString("%1 file%2 %3 unsaved database changes.\n"
                "Would you like to update the database before closing?")
            .arg(n).arg(n > 1 ? "s" : "").arg(n > 1 ? "have" : "has"),
        QMessageBox::Yes | QMessageBox::No | QMessageBox::Cancel);

    if (result == QMessageBox::Cancel) return false;
    if (result == QMessageBox::Yes) onUpdateDatabase();
    return true;
}

QString MainWindow::lastBrowseDir() const
{
    if (!m_lastBrowseDir.isEmpty() && QDir(m_lastBrowseDir).exists())
        return m_lastBrowseDir;

    // Default: Weekly_Reports_Transfer
    const QString weeklyReports =
        "C:/Users/S1134987/OneDrive - Shenzhen Smoore Technology Limited"
        "/Shared Files Between Computers/Weekly_Reports_Transfer";
    if (QDir(weeklyReports).exists())
        return weeklyReports;

    // Fallback: user's Documents folder
    return QDir::homePath() + "/Documents";
}

void MainWindow::setLastBrowseDir(const QString& filePath)
{
    m_lastBrowseDir = QFileInfo(filePath).absolutePath();
}

// ─── Help ─────────────────────────────────────────────────────────────────────
void MainWindow::onHelp()
{
    QDesktopServices::openUrl(QUrl("https://github.com/SDR/DataViewerEnterprise"));
}
void MainWindow::onAbout()
{
    QMessageBox::about(this, "About DataViewer Enterprise",
        "<b>DataViewer Enterprise</b><br/>Version 1.0.0<br/><br/>"
        "Professional engineering data analysis tool.<br/>"
        "Built with Qt 6 and C++17.<br/><br/>"
        "\xC2\xA9 2025 SDR. All rights reserved.");
}

// ─── Drag & Drop ─────────────────────────────────────────────────────────────
void MainWindow::dragEnterEvent(QDragEnterEvent* e)
{
    if (e->mimeData()->hasUrls()) e->acceptProposedAction();
}
void MainWindow::dropEvent(QDropEvent* e)
{
    for (const QUrl& url : e->mimeData()->urls()) {
        QString p = url.toLocalFile();
        if (p.endsWith(".xlsx", Qt::CaseInsensitive) || p.endsWith(".xls", Qt::CaseInsensitive))
            loadFile(p);
    }
}

// ─── Settings ─────────────────────────────────────────────────────────────────
void MainWindow::restoreSettings()
{
    QSettings s("SDR", "DataViewerEnterprise");
    restoreGeometry(s.value("geometry").toByteArray());
    restoreState(s.value("windowState").toByteArray());

    m_inboxPath = s.value("inboxPath").toString();
    {
        // Re-detect if: no saved path, path no longer exists, or still on old DataViewer Inbox fallback
        bool isOldFallback = m_inboxPath.endsWith("DataViewer Inbox", Qt::CaseInsensitive);
        if (m_inboxPath.isEmpty() || !QDir(m_inboxPath).exists() || isOldFallback)
            m_inboxPath = defaultInboxPath();
    }
    if (!m_inboxPath.isEmpty() && QDir(m_inboxPath).exists())
        m_inboxWatcher->addPath(m_inboxPath);
    // Update badge on startup
    onInboxFolderChanged(m_inboxPath);
}
void MainWindow::saveSettings()
{
    QSettings s("SDR", "DataViewerEnterprise");
    s.setValue("geometry", saveGeometry());
    s.setValue("windowState", saveState());
    s.setValue("inboxPath", m_inboxPath);
}
void MainWindow::closeEvent(QCloseEvent* e)
{
    if (!promptSaveDatabase()) { e->ignore(); return; }
    saveSettings();
    e->accept();
}

// ─── Accessors ────────────────────────────────────────────────────────────────
FileResult* MainWindow::currentFile() const
{
    if (m_currentFileIndex < 0 || m_currentFileIndex >= m_loadedFiles.size()) return nullptr;
    return const_cast<FileResult*>(&m_loadedFiles[m_currentFileIndex]);
}
SheetResult* MainWindow::currentSheet() const
{
    const auto* f = currentFile();
    if (!f || m_currentSheetIndex < 0 || m_currentSheetIndex >= f->sheets.size()) return nullptr;
    return const_cast<SheetResult*>(&f->sheets[m_currentSheetIndex]);
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
void MainWindow::updateStatusBar(const QString& msg) { m_statusLabel->setText(msg); }
void MainWindow::setProgress(int pct, const QString& msg)
{
    m_progressBar->setVisible(true);
    m_progressBar->setValue(pct);
    m_statusLabel->setText(msg);
    if (pct >= 100) m_progressBar->setVisible(false);
    QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
}
void MainWindow::showError(const QString& t, const QString& m) { QMessageBox::critical(this, t, m); }
void MainWindow::showInfo(const QString& t, const QString& m)  { QMessageBox::information(this, t, m); }

QString MainWindow::resourcePath() const
{
    QStringList candidates = {
        QCoreApplication::applicationDirPath() + "/resources",
        QCoreApplication::applicationDirPath() + "/../resources",
        "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer/resources",
        "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/resources"
    };
    for (const QString& p : candidates)
        if (QDir(p).exists()) return p;
    return candidates.first();
}

QString MainWindow::templatePath() const
{
    static const QString kTemplateName =
        "Standardized Test Template - December 2025.xlsx";
    QStringList candidates = {
        resourcePath() + "/templates/" + kTemplateName,
        QCoreApplication::applicationDirPath() + "/resources/templates/" + kTemplateName,
        "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/resources/templates/" + kTemplateName
    };
    for (const QString& p : candidates)
        if (QFile::exists(p)) return p;
    return QString();
}

QString MainWindow::findPython() const
{
    if (m_pythonProbed) return m_cachedPython;
    m_pythonProbed = true;

    for (const QString& exe : { QString("python"), QString("python3"), QString("py") }) {
        QProcess p;
        p.start(exe, { "--version" });
        if (p.waitForFinished(5000) && p.exitCode() == 0) {
            m_cachedPython = exe;
            return m_cachedPython;
        }
    }
    return m_cachedPython;  // empty
}

QString MainWindow::runPython(const QString& python,
                              const QString& script,
                              const QStringList& args,
                              QString& errOut)
{
    // Write script to a temporary file
    QString tempDir = QStandardPaths::writableLocation(QStandardPaths::TempLocation);
    QString scriptPath = tempDir + "/dve_script.py";

    QFile f(scriptPath);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Text)) {
        errOut = "Cannot write temp script: " + scriptPath;
        return QString();
    }
    f.write(script.toUtf8());
    f.close();

    QProcess proc;
    QProcessEnvironment env = QProcessEnvironment::systemEnvironment();
    env.insert("PYTHONIOENCODING", "utf-8");
    env.insert("PYTHONUTF8", "1");
    proc.setProcessEnvironment(env);

    QStringList fullArgs;
    fullArgs << scriptPath;
    fullArgs << args;
    proc.start(python, fullArgs);

    if (!proc.waitForFinished(30000)) {
        proc.kill();
        QFile::remove(scriptPath);
        errOut = "Python script timed out.";
        return QString();
    }
    QFile::remove(scriptPath);

    if (proc.exitCode() != 0) {
        errOut = QString::fromUtf8(proc.readAllStandardError()).trimmed();
        return QString();
    }
    return QString::fromUtf8(proc.readAllStandardOutput()).trimmed();
}
QString MainWindow::defaultDbPath() const
{
    // 1. Synology Drive local sync folder (primary — offline-capable, per-user)
    QString synoDir = QDir::homePath() + "/SynologyDrive/SDR/Device Group/DVE_Database";
    if (QDir(synoDir).exists())
        return synoDir + "/dataviewer.db";

    // 2. Local AppData fallback (if Synology Drive is not installed/synced)
    QString dir = QStandardPaths::writableLocation(QStandardPaths::AppDataLocation);
    QDir().mkpath(dir);
    return dir + "/dataviewer.db";
}
// ─── Data Cleanup ─────────────────────────────────────────────────────────────

void MainWindow::onCleanData()
{
    const SheetResult* sheet = currentSheet();
    if (!sheet || !sheet->hasSamples()) {
        showInfo("No Data", "Load a file with samples first.");
        return;
    }

    DataCleanupDialog dlg(*sheet, currentSheetExclusions(), this);
    if (dlg.exec() != QDialog::Accepted) return;

    const QMap<int, QSet<int>> result = dlg.exclusions();
    // Store the returned exclusions (empty sets are removed)
    for (int si = 0; si < sheet->samples.size(); ++si) {
        const QString key = cleanupKey(m_currentFileIndex, m_currentSheetIndex, si);
        if (result.contains(si) && !result[si].isEmpty())
            m_excludedRows[key] = result[si];
        else
            m_excludedRows.remove(key);
    }

    displayCurrentSample();
}

void MainWindow::onResetCleanup()
{
    const SheetResult* sheet = currentSheet();
    if (!sheet) return;
    for (int si = 0; si < sheet->samples.size(); ++si)
        m_excludedRows.remove(cleanupKey(m_currentFileIndex, m_currentSheetIndex, si));
    displayCurrentSample();
}

QString MainWindow::cleanupKey(int fileIdx, int sheetIdx, int sampleIdx) const
{
    return QString("%1:%2:%3").arg(fileIdx).arg(sheetIdx).arg(sampleIdx);
}

QSet<int> MainWindow::exclusionsFor(int fileIdx, int sheetIdx, int sampleIdx) const
{
    return m_excludedRows.value(cleanupKey(fileIdx, sheetIdx, sampleIdx));
}

bool MainWindow::currentSheetHasCleanup() const
{
    const SheetResult* sheet = currentSheet();
    if (!sheet) return false;
    for (int si = 0; si < sheet->samples.size(); ++si) {
        if (!exclusionsFor(m_currentFileIndex, m_currentSheetIndex, si).isEmpty())
            return true;
    }
    return false;
}

QMap<int, QSet<int>> MainWindow::currentSheetExclusions() const
{
    QMap<int, QSet<int>> result;
    const SheetResult* sheet = currentSheet();
    if (!sheet) return result;
    for (int si = 0; si < sheet->samples.size(); ++si) {
        const QSet<int> ex = exclusionsFor(m_currentFileIndex, m_currentSheetIndex, si);
        if (!ex.isEmpty()) result[si] = ex;
    }
    return result;
}

void MainWindow::updateCleanupButtons()
{
    if (m_resetCleanupBtn)
        m_resetCleanupBtn->setEnabled(currentSheetHasCleanup());
}

SampleResult MainWindow::buildCleanedSample(const SampleResult& sr,
                                             const QSet<int>& excluded) const
{
    if (excluded.isEmpty()) return sr;

    SampleResult cleaned = sr;

    // Build a cleanup note listing every excluded data point
    QStringList parts;
    int rowNum = 0;
    for (int i = 0; i < sr.rows.size(); ++i) {
        const DataRow& dr = sr.rows[i];
        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
        ++rowNum;
        if (excluded.contains(i))
            parts << QString("Puff %1 (TPM=%2)")
                     .arg(dr.puffs, 0, 'f', 0)
                     .arg(dr.tpm,   0, 'f', 3);
    }
    if (!parts.isEmpty())
        cleaned.extra["cleanupNote"] =
            QString("Data cleanup: %1 row(s) excluded [%2]")
            .arg(excluded.size())
            .arg(parts.join(", "));

    // Remove excluded rows from the copy
    QVector<DataRow> kept;
    kept.reserve(sr.rows.size() - excluded.size());
    for (int i = 0; i < sr.rows.size(); ++i)
        if (!excluded.contains(i))
            kept.append(sr.rows[i]);
    cleaned.rows = kept;

    // Recalculate derived metrics from the surviving rows
    GenericSheetProcessor proc;
    proc.calculateMetrics(cleaned);
    return cleaned;
}

SheetResult MainWindow::buildCleanedSheet(const SheetResult& sheet,
                                          int fileIdx, int sheetIdx) const
{
    SheetResult cleaned = sheet;
    for (int si = 0; si < sheet.samples.size(); ++si) {
        const QSet<int> ex = exclusionsFor(fileIdx, sheetIdx, si);
        if (!ex.isEmpty())
            cleaned.samples[si] = buildCleanedSample(sheet.samples[si], ex);
    }
    GenericSheetProcessor proc;
    proc.computeSheetAggregates(cleaned);
    return cleaned;
}

FileResult MainWindow::buildCleanedFile(const FileResult& file) const
{
    if (m_currentFileIndex < 0) return file;
    FileResult cleaned = file;
    for (int si = 0; si < file.sheets.size(); ++si) {
        SheetResult& cs = cleaned.sheets[si];
        bool anyChanged = false;
        for (int sampleIdx = 0; sampleIdx < file.sheets[si].samples.size(); ++sampleIdx) {
            const QSet<int> ex = exclusionsFor(m_currentFileIndex, si, sampleIdx);
            if (!ex.isEmpty()) {
                cs.samples[sampleIdx] =
                    buildCleanedSample(file.sheets[si].samples[sampleIdx], ex);
                anyChanged = true;
            }
        }
        if (anyChanged && cs.hasSamples()) {
            GenericSheetProcessor proc;
            proc.computeSheetAggregates(cs);
        }
    }
    return cleaned;
}

// ─────────────────────────────────────────────────────────────────────────────

void MainWindow::recalculateSampleMetrics(SheetResult& sheet)
{
    GenericSheetProcessor proc;
    for (SampleResult& sr : sheet.samples)
        proc.calculateMetrics(sr);
    proc.computeSheetAggregates(sheet);
}

void MainWindow::onAddRow()
{
    SheetResult* sheet = currentSheet();
    FileResult*  file  = currentFile();
    if (!sheet || !file || sheet->samples.isEmpty()) return;
    if (m_currentSampleIndex >= sheet->samples.size()) return;

    SampleResult& sample = sheet->samples[m_currentSampleIndex];

    // Determine puff increment and last weights from visible rows
    double lastPuffs  = 0.0;
    double lastAfter  = 0.0;
    double increment  = 10.0;
    int    visCount   = 0;
    double prevPuffs  = 0.0;
    for (const DataRow& dr : sample.rows) {
        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
        if (visCount == 1) increment = dr.puffs - prevPuffs;
        prevPuffs = dr.puffs;
        lastPuffs = dr.puffs;
        lastAfter = dr.afterWeight;
        ++visCount;
    }
    if (visCount == 0) {
        // No visible rows yet — use puffs from first row if available
        if (!sample.rows.isEmpty()) lastPuffs = sample.rows.first().puffs;
    }
    if (increment <= 0.0) increment = 10.0;

    DataRow nr;
    nr.puffs        = lastPuffs + increment;
    nr.beforeWeight = lastAfter;      // template pattern: before = previous after
    nr.afterWeight  = lastAfter;      // user will edit this
    sample.rows.append(nr);

    recalculateSampleMetrics(*sheet);
    displayCurrentSample();
    markFileModified();

    // Write new row to Excel in a single batch call (data starts at Excel row 5)
    const int excelRow = static_cast<int>(sample.rows.size()) - 1 + 5;
    const int colBase  = m_currentSampleIndex * 12 + 1;
    writeCellsToExcel(file->filePath, sheet->sheetName, {
        { excelRow, colBase + 0, QString::number(nr.puffs, 'f', 0) },
        { excelRow, colBase + 1, QString::number(nr.beforeWeight, 'f', 4) },
        { excelRow, colBase + 2, QString::number(nr.afterWeight, 'f', 4) },
    });
}

void MainWindow::onRemoveRow()
{
    const int tableRow = m_dataTable->currentRow();
    if (tableRow < 0) return;

    SheetResult* sheet = currentSheet();
    FileResult*  file  = currentFile();
    if (!sheet || !file || sheet->samples.isEmpty()) return;
    if (m_currentSampleIndex >= sheet->samples.size()) return;

    SampleResult& sample = sheet->samples[m_currentSampleIndex];

    // Map table row → data row index (skip rows where either weight is 0)
    int dataRow = -1, tIdx = 0;
    for (int i = 0; i < sample.rows.size(); ++i) {
        if (sample.rows[i].beforeWeight == 0.0 || sample.rows[i].afterWeight == 0.0)
            continue;
        if (tIdx == tableRow) { dataRow = i; break; }
        ++tIdx;
    }
    if (dataRow < 0) return;

    const int excelRow = dataRow + 5;
    sample.rows.removeAt(dataRow);

    recalculateSampleMetrics(*sheet);
    displayCurrentSample();
    markFileModified();

    deleteRowFromExcel(file->filePath, sheet->sheetName, excelRow);
}

void MainWindow::deleteRowFromExcel(const QString& filePath,
                                     const QString& sheetName,
                                     int excelRow1)
{
    const QString python = findPython();
    if (python.isEmpty()) return;

    static const char* kDeleteRow = R"PY(
import sys
from openpyxl import load_workbook
path, sheet, row_s = sys.argv[1], sys.argv[2], sys.argv[3]
wb = load_workbook(path)
ws = wb[sheet]
ws.delete_rows(int(row_s), 1)
wb.save(path)
print("OK")
)PY";

    QString err;
    runPython(python, kDeleteRow,
              { filePath, sheetName, QString::number(excelRow1) },
              err);
}

void MainWindow::writeCellToExcel(const QString& filePath, const QString& sheetName,
                                   int excelRow1, int excelCol1, const QString& value)
{
    writeCellsToExcel(filePath, sheetName, { { excelRow1, excelCol1, value } });
}

void MainWindow::writeCellsToExcel(const QString& filePath, const QString& sheetName,
                                    const QVector<CellWrite>& cells)
{
    if (cells.isEmpty()) return;
    const QString python = findPython();
    if (python.isEmpty()) return;

    // Batch Python script: writes all cells in one load-save cycle.
    // Arguments after path+sheet are triplets: row col value row col value ...
    static const char* kWriteCells = R"PY(
import sys
from openpyxl import load_workbook
path, sheet = sys.argv[1], sys.argv[2]
wb = load_workbook(path)
ws = wb[sheet]
args = sys.argv[3:]
i = 0
while i + 2 < len(args):
    r, c, val = int(args[i]), int(args[i+1]), args[i+2]
    try:
        ws.cell(row=r, column=c).value = float(val) if val.strip() else None
    except ValueError:
        ws.cell(row=r, column=c).value = val if val.strip() else None
    i += 3
wb.save(path)
print("OK")
)PY";

    QStringList args = { filePath, sheetName };
    for (const CellWrite& cw : cells) {
        args << QString::number(cw.row) << QString::number(cw.col) << cw.value;
    }

    QString err;
    runPython(python, kWriteCells, args, err);
}

void MainWindow::queueExcelWrite(const QString& filePath, const QString& sheetName,
                                  int excelRow1, int excelCol1, const QString& value)
{
    // If file/sheet changed from what's pending, flush first
    if (!m_pendingWrites.isEmpty() &&
        (m_pendingWriteFile != filePath || m_pendingWriteSheet != sheetName)) {
        flushExcelWrites();
    }

    m_pendingWriteFile  = filePath;
    m_pendingWriteSheet = sheetName;

    // Overwrite any existing write to the same cell
    for (CellWrite& cw : m_pendingWrites) {
        if (cw.row == excelRow1 && cw.col == excelCol1) {
            cw.value = value;
            m_excelWriteTimer->start();  // restart timer
            return;
        }
    }
    m_pendingWrites.append({ excelRow1, excelCol1, value });
    m_excelWriteTimer->start();  // restart timer
}

void MainWindow::flushExcelWrites()
{
    if (m_pendingWrites.isEmpty()) return;

    writeCellsToExcel(m_pendingWriteFile, m_pendingWriteSheet, m_pendingWrites);
    m_pendingWrites.clear();
}

QStringList MainWindow::dataTableHeaders()
{
    return {"Puffs","Before (g)","After (g)","Pressure","Resistance",
            "Smell","Clog","Notes","TPM (mg/puff)","TPM Pwr Density","Variation (%)","Oil Consumed (mg)"};
}

// ─── Sample image loading / viewing ───────────────────────────────────────────

void MainWindow::onLoadImages()
{
    // ── Sensory mode: load images for the current sensory session ──
    if (m_sensoryMode && m_sensoryPanel) {
        SensorySession* sess = m_sensoryPanel->currentSession();
        if (!sess) return;

        QStringList paths = QFileDialog::getOpenFileNames(
            this, "Load Session Images", lastBrowseDir(),
            "Images (*.png *.jpg *.jpeg *.bmp *.gif);;All Files (*)");
        if (paths.isEmpty()) return;
        setLastBrowseDir(paths.first());

        for (const QString& p : paths) {
            if (!sess->imagePaths.contains(p))
                sess->imagePaths.append(p);
        }
        updateImageButton();
        return;
    }

    // ── TPM mode ──
    SheetResult* sheet = currentSheet();
    if (!sheet || m_currentSampleIndex < 0 ||
        m_currentSampleIndex >= sheet->samples.size())
        return;

    QStringList paths = QFileDialog::getOpenFileNames(
        this, "Load Sample Images", lastBrowseDir(),
        "Images (*.png *.jpg *.jpeg *.bmp *.gif);;All Files (*)");

    if (paths.isEmpty()) return;
    setLastBrowseDir(paths.first());

    SampleResult& sample = sheet->samples[m_currentSampleIndex];
    for (const QString& p : paths) {
        if (!sample.imagePaths.contains(p))
            sample.imagePaths.append(p);
    }

    updateImageButton();
    markFileModified();
}

void MainWindow::onViewImages()
{
    // ── Sensory mode: view images for the current sensory session ──
    if (m_sensoryMode && m_sensoryPanel) {
        SensorySession* sess = m_sensoryPanel->currentSession();
        if (!sess || sess->imagePaths.isEmpty()) return;

        QString displayName = m_sensoryPanel->sessionLabel(*sess);
        ImageViewDialog dlg(sess->imagePaths, sess->imageLayouts, sess->imageCrops, displayName, this);
        if (dlg.exec() == QDialog::Accepted) {
            sess->imagePaths   = dlg.imagePaths();
            sess->imageLayouts = dlg.imageLayouts();
            sess->imageCrops   = dlg.imageCrops();
            updateImageButton();
        }
        return;
    }

    // ── TPM mode ──
    SheetResult* sheet = currentSheet();
    if (!sheet || m_currentSampleIndex < 0 ||
        m_currentSampleIndex >= sheet->samples.size())
        return;

    SampleResult& sample = sheet->samples[m_currentSampleIndex];
    if (sample.imagePaths.isEmpty()) return;

    QString displayName = sample.sampleName.isEmpty() ? sample.sampleID : sample.sampleName;
    ImageViewDialog dlg(sample.imagePaths, sample.imageLayouts, sample.imageCrops, displayName, this);
    if (dlg.exec() == QDialog::Accepted) {
        sample.imagePaths   = dlg.imagePaths();
        sample.imageLayouts = dlg.imageLayouts();
        sample.imageCrops   = dlg.imageCrops();
        updateImageButton();
        markFileModified();
    }
}

void MainWindow::updateImageButton()
{
    if (!m_loadImagesBtn || !m_viewImagesBtn) return;

    int count = 0;
    if (m_sensoryMode && m_sensoryPanel) {
        SensorySession* sess = m_sensoryPanel->currentSession();
        if (sess) count = sess->imagePaths.size();
    } else {
        SheetResult* sheet = currentSheet();
        if (sheet && m_currentSampleIndex >= 0 &&
            m_currentSampleIndex < sheet->samples.size())
            count = sheet->samples[m_currentSampleIndex].imagePaths.size();
    }

    m_viewImagesBtn->setText(QString("View Images (%1)").arg(count));
    m_viewImagesBtn->setEnabled(count > 0);
}

// ─── Image Inbox ──────────────────────────────────────────────────────────────

QString MainWindow::defaultInboxPath() const
{
    QString docs = QStandardPaths::writableLocation(QStandardPaths::DocumentsLocation);
    QString monthFolder = QDate::currentDate().toString("yyyy-MM");

    // ── 1. WeCom Pro: [root]\WXWorkLocalPro\[userid]\Cache\Image\YYYY-MM ──
    // Search common install locations for WXWorkLocalPro
    QString home = QStandardPaths::writableLocation(QStandardPaths::HomeLocation);
    for (const QString& searchRoot : { docs, home }) {
        QDir weComRoot(searchRoot + "/WXWorkLocalPro");
        if (!weComRoot.exists()) continue;
        QStringList profiles = weComRoot.entryList(QDir::Dirs | QDir::NoDotAndDotDot);
        for (const QString& profile : profiles) {
            QDir imgDir(weComRoot.filePath(profile + "/Cache/Image"));
            if (!imgDir.exists()) continue;
            // Prefer current month subfolder; fall back to Image\ parent
            if (imgDir.exists(monthFolder))
                return imgDir.filePath(monthFolder);
            return imgDir.absolutePath();
        }
    }

    // ── 2. WhatsApp UWP ──────────────────────────────────────────────────────
    QString localApp = QProcessEnvironment::systemEnvironment().value("LOCALAPPDATA");
    if (!localApp.isEmpty()) {
        QString waPath = localApp +
            "/Packages/5319275A.WhatsAppDesktop_cv1g1gvanyjgm"
            "/LocalCache/Roaming/WhatsApp/Media/WhatsApp Images";
        if (QDir(waPath).exists()) return waPath;
    }

    // ── 3. WhatsApp in Pictures ───────────────────────────────────────────────
    QString pics = QStandardPaths::writableLocation(QStandardPaths::PicturesLocation);
    QString waPics = pics + "/WhatsApp Images";
    if (QDir(waPics).exists()) return waPics;

    return {};
}

void MainWindow::onOpenImageInbox()
{
    // ── Sensory mode: open inbox dialog with session targets ──
    if (m_sensoryMode && m_sensoryPanel) {
        QVector<SensorySession> sessions = m_sensoryPanel->allSessions();
        if (sessions.isEmpty()) return;

        // Build session labels: "testTitle - testerName - date"
        QStringList labels;
        for (const SensorySession& s : sessions)
            labels << m_sensoryPanel->sessionLabel(s);

        int curIdx = m_sensoryPanel->currentSessionIndex();

        ImageInboxDialog dlg(labels, curIdx, m_inboxPath, this);
        connect(&dlg, &ImageInboxDialog::watchFolderChanged,
                this, &MainWindow::onInboxFolderChanged);
        if (dlg.exec() != QDialog::Accepted) return;

        // Apply assignments to sensory sessions.
        // For each assignment, temporarily select the target session so that
        // currentSession() returns a live pointer we can modify.
        int origIdx = m_sensoryPanel->currentSessionIndex();
        int sessionCount = sessions.size();  // snapshot count (stable)

        for (const auto& a : dlg.assignments()) {
            int sessIdx = a.sampleIdx;
            if (sessIdx < 0 || sessIdx >= sessionCount) continue;

            // Select the target session (saves current first internally)
            m_sensoryPanel->selectSession(sessIdx);
            SensorySession* sess = m_sensoryPanel->currentSession();
            if (!sess) continue;

            for (const QString& p : a.imagePaths)
                if (!sess->imagePaths.contains(p))
                    sess->imagePaths.append(p);
        }

        // Restore the originally selected session
        if (origIdx >= 0 && origIdx < sessionCount)
            m_sensoryPanel->selectSession(origIdx);

        // Update watch folder if changed
        if (!dlg.watchFolder().isEmpty() && dlg.watchFolder() != m_inboxPath)
            onInboxFolderChanged(dlg.watchFolder());

        updateImageButton();
        return;
    }

    // ── TPM mode: full inbox dialog ──
    ImageInboxDialog dlg(m_loadedFiles,
                         m_currentFileIndex, m_currentSheetIndex, m_currentSampleIndex,
                         m_inboxPath, this);
    connect(&dlg, &ImageInboxDialog::watchFolderChanged,
            this, &MainWindow::onInboxFolderChanged);
    if (dlg.exec() != QDialog::Accepted) return;

    // Apply assignments
    for (const auto& a : dlg.assignments()) {
        if (a.fileIdx < 0 || a.fileIdx >= m_loadedFiles.size()) continue;
        FileResult& file = m_loadedFiles[a.fileIdx];
        if (a.sheetIdx < 0 || a.sheetIdx >= file.sheets.size()) continue;
        SheetResult& sheet = file.sheets[a.sheetIdx];
        if (a.sampleIdx < 0 || a.sampleIdx >= sheet.samples.size()) continue;
        SampleResult& sample = sheet.samples[a.sampleIdx];
        for (const QString& p : a.imagePaths)
            if (!sample.imagePaths.contains(p))
                sample.imagePaths.append(p);
        markFileModified();
    }

    // Update watch folder if changed
    if (!dlg.watchFolder().isEmpty() && dlg.watchFolder() != m_inboxPath)
        onInboxFolderChanged(dlg.watchFolder());

    displayCurrentSample();
}

void MainWindow::onInboxFolderChanged(const QString& path)
{
    if (!m_inboxWatcher->directories().isEmpty())
        m_inboxWatcher->removePaths(m_inboxWatcher->directories());
    m_inboxPath = path;
    if (!path.isEmpty() && QDir(path).exists())
        m_inboxWatcher->addPath(path);
    // Update inbox button badge (count image files in folder)
    if (m_inboxBtn) {
        QDir dir(path);
        QStringList imgs = dir.entryList(
            {"*.jpg", "*.jpeg", "*.png", "*.bmp", "*.webp", "*.gif"}, QDir::Files);
        int n = imgs.size();
        m_inboxBtn->setText(n > 0 ? QString("Images\n(%1)").arg(n) : "Images");
    }
}

} // namespace DVE
