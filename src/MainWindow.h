#pragma once

#include <QMainWindow>
#include <QDockWidget>
#include <QSplitter>
#include <QTabWidget>
#include <QTableWidget>
#include <QTreeWidget>
#include <QLabel>
#include <QProgressBar>
#include <QComboBox>
#include <QToolButton>
#include <QPushButton>
#include <QStatusBar>
#include <QMap>
#include <QTimer>
#include <QSet>
#include <QFuture>
#include <QFutureWatcher>

#include "ExcelReader.h"
#include "pipeline/ReportData.h"
#include "ui/NewFileDialog.h"
#include "ui/HeaderEditDialog.h"
#include "ui/ImageViewDialog.h"
#include "pipeline/DataProcessor.h"
#include "reporting/ReportGenerator.h"
#include "plotting/PlotWidget.h"
#include "widgets/RibbonWidget.h"
#include "database/DatabaseManager.h"
#include "ui/DatabaseBrowserDialog.h"
#include "ui/DataCleanupDialog.h"
#include "ui/ImageInboxDialog.h"
#include <QFileSystemWatcher>
#include <QStackedWidget>
#include <QListWidget>
#include "ui/DetailedSensoryPanel.h"
#include "utils/UpdateChecker.h"

namespace DVE {

// ─── Forward decls ────────────────────────────────────────────────────────────
class PlotWidget;
class SensoryPanel;
class DetailedSensoryPanel;

// ─── Main application window ──────────────────────────────────────────────────
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

    /// Load a file by path (used by CLI argument handling in main.cpp)
    void loadFile(const QString& path);

protected:
    void closeEvent(QCloseEvent* event) override;
    void dragEnterEvent(QDragEnterEvent* event) override;
    void dropEvent(QDropEvent* event) override;

private slots:
    // ── File menu ──
    void onNewFile();
    void onLoadFile();
    void onCloseFile();
    void onRecentFileTriggered(const QString& path);

    // ── Reports ──
    void onGenerateTestReport();
    void onGenerateFullReport();
    // ── Navigation ──
    void onFileSelected(int index);
    void onSheetSelected(int index);
    void onPrevSample();
    void onNextSample();

    // ── Background worker ──
    void onFileLoadFinished();
    void onReportFinished(bool success, const QString& path);

    // ── View ──
    void onViewDataTable();
    void onViewPlots();
    void onViewBoth();
    void onZoomIn();
    void onZoomOut();
    void onFitToWindow();

    // ── Tools / Sensory mode ──
    void toggleSensoryMode(bool checked);
    void toggleDetailedSensoryMode(bool checked);
    void onOpenDatabaseBrowser();

    // ── Data cleanup ──
    void onCleanData();
    void onResetCleanup();

    // ── Database ──
    void onUpdateDatabase();

    // ── Edit headers ──
    void onEditHeaders();
    void onTableCellChanged(int row, int col);
    void onPropCellChanged(int row, int col);
    void onAddRow();
    void onRemoveRow();

    // ── Help ──
    void onHelp();
    void onAbout();

    // ── Sample images ──
    void onLoadImages();
    void onViewImages();

    // ── Image Inbox ──
    void onOpenImageInbox();
    void onInboxFolderChanged(const QString& path);

    // ── Document Translator ──
    void onLaunchTranslator();

private:
    static bool writeTranslatorConfig(const QString& apiKey);
    QString loadApiKey();
    bool    saveApiKey(const QString& key);

    // ── File type detection & routing ──────────────────────────────────────
    enum class FileType { TPM, Sensory, DetailedSensory, Unknown };
    FileType detectFileType(const QString& path) const;
    void     routeFile(const QString& path);

    // ── Setup ────────────────────────────────────────────────────────────────
    void setupUI();
    void setupRibbon();
    void setupCentralWidget();
    void setupDockPanels();
    void setupStatusBar();
    void setupConnections();
    void restoreSettings();
    void saveSettings();

    // ── Ribbon tabs ──────────────────────────────────────────────────────────
    void buildHomeTab(RibbonTab* tab);
    void buildReportsTab(RibbonTab* tab);
    void buildViewTab(RibbonTab* tab);
    void buildToolsTab(RibbonTab* tab);

    // ── Left dock: File/Sheet browser ────────────────────────────────────────
    QDockWidget*  m_fileDock;
    QTreeWidget*  m_fileTree;          // shows loaded files and their sheets
    QComboBox*    m_fileCombo;         // quick-select file
    QComboBox*    m_sheetCombo;        // quick-select sheet

    // ── Right dock: Sample Properties ────────────────────────────────────────
    QDockWidget*  m_propDock;
    QWidget*      m_propPanel;
    QTableWidget* m_propTable;         // inline-editable sample properties
    QPushButton*  m_loadImagesBtn = nullptr;
    QPushButton*  m_viewImagesBtn = nullptr;

    // ── Central area ────────────────────────────────────────────────────────
    QStackedWidget* m_centralStack = nullptr;   // index 0=TPM, index 1=sensory
    QSplitter*      m_centralSplitter;          // TPM: table + plot

    // ── Data table panel ─────────────────────────────────────────────────────
    QWidget*      m_tablePanel;
    QTableWidget* m_dataTable;
    QWidget*      m_sampleNavBar;
    QPushButton*  m_prevBtn;
    QPushButton*  m_nextBtn;
    QPushButton*  m_addRowBtn    = nullptr;
    QPushButton*  m_removeRowBtn = nullptr;
    QLabel*       m_sampleCountLabel;

    // ── Plot panel ───────────────────────────────────────────────────────────
    PlotWidget*   m_plotWidget;

    // ── Ribbon ───────────────────────────────────────────────────────────────
    RibbonWidget*  m_ribbon;
    QToolButton*   m_resetCleanupBtn = nullptr;  // enabled only when cleanup is active
    QToolButton*   m_inboxBtn        = nullptr;

    // ── Ribbon button references (for mode switching) ────────────────────────
    // Home tab — TPM buttons
    QToolButton*   m_homeNewBtn   = nullptr;
    QToolButton*   m_homeLoadBtn  = nullptr;
    QToolButton*   m_homeCloseBtn = nullptr;
    // Home tab — Sensory buttons (hidden in TPM mode)
    QToolButton*   m_homeSensNewBtn   = nullptr;
    QToolButton*   m_homeSensSaveBtn  = nullptr;
    QToolButton*   m_homeSensLoadXlBtn = nullptr;
    QToolButton*   m_homeSensCloseBtn  = nullptr;
    // Home tab — Detailed Sensory buttons (hidden in TPM/Sensory mode)
    QToolButton*   m_homeDetSensNewBtn    = nullptr;
    QToolButton*   m_homeDetSensSaveBtn   = nullptr;
    QToolButton*   m_homeDetSensLoadXlBtn = nullptr;
    QToolButton*   m_homeDetSensCloseBtn  = nullptr;
    // Reports tab
    QToolButton*   m_reportBtn1 = nullptr;  // "Test Report" / "Sensory Report"
    QToolButton*   m_reportBtn2 = nullptr;  // "Full Report" (hidden in sensory mode)
    // Reports tab — groups to show/hide
    RibbonGroup*   m_cleanupGroup = nullptr;
    // Tools tab
    QToolButton*   m_sensoryBtn = nullptr;  // checkable toggle
    QToolButton*   m_detailedSensoryBtn = nullptr;  // checkable toggle

    // ── Status bar ───────────────────────────────────────────────────────────
    QLabel*       m_statusLabel;
    QProgressBar* m_progressBar;
    QLabel*       m_fileInfoLabel;
    QLabel*       m_dbSyncLabel = nullptr;

    // ── Data ─────────────────────────────────────────────────────────────────
    DataProcessor*    m_processor = nullptr;
    ReportGenerator*  m_reportGen;
    DatabaseManager*  m_db;

    QVector<FileResult> m_loadedFiles;   // all loaded files
    QSet<QString>       m_modifiedFilePaths;  // files with unsaved edits

    // Data cleanup: key = "fileIdx:sheetIdx:sampleIdx" → set of excluded row indices
    QMap<QString, QSet<int>> m_excludedRows;
    int m_currentFileIndex    = -1;
    int m_currentSheetIndex   = -1;
    int m_currentSampleIndex  = 0;

    FileResult*  currentFile()  const;
    SheetResult* currentSheet() const;

    // ── Auto-updater ─────────────────────────────────────────────────────────
    UpdateChecker*      m_updateChecker = nullptr;

    // ── Image Inbox ──────────────────────────────────────────────────────────
    QFileSystemWatcher* m_inboxWatcher = nullptr;
    QString             m_inboxPath;

    // ── Background load ──────────────────────────────────────────────────────
    QFutureWatcher<FileResult>* m_loadWatcher;
    bool m_loading = false;
    QStringList m_pendingLoadPaths;

    // ── Helpers ──────────────────────────────────────────────────────────────
    void populateFileTree();
    void populateSheetCombo();
    void displayCurrentSample();
    void updateSampleNav();
    void updateProperties(const SampleResult& sample);
    void recalculateSampleMetrics(SheetResult& sheet);
    void writeCellToExcel(const QString& filePath, const QString& sheetName,
                          int excelRow1, int excelCol1, const QString& value);
    struct CellWrite { int row; int col; QString value; };
    void writeCellsToExcel(const QString& filePath, const QString& sheetName,
                           const QVector<CellWrite>& cells);
    void queueExcelWrite(const QString& filePath, const QString& sheetName,
                         int excelRow1, int excelCol1, const QString& value);
    void flushExcelWrites();
    void deleteRowFromExcel(const QString& filePath, const QString& sheetName,
                            int excelRow1);
    void updateStatusBar(const QString& msg);
    void setProgress(int pct, const QString& msg);
    void showError(const QString& title, const QString& msg);
    void showInfo(const QString& title, const QString& msg);

    void updateImageButton();
    void markFileModified();
    void updateDbSyncIndicator();

    // ── Data cleanup helpers ──────────────────────────────────────────────────
    QString cleanupKey(int fileIdx, int sheetIdx, int sampleIdx) const;
    QSet<int> exclusionsFor(int fileIdx, int sheetIdx, int sampleIdx) const;
    bool currentSheetHasCleanup() const;
    QMap<int, QSet<int>> currentSheetExclusions() const;
    void updateCleanupButtons();
    SampleResult buildCleanedSample(const SampleResult& sr, const QSet<int>& excluded) const;
    SheetResult  buildCleanedSheet(const SheetResult& sheet, int fileIdx, int sheetIdx) const;
    FileResult   buildCleanedFile(const FileResult& file) const;

    QString resourcePath() const;
    QString defaultDbPath();
    QString defaultInboxPath() const;
    QString templatePath() const;
    QString findPython() const;
    mutable QString m_cachedPython;
    mutable bool    m_pythonProbed = false;

    // Run a one-shot Python script (writes to temp file, executes, returns stdout).
    // Returns empty string on error and sets lastError via errOut.
    static QString runPython(const QString& python,
                             const QString& script,
                             const QStringList& args,
                             QString& errOut);

    // Remembers last-used directory for file dialogs
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
    mutable QString m_lastBrowseDir;

    // ── Debounced Excel write queue ──────────────────────────────────────────
    QTimer*              m_excelWriteTimer = nullptr;
    QTimer*              m_dbSaveTimer     = nullptr;  // auto-save after inactivity
    QString              m_pendingWriteFile;
    QString              m_pendingWriteSheet;
    QVector<CellWrite>   m_pendingWrites;

    // Prompt user to save DB if there are unsaved changes; returns false if user cancels
    bool promptSaveDatabase();

    // Column headers for data table
    static QStringList dataTableHeaders();

    // Returns desiredPath if it doesn't exist; otherwise appends (2), (3), …
    static QString uniqueFilename(const QString& desiredPath);

    // ── Sensory mode ────────────────────────────────────────────────────────
    bool            m_sensoryMode = false;
    bool            m_sensorySessionsDirty = false;
    SensoryPanel*   m_sensoryPanel = nullptr;
    bool                    m_detailedSensoryMode = false;
    bool                    m_detailedSensorySessionsDirty = false;
    DetailedSensoryPanel*   m_detailedSensoryPanel = nullptr;
    QListWidget*            m_detailedSensoryNav = nullptr;

    // Navigator stack inside left dock (index 0 = file tree, index 1 = sensory sessions)
    QStackedWidget* m_navStack     = nullptr;
    QListWidget*    m_sensoryNav   = nullptr;
    QLabel*         m_navLabel     = nullptr;   // "Loaded Files:" / "Sessions:"

    void initSensoryPanel();
    void updateRibbonForMode();
    void refreshSensoryNavigator();
    void updateSensoryProperties();
    void initDetailedSensoryPanel();
    void refreshDetailedSensoryNavigator();
    void updateDetailedSensoryProperties();

    // ── Test Averages panel (sensory mode only) ──────────────────────────────
    QWidget*      m_testAvgPanel    = nullptr;
    QListWidget*  m_testAvgList     = nullptr;   // unique test titles
    QTableWidget* m_testAvgTable    = nullptr;   // metric averages
    QLabel*       m_testAvgAssessors = nullptr;
    QLabel*       m_testAvgTesters   = nullptr;
    QLabel*       m_testAvgCount     = nullptr;
    void refreshSensoryAverages();
    void onTestAvgSelectionChanged();
};

} // namespace DVE
