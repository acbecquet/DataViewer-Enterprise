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

namespace DVE {

// ─── Forward decls ────────────────────────────────────────────────────────────
class PlotWidget;

// ─── Main application window ──────────────────────────────────────────────────
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

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

    // ── Tools ──
    void onOpenSensory();
    void onOpenDatabaseBrowser();
    void onExportToExcel();

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

private:
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

    // ── Central splitter ─────────────────────────────────────────────────────
    QSplitter*    m_centralSplitter;

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
    RibbonWidget* m_ribbon;

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
    int m_currentFileIndex    = -1;
    int m_currentSheetIndex   = -1;
    int m_currentSampleIndex  = 0;

    FileResult*  currentFile()  const;
    SheetResult* currentSheet() const;

    // ── Background load ──────────────────────────────────────────────────────
    QFutureWatcher<FileResult>* m_loadWatcher;
    bool m_loading = false;

    // ── Helpers ──────────────────────────────────────────────────────────────
    void loadFile(const QString& path);
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
    void deleteRowFromExcel(const QString& filePath, const QString& sheetName,
                            int excelRow1);
    void updateStatusBar(const QString& msg);
    void setProgress(int pct, const QString& msg);
    void showError(const QString& title, const QString& msg);
    void showInfo(const QString& title, const QString& msg);

    void updateImageButton();
    void markFileModified();
    void updateDbSyncIndicator();

    QString resourcePath() const;
    QString defaultDbPath() const;
    QString templatePath() const;
    QString findPython() const;

    // Run a one-shot Python script (writes to temp file, executes, returns stdout).
    // Returns empty string on error and sets lastError via errOut.
    static QString runPython(const QString& python,
                             const QString& script,
                             const QStringList& args,
                             QString& errOut);

    // Column headers for data table
    static QStringList dataTableHeaders();
};

} // namespace DVE
