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

#include <QMap>
#include <QTimer>
#include <QSet>
#include <QFuture>
#include <QFutureWatcher>
#include <QVector>
#include <QDateTime>

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
#include "pipeline/SensoryData.h"
#include "utils/UpdateChecker.h"

namespace DVE {

// ─── Forward decls ────────────────────────────────────────────────────────────
class PlotWidget;
class SensoryPanel;
class DetailedSensoryPanel;
class IdentityManager;
class PostgresConnection;
class NotificationListener;
class PresenceManager;
class LiveSync;
class PresenceDotsDelegate;
class CellFocusDelegate;
class RegimeComboDelegate;
class PresenceAvatarBar;
class RowDeletedBanner;
class OfflineSnapshot;
class ConnectionMonitor;
class OfflineBanner;
struct PresenceChange;
struct RowChange;

// ─── Main application window ──────────────────────────────────────────────────
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget* parent = nullptr);
    ~MainWindow() override;

    enum FileStatus { FileStatusOk, FileStatusModified, FileStatusClosed };
    enum DbStatus   { DbStatusOk, DbStatusModified, DbStatusDisconnected };

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

    // ── Export to Excel (manual flush of debounced write-back) ──
    void onExportToExcelTriggered();

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

    // ── Presence UI (Plan B Phase 5) ──
    // Refresh nav-item dots + avatar bar for a given resource. Called both
    // by NotificationListener::presenceChanged (live updates) and locally
    // after activating a new resource (initial paint).
    void refreshPresenceFor(const QString& resourceType, qint64 resourceId);
    // Iterate every navigator and seed presence data for all known resources.
    // Used after populating the nav widgets so dots show on first paint.
    void refreshAllPresence();
    // Drop the current resource activation (mode switch, file close, etc.).
    // Empties the avatar bar and tells PresenceManager to deactivate so other
    // clients see this user leave the resource.
    void clearActivePresence();

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

    // ── Status bar helpers (v2.0.9) ──────────────────────────────────────────
    void buildStatusBar();
    void setStatusFile(const QString& text, FileStatus status);
    void setStatusDb(const QString& text, DbStatus status);
    void setStatusBreadcrumb(const QStringList& segments);

    // ── Ribbon tabs ──────────────────────────────────────────────────────────
    void buildHomeTab(RibbonTab* tab);
    void buildReportsTab(RibbonTab* tab);
    void buildViewTab(RibbonTab* tab);
    void buildToolsTab(RibbonTab* tab);
    void buildSettingsTab(RibbonTab* tab);

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

    // ── Status bar (v2.0.9 redesign) ────────────────────────────────────────
    QWidget*      m_statusBarWidget  = nullptr;  // custom QWidget replacing QStatusBar
    QLabel*       m_statusFileDot    = nullptr;  // ● colored by file state
    QLabel*       m_statusFileText   = nullptr;  // "File closed", "Loaded: …"
    QLabel*       m_statusDbDot      = nullptr;  // ● colored by db state
    QLabel*       m_statusDbText     = nullptr;  // "Local DB: Synced" / "Disconnected"
    QLabel*       m_statusBreadcrumb = nullptr;  // filename › sheet › sample
    QStringList   m_lastBreadcrumbSegments;      // re-render on breakpoint change
    QProgressBar* m_progressBar      = nullptr;  // retained for setProgress()
    // Legacy aliases kept so call sites compile while being migrated:
    QLabel*       m_dbSyncLabel      = nullptr;  // replaced by m_statusDbText

    // ── Data ─────────────────────────────────────────────────────────────────
    DataProcessor*    m_processor = nullptr;
    ReportGenerator*  m_reportGen;
    DatabaseManager*  m_db;

    QVector<FileResult> m_loadedFiles;   // all loaded files
    QSet<QString>       m_modifiedFilePaths;  // files with unsaved edits

    // v2.0.2 H6: belt-and-suspenders echo guard. onRemoteCellChanged
    // already wraps setText in a QSignalBlocker, but other paths that
    // synthesize cell text changes during remote application can leak
    // through. The flag is set true while a remote write is being
    // applied; onDataTableItemChanged early-returns when it sees it.
    bool m_applyingRemote = false;

    // Data cleanup: key = "fileIdx:sheetIdx:sampleIdx" → set of excluded row indices
    QMap<QString, QSet<int>> m_excludedRows;
    int m_currentFileIndex    = -1;
    int m_currentSheetIndex   = -1;
    int m_currentSampleIndex  = 0;

    FileResult*  currentFile()  const;
    SheetResult* currentSheet() const;

    // ── Auto-updater ─────────────────────────────────────────────────────────
    UpdateChecker*      m_updateChecker = nullptr;

    // ── Identity manager ──────────────────────────────────────────────────────
    DVE::IdentityManager* m_identity = nullptr;

    // ── Postgres-backed concurrency stack ────────────────────────────────────
    // Owned by MainWindow (this is parent); destruction order is reverse
    // construction. m_pgConn is a SEPARATE connection from DatabaseManager's
    // internal m_pg — NOTIFY/heartbeat workload is isolated from main queries.
    DVE::PostgresConnection*    m_pgConn   = nullptr;
    DVE::NotificationListener*  m_notify   = nullptr;
    DVE::PresenceManager*       m_presence = nullptr;
    DVE::LiveSync*              m_liveSync = nullptr;

    // ── Presence UI (Plan B Phase 5) ─────────────────────────────────────────
    // Delegate is shared by m_fileTree, m_sensoryNav, m_detailedSensoryNav.
    // Cheap (one QObject) and keeps all three widgets consistent.
    DVE::PresenceDotsDelegate*  m_presenceDelegate = nullptr;
    // v2.0.1: paints remote-focus border + name flag and remote-change
    // flash on the TPM data table cells. One per MainWindow.
    DVE::CellFocusDelegate*     m_cellFocusDelegate = nullptr;
    // Combo editor for column 4 when the sheet has per-row puffing regime.
    DVE::RegimeComboDelegate*   m_regimeDelegate    = nullptr;
    // Avatar bar sits at the top of the central editor area.
    DVE::PresenceAvatarBar*     m_avatarBar        = nullptr;
    // Banner shown above the central editor when a currently-open resource
    // is deleted by another user (T19). Lives between the avatar bar and
    // the central stack; hidden by default.
    DVE::RowDeletedBanner*      m_rowDeletedBanner = nullptr;
    // Currently-open resource — drives the avatar bar refresh logic and the
    // "this resource changed → refresh the bar" check in refreshPresenceFor.
    QString                     m_currentResourceType;
    qint64                      m_currentResourceId  = -1;

    // ── Offline mode (Plan C T7-T9) ──────────────────────────────────────────
    // OfflineSnapshot: SQLite mirror used as a read-only data source when PG
    //                  is unreachable. Lifetime owned here (passed as raw ptr
    //                  to DatabaseManager via setOfflineSnapshot()).
    // ConnectionMonitor: wraps m_pgConn and emits wentOffline/cameOnline.
    // OfflineBanner: top-of-window widget. Hidden by default; shown when the
    //                monitor flips us to offline.
    DVE::OfflineSnapshot*   m_snapshot      = nullptr;
    DVE::ConnectionMonitor* m_monitor       = nullptr;
    DVE::OfflineBanner*     m_offlineBanner = nullptr;

    // Pending TPM cell edits captured while offline. Replayed by
    // flushPendingEdits() when the monitor signals cameOnline().
    //
    // Scope: v1 captures TPM cell edits only. Sensory / detailed sensory
    // edits are not queued — saves attempted while offline surface a status
    // bar message and the user is expected to retry once reconnected.
    // Deferred to v1.1: per-cell yellow-dot badge on the data table for
    // rows with a queued edit. The OfflineBanner pending-count gives the
    // user-visible MVP signal.
    struct PendingEdit {
        int       fileIdx     = -1;
        int       sheetIdx    = -1;
        int       sampleIdx   = -1;
        QString   filePath;     // captured to survive m_loadedFiles reshuffle
        QString   sheetName;
        QDateTime capturedAt;
    };
    QVector<PendingEdit> m_pendingEdits;

    void onConnectionWentOffline();
    void onConnectionCameOnline();
    void onOfflineRetryClicked();
    void onRefreshSnapshotTriggered();
    void flushPendingEdits();

    // Plan B Phase 6 — don't-yank-in-progress-edits machinery.
    // Per-cell roles used on m_dataTable QTableWidgetItems:
    //   UserRole + 2 : baseline value (set at populate; compared on edit).
    //   UserRole + 3 : bool dirty (true → user is actively editing).
    //   UserRole + 4 : pending remote value (set when NOTIFY arrives for a
    //                  cell currently dirty; cleared when the user clicks
    //                  the yellow-decorated cell to accept).
    // Verticalheader item on m_dataTable stores the DataRow's database id
    // at Qt::UserRole so incoming NOTIFY on data_rows can find the row.
    void onDataTableItemChanged(QTableWidgetItem* it);
    void onDataTableItemClicked(QTableWidgetItem* it);
    // Apply / clear the yellow-border decoration on a cell. updateBaseline=true
    // means accept the remote value as the new baseline (clears dirty).
    void clearRemoteDecoration(QTableWidgetItem* it, bool updateBaseline);
    // Look up the on-screen QTableWidgetItem that maps to the given
    // data_rows.id. Returns nullptr if the id isn't present in the current
    // table (different sample open, fresh row, or post-save id churn).
    int findTableRowForDataRowId(qint64 dataRowId) const;
    // Best-effort: resolve a UUID to a display name via PresenceManager's
    // currently-active rows. Falls back to "another user".
    QString resolveUserName(const QString& uuid) const;
    // Phase 6 row-change handler. Encapsulates the data_rows decoration and
    // the row-deleted banner logic so the rowChanged lambda stays tiny.
    void handleRemoteRowChange(const DVE::RowChange& c);

    // v2.0.1 LiveSync inbound handlers — column-aware single-cell payloads
    // arrive here. Sensory tables filter out via the table-name guard.
    void onRemoteCellChanged(const QString& table, qint64 rowId,
                             const QString& column, const QVariant& newValue);
    void onRemoteCellFocused(const QString& table, qint64 rowId,
                             const QString& column,
                             const QString& userName,
                             const QString& userColor);
    void onRemoteCellBlurred(const QString& table, qint64 rowId,
                             const QString& column);

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
    // v2.0.2 H3: returns true on success, false if the openpyxl invocation
    // emitted a non-zero exit code or otherwise failed. flushExcelWrites
    // uses the return to decide whether to clear m_pendingWrites or retain
    // the entries for a later retry, and to gate the user-visible warning.
    bool writeCellsToExcel(const QString& filePath, const QString& sheetName,
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

    // ── Sheet-aware LiveSync column helpers ───────────────────────────────────
    // Column 4 is dual-purpose (resistance vs. puffing_regime). These
    // wrappers check the current sheet's hasPerRowRegime flag and return
    // the correct DB column name / visual column index.
    QString liveColumnForDataCol(int col) const;
    int     dataColForLiveColumn(const QString& dbColumn) const;
    QStringList currentFileRegimes() const;

    // Refresh the regime picker in the plot widget for the current file.
    void refreshPlotRegimes();

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

    // ── Debounced LiveSync focus broadcast ────────────────────────────────────
    // Arrow-keying through 50 cells in a few seconds would otherwise
    // fire 50 cell_focus DELETE+INSERT round-trips and 50 NOTIFY events.
    // 120 ms single-shot coalesces a scrub into a single focus broadcast.
    QTimer*              m_focusCommitTimer = nullptr;
    QString              m_pendingFocusTable;
    qint64               m_pendingFocusRowId  = -1;
    QString              m_pendingFocusColumn;
    bool                 m_pendingFocusBlur   = false;
    QString              m_pendingWriteFile;
    QString              m_pendingWriteSheet;
    QVector<CellWrite>   m_pendingWrites;

    // v2.0.2 H3: rate-limit the user-facing warning so a persistent
    // Python failure (e.g., openpyxl missing, file locked by Excel) does
    // not spam a dialog on every 500 ms flush tick. Reset on each
    // successful flush.
    bool m_excelWriteFailureShown = false;

    // Prompt user to save DB if there are unsaved changes; returns false if user cancels
    bool promptSaveDatabase();

    // Column headers for data table
    static QStringList dataTableHeaders(bool perRowRegime = false);

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

    // Sidebar compact mode (Task 7)
    QWidget*        m_sidebarFullPanel  = nullptr;  // the full left-dock splitter
    QWidget*        m_sidebarIconStrip  = nullptr;  // 32 px wide icon strip
    QStackedWidget* m_sidebarStack      = nullptr;  // index 0=full, index 1=strip
    void showSidebarOverlay();  // switches back to full panel; called by icon strip buttons

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
    QLabel*       m_showingLabel     = nullptr;   // "(Showing 1+2)" when multiple sessions selected
    void refreshSensoryAverages();
    void onTestAvgSelectionChanged();
    // Compute per-device averages across a set of sessions and populate the
    // averaged-table overlay + radar chart. Used by both Ctrl+click multi-select
    // and the Test Averages list. Pass the actual sessions so the caller can
    // either filter by index or by test title without each path duplicating
    // the math.
    void showSensoryAveragesFor(const QVector<SensorySession>& sessions,
                                const QVector<int>& sourceIndices);
};

} // namespace DVE
