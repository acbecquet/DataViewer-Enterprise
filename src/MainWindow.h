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

class QProgressDialog;   // pointer-only members (v2.4.16); full def in the .cpp

#include "ExcelReader.h"
#include "utils/ExcelWritePayload.h"
#include "pipeline/ReportData.h"
#include "ui/NewFileDialog.h"
#include "ui/HeaderEditDialog.h"
#include "ui/ImageViewDialog.h"
#include "pipeline/DataProcessor.h"
#include "reporting/ReportGenerator.h"
#include "plotting/PlotWidget.h"
#include "widgets/RibbonWidget.h"
#include "database/DatabaseManager.h"
#include "database/PersistWorker.h"        // SP4.5 Stage 2a (PersistJob)
#include "database/SnapshotRegenWorker.h"  // SP4.5 Stage 2a
#include <QThread>
#include <QHash>
#include <QTimer>
#include "ui/DatabaseBrowserDialog.h"
#include "ui/DataCleanupDialog.h"
#include "ui/ImageInboxDialog.h"
#include <QFileSystemWatcher>
#include <QStackedWidget>
#include <QListWidget>
#include "ui/DetailedSensoryPanel.h"
#include "pipeline/SensoryData.h"
#include "utils/UpdateChecker.h"
#include "utils/OutputPaths.h"
#include "utils/RecoveryManager.h"

namespace DVE {

// ─── Forward decls ────────────────────────────────────────────────────────────
class PlotWidget;
class ScrollHost;
class NotesStoryPanel;
class SensoryPanel;
class DetailedSensoryPanel;
class IdentityManager;
class PostgresConnection;
class NotificationListener;
class PresenceManager;
class LiveSync;
class PresenceDotsDelegate;
class PresenceAvatarBar;
class RowDeletedBanner;
class OfflineSnapshot;
class ConnectionMonitor;
class OfflineBanner;
class IncompleteDataBanner;
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

    // --- Test-support accessors (v2.7.0 responsive-UI verification) ---
    // Return the per-region ScrollHost / content widget that the --ui-stress
    // harness inspects to compute its closed-loop no-clip pass/fail. regionKey
    // is one of: "central", "navigator", "notes", "ribbonGroups".
    // Returns nullptr for an unknown key or a not-yet-constructed lazy panel.
    // These expose existing pointers only; they create nothing and have no
    // side effects.
    DVE::ScrollHost* scrollHostFor(const QString& regionKey) const;
    QWidget*         regionWidget(const QString& regionKey) const;

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

    // ── SP4.5 Stage 2a: background persist + debounced snapshot regen ──
    void onPersistFinished(DVE::PersistJob job, DVE::WriteResult wr);
    void onSnapshotRegenFinished(bool ok, const QString& error);
    void onSnapshotRegenRequired();

    // ── Plan C auto-recovery (Bug 1) ──
    // Startup prompt: if Recovery_prev/ holds a recoverable set (the prior
    // instance died uncleanly), offer to reload it. Triggered once after the
    // window is shown so the dialog has a parent and the panels exist. Safe to
    // call regardless of m_recoveryArmed — it reads recoverableItems() ONCE and
    // branches on whether that read was empty and on lastReadFailed() (empty +
    // failed => present-but-undecodable, warn; empty + not-failed => nothing to
    // recover, silent; non-empty => offer reopen).
    void maybeOfferRecovery();

    // Tools->Recover: manual, selective reload of the previous session's items.
    // Works any time this session (incl. after declining the startup prompt),
    // since Recovery_prev/ persists until the next clean close. Opens
    // RecoverDialog for per-item selection, then hands the chosen subset to
    // restoreItems().
    void onRecover();
    void onViewRawData();   // DV-17: Tools > View Raw Data — open the source Excel

    // ── Reports ──
    void onGenerateTestReport();
    void onGenerateFullReport();
    // ── Navigation ──
    void onFileSelected(int index);
    void onSheetSelected(int index);
    void onPrevSample();
    void onNextSample();

    // ── Editable notes-story panel (TPM) ──
    void onStoryCellEdited(int dataRow, int col, const QString& text);
    void onStoryNoteActivated(int dataRow);

    // ── Background worker ──
    void onFileLoadFinished();
    void onReportFinished(bool success, const QString& path);

    // ── Tools / Sensory mode ──
    void toggleSensoryMode(bool checked);
    void toggleDetailedSensoryMode(bool checked);
    void onOpenDatabaseBrowser();

    // ── Data cleanup ──
    void onCleanData();
    void onResetCleanup();
    void onUndoAllCleanup();   // DATAVIEWER-16: clear ALL exclusions, every file

    // ── Database ──
    // flushPending=true at DELIBERATE save points (Ctrl+U, program-close): drains
    // LiveSync to the DB first so the whole-session read-merge-write sees this
    // client's latest per-cell scores. The 5 s auto-save timer calls with false
    // (stays fully async; never blocks typing). (DATAVIEWER-4)
    void onUpdateDatabase(bool flushPending = false);

    // ── Export to Excel (manual flush of debounced write-back) ──
    void onExportToExcelTriggered();

    // ── Edit headers ──
    void onEditHeaders();
    void onPropCellChanged(int row, int col);

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
    // Master reset for the Navigator/Notes docks (Settings ▸ Panels ▸ Reset
    // Panels): un-float + re-dock to defaults, un-collapse, re-balance, persist.
    void resetPanelLayout();

    // ── Status bar helpers (v2.0.9) ──────────────────────────────────────────
    void buildStatusBar();
    void setStatusFile(const QString& text, FileStatus status);
    void setStatusDb(const QString& text, DbStatus status);
    void setStatusBreadcrumb(const QStringList& segments);

    // ── Ribbon tabs ──────────────────────────────────────────────────────────
    void buildHomeTab(RibbonTab* tab);
    void buildReportsTab(RibbonTab* tab);
    void buildToolsTab(RibbonTab* tab);
    void buildSettingsTab(RibbonTab* tab);

    // DATAVIEWER-13 (MS-5): modal key-capture for rebinding the Sensory stopwatch
    // hotkey; returns the Qt key code, or 0 if cancelled.
    int captureStopwatchKey();

    // ── Left dock: File/Sheet browser ────────────────────────────────────────
    QDockWidget*  m_fileDock;
    // True while the VeryNarrow breakpoint has force-hidden the side docks, so
    // the exit transition only restores docks that VeryNarrow itself collapsed
    // (and never resurrects the TPM-only Notes dock while in a sensory mode).
    bool          m_docksAutoCollapsed = false;
    // v2.7.0: per-region scroll wrappers exposed to the verification sweep via
    // scrollHostFor() (Task 16). Store the docks' content hosts so the
    // --ui-stress harness can assert the fits-or-scrolls guarantee per region.
    DVE::ScrollHost* m_navScrollHost = nullptr;     // wraps m_sidebarStack
    DVE::ScrollHost* m_notesScrollHost = nullptr;   // wraps the Notes pane
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
    QSplitter*      m_centralSplitter;          // TPM: plot (one-child splitter)
    QDockWidget*    m_notesDock = nullptr;      // TPM-only floatable Notes dock (mirrors Navigator)

    // ── Sample-nav bar (sits atop the Notes panel) ───────────────────────────
    QWidget*      m_sampleNavBar;
    QPushButton*  m_prevBtn;
    QPushButton*  m_nextBtn;
    QLabel*       m_sampleCountLabel;

    // ── Plot panel ───────────────────────────────────────────────────────────
    PlotWidget*   m_plotWidget;
    NotesStoryPanel* m_storyPanel = nullptr;

    // ── Ribbon ───────────────────────────────────────────────────────────────
    RibbonWidget*  m_ribbon;
    QToolButton*   m_resetCleanupBtn = nullptr;  // enabled only when cleanup is active
    QToolButton*   m_undoAllCleanupBtn = nullptr;  // clears ALL exclusions, every file
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
    QToolButton*   m_viewRawDataBtn = nullptr;  // DV-17: Tools > View Raw Data (disabled in Detailed)
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
    // v2.4.16: one modal progress dialog spanning the close -> save -> snapshot
    // sequence, so it shows ONE continuous bar from the start (no frozen pre-bar
    // gap). Reports keep their existing live status-bar bar (driven by setProgress).
    QProgressDialog* m_closeProgress  = nullptr;
    // Legacy aliases kept so call sites compile while being migrated:
    QLabel*       m_dbSyncLabel      = nullptr;  // replaced by m_statusDbText

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

    // ── SP4.5 Stage 2a: background persist + debounced snapshot regen ──────────
    // m_dbConfig is captured at DB-open so the worker threads can open their OWN
    // QSqlDatabase / PostgresConnection (Qt forbids sharing a connection across
    // threads). Both workers have no QObject parent (a parented QObject cannot be
    // moved to another thread); the QThreads are parented to MainWindow and are
    // quit()/wait()'d in the destructor before Qt auto-deletes the children.
    DVE::DbConfig               m_dbConfig;
    QThread*                    m_persistThread      = nullptr;
    DVE::PersistWorker*         m_persistWorker      = nullptr;
    quint64                     m_persistGeneration  = 0;   // bumped per enqueue
    QHash<QString, quint64>     m_lastEnqueuedGen;          // filePath -> latest gen
    QSet<QString>               m_backgroundSaveInFlight;   // files with a pending bg persist
    QSet<QString>               m_dirtiedDuringPersist;     // edited while a bg persist ran
    QThread*                    m_regenThread        = nullptr;
    DVE::SnapshotRegenWorker*   m_regenWorker        = nullptr;
    QTimer*                     m_snapshotRegenTimer = nullptr;
    bool                        m_snapshotRegenInFlight = false;
    bool                        m_persistDraining       = false;  // close drain re-entrancy
    int                         m_regenFailStreak       = 0;      // bound the retry loop

    // v2.5.0 Task 4 (RC3): mirror of LiveSync::unsyncedEditCount(), updated via
    // the unsyncedEditsChanged signal. Non-zero drives the DB sync indicator
    // into a warning state so silent per-cell commit loss can't hide. These
    // edits are reconciled on the next whole-session save (dirty-cell merge).
    int                         m_unsyncedEdits = 0;

    // v2.4.4 R5: one-shot guard so the "offline edit could not be queued"
    // warning (LiveSync::offlineEnqueueFailed -- a MIP-undecodable local queue)
    // is shown at most once per session. The indicator still updates every time;
    // only the interruptive popup is rate-limited to one.
    bool                        m_offlineEnqueueWarningShown = false;

    // v2.4.4 R5: one-shot guard for the "offline cache is MIP-undecodable"
    // warning shown when openReadOnly() fails specifically because snapshot.sqlite
    // is encrypted at rest (vs. mere first-run absence, which stays silent).
    bool                        m_offlineSnapshotDecodeWarningShown = false;

    // ── Presence UI (Plan B Phase 5) ─────────────────────────────────────────
    // Delegate is shared by m_fileTree, m_sensoryNav, m_detailedSensoryNav.
    // Cheap (one QObject) and keeps all three widgets consistent.
    DVE::PresenceDotsDelegate*  m_presenceDelegate = nullptr;
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
    DVE::OfflineSnapshot*        m_snapshot              = nullptr;
    DVE::ConnectionMonitor*      m_monitor               = nullptr;
    DVE::OfflineBanner*          m_offlineBanner         = nullptr;
    // Shown when the active TPM FileResult has any sheet with dbDataIncomplete.
    // Dismissed by the user via the X button; re-shown on next incompete load.
    DVE::IncompleteDataBanner*   m_incompleteDataBanner  = nullptr;

    void onConnectionWentOffline();
    void onConnectionCameOnline();
    void onOfflineRetryClicked();
    void onRefreshSnapshotTriggered();
    // SP4.5 audit fix: run the synchronous snapshot regen (close + manual refresh)
    // on a DEDICATED PG connection behind a determinate progress dialog, so a
    // LiveSync/presence timer firing during the progress pump can't issue a write
    // inside the regen's read-only txn on the shared m_pgConn (SQLSTATE 25P02).
    bool regenerateSnapshotWithProgress(const QString& title);
    // v2.4.16 progress-feedback helpers (blocking-op indicators; see makeBusyDialog
    // in the .cpp). Owner directive: show a bar from the START of anything that
    // blocks the UI -- never a frozen window with no feedback.
    void closeProgressEnd();                        // tear down m_closeProgress

    // v2.4.2 R3 (reset-to-5 keystone): after a reconnect, pull the
    // authoritative DB state for the currently-open resource and merge it into
    // memory so a stale in-memory save can never clobber freshly-normalized DB
    // values, while never losing the user's own unsaved (dirty) edits. Called
    // at the END of onConnectionCameOnline(), AFTER the per-cell LiveSync drain
    // — the outbound drain must land local edits first, THEN this pulls the
    // remote changes. Best-effort + logged; guarded on a live online m_db.
    void reloadOpenResourceAfterReconnect();

    // Show/hide m_incompleteDataBanner based on whether the currently-active
    // TPM FileResult has any sheet with dbDataIncomplete == true.
    // Call whenever the active file or its data changes.
    void updateIncompleteDataBanner();

    // Best-effort: resolve a UUID to a display name via PresenceManager's
    // currently-active rows. Falls back to "another user".
    QString resolveUserName(const QString& uuid) const;
    // Row-change handler for inbound NOTIFY. Drives the row-deleted banner (T19)
    // for the currently-open file / sensory / detailed-sensory resource.
    void handleRemoteRowChange(const DVE::RowChange& c);

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
    // SP3-T4 (R6): the cell-edit struct now lives in ExcelWritePayload so the
    // off-thread worker, the synchronous path, and the equivalence test share
    // one definition. CellWrite stays the in-class name everything already uses.
    using CellWrite = DVE::ExcelCellWrite;
    // v2.0.2 H3: returns true on success, false if the openpyxl invocation
    // emitted a non-zero exit code or otherwise failed. flushExcelWrites
    // uses the return to decide whether to clear m_pendingWrites or retain
    // the entries for a later retry, and to gate the user-visible warning.
    // SP3-T4: this is now the SYNCHRONOUS path, used only by bulk row ops
    // (add/remove row) and by finishExcelWritesBlocking on the close paths.
    // The debounced single-cell flush goes through the off-thread worker below.
    bool writeCellsToExcel(const QString& filePath, const QString& sheetName,
                           const QVector<CellWrite>& cells,
                           int timeoutMs = kExcelBatchTimeoutMs);
    void queueExcelWrite(const QString& filePath, const QString& sheetName,
                         int excelRow1, int excelCol1, const QString& value);
    void flushExcelWrites();

    // ── SP3-T4 (R6): off-thread Excel write-back ──────────────────────────────
    // The openpyxl save subprocess used to run on the UI thread via
    // QProcess::waitForFinished(30000), producing an up-to-30 s "Not Responding"
    // freeze on a large workbook. flushExcelWrites() now dispatches a
    // QtConcurrent worker that takes COPIES only (no QWidget, no shared member),
    // and m_excelFlushWatcher delivers the result back on the UI thread. The
    // re-entrancy guard mirrors RecoveryManager: only one python writer to a
    // file at a time; a flush requested mid-flight sets m_excelFlushPending and
    // re-fires on completion.

    // Tiered timeout budgets passed to the worker / synchronous path.
    static constexpr int kExcelInteractiveTimeoutMs = 8000;   // debounced single-cell flush
    static constexpr int kExcelBatchTimeoutMs       = 30000;  // bulk ops (add/remove row, multi-cell)

    // Small value result the worker hands back. No QWidget / no member refs.
    struct ExcelWriteResult { bool ok = false; QString error; };

    // The pure off-thread worker: runs the openpyxl subprocess for a prebuilt
    // (script, args) payload with the given timeout. Static + self-contained so
    // it can run on a QtConcurrent thread without touching `this`.
    static ExcelWriteResult runExcelWriteWorker(const QString& python,
                                                const QString& script,
                                                const QStringList& args,
                                                int timeoutMs);

    // Slot-equivalent invoked on the UI thread when the worker finishes.
    void onExcelFlushFinished();

    // Close-path drain: if a flush is in flight, wait it out, then run any
    // still-pending writes synchronously so the app never exits with a write
    // abandoned on a background thread (and never crashes on a running future).
    void finishExcelWritesBlocking();
    void updateStatusBar(const QString& msg);
    void setProgress(int pct, const QString& msg);
    void showError(const QString& title, const QString& msg);
    void showInfo(const QString& title, const QString& msg);

    void updateImageButton();
    void markFileModified();
    void updateDbSyncIndicator();

    // DATAVIEWER-3: persist a freshly loaded/refreshed TPM file and reflect the
    // WriteResult (keep dirty + surface on failure, clear on success).
    void persistLoadedFile(int fileIndex);
    // SP4.5 Stage 2a: background persist-on-load + close-time drain guard.
    void enqueuePersist(int fileIndex);
    void drainPersistWorkerBlocking(int timeoutMs = 15000);

    // DATAVIEWER-4: authoritatively persist the given sensory sessions (panel
    // indices) before they're closed, so Close never drops edits. Returns the
    // indices that FAILED to save (caller keeps those open). Mirrors the
    // per-session save in onUpdateDatabase (rename->INSERT, UniqueViolation skip),
    // scoped to the closing set and quiet on success.
    QVector<int> saveSensorySessionsBeforeClose(const QVector<int>& indices);
    QVector<int> saveDetailedSensorySessionsBeforeClose(const QVector<int>& indices);

    // ── v2.5.0 RC5: never hard-block an unnamed session on close ───────────────
    // What the user can do with an unnamed (un-keyable) session at close time.
    enum class SessionCloseChoice { NameIt, Discard, Cancel };

    // Just enough about an offending session to fill the close dialog without a
    // mode-specific overload of the dialog itself. `hasDiskFile` lets the dialog
    // mention an autosave .xlsx that will be left on disk (we never delete it).
    struct SessionCloseInfo {
        QString label;          // human label (title - tester, or "(unnamed)")
        QString missing;        // "a test name and a tester" etc.
        int     sampleCount = 0;
        bool    hasDiskFile = false;
    };
    SessionCloseInfo sessionCloseInfoFor(const SensorySession& s) const;
    SessionCloseInfo sessionCloseInfoFor(const DetailedSensorySession& s) const;

    // Show the three-option close dialog for one unnamed session and return the
    // user's choice. Shared by both sensory and the program-close path.
    SessionCloseChoice promptUnnamedSessionOnClose(const SessionCloseInfo& info);

    // v2.5.0 RC5: at program close (promptSaveDatabase "Save All"), route every
    // unnamed non-placeholder session through the per-session Name/Discard/Cancel
    // dialog. Returns false (abort the close) if the user chose Name It Now or
    // Cancel for any session; true (proceed) once all unnamed sessions are
    // discarded or there are none. Discard removes the DB row + drops the panel
    // session in place (no follow-up closeSessions on this path).
    bool resolveUnnamedSessionsForProgramClose();

    // ── Sheet-aware LiveSync column helper ────────────────────────────────────
    // Maps a DVE::Cols index to its data_rows DB column for the per-cell
    // LiveSync commit. Column 4 is dual-purpose (resistance vs. puffing_regime),
    // resolved via the current sheet's hasPerRowRegime flag.
    QString liveColumnForDataCol(int col) const;
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
    // fileIdx must be the file's ACTUAL index in m_loadedFiles (GAP-A), not the
    // currently-selected file, so multi-file/Combined reports apply each file's
    // own exclusions.
    FileResult   buildCleanedFile(const FileResult& file, int fileIdx) const;

    // DATAVIEWER-16: persist exclusions across restarts, keyed by file PATH so
    // they survive file reordering between sessions. (onUndoAllCleanup is a slot,
    // declared in the slots section above.)
    void saveExclusions() const;
    void restoreExclusionsForFile(int fileIdx);

    QString resourcePath() const;
    QString defaultInboxPath() const;
    QString templatePath() const;
    QString findPython() const;
    mutable QString m_cachedPython;
    mutable bool    m_pythonProbed = false;

    // Run a one-shot Python script (writes to temp file, executes, returns stdout).
    // Returns empty string on error and sets lastError via errOut.
    // SP3-T4 (R6): timeoutMs is the QProcess::waitForFinished budget. It defaults
    // to 30 s so every pre-existing caller is unchanged; the off-thread Excel
    // worker passes a tiered budget (interactive vs batch). runPython remains a
    // pure static with no member/QWidget access, so the worker calls it directly
    // off-thread — the exact same invocation the synchronous path uses.
    static QString runPython(const QString& python,
                             const QString& script,
                             const QStringList& args,
                             QString& errOut,
                             int timeoutMs = 30000);

    // Remembers last-used directory for file dialogs
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
    mutable QString m_lastBrowseDir;

    // ── Debounced Excel write queue ──────────────────────────────────────────
    QTimer*              m_excelWriteTimer = nullptr;
    QTimer*              m_dbSaveTimer     = nullptr;  // auto-save after inactivity

    // ── Notes-panel edit → coalesced plot re-render ──────────────────────────
    QTimer* m_storyPlotTimer    = nullptr;  // debounces the heavy setSheetData per edit
    bool    m_inStoryCellEdit   = false;    // re-entrancy guard for onStoryCellEdited
    bool    m_storyRegimeDirty  = false;    // a regime cell changed → refresh the plot's regime picker

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

    // ── SP3-T4 (R6): off-thread flush state ───────────────────────────────────
    // The watcher delivers the worker's result on the UI thread. The in-flight
    // copies are what the worker is persisting RIGHT NOW; on failure they are
    // re-merged into m_pendingWrites via mergePendingWithInFlight — the older
    // in-flight cells SEED the result and any newer pending edit to the same
    // cell OVERLAYS it (newest wins). Only one worker runs at a time (the
    // guard), and the synchronous bulk ops (onAddRow / onRemoveRow) drain that
    // worker via finishExcelWritesBlocking before they write, so exactly one
    // openpyxl process ever touches a given workbook.
    QFutureWatcher<ExcelWriteResult>* m_excelFlushWatcher = nullptr;
    bool                 m_excelFlushInFlight = false;  // a worker flush is running
    bool                 m_excelFlushPending  = false;  // another flush requested mid-flight
    QString              m_inFlightWriteFile;
    QString              m_inFlightWriteSheet;
    QVector<CellWrite>   m_inFlightWrites;

    // Prompt user to save DB if there are unsaved changes; returns false if user cancels
    bool promptSaveDatabase();

    // Plan C C10: human-readable list of every in-memory item that holds
    // unsaved work — each modified TPM file (by name), each dirty Sensory
    // session, and each dirty Detailed-sensory session (placeholders
    // excluded). Drives the consolidated close prompt; empty ⇒ nothing to
    // save. const because it only reads state.
    QVector<QString> unsavedInventory() const;

    // Returns desiredPath if it doesn't exist; otherwise appends (2), (3), …
    static QString uniqueFilename(const QString& desiredPath);

    // Active report mode (TPM / Sensory / Detailed), used to resolve the
    // per-mode default folder for save / load / new-file dialogs.
    ReportMode currentReportMode() const;

    // ── Sensory mode ────────────────────────────────────────────────────────
    bool            m_sensoryMode = false;
    bool            m_sensorySessionsDirty = false;
    SensoryPanel*   m_sensoryPanel = nullptr;
    bool                    m_detailedSensoryMode = false;
    bool                    m_detailedSensorySessionsDirty = false;
    DetailedSensoryPanel*   m_detailedSensoryPanel = nullptr;
    QListWidget*            m_detailedSensoryNav = nullptr;

    // ── Plan C auto-recovery (Bug 1) ──────────────────────────────────────────
    // Rolling crash/auto-update snapshot of all three in-memory stores
    // (m_loadedFiles + both sensory panels' sessions). Constructed in the ctor;
    // adoptPreviousSession() runs there too, before any flush can fire.
    RecoveryManager* m_recovery = nullptr;
    // True iff adoptPreviousSession() cleanly promoted the prior live store. Only
    // then do we wire the state provider + noteDirty() hooks — a false return
    // means crash data is stranded in the live dir and MUST NOT be overwritten.
    bool             m_recoveryArmed = false;
    // State provider for m_recovery: captures the full current set of every store
    // as self-contained value copies (entries WITH payloads). const + non-blocking
    // (thin glue over the already-tested serializers); runs on the UI thread.
    QVector<RecoveryEntry> captureRecoveryState() const;
    // Reload a set of recovered items (from Recovery_prev/, or — in C9 — a
    // user-selected subset) back into the live stores. Deserializes each entry,
    // recomputes the TPM plot series (intentionally not serialized), constructs
    // the sensory panels if a recovered session needs one before the user
    // toggled the mode, marks everything dirty so it re-saves through the normal
    // optimistic-concurrency path, refreshes the UI, switches to the first
    // restored item's mode, then re-arms the snapshot via noteDirty().
    void restoreItems(const QVector<RecoveryEntry>& items);

    // Navigator stack inside left dock (index 0 = file tree, index 1 = sensory sessions)
    QStackedWidget* m_navStack     = nullptr;
    QListWidget*    m_sensoryNav   = nullptr;
    QLabel*         m_navLabel     = nullptr;   // "Loaded Files:" / "Sessions:"

    // Sidebar compact mode (Task 7)
    QWidget*        m_sidebarFullPanel  = nullptr;  // the full left-dock splitter
    QWidget*        m_sidebarIconStrip  = nullptr;  // 32 px wide icon strip
    QStackedWidget* m_sidebarStack      = nullptr;  // index 0=full, index 1=strip
    void showSidebarOverlay();  // switches back to full panel; called by icon strip buttons

    // VeryNarrow responsive rule: collapse both side docks when veryNarrow is
    // true (reclaiming width for the central ScrollHost), and restore them to
    // their current mode-appropriate visibility when it is false. Idempotent.
    void applyVeryNarrowDockState(bool veryNarrow);

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
