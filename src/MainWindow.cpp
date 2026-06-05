#include "MainWindow.h"
#include "utils/AppTheme.h"
#include "utils/OutputPaths.h"
#include "utils/ResponsiveLayout.h"
#include "pipeline/SheetProcessors.h"
#include "ui/SensoryPanel.h"
#include "ui/DetailedSensoryPanel.h"
#include "ui/SopDialog.h"
#include "ui/RecoverDialog.h"
#include "database/IdentityManager.h"
#include "database/IdentityPromptDialog.h"
#include "database/ConfigLoader.h"
#include "database/PostgresConnection.h"
#include "database/NotificationListener.h"
#include "database/PresenceManager.h"
#include "database/LiveSync.h"
#include "database/VersionLookup.h"
#include "RemoteCellHelpers.h"
#include "database/OfflineSnapshot.h"
#include "database/ConnectionMonitor.h"
#include "widgets/CellFocusDelegate.h"
#include "widgets/RegimeComboDelegate.h"
#include "widgets/PresenceDotsDelegate.h"
#include "widgets/PresenceAvatarBar.h"
#include "widgets/RowDeletedBanner.h"
#include "widgets/OfflineBanner.h"
#include "widgets/IncompleteDataBanner.h"
#include "pipeline/RegimeUtils.h"

#include <QApplication>
#include <QFileDialog>
#include <QMessageBox>
#include <QInputDialog>
#include <QSettings>
#include <QTimer>
#include <QCloseEvent>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QMimeData>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QDialog>
#include <QLabel>
#include <QListWidget>
#include <QListWidgetItem>
#include <QPushButton>
#include <QHeaderView>
#include <QDesktopServices>
#include <QUrl>
#include <QDir>
#include <QStandardPaths>
#include <QTreeWidgetItem>
#include <QtConcurrent/QtConcurrent>
#include <QDebug>
#include <QRegularExpression>
#include <QFrame>
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
#include <QTextStream>
#include <QSqlDatabase>
#include <QSqlQuery>
#include <QProgressDialog>
#include <QAction>
#include <QMenu>
#include <algorithm>
#include "xlsxdocument.h"
#include "pipeline/SensoryData.h"
#include "pipeline/DetailedSensoryData.h"
#include "pipeline/ReportDataJson.h"   // fileResultToJson for the recovery snapshot

namespace {

// v2.0.1: maps a TPM data-table visual column to its data_rows DB column.
// Order MUST match MainWindow::dataTableHeaders() — columns 0..7 are the
// editable inputs; 8..11 are calculated and have no DB counterpart.
const QStringList& kDataTableColumns()
{
    static const QStringList cols = {
        QStringLiteral("puffs"),
        QStringLiteral("before_weight"),
        QStringLiteral("after_weight"),
        QStringLiteral("draw_pressure"),
        QStringLiteral("resistance"),
        QStringLiteral("smell"),
        QStringLiteral("clog"),
        QStringLiteral("notes"),
    };
    return cols;
}

QString columnNameForDataTableColumn(int col)
{
    const QStringList& cols = kDataTableColumns();
    return (col >= 0 && col < cols.size()) ? cols.at(col) : QString();
}

int dataTableColumnForColumnName(const QString& name)
{
    return kDataTableColumns().indexOf(name);
}

} // namespace

namespace DVE {

// ─── Construction / Destruction ───────────────────────────────────────────────
MainWindow::MainWindow(QWidget* parent)
    : QMainWindow(parent)
    , m_processor(new DataProcessor())
    , m_reportGen(new ReportGenerator(this))
    , m_db(new DatabaseManager(this))
    , m_loadWatcher(new QFutureWatcher<FileResult>(this))
{
    setWindowTitle(QStringLiteral("DataViewer Enterprise  v%1")
                       .arg(QApplication::applicationVersion()));
    setMinimumSize(1280, 800);
    resize(1600, 900);
    setAcceptDrops(true);

    // Identity must exist before opening the DB — DatabaseManager::open()
    // takes the identity pointer so subsequent operations can stamp rows
    // with the current user's UUID/display name.
    m_identity = new DVE::IdentityManager(this);

    {
        DVE::DbConfig cfg;
        QString err;
        const QString confPath = QDir::cleanPath(
            QString::fromLocal8Bit(qgetenv("LOCALAPPDATA")) + "/DataViewer/db.conf");

        // Plan C T8: open the offline snapshot best-effort up front. The
        // snapshot is the read-only data source DatabaseManager routes to
        // when m_online == false. If the PG connection succeeds below, the
        // snapshot just sits there until the user goes offline.
        m_snapshot = new DVE::OfflineSnapshot(this);
        m_snapshot->openReadOnly();  // best-effort; absence is fine on first run

        if (!DVE::ConfigLoader::load(confPath, cfg, &err)) {
            QMessageBox::critical(this, tr("Database config error"),
                                  tr("Could not read database configuration from\n%1\n\n%2")
                                      .arg(confPath, err));
        } else if (!m_db->open(cfg, m_identity)) {
            // Plan C T8: PG unreachable at startup. If we have a snapshot,
            // boot in offline read-only mode. Otherwise, this is fatal —
            // the app has nothing to show.
            //
            // Deferred to v1.1: a full offline-boot UX that subsequently
            // attempts reconnects via ConnectionMonitor (currently this path
            // shows the banner with no monitor — user has no way to retry
            // beyond closing + reopening the app). Acceptable trade-off
            // given the existing constructor's structure; the retry path
            // is exercised by the *during-session* offline detection.
            if (m_snapshot->isOpen()) {
                m_db->setOfflineSnapshot(m_snapshot);
                m_db->setOnline(false);
                QMessageBox::warning(this, tr("Database unreachable"),
                                     tr("Could not connect to the Postgres database. "
                                        "Booting in read-only offline mode.\n\n%1")
                                         .arg(m_db->lastError()));
            } else {
                QMessageBox::critical(this, tr("Database unreachable"),
                                      tr("Cannot reach the database and no offline copy "
                                         "is available.\n\nConnect to the network and "
                                         "try again.\n\n%1")
                                          .arg(m_db->lastError()));
            }
        } else {
            // DB open succeeded. Bring up the rest of the concurrency stack.
            m_db->setOfflineSnapshot(m_snapshot);  // shared lifetime; DM holds raw ptr

            m_pgConn = new DVE::PostgresConnection(this);
            if (!m_pgConn->open(cfg)) {
                QMessageBox::warning(this, tr("Live updates unavailable"),
                                     tr("Connected to the database but could not open the "
                                        "live-updates connection. Other users' changes will "
                                        "not appear automatically until you reopen the app.\n\n%1")
                                         .arg(m_pgConn->lastError()));
                // m_pgConn stays as an owned but unopened object; m_notify/m_presence will
                // not subscribe/activate. The app keeps working in single-user mode.
            } else {
                m_notify = new DVE::NotificationListener(m_pgConn, this);
                if (!m_notify->subscribe()) {
                    qWarning() << "NotificationListener::subscribe() failed:"
                               << "live updates disabled for this session";
                }
                m_presence = new DVE::PresenceManager(m_pgConn, m_identity, this);

                // Plan C T8: ConnectionMonitor wraps the live-updates conn.
                // Wiring up wentOffline/cameOnline must happen here (we need
                // both m_pgConn and cfg in scope). Banner connection happens
                // after setupUI() runs since m_offlineBanner is set up there.
                m_monitor = new DVE::ConnectionMonitor(m_pgConn, cfg, this);
                connect(m_monitor, &DVE::ConnectionMonitor::wentOffline, this,
                        &MainWindow::onConnectionWentOffline);
                connect(m_monitor, &DVE::ConnectionMonitor::cameOnline, this,
                        &MainWindow::onConnectionCameOnline);
                // Don't start() yet — let setupUI() construct the banner first
                // so the wentOffline handler can show it without a null deref.
            }

            // v2.0.1: SaveCoordinator + ConflictResolver retired — LiveSync
            // now owns per-cell DB persistence and conflict handling. The
            // UniqueViolationDialog is still surfaced from save flows that
            // INSERT new rows (file paths, session names).

            // Own-UUID echo filter (Task 22): early-return when a NOTIFY says
            // we caused the change ourselves, so we don't trigger a UI refresh on
            // our own writes. The real UI refresh slot is filled in by Phase 5/6;
            // for 7a we only set up the filter + a qDebug stub.
            if (m_notify && m_pgConn && m_pgConn->isOpen()) {
                m_liveSync = new DVE::LiveSync(m_pgConn, m_identity, this);
                // v2.0.1 polish-2: spin up the background writer thread so
                // every cell commit is async and never blocks the UI.
                m_liveSync->setWorkerConfig(cfg);
                // v2.0.1 Task 9: hand the snapshot to LiveSync so
                // commitCell() can queue per-cell edits when the
                // connection drops mid-session.
                if (m_snapshot) m_liveSync->setOfflineSnapshot(m_snapshot);
                // v2.0.11: optimistic-concurrency disabled. The v2.0.2
                // VersionLookup callback supplied an expected version on every
                // per-cell commit, but the server-side stored proc bumps the
                // row's version on every successful write and there is no
                // back-propagation path that updates the in-memory cache.
                // Result: the first LiveSync write succeeds, the second uses
                // the now-stale local version, and the server rejects it with
                // an OCC miss — and the commitConflict signal goes nowhere
                // (no MainWindow slot listens), so the edit is silently lost.
                // Without the lookup, currentVersionFor returns -1 and the
                // worker binds NULL for expectedVersion, which puts the stored
                // proc on its pre-v2.0.2 no-OCC (last-writer-wins) path. The
                // project's "we don't have to worry about merging" stance
                // makes LWW the correct semantics anyway.
            }

            if (m_notify) {
                // v2.0.2 C7: single rowChanged subscriber. The previous
                // arrangement had two disjoint connect() lambdas — one to
                // paint the yellow remote-conflict decoration via
                // handleRemoteRowChange(), the other to dispatch through
                // LiveSync — and Qt's signal/slot dispatch order between
                // them was undefined, so the LiveSync overwrite could race
                // ahead of the dirty-cell decoration. Consolidating into
                // one slot fixes the ordering: decorate first, then apply.
                // M8: read selfUuid inside the lambda body each call so we
                // don't capture a value that becomes stale if the identity
                // changes mid-session.
                connect(m_notify, &DVE::NotificationListener::rowChanged, this,
                        [this](const DVE::RowChange& c) {
                            const QString selfUuid = m_identity
                                ? m_identity->uuid().toString(QUuid::WithoutBraces)
                                : QString();
                            if (c.updatedBy == selfUuid) return;
                            qDebug().noquote()
                                << "[MainWindow] rowChanged from another user:"
                                << c.table << c.op << "id=" << c.id
                                << "by=" << c.updatedBy;
                            // 1. Paint yellow remote-conflict decoration on
                            //    dirty data_rows cells (T18) and dispatch
                            //    row-deleted toast for the open resource (T19).
                            handleRemoteRowChange(c);
                            // 2. Forward to LiveSync so its
                            //    onRemoteCellChanged slot applies the value
                            //    to clean cells — the C6 dirty-guard inside
                            //    applyRemoteValueToCell ensures it won't
                            //    clobber what step 1 just decorated.
                            if (m_liveSync) m_liveSync->onRowChanged(c);
                        });
                connect(m_notify, &DVE::NotificationListener::presenceChanged, this,
                        [this](const DVE::PresenceChange& p) {
                            const QString selfUuid = m_identity
                                ? m_identity->uuid().toString(QUuid::WithoutBraces)
                                : QString();
                            if (p.userUuid.toString(QUuid::WithoutBraces) == selfUuid) return;
                            refreshPresenceFor(p.resourceType, p.resourceId);
                        });
            }

            if (m_liveSync && m_notify) {
                connect(m_notify, &DVE::NotificationListener::cellFocusChanged, m_liveSync,
                        [this](const DVE::CellFocusChange& f) {
                            const QString selfUuid = m_identity
                                ? m_identity->uuid().toString(QUuid::WithoutBraces)
                                : QString();
                            if (f.userUuid.toString(QUuid::WithoutBraces) == selfUuid) return;
                            m_liveSync->onCellFocusChanged(f);
                        });
                // v2.0.1: outbound LiveSync signals → MainWindow handlers that
                // repaint cells with remote values / focus borders.
                connect(m_liveSync, &DVE::LiveSync::cellChanged, this,
                        &MainWindow::onRemoteCellChanged);
                connect(m_liveSync, &DVE::LiveSync::cellFocused, this,
                        &MainWindow::onRemoteCellFocused);
                connect(m_liveSync, &DVE::LiveSync::cellBlurred, this,
                        &MainWindow::onRemoteCellBlurred);
            }
        }
    }

    setupUI();

    // Wire ResponsiveLayout: ribbon goes icons-only + breadcrumb truncates
    // when the window is snapped narrower than 1100 px.  Also collapses the
    // sidebar to a 32 px icon strip (Task 7).
    DVE::ResponsiveLayout::instance().beginTracking(this);
    connect(&DVE::ResponsiveLayout::instance(),
            &DVE::ResponsiveLayout::breakpointChanged,
            this, [this](DVE::ResponsiveLayout::Breakpoint bp, int) {
        if (m_ribbon) m_ribbon->setCompactMode(bp == DVE::ResponsiveLayout::Compact);
        setStatusBreadcrumb(m_lastBreadcrumbSegments);
        if (m_sidebarStack) {
            const bool compact = (bp == DVE::ResponsiveLayout::Compact);
            m_sidebarStack->setCurrentIndex(compact ? 1 : 0);
            // Dock min/max width: 32 px strip in compact, normal range otherwise.
            m_fileDock->setMinimumWidth(compact ? 32  : 220);
            m_fileDock->setMaximumWidth(compact ? 32  : 320);
        }
    });

    setupConnections();
    restoreSettings();

    // ── Plan C auto-recovery (Bug 1) ──────────────────────────────────────────
    // Own the rolling snapshot store. adoptPreviousSession() MUST run here,
    // before any flush could fire: the debounce/safety timers only tick once
    // this ctor returns and the event loop spins, and the state provider isn't
    // wired until just below, so nothing can write the live dir before adopt
    // promotes a crashed session to Recovery_prev/.
    //
    // C4 contract: adopt returns false when it could NOT cleanly promote the
    // prior live store (e.g. a locked file) — the crash data is then still in
    // liveDir() and wiring the rolling flush would overwrite it. So we arm the
    // capture hooks (the provider + every noteDirty() site) ONLY when it
    // returns true. The reopen prompt (C8) reads Recovery_prev/ either way.
    m_recovery = new RecoveryManager(this);
    m_recoveryArmed = m_recovery->adoptPreviousSession();
    if (m_recoveryArmed) {
        m_recovery->setStateProvider([this]() { return captureRecoveryState(); });
    } else {
        qWarning() << "RecoveryManager: adoptPreviousSession() could not promote the "
                      "prior live store; skipping rolling-flush wiring this session so "
                      "stranded crash data in the live dir is preserved for recovery.";
    }

    // Plan C T8: wire the offline banner to the monitor + start the monitor.
    // setupCentralWidget() constructs m_offlineBanner (hidden by default).
    if (m_offlineBanner) {
        connect(m_offlineBanner, &DVE::OfflineBanner::retryClicked, this,
                &MainWindow::onOfflineRetryClicked);
    }
    if (m_monitor) {
        m_monitor->start();
    }
    // If we booted offline (PG open failed but snapshot is available), show
    // the banner up-front so the user understands the read-only state.
    // Note: on the offline-boot path m_db->isOpen() is false (open() failed),
    // so the gate keys off m_snapshot->isOpen() + !m_db->isOnline() instead.
    if (m_db && !m_db->isOnline() && m_snapshot && m_snapshot->isOpen()
        && m_offlineBanner) {
        m_offlineBanner->setLastSync(m_snapshot->snapshotTakenAt());
        m_offlineBanner->setPendingCount(0);
        m_offlineBanner->setVisible(true);
    }

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

    // v2.0.1: debounce cell-focus broadcasts so arrow-key scrubbing
    // doesn't NOTIFY-storm the server. The latest selection wins; only
    // the final focusCell/blurCell call is sent.
    m_focusCommitTimer = new QTimer(this);
    m_focusCommitTimer->setSingleShot(true);
    m_focusCommitTimer->setInterval(120);
    connect(m_focusCommitTimer, &QTimer::timeout, this, [this]() {
        if (!m_liveSync) return;
        if (m_pendingFocusBlur) {
            m_liveSync->blurCell();
        } else {
            m_liveSync->focusCell(m_pendingFocusTable,
                                  m_pendingFocusRowId,
                                  m_pendingFocusColumn);
        }
    });

    m_updateChecker = new UpdateChecker(this);
    // Plan C C7: the updater's "Update Now" hard-exits via std::_Exit(0),
    // bypassing closeEvent and every destructor (identical to a crash). Flush a
    // complete, synchronous recovery snapshot just before that exit so an update
    // never loses in-flight work. m_recovery is constructed above, so it always
    // exists here; flushNow(true) is a no-op when no state provider is set (i.e.
    // when recovery isn't armed), so this is safe regardless of m_recoveryArmed.
    m_updateChecker->setPreExitHook([this] {
        if (m_recovery) m_recovery->flushNow(true);
    });
    QTimer::singleShot(1500, m_updateChecker, &UpdateChecker::start);

    // Identity bootstrap — the IdentityManager is constructed earlier so it
    // can be passed into DatabaseManager::open(). QTimer::singleShot(0, ...)
    // defers the first-launch dialog until after the main window has fully
    // shown, centering it cleanly.
    if (m_identity->firstLaunchPending()) {
        QTimer::singleShot(0, this, [this]() {
            DVE::IdentityPromptDialog dlg(m_identity, m_pgConn, this);
            dlg.exec();
        });
    }

    // Plan C C8: after the window is shown and the panels exist, offer to reload
    // a previous session that ended in a crash / hard-exit update. Deferred via
    // singleShot(0) so the modal dialog has a fully-constructed, visible parent.
    // This runs regardless of m_recoveryArmed: it only reads hasRecoverable(),
    // which is false unless adoptPreviousSession() moved a non-empty store to
    // Recovery_prev/. (When recovery wasn't armed, the crash data is still in the
    // *live* dir, not _prev, so hasRecoverable() is false and we don't double-offer.)
    QTimer::singleShot(0, this, &MainWindow::maybeOfferRecovery);

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
    buildSettingsTab(m_ribbon->addTab("Settings"));

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
        AppTheme::icon("file-plus"), "Create a new test file from template");
    m_homeLoadBtn  = fileGrp->addLargeButton("Load File",
        AppTheme::icon("folder-open"), "Open an Excel file (Ctrl+O)");
    m_homeCloseBtn = fileGrp->addLargeButton("Close",
        AppTheme::icon("x"), "Close current file");

    connect(m_homeNewBtn,   &QToolButton::clicked, this, &MainWindow::onNewFile);
    connect(m_homeLoadBtn,  &QToolButton::clicked, this, &MainWindow::onLoadFile);
    connect(m_homeCloseBtn, &QToolButton::clicked, this, &MainWindow::onCloseFile);

    // Sensory-mode buttons (initially hidden)
    m_homeSensNewBtn   = fileGrp->addLargeButton("New\nSession",
        AppTheme::icon("file-plus"), "Create a new sensory session");
    m_homeSensSaveBtn  = fileGrp->addLargeButton("Save",
        AppTheme::icon("save"), "Save session (Ctrl+S)");
    m_homeSensLoadXlBtn = fileGrp->addLargeButton("Load\nExcel",
        AppTheme::icon("folder-open"), "Load sensory data from Excel");
    m_homeSensCloseBtn  = fileGrp->addLargeButton("Close",
        AppTheme::icon("x"), "Close selected session(s)");

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
    // Route through onLoadFile so file-type detection drives the mode.
    // Loading a TPM file from sensory mode switches to TPM mode automatically.
    connect(m_homeSensLoadXlBtn, &QToolButton::clicked, this, &MainWindow::onLoadFile);
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

    // ── Detailed Sensory buttons (hidden by default) ────────────────────
    m_homeDetSensNewBtn   = fileGrp->addLargeButton("New\nSession",
        AppTheme::icon("file-plus"), "Create a new detailed sensory session");
    m_homeDetSensSaveBtn  = fileGrp->addLargeButton("Save",
        AppTheme::icon("save"), "Save session (Ctrl+S)");
    m_homeDetSensLoadXlBtn = fileGrp->addLargeButton("Load\nExcel",
        AppTheme::icon("folder-open"), "Load detailed sensory data from Excel");
    m_homeDetSensCloseBtn  = fileGrp->addLargeButton("Close",
        AppTheme::icon("x"), "Close selected session(s)");

    m_homeDetSensNewBtn->setVisible(false);
    m_homeDetSensSaveBtn->setVisible(false);
    m_homeDetSensLoadXlBtn->setVisible(false);
    m_homeDetSensCloseBtn->setVisible(false);

    connect(m_homeDetSensNewBtn, &QToolButton::clicked, this, [this]() {
        if (m_detailedSensoryPanel) m_detailedSensoryPanel->newSession();
    });
    connect(m_homeDetSensSaveBtn, &QToolButton::clicked, this, [this]() {
        if (m_detailedSensoryPanel) m_detailedSensoryPanel->save();
    });
    connect(m_homeDetSensLoadXlBtn, &QToolButton::clicked, this, &MainWindow::onLoadFile);
    connect(m_homeDetSensCloseBtn, &QToolButton::clicked, this, [this]() {
        if (!m_detailedSensoryPanel) return;
        QVector<int> indices;
        for (auto* item : m_detailedSensoryNav->selectedItems())
            indices.append(m_detailedSensoryNav->row(item));
        if (indices.isEmpty() && m_detailedSensoryPanel->currentSessionIndex() >= 0)
            indices.append(m_detailedSensoryPanel->currentSessionIndex());
        if (indices.isEmpty()) return;
        m_detailedSensoryPanel->closeSessions(indices);
        updateImageButton();
    });

    // ── SOPs group ───────────────────────────────────────────────────────────
    auto* sopGrp = tab->addGroup("Reference");
    auto* sopBtn = sopGrp->addLargeButton("SOPs",
        AppTheme::icon("info"),
        "View Standard Operating Procedures");
    connect(sopBtn, &QToolButton::clicked, this, [this]() {
        QString sopPath = resourcePath() + "/templates/Standardized Test Template - December 2025.xlsx";
        SopDialog dlg(sopPath, this);
        dlg.exec();
    });

    // ── Database group ────────────────────────────────────────────────────────
    auto* dbGrp = tab->addGroup("Database");
    auto* dbBtn = dbGrp->addLargeButton("Database",
        AppTheme::icon("database"), "Browse file database");
    connect(dbBtn, &QToolButton::clicked, this, &MainWindow::onOpenDatabaseBrowser);

    auto* refreshSnapBtn = dbGrp->addLargeButton("Refresh\nSnapshot",
        AppTheme::icon("refresh-cw"),
        "Regenerate the local read-only copy of the database for offline use");
    connect(refreshSnapBtn, &QToolButton::clicked,
            this, &MainWindow::onRefreshSnapshotTriggered);

    // ── Modes group ───────────────────────────────────────────────────────────
    auto* modeGrp = tab->addGroup("Modes");
    m_sensoryBtn = modeGrp->addLargeButton("Sensory",
        AppTheme::icon("sparkles"), "Toggle sensory evaluation mode");
    m_sensoryBtn->setCheckable(true);
    connect(m_sensoryBtn, &QToolButton::toggled, this, &MainWindow::toggleSensoryMode);

    m_detailedSensoryBtn = modeGrp->addLargeButton("Detailed\nSensory",
        AppTheme::icon("list-checks"),
        "Toggle detailed sensory evaluation mode (S2-1)");
    m_detailedSensoryBtn->setCheckable(true);
    connect(m_detailedSensoryBtn, &QToolButton::toggled,
            this, &MainWindow::toggleDetailedSensoryMode);

    // ── Images group ──────────────────────────────────────────────────────────
    auto* imgGrp = tab->addGroup("Images");
    m_inboxBtn = imgGrp->addLargeButton("Images",
        AppTheme::icon("image"),
        "Open Image Inbox to assign photos to samples");
    connect(m_inboxBtn, &QToolButton::clicked, this, &MainWindow::onOpenImageInbox);
}

void MainWindow::buildReportsTab(RibbonTab* tab)
{
    auto* rptGrp  = tab->addGroup("Generate");
    m_reportBtn1 = rptGrp->addLargeButton("Test Report",
        AppTheme::icon("file-text"),
        "Generate a PPTX report for the current sheet");
    m_reportBtn2 = rptGrp->addLargeButton("Full Report",
        AppTheme::icon("files"),
        "Generate a PPTX report for all sheets");

    connect(m_reportBtn1, &QToolButton::clicked, this, &MainWindow::onGenerateTestReport);
    connect(m_reportBtn2, &QToolButton::clicked, this, &MainWindow::onGenerateFullReport);

    // ── Data Cleanup group ────────────────────────────────────────────────────
    m_cleanupGroup = tab->addGroup("Data Cleanup");
    auto* cleanBtn   = m_cleanupGroup->addLargeButton("Clean Data",
        AppTheme::icon("eraser"),
        "Open the data cleanup dialog to exclude outliers from plots and reports");
    m_resetCleanupBtn = m_cleanupGroup->addLargeButton("Reset Cleanup",
        AppTheme::icon("rotate-ccw"),
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
    auto* extGrp = tab->addGroup("External Tools");
    auto* translatorBtn = extGrp->addLargeButton("Translator",
        AppTheme::icon("languages"),
        "Open Document Translator");
    connect(translatorBtn, &QToolButton::clicked, this, &MainWindow::onLaunchTranslator);

    auto* recGrp = tab->addGroup("Recovery");
    auto* recoverBtn = recGrp->addLargeButton("Recover",
        AppTheme::icon("rotate-ccw"),
        "Restore unsaved work from the last session");
    connect(recoverBtn, &QToolButton::clicked, this, &MainWindow::onRecover);
}

void MainWindow::buildSettingsTab(RibbonTab* tab)
{
    RibbonGroup* grp = tab->addGroup(QStringLiteral("Output Paths"));

    struct PathBtn { QString label; QString title; ReportMode mode; };
    const QVector<PathBtn> defs = {
        { QStringLiteral("Set TPM Output Path"),
          QStringLiteral("Select TPM Report Output Folder"),              ReportMode::Tpm },
        { QStringLiteral("Set Sensory Output Path"),
          QStringLiteral("Select Sensory Report Output Folder"),          ReportMode::Sensory },
        { QStringLiteral("Set Detailed Sensory Output Path"),
          QStringLiteral("Select Detailed Sensory Report Output Folder"), ReportMode::DetailedSensory },
    };

    auto tip = [](ReportMode m) -> QString {
        const QString configured = OutputPaths::configuredDir(m);
        return configured.isEmpty() ? QStringLiteral("Not set — defaults to Documents") : configured;
    };

    for (const PathBtn& d : defs) {
        QToolButton* btn = grp->addLargeButton(d.label,
                                               AppTheme::icon(QStringLiteral("folder-open")),
                                               tip(d.mode));
        const ReportMode mode = d.mode;
        const QString title = d.title;
        connect(btn, &QToolButton::clicked, this, [this, btn, mode, title, tip]() {
            const QString cur = OutputPaths::configuredDir(mode);
            const QString start = cur.isEmpty() ? OutputPaths::documentsDir() : cur;
            const QString dir = QFileDialog::getExistingDirectory(this, title, start);
            if (!dir.isEmpty()) {
                OutputPaths::setConfiguredDir(mode, dir);
                btn->setToolTip(tip(mode));
            }
        });
    }
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

    // Nav bar: Prev | count | Next | … | Add Row | Remove Row
    // Styled frame with accent-subtle background for visual hierarchy.
    m_sampleNavBar = new QWidget(this);
    m_sampleNavBar->setObjectName("sampleNavBar");
    m_sampleNavBar->setStyleSheet(QString(
        "#sampleNavBar { background-color: %1; border: 1px solid %2;"
        "  border-radius: 4px; padding: 4px 8px; }")
        .arg(AppTheme::accentSubtle().name(), AppTheme::borderSubtle().name()));
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
    const QString navBtnStyle = QString(
        "QPushButton { border:1px solid %1; border-radius:4px;"
        "  background:#E0E0E0; color:%2; font-size:11pt; padding:0px; }"
        "QPushButton:hover { background:%3; border-color:%4; color:#003388; }"
        "QPushButton:pressed { background:%5; }"
        "QPushButton:disabled { background:#EBEBEB; color:#AAAAAA; }")
        .arg(AppTheme::border().name(),
             AppTheme::textPrimary().name(),
             AppTheme::hoverBg().name(),
             AppTheme::accent().name(),
             AppTheme::selectBg().name());
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
        btn->setFont(AppTheme::fontSmall());
    }
    // Add Row gets primary styling via global QSS; Remove Row stays default.
    m_addRowBtn->setProperty("primary", true);
    m_addRowBtn->style()->unpolish(m_addRowBtn);
    m_addRowBtn->style()->polish(m_addRowBtn);
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
    // v2.0.1: paints remote-focus border + name flag + flash decoration.
    m_cellFocusDelegate = new DVE::CellFocusDelegate(this);
    m_dataTable->setItemDelegate(m_cellFocusDelegate);
    // Dual-mode combo editor for column 4 (Resistance / Puffing Regime).
    m_regimeDelegate = new DVE::RegimeComboDelegate(this);
    m_dataTable->setItemDelegateForColumn(DVE::Cols::RESISTANCE, m_regimeDelegate);
    m_dataTable->horizontalHeader()->setStretchLastSection(false);
    m_dataTable->verticalHeader()->setDefaultSectionSize(28);
    m_dataTable->setShowGrid(true);
    m_dataTable->setGridStyle(Qt::SolidLine);
    m_dataTable->setAlternatingRowColors(true);
    m_dataTable->setSelectionBehavior(QAbstractItemView::SelectRows);
    m_dataTable->setEditTriggers(QAbstractItemView::DoubleClicked | QAbstractItemView::AnyKeyPressed);
    m_dataTable->setWordWrap(false);
    connect(m_dataTable, &QTableWidget::cellChanged,
            this, &MainWindow::onTableCellChanged);
    // Phase 6 (T17): dirty tracking on edit. Wired once here; displayCurrentSample
    // stamps baseline values at populate time (signals blocked), so this only
    // fires on genuine user edits.
    connect(m_dataTable, &QTableWidget::itemChanged,
            this, &MainWindow::onDataTableItemChanged);
    // Phase 6 (T18): clicking a yellow-decorated cell accepts the remote value.
    connect(m_dataTable, &QTableWidget::itemClicked,
            this, &MainWindow::onDataTableItemClicked);
    connect(m_dataTable, &QTableWidget::itemSelectionChanged, this, [this]() {
        if (m_removeRowBtn)
            m_removeRowBtn->setEnabled(m_dataTable->currentRow() >= 0);
    });
    // v2.0.1: broadcast which cell this user is editing so peers can paint
    // a Variant-C border around it. data_rows.id lives on the vertical
    // header at Qt::UserRole (set during displayCurrentSample). The
    // broadcast is debounced by m_focusCommitTimer so arrow-keying
    // through cells doesn't fire one round-trip per cell.
    connect(m_dataTable, &QTableWidget::currentItemChanged, this,
            [this](QTableWidgetItem* curr, QTableWidgetItem* /*prev*/) {
                if (!m_liveSync || !m_focusCommitTimer) return;
                if (!curr) {
                    m_pendingFocusBlur = true;
                    m_focusCommitTimer->start();
                    return;
                }
                const QString column = liveColumnForDataCol(curr->column());
                const QTableWidgetItem* vh =
                    m_dataTable->verticalHeaderItem(curr->row());
                const qint64 rowId = vh ? vh->data(Qt::UserRole).toLongLong() : -1;
                if (rowId > 0 && !column.isEmpty()) {
                    m_pendingFocusBlur   = false;
                    m_pendingFocusTable  = QStringLiteral("data_rows");
                    m_pendingFocusRowId  = rowId;
                    m_pendingFocusColumn = column;
                } else {
                    m_pendingFocusBlur = true;
                }
                m_focusCommitTimer->start();
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
    m_centralSplitter->setStretchFactor(0, 40);
    m_centralSplitter->setStretchFactor(1, 60);

    // Wrap in a stacked widget (index 0 = TPM, index 1 = sensory, added lazily)
    m_centralStack = new QStackedWidget(this);
    m_centralStack->addWidget(m_centralSplitter);   // index 0

    // Avatar bar lives above the central stack. The container becomes the
    // window's central widget; the stack still drives mode-switching, the
    // bar just sits on top showing live presence for the active resource.
    m_avatarBar = new DVE::PresenceAvatarBar(this);

    // Row-deleted banner — sits between the avatar bar and the central
    // stack. Hidden by default; shown when a DELETE NOTIFY arrives for the
    // currently-open resource (Phase 6 T19).
    m_rowDeletedBanner = new DVE::RowDeletedBanner(this);

    // Plan C T8: OfflineBanner sits at the very top of the central area
    // (above the avatar bar). Hidden by default — ConnectionMonitor flips
    // it on/off via wentOffline/cameOnline.
    m_offlineBanner = new DVE::OfflineBanner(this);
    m_offlineBanner->setVisible(false);

    // Plan B B8: IncompleteDataBanner — shown just below OfflineBanner when
    // the active TPM FileResult has any sheet with dbDataIncomplete == true.
    // Hidden by default; updateIncompleteDataBanner() toggles it.
    m_incompleteDataBanner = new DVE::IncompleteDataBanner(this);

    auto* centralContainer = new QWidget(this);
    auto* centralVL        = new QVBoxLayout(centralContainer);
    centralVL->setContentsMargins(0, 0, 0, 0);
    centralVL->setSpacing(0);
    centralVL->addWidget(m_offlineBanner);
    centralVL->addWidget(m_incompleteDataBanner);
    centralVL->addWidget(m_avatarBar);
    centralVL->addWidget(m_rowDeletedBanner);
    centralVL->addWidget(m_centralStack, 1);
    setCentralWidget(centralContainer);

    // Recreate flow: the banner emits when the user clicks; we push the
    // currently-open resource directly through DatabaseManager with
    // id/version cleared so it becomes a fresh INSERT. v2.0.1 retired
    // SaveCoordinator — conflict dialogs no longer surface; UniqueViolation
    // is the only realistic failure path here (duplicate natural key) and
    // is surfaced via the status bar.
    connect(m_rowDeletedBanner, &DVE::RowDeletedBanner::recreateRequested,
            this, [this]() {
        if (!m_db) {
            m_rowDeletedBanner->dismiss();
            return;
        }
        // The "resource that was deleted" is whatever the user has open
        // right now. Snapshot before the save so we can locate the right
        // local struct on the resource-type branch.
        const QString resType = m_currentResourceType;
        const qint64  resId   = m_currentResourceId;
        bool savedOk = false;
        if (resType == QLatin1String("file")) {
            for (int i = 0; i < m_loadedFiles.size(); ++i) {
                if (qint64(m_loadedFiles[i].id) != resId) continue;
                // Reset id/version so the tryWrite path takes the INSERT
                // branch. The mutable-ref tryWriteFile overload (I1 fix)
                // writes the post-save id+version back into m_loadedFiles[i].
                m_loadedFiles[i].id      = -1;
                m_loadedFiles[i].version = 0;
                savedOk = (m_db->tryWriteFile(m_loadedFiles[i])
                           == DVE::WriteResult::Success);
                if (savedOk) {
                    const int newId = m_loadedFiles[i].id;
                    // I3: after a successful recreate, the new file row's
                    // child rows (samples / data_rows / images) all got
                    // fresh server-assigned ids. Stamped vertical-header
                    // DataRow ids on the table are now stale, so re-load
                    // the FileResult and re-populate the current sample.
                    if (newId > 0) {
                        FileResult fresh = m_db->loadFile(newId);
                        if (!fresh.filePath.isEmpty()) {
                            m_loadedFiles[i] = fresh;
                            if (i == m_currentFileIndex) {
                                displayCurrentSample();
                            }
                        }
                    }
                    m_currentResourceId = qint64(m_loadedFiles[i].id);
                    refreshPresenceFor(m_currentResourceType,
                                       m_currentResourceId);
                }
                break;
            }
        } else if (resType == QLatin1String("sensory_session")
                   && m_sensoryPanel) {
            // Scope limitation: recreate only works when the deleted
            // session is the currently-selected one. If the user has
            // navigated away (or no session is current), we log and
            // fail the banner.
            if (SensorySession* curr = m_sensoryPanel->currentSession();
                curr && qint64(curr->id) == resId) {
                curr->id      = -1;
                curr->version = 0;
                savedOk = (m_db->tryWriteSensorySession(*curr)
                           == DVE::WriteResult::Success);
                if (savedOk) {
                    // H10: refetch the new row so child image IDs / version
                    // populate from the server. Without this, the in-memory
                    // session keeps the post-INSERT id but image_ids stay
                    // -1, and the next OCC-protected commit against an image
                    // row would fail VersionMismatch against a stale anchor.
                    if (curr->id > 0) {
                        SensorySession fresh =
                            m_db->loadSensorySession(curr->id);
                        if (fresh.id > 0) *curr = fresh;
                    }
                    m_currentResourceId = qint64(curr->id);
                    refreshPresenceFor(m_currentResourceType,
                                       m_currentResourceId);
                    refreshSensoryNavigator();
                }
            } else {
                qWarning() << "Sensory recreate skipped: current session does "
                              "not match deleted resource id" << resId;
            }
        } else if (resType == QLatin1String("detailed_sensory_session")
                   && m_detailedSensoryPanel) {
            if (DetailedSensorySession* curr =
                    m_detailedSensoryPanel->currentSession();
                curr && qint64(curr->id) == resId) {
                curr->id      = -1;
                curr->version = 0;
                savedOk = (m_db->tryWriteDetailedSensorySession(*curr)
                           == DVE::WriteResult::Success);
                if (savedOk) {
                    // H10 (symmetric with the SensorySession branch):
                    // refetch the new row so child image IDs / version
                    // populate from the server before the next OCC commit.
                    if (curr->id > 0) {
                        DetailedSensorySession fresh =
                            m_db->loadDetailedSensorySession(curr->id);
                        if (fresh.id > 0) *curr = fresh;
                        m_currentResourceId = qint64(curr->id);
                        refreshPresenceFor(m_currentResourceType,
                                           m_currentResourceId);
                    }
                    refreshDetailedSensoryNavigator();
                }
            } else {
                qWarning() << "Detailed-sensory recreate skipped: current "
                              "session does not match deleted resource id"
                           << resId;
            }
        }
        if (savedOk) {
            updateStatusBar(tr("Resource recreated as a new row."));
            // H10: only dismiss the banner when the recreate landed. On
            // failure the user keeps the banner visible so they can try
            // again or copy their data out before navigating away.
            m_rowDeletedBanner->dismiss();
        } else {
            qWarning() << "Recreate failed for" << resType << resId;
            updateStatusBar(tr("Recreate failed — see log for details."));
        }
    });
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
    m_fileTree->setIndentation(16);
    m_fileTree->setUniformRowHeights(true);
    m_fileTree->setAlternatingRowColors(true);
    m_navStack->addWidget(m_fileTree);   // index 0

    m_sensoryNav = new QListWidget(filePanel);
    m_sensoryNav->setAlternatingRowColors(true);
    m_sensoryNav->setSelectionMode(QAbstractItemView::ExtendedSelection);
    m_sensoryNav->setToolTip(
        "Click a session to view it.\n"
        "Ctrl+Click to select multiple — the chart and table show the average\n"
        "of just those sessions.\n"
        "Double-click a label to rename it.");
    m_navStack->addWidget(m_sensoryNav); // index 1

    // Single PresenceDotsDelegate shared across all three nav widgets. Cheap
    // (one QObject) and keeps painting consistent. The detailed-sensory nav
    // is created later in initDetailedSensoryPanel — it picks up the same
    // delegate via the same member pointer.
    m_presenceDelegate = new DVE::PresenceDotsDelegate(this);
    m_fileTree->setItemDelegateForColumn(0, m_presenceDelegate);
    m_sensoryNav->setItemDelegate(m_presenceDelegate);

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
    avgHeader->setStyleSheet(QString(
        "background:%1; color:white; font-weight:600; font-size:8pt;")
        .arg(AppTheme::tableHeader().name()));

    m_testAvgList = new QListWidget(m_testAvgPanel);
    m_testAvgList->setAlternatingRowColors(true);
    m_testAvgList->setStyleSheet(
        "QListWidget { font-size: 8pt; } "
        "QListWidget::item { padding: 1px 2px; } ");
    m_testAvgList->setSpacing(0);
    m_testAvgList->setToolTip(
        "Click a test to average ALL sessions for that test.\n"
        "This clears any selection above and shows every sample.");

    // Hidden table — data shown in the main panel overlay instead
    m_testAvgTable = new QTableWidget(0, 2, m_testAvgPanel);
    m_testAvgTable->setVisible(false);

    m_testAvgAssessors = new QLabel("Assessors: —", m_testAvgPanel);
    m_testAvgAssessors->setWordWrap(true);
    const QString infoLabelStyle = QString("font-size: 7pt; padding: 1px 4px; color: %1;")
        .arg(AppTheme::textSec().name());
    m_testAvgAssessors->setStyleSheet(infoLabelStyle);
    m_testAvgTesters = new QLabel("Testers: —", m_testAvgPanel);
    m_testAvgTesters->setWordWrap(true);
    m_testAvgTesters->setStyleSheet(infoLabelStyle);
    m_testAvgCount = new QLabel("Sessions: 0", m_testAvgPanel);
    m_testAvgCount->setStyleSheet(infoLabelStyle);
    // "Showing 1+2" indicator — populated by Ctrl+click on the session nav.
    // Hidden when 0 or 1 session is selected.
    m_showingLabel = new QLabel(QString(), m_testAvgPanel);
    m_showingLabel->setWordWrap(true);
    m_showingLabel->setStyleSheet(QString(
        "font-size: 7pt; padding: 1px 4px; color: %1; font-style: italic;")
        .arg(AppTheme::accent().name()));
    m_showingLabel->setVisible(false);

    avgVL->addWidget(avgHeader);
    avgVL->addWidget(m_testAvgList, 1);
    avgVL->addWidget(m_showingLabel);
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
    propHeader->setStyleSheet(QString(
        "background:%1; color:white; font-weight:600; font-size:8pt;")
        .arg(AppTheme::tableHeader().name()));

    m_propTable = new QTableWidget(0, 2, m_propPanel);
    m_propTable->setHorizontalHeaderLabels({"Property", "Value"});
    m_propTable->horizontalHeader()->setSectionResizeMode(0, QHeaderView::Interactive);
    m_propTable->horizontalHeader()->setSectionResizeMode(1, QHeaderView::Stretch);
    m_propTable->setColumnWidth(0, 110);   // ~40% of typical 280 px dock width
    m_propTable->verticalHeader()->setVisible(false);
    m_propTable->verticalHeader()->setDefaultSectionSize(22);
    m_propTable->setShowGrid(true);
    m_propTable->setGridStyle(Qt::SolidLine);
    m_propTable->setAlternatingRowColors(true);
    m_propTable->setSelectionMode(QAbstractItemView::SingleSelection);
    m_propTable->setEditTriggers(QAbstractItemView::DoubleClicked | QAbstractItemView::AnyKeyPressed);
    // Styling via global QSS — no per-widget stylesheet needed.
    connect(m_propTable, &QTableWidget::cellChanged, this, &MainWindow::onPropCellChanged);

    propVL->addWidget(propHeader);
    propVL->addWidget(m_propTable, 1);

    // ── Image buttons below the properties table ──────────────────────────────
    QWidget*     imgBar    = new QWidget(m_propPanel);
    QHBoxLayout* imgLayout = new QHBoxLayout(imgBar);
    imgLayout->setContentsMargins(8, 4, 8, 4);
    imgLayout->setSpacing(8);

    m_loadImagesBtn = new QPushButton("Load Images", imgBar);
    m_viewImagesBtn = new QPushButton("View Images (0)", imgBar);
    m_viewImagesBtn->setEnabled(false);
    // Styling via global QSS QPushButton rule — no per-button stylesheets.
    m_loadImagesBtn->setFont(AppTheme::fontSmall());
    m_viewImagesBtn->setFont(AppTheme::fontSmall());

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

    // ── Sidebar collapse stack (Task 7) ───────────────────────────────────────
    // index 0 = full panel (leftSplitter), index 1 = 32 px icon strip
    m_sidebarFullPanel = leftSplitter;

    // Icon strip: three small tool buttons for navigator / properties / images
    m_sidebarIconStrip = new QWidget(this);
    m_sidebarIconStrip->setFixedWidth(32);
    m_sidebarIconStrip->setObjectName("sidebarIconStrip");
    m_sidebarIconStrip->setStyleSheet(
        "#sidebarIconStrip { background: " + AppTheme::surfacePanel().name() + ";"
        "  border-right: 1px solid " + AppTheme::borderSubtle().name() + "; }");
    auto* stripVL = new QVBoxLayout(m_sidebarIconStrip);
    stripVL->setContentsMargins(0, 4, 0, 4);
    stripVL->setSpacing(2);

    auto makeStripBtn = [&](const QString& iconName, const QString& tip) {
        auto* btn = new QToolButton(m_sidebarIconStrip);
        btn->setIcon(AppTheme::icon(iconName));
        btn->setIconSize(QSize(20, 20));
        btn->setFixedSize(32, 32);
        btn->setToolTip(tip);
        btn->setAutoRaise(true);
        connect(btn, &QToolButton::clicked, this, &MainWindow::showSidebarOverlay);
        return btn;
    };
    stripVL->addWidget(makeStripBtn("folder-open", "Navigator"));
    stripVL->addWidget(makeStripBtn("menu",         "Sample Properties"));
    stripVL->addWidget(makeStripBtn("image",        "Images"));
    stripVL->addStretch();

    m_sidebarStack = new QStackedWidget(this);
    m_sidebarStack->addWidget(m_sidebarFullPanel);  // index 0
    m_sidebarStack->addWidget(m_sidebarIconStrip);  // index 1

    m_fileDock->setWidget(m_sidebarStack);
    addDockWidget(Qt::LeftDockWidgetArea, m_fileDock);
    m_fileDock->setMinimumWidth(32);
    m_fileDock->setMaximumWidth(320);
}

void MainWindow::showSidebarOverlay()
{
    // Simple implementation: switch from icon-strip back to full panel.
    // The user sees the full sidebar until the window widens past the
    // breakpoint, which triggers the ResponsiveLayout lambda to reset the
    // correct index. If still compact, clicking elsewhere doesn't auto-
    // collapse; the user must resize. A true Qt::Popup overlay is a follow-up.
    if (m_sidebarStack)
        m_sidebarStack->setCurrentIndex(0);
}

void MainWindow::setupStatusBar()
{
    buildStatusBar();
    updateDbSyncIndicator();
}

void MainWindow::buildStatusBar()
{
    // Remove the default QStatusBar (frees its space and removes the grip).
    setStatusBar(nullptr);

    m_statusBarWidget = new QWidget(this);
    m_statusBarWidget->setObjectName("dvStatusBar");
    m_statusBarWidget->setStyleSheet(QString(
        "#dvStatusBar { background-color: %1; border-top: 1px solid %2; }"
        "#dvStatusBar QLabel { color: %3; padding: 0px 4px; font-size: 9pt; }"
    ).arg(AppTheme::surfaceStatusBar().name(),
          AppTheme::borderDefault().name(),
          AppTheme::textPrimary().name()));

    auto* hl = new QHBoxLayout(m_statusBarWidget);
    hl->setContentsMargins(8, 4, 8, 4);
    hl->setSpacing(6);

    auto makeDot = [this](const QColor& c) {
        auto* d = new QLabel(QStringLiteral("●"), m_statusBarWidget);
        d->setStyleSheet(QString("color: %1;").arg(c.name()));
        return d;
    };

    // ── File segment ──────────────────────────────────────────────────────────
    m_statusFileDot  = makeDot(AppTheme::textMuted());
    m_statusFileText = new QLabel("File closed", m_statusBarWidget);
    hl->addWidget(m_statusFileDot);
    hl->addWidget(m_statusFileText);

    auto* sep1 = new QLabel("|", m_statusBarWidget);
    sep1->setStyleSheet(QString("color: %1;").arg(AppTheme::borderDefault().name()));
    hl->addWidget(sep1);

    // ── DB segment ────────────────────────────────────────────────────────────
    m_statusDbDot  = makeDot(AppTheme::success());
    m_statusDbText = new QLabel("Local DB ready", m_statusBarWidget);
    hl->addWidget(m_statusDbDot);
    hl->addWidget(m_statusDbText);

    auto* sep2 = new QLabel("|", m_statusBarWidget);
    sep2->setStyleSheet(QString("color: %1;").arg(AppTheme::borderDefault().name()));
    hl->addWidget(sep2);

    // ── Breadcrumb segment ────────────────────────────────────────────────────
    m_statusBreadcrumb = new QLabel("", m_statusBarWidget);
    m_statusBreadcrumb->setStyleSheet(
        QString("color: %1;").arg(AppTheme::textSecondary().name()));
    hl->addWidget(m_statusBreadcrumb);

    // ── Progress bar (right-aligned, hidden until needed) ─────────────────────
    m_progressBar = new QProgressBar(m_statusBarWidget);
    m_progressBar->setRange(0, 100);
    m_progressBar->setVisible(false);
    m_progressBar->setMaximumWidth(200);
    m_progressBar->setMaximumHeight(16);
    hl->addStretch(1);
    hl->addWidget(m_progressBar);

    // Attach the custom bar at the bottom of the central container's VBoxLayout.
    if (auto* cw = centralWidget()) {
        if (auto* lay = qobject_cast<QVBoxLayout*>(cw->layout())) {
            lay->addWidget(m_statusBarWidget);
        }
    }
}

void MainWindow::setStatusFile(const QString& text, FileStatus status)
{
    QColor c;
    switch (status) {
        case FileStatusOk:       c = AppTheme::success();   break;
        case FileStatusModified: c = AppTheme::warning();   break;
        case FileStatusClosed:   c = AppTheme::textMuted(); break;
    }
    if (m_statusFileDot)  m_statusFileDot->setStyleSheet(
        QString("color: %1;").arg(c.name()));
    if (m_statusFileText) m_statusFileText->setText(text);
}

void MainWindow::setStatusDb(const QString& text, DbStatus status)
{
    QColor c;
    switch (status) {
        case DbStatusOk:           c = AppTheme::success(); break;
        case DbStatusModified:     c = AppTheme::warning(); break;
        case DbStatusDisconnected: c = AppTheme::error();   break;
    }
    if (m_statusDbDot)  m_statusDbDot->setStyleSheet(
        QString("color: %1;").arg(c.name()));
    if (m_statusDbText) m_statusDbText->setText(text);
}

void MainWindow::setStatusBreadcrumb(const QStringList& segments)
{
    m_lastBreadcrumbSegments = segments;
    if (!m_statusBreadcrumb) return;
    // Use U+203A "›" as separator. In compact mode, truncate middle segments.
    if (DVE::ResponsiveLayout::instance().isCompact() && segments.size() > 2) {
        m_statusBreadcrumb->setText(
            segments.first() + QStringLiteral(" › … › ") + segments.last());
    } else {
        m_statusBreadcrumb->setText(segments.join(QStringLiteral(" › ")));
    }
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

    // Ctrl+U: flush dirty TPM files + sensory sessions to the database.
    // LiveSync handles per-cell persistence for sensory cells, but the
    // session-level safety net + TPM file writes still need an explicit
    // trigger — the 5 s auto-save tick was reintroducing UI freezes on
    // slower LANs, so we surface Ctrl+U again and let the user decide
    // when to flush.
    auto* dbUpdateAct = new QAction(this);
    dbUpdateAct->setShortcut(QKeySequence(Qt::CTRL | Qt::Key_U));
    connect(dbUpdateAct, &QAction::triggered, this, &MainWindow::onUpdateDatabase);
    addAction(dbUpdateAct);

    // Sensory navigator selection → switch session, or compute averages on multi-select.
    // The session list runs in ExtendedSelection so Ctrl+Click toggles inclusion;
    // we react here to keep the radar chart, the averaged table overlay, and the
    // "(Showing N+M)" hint label all in sync with whatever's currently selected.
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
            if (m_showingLabel) m_showingLabel->setVisible(false);

            // Phase 5: presence activate on session open. The item carries
            // the DB session id stamped by refreshSensoryNavigator.
            if (m_presence) {
                const qint64 sessId =
                    selected.first()->data(Qt::UserRole).toLongLong();
                if (sessId > 0) {
                    m_presence->activate(QStringLiteral("sensory_session"),
                                         sessId, QStringLiteral("viewing"));
                    m_currentResourceType = QStringLiteral("sensory_session");
                    m_currentResourceId   = sessId;
                    refreshPresenceFor(m_currentResourceType,
                                       m_currentResourceId);
                }
            }
        } else if (selected.size() > 1) {
            // Multi-selection (Ctrl+Click): show averaged radar chart AND
            // populate the left-side averaged-table overlay so the user sees
            // the per-device averages of just the selected sessions.
            QVector<int> indices;
            for (auto* item : selected)
                indices.append(m_sensoryNav->row(item));
            std::sort(indices.begin(), indices.end());

            const auto allSess = m_sensoryPanel->allSessions();
            QVector<SensorySession> picked;
            picked.reserve(indices.size());
            for (int idx : indices)
                if (idx >= 0 && idx < allSess.size())
                    picked.append(allSess[idx]);

            showSensoryAveragesFor(picked, indices);

            // Update the "(Showing 1+2)" hint label using 1-based numbering
            // matching the prefix shown in the navigator.
            if (m_showingLabel) {
                QStringList parts;
                for (int idx : indices) parts << QString::number(idx + 1);
                m_showingLabel->setText(QStringLiteral("(Showing ") +
                                        parts.join(QStringLiteral("+")) +
                                        QStringLiteral(")"));
                m_showingLabel->setVisible(true);
            }
            // Multi-select doesn't correspond to a single test title in the
            // Test Averages list — clear that selection so the two panels
            // don't appear to disagree.
            if (m_testAvgList) {
                QSignalBlocker b(m_testAvgList);
                m_testAvgList->clearSelection();
                m_testAvgList->setCurrentRow(-1);
            }
        } else if (selected.isEmpty()) {
            // All deselected — clear chart
            m_sensoryPanel->showAveragedChart({});
            if (m_showingLabel) m_showingLabel->setVisible(false);
        }
        updateImageButton();
    });

    // Editing a session label: strip the leading "N. " numbering before passing
    // the new name to the panel so users don't accidentally save "1. My Test".
    // The number is re-applied by refreshSensoryNavigator on the next refresh.
    connect(m_sensoryNav, &QListWidget::itemChanged, this, [this](QListWidgetItem* item) {
        if (!m_sensoryPanel) return;
        int row = m_sensoryNav->row(item);
        if (row < 0) return;
        QSignalBlocker blocker(m_sensoryNav);
        QString text = item->text();
        static const QRegularExpression kNumPrefix(
            QStringLiteral("^\\s*\\d+\\.\\s+"));
        text.remove(kNumPrefix);
        m_sensoryPanel->renameSession(row, text);
        // renameSession emits sessionsChanged → refreshSensoryNavigator may
        // clear m_sensoryNav, deleting `item`. Re-resolve via the row index
        // before touching the pointer again; the item at this row is the
        // freshly-created replacement and is safe to write through.
        QListWidgetItem* refreshedItem = m_sensoryNav->item(row);
        if (!refreshedItem) return;
        refreshedItem->setText(QStringLiteral("%1. %2").arg(row + 1).arg(text));
    });

    // Ctrl+S shortcut — routes to the active mode's save action.
    //  - TPM mode: flush the debounced Excel write-back queue. LiveSync
    //    already persists every cell commit to the DB, so this is purely
    //    a manual openpyxl flush.
    //  - Sensory mode: SensoryPanel::save() (writes .json / .xlsx as the
    //    user has configured + persists the session to Postgres).
    //  - Detailed Sensory mode: DetailedSensoryPanel::save() (same
    //    contract as SensoryPanel::save for the 14-question form).
    auto* saveAct = new QAction(this);
    saveAct->setShortcut(QKeySequence::Save);
    connect(saveAct, &QAction::triggered, this, [this]() {
        if (m_detailedSensoryMode && m_detailedSensoryPanel) {
            m_detailedSensoryPanel->save();
        } else if (m_sensoryMode && m_sensoryPanel) {
            m_sensoryPanel->save();
        } else {
            onExportToExcelTriggered();
        }
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
        OutputPaths::resolveDir(ReportMode::Tpm,m_lastBrowseDir) + "/New Test File.xlsx",
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
        case 4:
            if (sheet->hasPerRowRegime) dr.puffingRegime = text;
            else                         dr.resistance   = text.toDouble();
            break;
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
    if (sheet->hasPerRowRegime) refreshPlotRegimes();

    // Plan C T7: if we're offline, capture a PendingEdit so the change can
    // be retried after reconnect. The Excel write below still happens — the
    // local .xlsx is the source of truth for the cell's text. The PG save
    // would silently no-op (tryWriteFile returns OfflineReadOnly).
    //
    // We dedupe by (filePath, sheetIdx) — replaying a whole file save
    // covers every edit to it. v1.1 may switch to per-cell granularity
    // (and a yellow-dot decoration).
    if (m_db && !m_db->isOnline()) {
        bool alreadyQueued = false;
        for (const PendingEdit& p : m_pendingEdits) {
            if (p.filePath == file->filePath
                && p.sheetIdx == m_currentSheetIndex) {
                alreadyQueued = true;
                break;
            }
        }
        if (!alreadyQueued) {
            PendingEdit pe;
            pe.fileIdx    = m_currentFileIndex;
            pe.sheetIdx   = m_currentSheetIndex;
            pe.sampleIdx  = m_currentSampleIndex;
            pe.filePath   = file->filePath;
            pe.sheetName  = sheet->sheetName;
            pe.capturedAt = QDateTime::currentDateTime();
            m_pendingEdits.append(pe);
            if (m_offlineBanner) {
                m_offlineBanner->setPendingCount(m_pendingEdits.size());
            }
        }
        setStatusFile(tr("Working offline — change will retry when reconnected."),
                      FileStatusModified);
    }

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
                if      (label == "Control")                sess->control = value;
                else if (label == "Blind?")                 sess->isBlind = (value.toUpper() == "Y");
                else if (label == "Primary Difference(s)")  sess->primaryDifferences = value;

                if (m_db->isOpen()) {
                    // Plan C T7: surface offline state to the user — the
                    // in-memory mutation above is preserved but the PG save
                    // will return OfflineReadOnly. v1 doesn't queue sensory
                    // edits (TPM only); user must retry after reconnect.
                    if (!m_db->isOnline()) {
                        setStatusFile(tr("Working offline — sensory change will not "
                                         "save until reconnected."),
                                      FileStatusModified);
                    }
                    // v2.0.1: bulk save via DatabaseManager. LiveSync owns
                    // per-cell persistence; this prop-table autosave is a
                    // session-level fallback for fields not yet wired to
                    // LiveSync's commitCell path.
                    if (!m_db->saveSensorySession(*sess))
                        qDebug() << "[MainWindow] propTable autosave: save failed";
                }
                m_sensorySessionsDirty = true;
                // Plan C: prop-table edits (control/blind/primary-difference)
                // mutate the session directly without emitting sessionsChanged,
                // so the snapshot must be marked dirty here too.
                if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
                updateDbSyncIndicator();
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

// ─── File type detection ────────────────────────────────────────────────────
// Peek at the workbook's first row to determine if it's a Sensory, Detailed
// Sensory, or TPM file. JSON files are always Sensory. Accepts .xlsx and
// .xlsm (macro-enabled workbooks share the OOXML format; QXlsx and the
// openpyxl/COM reader treat them identically).
MainWindow::FileType MainWindow::detectFileType(const QString& path) const
{
    QString ext = QFileInfo(path).suffix().toLower();

    // JSON files are always sensory exports
    if (ext == "json")
        return FileType::Sensory;

    if (ext != "xlsx" && ext != "xlsm" && ext != "xls")
        return FileType::Unknown;

    // Quick-open with QXlsx to read first-row headers
    QXlsx::Document xlsx(path);
    if (!xlsx.load())
        return FileType::Unknown;

    QString a1 = xlsx.read(1, 1).toString().trimmed();

    // Detailed Sensory: A1 = "Sample Name" and row 1 contains detailed metric headers
    if (a1.compare("Sample Name", Qt::CaseInsensitive) == 0) {
        int matchCount = 0;
        for (int col = 2; col <= 15; ++col) {
            QString hdr = xlsx.read(1, col).toString().trimmed();
            if (hdr.isEmpty()) break;
            for (const QString& metric : kDetailedAllMetrics) {
                if (hdr.compare(metric, Qt::CaseInsensitive) == 0) {
                    ++matchCount;
                    break;
                }
            }
        }
        if (matchCount >= 3)
            return FileType::DetailedSensory;
    }

    // Sensory (saved format): A1 = "Sample" and row 1 contains sensory metric headers
    if (a1.compare("Sample", Qt::CaseInsensitive) == 0) {
        int matchCount = 0;
        for (int col = 2; col <= 8; ++col) {
            QString hdr = xlsx.read(1, col).toString().trimmed();
            for (const QString& metric : kSensoryMetrics) {
                if (hdr.compare(metric, Qt::CaseInsensitive) == 0) {
                    ++matchCount;
                    break;
                }
            }
        }
        if (matchCount >= 3)
            return FileType::Sensory;
    }

    // Sensory (standardized template): check for sensory metric names in row 3+
    // These templates have multi-tester blocks, not a flat "Sample" header
    // Look for sensory metric names in any cell of rows 1-5
    {
        int sensoryHits = 0;
        for (int r = 1; r <= 5; ++r) {
            for (int c = 1; c <= 20; ++c) {
                QString val = xlsx.read(r, c).toString().trimmed();
                for (const QString& metric : kSensoryMetrics) {
                    if (val.compare(metric, Qt::CaseInsensitive) == 0) {
                        ++sensoryHits;
                        break;
                    }
                }
            }
        }
        if (sensoryHits >= 3)
            return FileType::Sensory;
    }

    // ── TPM detection: check all known template formats ─────────────────────
    // Scan each sheet for TPM signatures.  A file is TPM if any sheet has a
    // header row starting with "puffs" (the universal TPM indicator), or if
    // it contains sheet names matching known TPM test types.
    QStringList sheetNames = xlsx.sheetNames();

    // Check sheet names for known TPM test types
    static const QStringList kTpmSheetIndicators = {
        "Lifetime", "Quick Screening", "Long Puff", "Rapid Puff",
        "Intense", "Extended", "User Test", "Horizontal",
        "Big Headspace", "Temperature Cycling", "Upside Down",
        "Viscosity Compat", "Various Oil", "Negative Pressure",
        "Low Temperature", "Device Life", "Vacuum", "Aerosol",
        "Test Plan"
    };
    for (const QString& sn : sheetNames) {
        for (const QString& ind : kTpmSheetIndicators) {
            if (sn.contains(ind, Qt::CaseInsensitive))
                return FileType::TPM;
        }
    }

    // Check each sheet for TPM data signatures
    for (const QString& sn : sheetNames) {
        xlsx.selectSheet(sn);

        // Format A-E: row 4 col 1 = "puffs"
        QString r4c1 = xlsx.read(4, 1).toString().trimmed();
        if (r4c1.compare("puffs", Qt::CaseInsensitive) == 0)
            return FileType::TPM;

        // Format C/D/E metadata: row 1 has "Date:" or "Sample ID:" labels
        for (int c = 1; c <= 8; ++c) {
            QString val = xlsx.read(1, c).toString().trimmed();
            if (val.compare("Date:", Qt::CaseInsensitive) == 0 ||
                val.compare("Sample ID:", Qt::CaseInsensitive) == 0 ||
                val.compare("Heating Technology:", Qt::CaseInsensitive) == 0)
                return FileType::TPM;
        }

        // Format A/B: row 2 has "Cart #" or "Ri (Ohms)"
        QString r2c1 = xlsx.read(2, 1).toString().trimmed();
        QString r2c3 = xlsx.read(2, 3).toString().trimmed();
        if (r2c1.compare("Cart #", Qt::CaseInsensitive) == 0 ||
            r2c3.compare("Ri (Ohms)", Qt::CaseInsensitive) == 0)
            return FileType::TPM;

        // Format C metadata: row 1 has "Tester:" or "Project:" labels
        for (int c = 1; c <= 8; ++c) {
            QString val = xlsx.read(1, c).toString().trimmed();
            if (val.compare("Tester:", Qt::CaseInsensitive) == 0 ||
                val.compare("Project:", Qt::CaseInsensitive) == 0)
                return FileType::TPM;
        }

        // Format D/E: row 2 has "Resistance" label
        QString r2c3b = xlsx.read(2, 3).toString().trimmed();
        if (r2c3b.contains("Resistance", Qt::CaseInsensitive))
            return FileType::TPM;
    }

    // No format matched — truly unknown
    return FileType::Unknown;
}

void MainWindow::routeFile(const QString& path)
{
    FileType type = detectFileType(path);

    switch (type) {
    case FileType::Sensory:
        // Switch to sensory mode and load the file
        if (!m_sensoryBtn->isChecked())
            m_sensoryBtn->setChecked(true);  // triggers toggleSensoryMode
        if (m_sensoryPanel)
            m_sensoryPanel->loadFile(path);
        break;

    case FileType::DetailedSensory:
        // Switch to detailed sensory mode and load the file
        if (!m_detailedSensoryBtn->isChecked())
            m_detailedSensoryBtn->setChecked(true);
        if (m_detailedSensoryPanel)
            m_detailedSensoryPanel->loadFile(path);
        break;

    case FileType::TPM:
        // Ensure we're in TPM mode, then load
        if (m_sensoryBtn->isChecked())
            m_sensoryBtn->setChecked(false);
        if (m_detailedSensoryBtn->isChecked())
            m_detailedSensoryBtn->setChecked(false);
        loadFile(path);
        break;

    case FileType::Unknown:
        QMessageBox::warning(this, "Unknown File Format",
            "Could not determine the file type for:\n" +
            QFileInfo(path).fileName() +
            "\n\nThe file will be loaded in TPM mode (default).");
        if (m_sensoryBtn->isChecked())
            m_sensoryBtn->setChecked(false);
        if (m_detailedSensoryBtn->isChecked())
            m_detailedSensoryBtn->setChecked(false);
        loadFile(path);
        break;
    }
}

ReportMode MainWindow::currentReportMode() const
{
    if (m_detailedSensoryMode) return ReportMode::DetailedSensory;
    if (m_sensoryMode)         return ReportMode::Sensory;
    return ReportMode::Tpm;
}

void MainWindow::onLoadFile()
{
    QStringList paths = QFileDialog::getOpenFileNames(
        this, "Open File(s)", OutputPaths::resolveDir(currentReportMode(), lastBrowseDir()),
        "Excel / JSON Files (*.xlsx *.xlsm *.xls *.json);;Excel Files (*.xlsx *.xlsm *.xls)"
        ";;JSON Files (*.json);;All Files (*)"
    );
    if (paths.isEmpty()) return;
    setLastBrowseDir(paths.first());
    // Queue remaining files — they'll be routed individually when each finishes
    m_pendingLoadPaths = paths.mid(1);
    routeFile(paths.first());
}


void MainWindow::onCloseFile()
{
    if (m_currentFileIndex < 0 || m_currentFileIndex >= m_loadedFiles.size()) return;

    // H2: drain any debounced Excel writes BEFORE m_loadedFiles is mutated
    // so the file being closed picks up the last burst of edits. The
    // 500 ms m_excelWriteTimer may not have fired yet when the user
    // clicks Close.
    flushExcelWrites();

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
            if (m_db && m_db->saveFile(m_loadedFiles[m_currentFileIndex])) {
                m_modifiedFilePaths.remove(fp);
            }
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
    // Plan C: a file left the working set; mark the snapshot dirty so the
    // recovery index prunes it on the next flush.
    if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();

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
        setStatusFile("File closed", FileStatusClosed);
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
        qWarning() << "Load error:" << m_processor->lastError();
        showError("Load Error", "Failed to load file. Check the log for details.");
        updateStatusBar("Load failed.");
        // Continue loading remaining queued files even if one fails
        if (!m_pendingLoadPaths.isEmpty()) {
            QString next = m_pendingLoadPaths.takeFirst();
            routeFile(next);
        }
        return;
    }

    // Inherit id+version from any existing record (in-memory or DB) so the
    // save below becomes an UPDATE instead of an INSERT. Without this, any
    // file already in the database (migrated data, or files that were saved
    // in a previous session) hits the files.file_path UNIQUE constraint on
    // INSERT and pops UniqueViolationDialog every time the user clicks Open.
    for (const FileResult& mem : m_loadedFiles) {
        if (mem.filePath == result.filePath && mem.id > 0) {
            result.id = mem.id;
            result.version = mem.version;
            break;
        }
    }
    if (result.id <= 0 && m_db && m_db->isOnline()) {
        const FileResult dbRow = m_db->loadFileByPath(result.filePath);
        if (dbRow.id > 0) {
            result.id = dbRow.id;
            result.version = dbRow.version;
        }
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
            if (m_db) m_db->saveFile(m_loadedFiles[i]);
            m_modifiedFilePaths.remove(result.filePath);
            updateDbSyncIndicator();
            // Plan C: the open set changed (a file was reloaded in place); keep
            // the recovery index in sync.
            if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
            updateStatusBar("Refreshed: " + result.fileName);
            // Re-paint the file tree once the DB save has stamped FileResult.id
            // so presence dots can be associated with the right row.
            populateFileTree();
            // Continue loading remaining queued files
            if (!m_pendingLoadPaths.isEmpty()) {
                QString next = m_pendingLoadPaths.takeFirst();
                routeFile(next);
            }
            return;
        }
    }

    m_loadedFiles.append(result);
    m_currentFileIndex   = m_loadedFiles.size() - 1;
    m_currentSheetIndex  = 0;
    m_currentSampleIndex = 0;

    if (m_db) m_db->saveFile(m_loadedFiles[m_currentFileIndex]);
    m_modifiedFilePaths.remove(result.filePath);
    updateDbSyncIndicator();
    // Plan C: a newly-opened file joined the working set; mark the snapshot
    // dirty so the recovery index tracks the open set.
    if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
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

    // ── Continue loading queued files ────────────────────────────────────────
    if (!m_pendingLoadPaths.isEmpty()) {
        QString next = m_pendingLoadPaths.takeFirst();
        routeFile(next);
    }
}

// ─── Navigation ───────────────────────────────────────────────────────────────
void MainWindow::onFileSelected(int index)
{
    if (index < 0 || index >= m_loadedFiles.size()) return;
    m_currentFileIndex   = index;
    m_currentSheetIndex  = 0;
    m_currentSampleIndex = 0;
    populateSheetCombo();

    // Activate presence for this file so other clients see us viewing it
    // and the local avatar bar updates with anyone else already here.
    if (m_presence) {
        const qint64 fileId = qint64(m_loadedFiles[index].id);
        if (fileId > 0) {
            m_presence->activate(QStringLiteral("file"), fileId,
                                 QStringLiteral("viewing"));
            m_currentResourceType = QStringLiteral("file");
            m_currentResourceId   = fileId;
            refreshPresenceFor(m_currentResourceType, m_currentResourceId);
        }
    }
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
        // Persist the DB row id on the tree item so refreshPresenceFor()
        // can locate the right item without re-scanning m_loadedFiles by
        // name. -1 means "not yet persisted" → no presence to show.
        fi->setData(0, Qt::UserRole, qlonglong(f.id));
        for (const auto& sheet : f.sheets) {
            auto* si = new QTreeWidgetItem(fi, {sheet.sheetName});
            si->setIcon(0, sheetIcon);
        }
        fi->setExpanded(true);
    }
    if (m_currentFileIndex >= 0 && m_currentFileIndex < m_fileCombo->count())
        m_fileCombo->setCurrentIndex(m_currentFileIndex);
    m_fileCombo->blockSignals(false);

    refreshAllPresence();
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

// ─── Presence UI (Plan B Phase 5) ────────────────────────────────────────────
void MainWindow::refreshPresenceFor(const QString& resourceType, qint64 resourceId)
{
    if (!m_presence) return;
    if (resourceId < 0) return;

    const auto rows = m_presence->activeFor(resourceType, resourceId);

    QStringList colors;
    QStringList intents;
    QStringList tooltipParts;
    colors.reserve(rows.size());
    intents.reserve(rows.size());
    tooltipParts.reserve(rows.size());
    for (const auto& r : rows) {
        colors << r.userColor;
        intents << r.intent;
        tooltipParts << QStringLiteral("%1 (%2)").arg(r.userName, r.intent);
    }
    const QString tooltip = tooltipParts.join(QLatin1Char('\n'));

    // Walk the right nav widget. We attach colors+intents+tooltip to the
    // matching item; the PresenceDotsDelegate reads the roles at paint time.
    //
    // Each setData below CHANGES item state, which fires {QListWidget,
    // QTreeWidget}::itemChanged. The sensory itemChanged handler treats
    // that as a user rename → emits sessionsChanged → refreshSensoryNavigator
    // → clears m_sensoryNav (deleting the very item* we're still iterating).
    // The next setData / setToolTip would then crash inside
    // QListWidgetItem::setText with an access violation on a freed item.
    //
    // Block signals on the affected widget while we update item roles —
    // these are programmatic changes, not user edits.
    if (resourceType == QLatin1String("file") && m_fileTree) {
        QSignalBlocker blocker(m_fileTree);
        for (int i = 0; i < m_fileTree->topLevelItemCount(); ++i) {
            QTreeWidgetItem* fi = m_fileTree->topLevelItem(i);
            if (!fi) continue;
            if (fi->data(0, Qt::UserRole).toLongLong() != resourceId) continue;
            fi->setData(0, DVE::PresenceDotsDelegate::kColorsRole,  colors);
            fi->setData(0, DVE::PresenceDotsDelegate::kIntentsRole, intents);
            fi->setToolTip(0, tooltip);
            break;
        }
        // #4: force immediate repaint -- setData() alone doesn't notify the
        // view, so dots wouldn't refresh until the next user-driven layout pass.
        m_fileTree->viewport()->update();
    } else if (resourceType == QLatin1String("sensory_session") && m_sensoryNav) {
        QSignalBlocker blocker(m_sensoryNav);
        for (int i = 0; i < m_sensoryNav->count(); ++i) {
            QListWidgetItem* it = m_sensoryNav->item(i);
            if (!it) continue;
            if (it->data(Qt::UserRole).toLongLong() != resourceId) continue;
            it->setData(DVE::PresenceDotsDelegate::kColorsRole,  colors);
            it->setData(DVE::PresenceDotsDelegate::kIntentsRole, intents);
            it->setToolTip(tooltip);
            break;
        }
        m_sensoryNav->viewport()->update();
    } else if (resourceType == QLatin1String("detailed_sensory_session") &&
               m_detailedSensoryNav) {
        QSignalBlocker blocker(m_detailedSensoryNav);
        for (int i = 0; i < m_detailedSensoryNav->count(); ++i) {
            QListWidgetItem* it = m_detailedSensoryNav->item(i);
            if (!it) continue;
            if (it->data(Qt::UserRole).toLongLong() != resourceId) continue;
            it->setData(DVE::PresenceDotsDelegate::kColorsRole,  colors);
            it->setData(DVE::PresenceDotsDelegate::kIntentsRole, intents);
            it->setToolTip(tooltip);
            break;
        }
        m_detailedSensoryNav->viewport()->update();
    }

    // If this is the resource the user is currently focused on, refresh
    // the avatar bar too. clear() handles the empty-rows case (hides the bar).
    if (resourceType == m_currentResourceType && resourceId == m_currentResourceId
        && m_avatarBar) {
        if (rows.isEmpty()) m_avatarBar->clear();
        else m_avatarBar->setPresence(rows,
                                      m_identity ? m_identity->uuid() : QUuid());
    }
}

void MainWindow::refreshAllPresence()
{
    if (!m_presence) return;

    if (m_fileTree) {
        for (int i = 0; i < m_fileTree->topLevelItemCount(); ++i) {
            QTreeWidgetItem* fi = m_fileTree->topLevelItem(i);
            if (!fi) continue;
            const qint64 id = fi->data(0, Qt::UserRole).toLongLong();
            if (id > 0) refreshPresenceFor(QStringLiteral("file"), id);
        }
    }
    if (m_sensoryNav) {
        for (int i = 0; i < m_sensoryNav->count(); ++i) {
            QListWidgetItem* it = m_sensoryNav->item(i);
            if (!it) continue;
            const qint64 id = it->data(Qt::UserRole).toLongLong();
            if (id > 0) refreshPresenceFor(QStringLiteral("sensory_session"), id);
        }
    }
    if (m_detailedSensoryNav) {
        for (int i = 0; i < m_detailedSensoryNav->count(); ++i) {
            QListWidgetItem* it = m_detailedSensoryNav->item(i);
            if (!it) continue;
            const qint64 id = it->data(Qt::UserRole).toLongLong();
            if (id > 0)
                refreshPresenceFor(QStringLiteral("detailed_sensory_session"), id);
        }
    }
}

void MainWindow::clearActivePresence()
{
    if (m_presence && !m_currentResourceType.isEmpty()) {
        m_presence->deactivate();
    }
    m_currentResourceType.clear();
    m_currentResourceId = -1;
    if (m_avatarBar) m_avatarBar->clear();
    if (m_rowDeletedBanner) m_rowDeletedBanner->dismiss();
}

// ── Phase 6 — don't-yank-in-progress edits ──────────────────────────────────
void MainWindow::onDataTableItemChanged(QTableWidgetItem* it)
{
    if (!it) return;
    // v2.0.2 H6: skip programmatic writes done by the remote-apply path.
    // onRemoteCellChanged blocks signals on m_dataTable while it writes,
    // but other paths (e.g. cell-clicked accept, populate refresh)
    // synthesize text changes that look like user edits. The flag is
    // toggled true only inside applyRemoteValueToCell-using callers.
    if (m_applyingRemote) return;
    // Baseline is stamped during populate; missing means this cell is
    // outside the editable per-row data (raw-table sheet) — skip silently.
    const QVariant baseline = it->data(Qt::UserRole + 2);
    if (!baseline.isValid()) return;
    const bool dirty = (it->text() != baseline.toString());
    it->setData(Qt::UserRole + 3, dirty);
    // If the user typed past a pending remote-change decoration without
    // accepting it, drop the yellow tag — they're deliberately keeping
    // their own value over the remote one.
    if (it->data(Qt::UserRole + 4).isValid()) {
        clearRemoteDecoration(it, /*updateBaseline=*/false);
    }

    // v2.0.1 live sync: push the same edit to the DB immediately so every
    // other client sees it without waiting for a Save. data_rows.id lives
    // on the vertical header (same source findTableRowForDataRowId reads).
    if (m_liveSync) {
        const QString column = liveColumnForDataCol(it->column());
        if (!column.isEmpty()) {
            const QTableWidgetItem* vh =
                m_dataTable->verticalHeaderItem(it->row());
            const qint64 rowId = vh ? vh->data(Qt::UserRole).toLongLong() : -1;
            if (rowId > 0) {
                m_liveSync->commitCell(QStringLiteral("data_rows"),
                                       rowId, column, it->text());
            }
        }
    }
}

void MainWindow::onDataTableItemClicked(QTableWidgetItem* it)
{
    if (!it) return;
    const QVariant pending = it->data(Qt::UserRole + 4);
    if (!pending.isValid()) return;  // no remote decoration to accept
    // Accept the remote value: write the text and treat it as the new
    // baseline so the cell is clean afterward.
    m_dataTable->blockSignals(true);
    it->setText(pending.toString());
    m_dataTable->blockSignals(false);
    clearRemoteDecoration(it, /*updateBaseline=*/true);
    // I2 fix: signals were blocked during setText above, so
    // onTableCellChanged didn't fire and the underlying DataRow in
    // m_loadedFiles still holds the user's stale local value. Flush
    // explicitly so the next save persists the accepted remote value
    // instead of clobbering it.
    onTableCellChanged(it->row(), it->column());
    updateStatusBar(tr("Accepted remote value."));
}

void MainWindow::clearRemoteDecoration(QTableWidgetItem* it,
                                       bool updateBaseline)
{
    if (!it) return;
    m_dataTable->blockSignals(true);
    if (updateBaseline) {
        it->setData(Qt::UserRole + 2, it->text());
        it->setData(Qt::UserRole + 3, false);
    }
    it->setData(Qt::UserRole + 4, QVariant());
    it->setToolTip(QString());
    // Restore the original background — calculated cells stayed white
    // through the decoration, editable cells likewise. Excluded-row
    // styling is reapplied by displayCurrentSample on the next repaint.
    it->setBackground(QColor(Qt::white));
    m_dataTable->blockSignals(false);
}

int MainWindow::findTableRowForDataRowId(qint64 dataRowId) const
{
    if (!m_dataTable || dataRowId < 0) return -1;
    for (int r = 0; r < m_dataTable->rowCount(); ++r) {
        const QTableWidgetItem* h = m_dataTable->verticalHeaderItem(r);
        if (!h) continue;
        if (h->data(Qt::UserRole).toLongLong() == dataRowId) return r;
    }
    return -1;
}

QString MainWindow::resolveUserName(const QString& uuid) const
{
    if (uuid.isEmpty()) return tr("another user");
    if (!m_presence) {
        // No presence info — fall back to the first 8 chars of the UUID.
        return uuid.left(8);
    }
    // Sweep all known active resources for a matching UUID. PresenceManager
    // doesn't expose a UUID-keyed cache; the volume is tiny (a handful of
    // active users at peak) so the linear scan is fine.
    auto match = [&](const QVector<PresenceRow>& rows) -> QString {
        for (const PresenceRow& r : rows) {
            if (r.userUuid.toString(QUuid::WithoutBraces) == uuid)
                return r.userName;
        }
        return QString();
    };
    QString name;
    if (m_fileTree) {
        for (int i = 0; i < m_fileTree->topLevelItemCount() && name.isEmpty(); ++i) {
            const qint64 id = m_fileTree->topLevelItem(i)->data(0, Qt::UserRole).toLongLong();
            if (id > 0)
                name = match(m_presence->activeFor(QStringLiteral("file"), id));
        }
    }
    if (name.isEmpty() && m_sensoryNav) {
        for (int i = 0; i < m_sensoryNav->count() && name.isEmpty(); ++i) {
            const qint64 id = m_sensoryNav->item(i)->data(Qt::UserRole).toLongLong();
            if (id > 0)
                name = match(m_presence->activeFor(
                    QStringLiteral("sensory_session"), id));
        }
    }
    if (name.isEmpty() && m_detailedSensoryNav) {
        for (int i = 0; i < m_detailedSensoryNav->count() && name.isEmpty(); ++i) {
            const qint64 id = m_detailedSensoryNav->item(i)->data(Qt::UserRole).toLongLong();
            if (id > 0)
                name = match(m_presence->activeFor(
                    QStringLiteral("detailed_sensory_session"), id));
        }
    }
    return name.isEmpty() ? uuid.left(8) : name;
}

void MainWindow::handleRemoteRowChange(const DVE::RowChange& c)
{
    // ── T19: row-deleted toast ────────────────────────────────────────────
    // A DELETE on the currently-open resource (file / sensory session /
    // detailed sensory session) drops the banner. Other DELETEs are
    // ignored — the user might have the resource in a list but not open.
    if (c.op == QLatin1String("DELETE") && m_rowDeletedBanner) {
        const bool isOpenFile  = c.table == QLatin1String("files")
                              && m_currentResourceType == QLatin1String("file")
                              && c.id == m_currentResourceId;
        const bool isOpenSens  = c.table == QLatin1String("sensory_sessions")
                              && m_currentResourceType
                                  == QLatin1String("sensory_session")
                              && c.id == m_currentResourceId;
        const bool isOpenDSens = c.table
                                  == QLatin1String("detailed_sensory_sessions")
                              && m_currentResourceType
                                  == QLatin1String("detailed_sensory_session")
                              && c.id == m_currentResourceId;
        if (isOpenFile || isOpenSens || isOpenDSens) {
            QString label;
            if (isOpenFile)        label = tr("file");
            else if (isOpenSens)   label = tr("sensory session");
            else                   label = tr("detailed sensory session");
            m_rowDeletedBanner->showFor(label,
                                        resolveUserName(c.updatedBy));
        }
        return;
    }

    // ── T18: yellow decoration on a remotely-edited dirty cell ────────────
    if (c.table != QLatin1String("data_rows")) return;
    if (c.op   == QLatin1String("DELETE"))     return;
    if (!m_dataTable) return;
    const int tableRow = findTableRowForDataRowId(c.id);
    if (tableRow < 0) return;  // row not on screen — nothing to do

    // Best-effort: pull the new row from Postgres so we can show the
    // pending value in the tooltip. If the read fails, drop a generic
    // tooltip and skip the per-column pending-value stamp.
    QVector<QString> remoteVals;
    bool haveRemote = false;
    if (m_pgConn && m_pgConn->isOpen()) {
        QSqlDatabase& db = m_pgConn->queryDb();
        QSqlQuery q(db);
        q.prepare("SELECT puffs, before_weight, after_weight, draw_pressure, "
                  "resistance, smell, clog, notes, tpm, tpm_power_density, "
                  "variation_tpm, oil_consumed, puffing_regime "
                  "FROM data_rows WHERE id = ?");
        q.addBindValue(qlonglong(c.id));
        if (q.exec() && q.next()) {
            remoteVals.reserve(12);
            for (int i = 0; i < 12; ++i) remoteVals << q.value(i).toString();
            // Column index 4 is dual-purpose: on a new-template sheet it holds
            // the per-row Puffing Regime, not Resistance. Stamp the regime so
            // "take their value" (onDataTableItemClicked) applies the right
            // field instead of silently overwriting the regime with resistance.
            const SheetResult* sh = currentSheet();
            if (sh && sh->hasPerRowRegime && remoteVals.size() > DVE::Cols::RESISTANCE)
                remoteVals[DVE::Cols::RESISTANCE] = q.value(12).toString();
            haveRemote = true;
        }
    }

    const QString who = resolveUserName(c.updatedBy);
    const QString tooltip = haveRemote
        ? tr("Changed by %1 — click to take their value").arg(who)
        : tr("Changed by %1 — reload to see the new value").arg(who);

    // Only decorate cells the user is actively editing (dirty). Calculated
    // cells (TPM, power density, variation, oil consumed) aren't directly
    // editable so they're skipped — their values follow from the edited
    // primary measurements anyway.
    bool decoratedAny = false;
    m_dataTable->blockSignals(true);
    for (int col = 0; col < m_dataTable->columnCount(); ++col) {
        QTableWidgetItem* it = m_dataTable->item(tableRow, col);
        if (!it) continue;
        if (!it->data(Qt::UserRole + 3).toBool()) continue;  // not dirty
        decoratedAny = true;
        it->setBackground(QColor("#fef3c7"));
        it->setToolTip(tooltip);
        if (haveRemote && col < remoteVals.size())
            it->setData(Qt::UserRole + 4, remoteVals[col]);
        else
            it->setData(Qt::UserRole + 4, QVariant());  // generic tooltip only
    }
    m_dataTable->blockSignals(false);
    if (decoratedAny)
        updateStatusBar(tr("Remote change on row you're editing — click a "
                           "yellow cell to accept."));
}

// ── v2.0.1 LiveSync inbound handlers ────────────────────────────────────────
// data_rows-only here; sensory tables register their own handlers in Task 8.

void MainWindow::onRemoteCellChanged(const QString& table, qint64 rowId,
                                     const QString& column,
                                     const QVariant& newValue)
{
    if (table != QLatin1String("data_rows")) return;
    const int row = findTableRowForDataRowId(rowId);
    if (row < 0) return;
    const int col = dataColForLiveColumn(column);
    if (col < 0) return;

    // C6: applyRemoteValueToCell skips dirty cells — the yellow
    // decoration painted by handleRemoteRowChange already surfaces the
    // remote-conflict; onDataTableItemClicked accepts the value if the
    // user wants it. H6: gate onDataTableItemChanged on m_applyingRemote
    // while we write so it doesn't echo the remote value back through
    // LiveSync as a fresh local edit.
    m_applyingRemote = true;
    const bool applied = DVE::applyRemoteValueToCell(
        m_dataTable, row, col, newValue.toString());
    m_applyingRemote = false;
    if (!applied) return;

    QTableWidgetItem* it = m_dataTable->item(row, col);
    if (!it) return;

    // Flash for 200 ms.
    it->setData(DVE::CellFocusDelegate::kFlashRole, true);
    m_dataTable->viewport()->update();
    QTimer::singleShot(200, this, [this, rowId, column]() {
        if (!m_dataTable) return;
        const int r = findTableRowForDataRowId(rowId);
        if (r < 0) return;
        const int c = dataColForLiveColumn(column);
        if (c < 0) return;
        QTableWidgetItem* it2 = m_dataTable->item(r, c);
        if (!it2) return;
        it2->setData(DVE::CellFocusDelegate::kFlashRole, QVariant());
        m_dataTable->viewport()->update();
    });
}

void MainWindow::onRemoteCellFocused(const QString& table, qint64 rowId,
                                     const QString& column,
                                     const QString& userName,
                                     const QString& userColor)
{
    if (table != QLatin1String("data_rows")) return;
    const int row = findTableRowForDataRowId(rowId);
    if (row < 0) return;
    const int col = dataColForLiveColumn(column);
    if (col < 0) return;
    QTableWidgetItem* it = m_dataTable->item(row, col);
    if (!it) return;

    QSignalBlocker b(m_dataTable);
    it->setData(DVE::CellFocusDelegate::kFocusColorRole, userColor);
    it->setData(DVE::CellFocusDelegate::kFocusNameRole,  userName);
    m_dataTable->viewport()->update();
}

void MainWindow::onRemoteCellBlurred(const QString& table, qint64 rowId,
                                     const QString& column)
{
    if (table != QLatin1String("data_rows")) return;
    const int row = findTableRowForDataRowId(rowId);
    if (row < 0) return;
    const int col = dataColForLiveColumn(column);
    if (col < 0) return;
    QTableWidgetItem* it = m_dataTable->item(row, col);
    if (!it) return;

    QSignalBlocker b(m_dataTable);
    it->setData(DVE::CellFocusDelegate::kFocusColorRole, QVariant());
    it->setData(DVE::CellFocusDelegate::kFocusNameRole,  QVariant());
    m_dataTable->viewport()->update();
}

void MainWindow::displayCurrentSample()
{
    // Update the incomplete-data banner whenever the displayed content changes.
    // This must run before the early-return paths so the banner stays accurate
    // even when switching to raw/SOP sheets or empty files.
    updateIncompleteDataBanner();

    const SheetResult* sheet = currentSheet();

    // ── Raw table (SOP / instruction sheets) ──────────────────────────────────
    if (sheet && sheet->isRawTable) {
        // Raw/SOP sheets have no per-row regime column — deactivate the combo
        // delegate so a previously-viewed new-template sheet doesn't leave a
        // regime dropdown on column 4 of this raw table.
        if (m_regimeDelegate) m_regimeDelegate->setActive(false);
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
        setStatusBreadcrumb({ f ? f->fileName : QString(), sheet->sheetName });
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

    // Relabel column 4 header based on whether this sheet uses per-row regime.
    const bool perRowRegime = sheet && sheet->hasPerRowRegime;
    m_dataTable->setHorizontalHeaderLabels(dataTableHeaders(perRowRegime));

    // Keep the regime-combo delegate in sync with the current sheet mode.
    if (m_regimeDelegate) {
        m_regimeDelegate->setActive(perRowRegime);
        m_regimeDelegate->setRegimes(currentFileRegimes());
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
        if (perRowRegime) setStr(dr.puffingRegime);
        else (dr.resistance == 0.0) ? setEmpty() : setNum(dr.resistance, 3);
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

        // ── Phase 6 (T17) — baseline stamp + data_rows id ────────────────────
        // Every editable cell stamps its current text as the baseline at
        // UserRole+2 so the itemChanged handler can detect divergence on
        // edit. Signals are still blocked here, so no spurious dirty events
        // fire during populate.
        for (int c = 0; c < m_dataTable->columnCount(); ++c) {
            auto* itm = m_dataTable->item(tRow, c);
            if (!itm) continue;
            itm->setData(Qt::UserRole + 2, itm->text());
            itm->setData(Qt::UserRole + 3, false);  // clean
            itm->setData(Qt::UserRole + 4, QVariant());  // no pending remote
        }
        // Stash the data_rows.id on the vertical header item so a later
        // NOTIFY for data_rows can map back to this visible row. -1 means
        // "not yet persisted"; the lookup gracefully skips those.
        {
            QTableWidgetItem* header = m_dataTable->verticalHeaderItem(tRow);
            if (!header) {
                header = new QTableWidgetItem();
                m_dataTable->setVerticalHeaderItem(tRow, header);
            }
            header->setData(Qt::UserRole, dr.id);
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
    // File-scoped per spec: the picker lists every regime in the file (across
    // all sheets) so a selection persists as the user flips between sheets.
    // (This re-renders the plot a second time after setSheetData — cheap, and
    // kept separate because regimes are file-scoped, not sheet-scoped.)
    refreshPlotRegimes();

    updateSampleNav();
    updateImageButton();
    updateCleanupButtons();

    // ── Breadcrumb ────────────────────────────────────────────────────────────
    const auto* f = currentFile();
    setStatusBreadcrumb({
        f ? f->fileName : QString(),
        sheet->sheetName,
        QString("Sample %1 / %2").arg(m_currentSampleIndex + 1)
                                  .arg(sheet->samples.size())
    });
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
        it->setForeground(QColor(0x1F, 0x4E, 0x79));
        QFont f = it->font(); f.setBold(true); f.setPointSize(8); it->setFont(f);
        m_propTable->setItem(row, 0, it);
        QTableWidgetItem* it2 = new QTableWidgetItem();
        it2->setFlags(Qt::ItemIsEnabled);
        m_propTable->setItem(row, 1, it2);
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

    QString path = QFileDialog::getSaveFileName(
        this, "Save Test Report",
        OutputPaths::resolveDir(ReportMode::Tpm, m_lastBrowseDir) + "/" +
            OutputPaths::reportFileName(QFileInfo(file->filePath).completeBaseName(), sheet->sheetName),
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

QString MainWindow::uniqueFilename(const QString& desiredPath)
{
    if (!QFile::exists(desiredPath)) return desiredPath;
    QFileInfo fi(desiredPath);
    QString stem = fi.completeBaseName();
    QString ext  = fi.suffix();
    QString dir  = fi.absolutePath();
    for (int i = 2; i < 1000; ++i) {
        QString candidate = QString("%1/%2 (%3).%4").arg(dir, stem).arg(i).arg(ext);
        if (!QFile::exists(candidate)) return candidate;
    }
    return desiredPath;
}

void MainWindow::onGenerateFullReport()
{
    if (m_loadedFiles.isEmpty()) {
        QMessageBox::warning(this, "No Data",
                             "Load at least one Excel file before generating a Full Report.");
        return;
    }

    QDialog picker(this);
    picker.setWindowTitle("Select Files for Full Report");
    picker.setMinimumSize(500, 400);
    picker.resize(550, 450);

    auto* layout = new QVBoxLayout(&picker);
    layout->addWidget(new QLabel("Select files to include (Ctrl+Click or Shift+Click):"));

    auto* list = new QListWidget;
    list->setSelectionMode(QAbstractItemView::ExtendedSelection);
    for (int i = 0; i < m_loadedFiles.size(); ++i) {
        auto* item = new QListWidgetItem(m_loadedFiles[i].fileName);
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

    QVector<FileResult> files;
    for (QListWidgetItem* item : list->selectedItems()) {
        int idx = item->data(Qt::UserRole).toInt();
        if (idx >= 0 && idx < m_loadedFiles.size())
            files.append(buildCleanedFile(m_loadedFiles[idx]));
    }

    if (files.isEmpty()) {
        QMessageBox::warning(this, "No Selection", "No files selected.");
        return;
    }

    QStringList reportFailures;

    if (files.size() == 1) {
        const QString def = OutputPaths::resolveDir(ReportMode::Tpm, m_lastBrowseDir) + "/" +
            OutputPaths::reportFileName(QFileInfo(files.first().filePath).completeBaseName());
        const QString path = QFileDialog::getSaveFileName(this, "Save Full Report", def, "PowerPoint (*.pptx)");
        if (path.isEmpty()) return;
        setLastBrowseDir(path);

        ReportConfig cfg;
        cfg.outputPath = path;
        m_reportGen->setResourcePath(resourcePath());
        m_reportGen->generateFullReport(files.first(), cfg);
        return;
    }

    const QString outDir = QFileDialog::getExistingDirectory(
        this, "Select Output Folder",
        OutputPaths::resolveDir(ReportMode::Tpm, m_lastBrowseDir));
    if (outDir.isEmpty()) return;
    setLastBrowseDir(outDir);

    m_reportGen->setResourcePath(resourcePath());

    int succeeded = 0;
    const int total = files.size() + 1;

    for (int i = 0; i < files.size(); ++i) {
        const FileResult& f = files[i];
        QString outPath = outDir + "/" + OutputPaths::reportFileName(QFileInfo(f.filePath).completeBaseName());
        outPath = uniqueFilename(outPath);
        ReportConfig cfg;
        cfg.outputPath = outPath;
        setProgress((100 * i) / total, QString("Generating %1 of %2").arg(i+1).arg(total));
        bool ok = m_reportGen->generateFullReport(f, cfg);
        if (!ok) reportFailures << f.fileName;
        else     ++succeeded;
    }

    {
        QString combinedPath = outDir + "/Combined_" +
            QDate::currentDate().toString("yyyy-MM-dd") + "_report.pptx";
        combinedPath = uniqueFilename(combinedPath);
        setProgress((100 * files.size()) / total, "Generating combined report");
        ReportConfig cfg;
        cfg.outputPath = combinedPath;
        bool ok = m_reportGen->generateCombinedFullReport(files, cfg, combinedPath);
        if (!ok) reportFailures << QFileInfo(combinedPath).fileName();
        else     ++succeeded;
    }

    setProgress(100, QString());

    QMessageBox box(this);
    box.setWindowTitle("Reports Generated");
    QString text = QString("Generated %1 of %2 reports.").arg(succeeded).arg(total);
    if (!reportFailures.isEmpty())
        text += "\n\nFailed:\n  " + reportFailures.join("\n  ");
    box.setText(text);
    auto* openBtn = box.addButton("Open Folder", QMessageBox::ActionRole);
    box.addButton(QMessageBox::Ok);
    box.exec();
    if (box.clickedButton() == openBtn)
        QDesktopServices::openUrl(QUrl::fromLocalFile(outDir));
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
        qWarning() << "Report generation error:" << m_reportGen->lastError();
        showError("Report Failed", "Could not generate report. Check the log for details.");
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
    // Mode change ends focus on the prior resource. The new mode's first
    // selection will re-activate; until then, avatar bar is empty and no
    // stray NOTIFY can refresh a stale resource.
    clearActivePresence();

    if (checked) {
        // Uncheck detailed sensory mode if active
        if (m_detailedSensoryMode && m_detailedSensoryBtn) {
            m_detailedSensoryBtn->blockSignals(true);
            m_detailedSensoryBtn->setChecked(false);
            m_detailedSensoryBtn->blockSignals(false);
            m_detailedSensoryMode = false;
        }

        if (!m_sensoryPanel) {
            initSensoryPanel();
        }
        m_centralStack->setCurrentWidget(m_sensoryPanel);
        m_navStack->setCurrentWidget(m_sensoryNav);
        m_navLabel->setText("Sessions:  <span style='color:gray; font-size:11px;'>select multiple to show average sensory score</span>");
        refreshSensoryNavigator();
        if (m_testAvgPanel) m_testAvgPanel->setVisible(true);
        refreshSensoryAverages();
        updateSensoryProperties();
    } else {
        m_centralStack->setCurrentWidget(m_centralSplitter);
        m_navStack->setCurrentWidget(m_fileTree);
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

void MainWindow::toggleDetailedSensoryMode(bool checked)
{
    m_detailedSensoryMode = checked;
    clearActivePresence();

    if (checked) {
        // Uncheck sensory mode if active
        if (m_sensoryMode) {
            m_sensoryBtn->blockSignals(true);
            m_sensoryBtn->setChecked(false);
            m_sensoryBtn->blockSignals(false);
            m_sensoryMode = false;
        }

        if (!m_detailedSensoryPanel) {
            initDetailedSensoryPanel();
        }
        m_centralStack->setCurrentWidget(m_detailedSensoryPanel);
        m_navStack->setCurrentWidget(m_detailedSensoryNav);
        m_navLabel->setText("Sessions:  <span style='color:gray; font-size:11px;'>select multiple to show average score</span>");
        refreshDetailedSensoryNavigator();
        if (m_testAvgPanel) m_testAvgPanel->setVisible(true);
        updateDetailedSensoryProperties();
    } else {
        m_centralStack->setCurrentWidget(m_centralSplitter);
        m_navStack->setCurrentWidget(m_fileTree);
        m_navLabel->setText("Loaded Files:");
        if (m_testAvgPanel) m_testAvgPanel->setVisible(false);
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
    m_sensoryPanel->setLiveSync(m_liveSync);                // nullptr-safe
    m_centralStack->addWidget(m_sensoryPanel);   // index 1

    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::refreshSensoryNavigator);
    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::refreshSensoryAverages);
    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::updateImageButton);
    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, &MainWindow::updateSensoryProperties);
    connect(m_sensoryPanel, &SensoryPanel::sessionsChanged,
            this, [this]() {
        m_sensorySessionsDirty = true;
        // Plan C: mark the recovery snapshot dirty on any sensory-session change.
        if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
        updateDbSyncIndicator();
        // Sensory persistence is handled by LiveSync per-cell; the
        // session-level fallback in onUpdateDatabase is reserved for
        // Ctrl+U + the on-close prompt. Kicking the 5 s timer here
        // produced a UI freeze on slower LANs because the full save
        // loop ran synchronously on the UI thread.
    });

    // Plan C (C6 fix): sessionsChanged fires only on STRUCTURAL ops
    // (new/close/rename/add/removeSample/load). Per-field value edits — score
    // sliders, comments, sample names, header fields — emit dataEdited()
    // instead and would otherwise never reach the crash snapshot. Wire it to
    // the same noteDirty() so routine data entry is captured. (We deliberately
    // do NOT also run the structural consumers here — those are pure repaint
    // overhead on every keystroke; the snapshot only needs noteDirty().)
    connect(m_sensoryPanel, &SensoryPanel::dataEdited, this, [this]() {
        m_sensorySessionsDirty = true;
        if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
    });
}

void MainWindow::initDetailedSensoryPanel()
{
    m_detailedSensoryPanel = new DetailedSensoryPanel(m_db, this);
    m_detailedSensoryPanel->setLiveSync(m_liveSync);                // nullptr-safe
    m_centralStack->addWidget(m_detailedSensoryPanel);

    // Add navigator list for detailed sensory sessions
    m_detailedSensoryNav = new QListWidget(this);
    m_detailedSensoryNav->setSelectionMode(QAbstractItemView::ExtendedSelection);
    m_navStack->addWidget(m_detailedSensoryNav);  // index 2
    if (m_presenceDelegate)
        m_detailedSensoryNav->setItemDelegate(m_presenceDelegate);

    connect(m_detailedSensoryNav, &QListWidget::currentRowChanged, this, [this](int row) {
        if (m_detailedSensoryPanel && row >= 0) {
            m_detailedSensoryPanel->selectSession(row);
            updateDetailedSensoryProperties();

            // Phase 5: presence activate on detailed-sensory session open.
            if (m_presence) {
                QListWidgetItem* it = m_detailedSensoryNav->item(row);
                const qint64 sessId =
                    it ? it->data(Qt::UserRole).toLongLong() : -1;
                if (sessId > 0) {
                    m_presence->activate(
                        QStringLiteral("detailed_sensory_session"),
                        sessId, QStringLiteral("viewing"));
                    m_currentResourceType =
                        QStringLiteral("detailed_sensory_session");
                    m_currentResourceId = sessId;
                    refreshPresenceFor(m_currentResourceType,
                                       m_currentResourceId);
                }
            }
        }
    });

    connect(m_detailedSensoryNav, &QListWidget::itemSelectionChanged, this, [this]() {
        auto items = m_detailedSensoryNav->selectedItems();
        if (items.size() > 1) {
            QVector<int> indices;
            for (auto* item : items)
                indices.append(m_detailedSensoryNav->row(item));
            m_detailedSensoryPanel->showAveragedChart(indices);
        }
    });

    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, &MainWindow::refreshDetailedSensoryNavigator);
    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, &MainWindow::updateDetailedSensoryProperties);
    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, &MainWindow::updateImageButton);
    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::sessionsChanged,
            this, [this]() {
        m_detailedSensorySessionsDirty = true;
        // Plan C: mark the recovery snapshot dirty on any detailed-session change.
        if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
        updateDbSyncIndicator();
    });

    // Plan C (C6 fix): same per-field gap as SensoryPanel — sessionsChanged
    // misses score/combo/comment/name and session-field (header + oil-smell/
    // clog/mouthpiece) value edits. dataEdited() fires on those; route it to
    // noteDirty() so detailed-sensory data entry reaches the crash snapshot.
    connect(m_detailedSensoryPanel, &DetailedSensoryPanel::dataEdited, this, [this]() {
        m_detailedSensorySessionsDirty = true;
        if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
    });
}

void MainWindow::updateRibbonForMode()
{
    bool sensory = m_sensoryMode;
    bool detailedSensory = m_detailedSensoryMode;
    bool tpm = !sensory && !detailedSensory;

    // Home tab: show/hide TPM vs sensory vs detailed sensory buttons
    m_homeNewBtn->setVisible(tpm);
    m_homeLoadBtn->setVisible(tpm);
    m_homeCloseBtn->setVisible(tpm);

    m_homeSensNewBtn->setVisible(sensory);
    m_homeSensSaveBtn->setVisible(sensory);
    m_homeSensLoadXlBtn->setVisible(sensory);
    m_homeSensCloseBtn->setVisible(sensory);

    if (m_homeDetSensNewBtn) {
        m_homeDetSensNewBtn->setVisible(detailedSensory);
        m_homeDetSensSaveBtn->setVisible(detailedSensory);
        m_homeDetSensLoadXlBtn->setVisible(detailedSensory);
        m_homeDetSensCloseBtn->setVisible(detailedSensory);
    }

    // Reports tab: swap labels and connections
    disconnect(m_reportBtn1, &QToolButton::clicked, nullptr, nullptr);
    disconnect(m_reportBtn2, &QToolButton::clicked, nullptr, nullptr);

    if (sensory) {
        m_reportBtn1->setText("Sensory\nReport");
        m_reportBtn1->setIcon(QIcon(resourcePath() + "/images/ccell_icon.png"));
        m_reportBtn1->setToolTip("Generate PPTX report for selected sensory sessions");
        m_reportBtn2->setVisible(false);
        connect(m_reportBtn1, &QToolButton::clicked, this, [this]() {
            if (m_sensoryPanel) m_sensoryPanel->generateFullReport();
        });
        if (m_cleanupGroup) m_cleanupGroup->setVisible(false);

    } else if (detailedSensory) {
        m_reportBtn1->setText("Detailed\nSensory Report");
        m_reportBtn1->setIcon(QIcon(resourcePath() + "/images/ccell_icon_black.png"));
        m_reportBtn1->setToolTip("Generate PPTX report for selected detailed sensory sessions");
        m_reportBtn2->setVisible(false);
        connect(m_reportBtn1, &QToolButton::clicked, this, [this]() {
            if (m_detailedSensoryPanel) m_detailedSensoryPanel->generateFullReport();
        });
        if (m_cleanupGroup) m_cleanupGroup->setVisible(false);

    } else {
        m_reportBtn1->setText("Test Report");
        m_reportBtn1->setIcon(style()->standardIcon(QStyle::SP_FileDialogDetailedView));
        m_reportBtn1->setToolTip("Generate a PPTX report for the current sheet");
        m_reportBtn2->setVisible(true);
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
            // Prefix each row with its 1-based index so users have something
            // to refer to in the "(Showing 1+2)" hint. The itemChanged handler
            // strips this back off when the user renames a session.
            const QString labelText = QStringLiteral("%1. %2")
                                          .arg(i + 1)
                                          .arg(m_sensoryPanel->sessionLabel(sessions[i]));
            auto* navItem = new QListWidgetItem(labelText);
            navItem->setFlags(navItem->flags() | Qt::ItemIsEditable);
            // Stash the DB session id for refreshPresenceFor() lookups.
            navItem->setData(Qt::UserRole, qlonglong(sessions[i].id));
            m_sensoryNav->addItem(navItem);
        }

        int cur = m_sensoryPanel->currentSessionIndex();
        if (cur >= 0 && cur < m_sensoryNav->count())
            m_sensoryNav->setCurrentRow(cur);
    }   // QSignalBlocker restores prior blocked state here

    refreshAllPresence();
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
    m_propTable->setRowCount(14);
    m_propTable->setColumnCount(2);

    // ── Helper lambdas (same style as updateProperties) ──
    auto makeHeader = [&](int row, const QString& title) {
        QTableWidgetItem* it = new QTableWidgetItem(title);
        it->setFlags(Qt::ItemIsEnabled);
        it->setForeground(QColor(0x1F, 0x4E, 0x79));
        QFont f = it->font(); f.setBold(true); f.setPointSize(8); it->setFont(f);
        m_propTable->setItem(row, 0, it);
        QTableWidgetItem* it2 = new QTableWidgetItem();
        it2->setFlags(Qt::ItemIsEnabled);
        m_propTable->setItem(row, 1, it2);
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

    // ── Section: Test Properties ──
    makeHeader(7, "  Test Properties");
    makeEditable(8,  "Control",              sess->control);
    makeEditable(9,  "Blind?",               sess->isBlind ? "Y" : "N");
    makeEditable(10, "Primary Difference(s)", sess->primaryDifferences);

    // ── Section: Computed ──
    // Compute highest/lowest rated by "Overall Liking"
    QString highestRated, lowestRated;
    if (!sess->samples.isEmpty()) {
        double maxScore = -1.0;
        double minScore = 10.0;
        QStringList maxNames, minNames;

        for (const SensorySample& samp : sess->samples) {
            double score = samp.scores.value("Overall Liking", -1.0);
            if (score < 0) continue;
            if (score > maxScore) {
                maxScore = score;
                maxNames.clear();
                maxNames.append(samp.name);
            } else if (qAbs(score - maxScore) < 0.01) {
                maxNames.append(samp.name);
            }
            if (score < minScore) {
                minScore = score;
                minNames.clear();
                minNames.append(samp.name);
            } else if (qAbs(score - minScore) < 0.01) {
                minNames.append(samp.name);
            }
        }
        highestRated = maxNames.join(", ") + QString(" (%1)").arg(maxScore, 0, 'f', 1);
        lowestRated  = minNames.join(", ") + QString(" (%1)").arg(minScore, 0, 'f', 1);
    }

    makeHeader(11, "  Computed");
    makeReadOnly(12, "Highest Rated", highestRated);
    makeReadOnly(13, "Lowest Rated",  lowestRated);

    m_propTable->blockSignals(false);
}

// ─── Detailed Sensory Navigator + Properties ─────────────────────────────────
void MainWindow::refreshDetailedSensoryNavigator()
{
    if (!m_detailedSensoryNav || !m_detailedSensoryPanel) return;
    m_detailedSensoryNav->blockSignals(true);
    m_detailedSensoryNav->clear();
    auto sessions = m_detailedSensoryPanel->allSessions();
    for (const auto& s : sessions) {
        auto* item = new QListWidgetItem(m_detailedSensoryPanel->sessionLabel(s));
        // Same pattern as refreshSensoryNavigator: stash DB session id for
        // refreshPresenceFor() lookups.
        item->setData(Qt::UserRole, qlonglong(s.id));
        m_detailedSensoryNav->addItem(item);
    }
    int idx = m_detailedSensoryPanel->currentSessionIndex();
    if (idx >= 0 && idx < m_detailedSensoryNav->count())
        m_detailedSensoryNav->setCurrentRow(idx);
    m_detailedSensoryNav->blockSignals(false);

    refreshAllPresence();
}

void MainWindow::updateDetailedSensoryProperties()
{
    if (!m_detailedSensoryPanel) return;
    auto* sess = m_detailedSensoryPanel->currentSession();
    if (!sess) { m_propTable->setRowCount(0); return; }

    m_propTable->blockSignals(true);
    m_propTable->setRowCount(0);

    auto addRow = [&](const QString& prop, const QString& val) {
        int r = m_propTable->rowCount();
        m_propTable->insertRow(r);
        auto* pItem = new QTableWidgetItem(prop);
        pItem->setFlags(pItem->flags() & ~Qt::ItemIsEditable);
        m_propTable->setItem(r, 0, pItem);
        m_propTable->setItem(r, 1, new QTableWidgetItem(val));
    };

    // Session Info header
    int r = m_propTable->rowCount();
    m_propTable->insertRow(r);
    auto* hdr = new QTableWidgetItem("Session Info");
    hdr->setFlags(hdr->flags() & ~Qt::ItemIsEditable);
    hdr->setForeground(QColor(0, 120, 215));
    QFont f = hdr->font(); f.setBold(true); hdr->setFont(f);
    m_propTable->setItem(r, 0, hdr);
    m_propTable->setItem(r, 1, new QTableWidgetItem());

    addRow("Test Title", sess->testTitle);
    addRow("Assessor", sess->assessorName);
    addRow("Tester", sess->testerName);
    addRow("Media", sess->media);
    addRow("Date", sess->date);
    addRow("Samples", QString::number(sess->samples.size()));
    addRow("Viscosity", sess->viscosity);
    addRow("Oil Smell Liking", QString::number(sess->oilSmellLiking));
    addRow("Clog", sess->clog ? "Yes" : "No");
    if (sess->clog) addRow("Clog Oil Level", sess->clogOilLevel);
    addRow("Device Return Date", sess->deviceReturnDate);
    addRow("Facilitator", sess->facilitatorName);
    addRow("Facilitator Comment", sess->facilitatorComment);

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

    // Clicking a Test Average always means "show the aggregate, not a session".
    // Drop any session multi-selection so the navigator doesn't suggest the
    // averaged view is filtered to a subset, and hide the multi-select hint.
    if (m_sensoryNav) {
        QSignalBlocker b(m_sensoryNav);
        m_sensoryNav->clearSelection();
        m_sensoryNav->setCurrentRow(-1);
    }
    if (m_showingLabel) m_showingLabel->setVisible(false);

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

// Compute per-device averages across `sessions` and push the result to both
// the right-side Test Averages widget and the left-side averaged-table overlay.
// `sourceIndices` are the navigator row indices the sessions came from — used
// only so showAveragedChart() colours the radar consistently with the user's
// selection. Mirrors onTestAvgSelectionChanged() but is parameterised by a
// session set rather than by test title so the multi-select code path can
// share the math.
void MainWindow::showSensoryAveragesFor(const QVector<SensorySession>& sessions,
                                        const QVector<int>& sourceIndices)
{
    if (!m_sensoryPanel || !m_testAvgTable) return;

    m_testAvgTable->setRowCount(0);

    if (sessions.isEmpty()) {
        m_sensoryPanel->showAveragedChart(sourceIndices);
        return;
    }

    struct DeviceAccum { QMap<QString, double> sums; int count = 0; };
    QMap<QString, DeviceAccum> deviceMap;
    QStringList deviceOrder;
    QSet<QString> assessors, testers;

    for (const auto& sess : sessions) {
        if (!sess.assessorName.isEmpty()) assessors.insert(sess.assessorName);
        if (!sess.testerName.isEmpty())   testers.insert(sess.testerName);
        for (const auto& sample : sess.samples) {
            QString key = sample.name.isEmpty() ? QStringLiteral("Sample") : sample.name;
            if (!deviceMap.contains(key)) deviceOrder.append(key);
            DeviceAccum& acc = deviceMap[key];
            for (const QString& m : kSensoryMetrics)
                acc.sums[m] += sample.scores.value(m, 5);
            acc.count++;
        }
    }

    // Update the labels under the Test Averages list so the user sees the
    // assessors/testers/session count for the multi-select aggregation.
    QStringList assessorList = assessors.values();
    assessorList.sort(Qt::CaseInsensitive);
    QStringList testerList = testers.values();
    testerList.sort(Qt::CaseInsensitive);
    if (m_testAvgAssessors)
        m_testAvgAssessors->setText("Assessors: " +
            (assessorList.isEmpty() ? QString::fromUtf8("\xE2\x80\x94")
                                    : assessorList.join(", ")));
    if (m_testAvgTesters)
        m_testAvgTesters->setText("Testers: " +
            (testerList.isEmpty() ? QString::fromUtf8("\xE2\x80\x94")
                                  : testerList.join(", ")));
    if (m_testAvgCount)
        m_testAvgCount->setText("Sessions: " + QString::number(sessions.size()));

    // Push to the radar chart and the left-side table overlay.
    m_sensoryPanel->showAveragedChart(sourceIndices);

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

    // ── Detailed Sensory selection: switch to detailed sensory mode and load ──
    if (dlg.isDetailedSensorySelection()) {
        const QVector<int> detSensIds = dlg.selectedDetailedSensoryIds();
        if (detSensIds.isEmpty()) return;

        QVector<DetailedSensorySession> sessions;
        for (int id : detSensIds) {
            DetailedSensorySession sess = m_db->loadDetailedSensorySession(id);
            if (!sess.samples.isEmpty())
                sessions.append(sess);
        }

        if (sessions.isEmpty()) {
            showError("Database Load", "Could not load detailed sensory session(s) from the database.");
            return;
        }

        if (!m_detailedSensoryMode) {
            m_detailedSensoryBtn->setChecked(true);
        }
        m_detailedSensoryPanel->loadSessions(sessions);
        return;
    }

    // ── TPM file selection ──
    const QVector<int> ids = dlg.selectedFileIds();
    if (ids.isEmpty()) return;

    int loaded = 0;
    for (int id : ids) {
        FileResult dbResult = m_db->loadFile(id);
        if (dbResult.filePath.isEmpty()) continue;

        // If the original file still exists on disk, re-process it through
        // the full pipeline so legacy-format normalization and any other
        // updates are applied.  Fall back to DB cache if the file is gone.
        FileResult result;
        if (QFile::exists(dbResult.filePath)) {
            result = m_processor->processFile(dbResult.filePath);
            if (result.filePath.isEmpty()) {
                result = dbResult;   // processing failed — use DB cache
            } else {
                // Inherit id+version from the DB row so the save below is
                // an UPDATE, not an INSERT. processFile reads only the
                // .xlsx so the returned FileResult has id=-1; without this
                // the save trips the files.file_path UNIQUE constraint and
                // pops UniqueViolationDialog on every Load-from-Database.
                result.id      = dbResult.id;
                result.version = dbResult.version;

                // Update DB with fresh data. v2.0.1: LiveSync owns per-cell
                // sync; this is a one-shot full-row save for the freshly
                // loaded file.
                if (m_db) m_db->saveFile(result);
            }
        } else {
            result = dbResult;
        }

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
        // Restore-last-session: DB-loaded files joined the working set, so mark
        // the snapshot dirty (they would otherwise be missing from the recovery
        // store until later edited). Guarded like every other noteDirty() site.
        if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
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
    int cancelled = 0;

    // v2.0.6: inheritExistingIdsAndVersions() used to run here on every
    // Ctrl+U and every 5-second auto-save tick (the m_dbSaveTimer slot).
    // That fired N synchronous Postgres SELECTs on the UI thread per
    // tick, which produced "Not Responding" freezes on slower LAN
    // segments. The call is now made once at load time
    // (SensoryPanel::loadSessions / DetailedSensoryPanel::loadSessions),
    // which is the only point where new id==-1 sessions actually enter
    // the panel state — after that, ids stay set and the lookup would
    // be a no-op anyway.

    // ── Save TPM files ──
    // v2.0.1: LiveSync persists per-cell, so this batch-save path is mostly
    // a safety net for files loaded fresh from disk that haven't yet sync'd.
    // v2.0.5: use tryWriteFile directly so VersionMismatch / RowDeleted
    // from rows LiveSync already wrote can be skipped without counting
    // as failures — the row in the DB is already current.
    int filesSkipped = 0;
    for (int i = 0; i < m_loadedFiles.size(); ++i) {
        FileResult& fr = m_loadedFiles[i];
        if (!m_modifiedFilePaths.contains(fr.filePath)) continue;
        const QString oldPath = fr.filePath;
        if (!m_db) { ++failed; continue; }
        const DVE::WriteResult r = m_db->tryWriteFile(fr);
        if (r == DVE::WriteResult::Success) {
            m_modifiedFilePaths.remove(oldPath);
            ++saved;
        } else if (r == DVE::WriteResult::VersionMismatch
                || r == DVE::WriteResult::RowDeleted) {
            // LiveSync already wrote this file's rows per-cell. Drop
            // the dirty flag so we don't keep retrying on every tick.
            m_modifiedFilePaths.remove(oldPath);
            ++filesSkipped;
            qInfo().noquote()
                << "[onUpdateDatabase] TPM file" << fr.fileName
                << "skipped — already up to date via LiveSync (result="
                << static_cast<int>(r) << ")";
        } else {
            ++failed;
        }
    }

    // ── Save sensory sessions ──
    int sensSaved = 0;
    int sensSkipped = 0;
    if (m_sensoryPanel) {
        auto sessions = m_sensoryPanel->allSessions();
        for (SensorySession& sess : sessions) {
            if (DVE::isPlaceholderSession(sess)) continue;
            if (!m_db) { ++failed; continue; }

            // v2.1.0+: Test Title rename → new DB row. If the user changed
            // the Test Title (and therefore sessionName) since this session
            // was loaded, we don't UPDATE the existing row in place — that
            // would silently overwrite the original and the user has no way
            // to retrieve the pre-rename data. Instead we force INSERT so
            // both rows live in the DB. The user can clean up the old row
            // later if they want. originalSessionName is empty for sessions
            // that have never been persisted; those still go through INSERT
            // via id == -1 as before.
            const bool isRename = sess.id > 0
                && !sess.originalSessionName.isEmpty()
                && sess.originalSessionName != sess.sessionName;
            if (isRename) {
                qInfo().noquote() << "[onUpdateDatabase] sensory rename detected:"
                                  << sess.originalSessionName << "→"
                                  << sess.sessionName
                                  << "— routing to INSERT (old row preserved)";
                sess.id      = -1;
                sess.version = 0;
            }

            DVE::WriteResult r = m_db->tryWriteSensorySession(sess);

            // v2.1.0+: name collision on a fresh INSERT (either a brand-new
            // session or a rename whose target name another session already
            // owns). Old behavior was to offer "override — delete the other
            // row" but the user prefers to never delete; show an informative
            // message and skip this save so the user picks a different name.
            if (r == DVE::WriteResult::UniqueViolation) {
                QMessageBox::information(
                    this,
                    tr("Sensory Session Name Taken"),
                    tr("Another sensory session named \"%1\" already exists "
                       "for tester \"%2\" on %3.\n\n"
                       "Pick a different Test Title and save again — the "
                       "rename was not applied to the database.")
                        .arg(sess.sessionName, sess.testerName, sess.date));
                ++cancelled;
                continue;
            }

            if (r == DVE::WriteResult::Success) {
                ++sensSaved;
            } else if (r == DVE::WriteResult::VersionMismatch
                    || r == DVE::WriteResult::RowDeleted) {
                ++sensSkipped;
                qInfo().noquote()
                    << "[onUpdateDatabase] sensory session"
                    << sess.sessionName
                    << "skipped — already up to date via LiveSync (result="
                    << static_cast<int>(r) << ")";
            } else {
                ++failed;
            }
        }
        if (sensSaved > 0)
            m_sensorySessionsDirty = false;

        // v2.1.0+: merge id/version/originalSessionName back into panel
        // state from the local `sessions` copy that tryWriteSensorySession's
        // byRef back-fill updated. This is the authoritative path post-save
        // — by-index sync that also handles renames (where the new row's id
        // wouldn't be reachable from inheritExistingIdsAndVersions's
        // natural-key lookup against the OLD m_sessions[i].sessionName).
        m_sensoryPanel->syncSavedSessionState(sessions);
        // v2.0.10 carry-over: also run the natural-key reconciliation for
        // sessions that came in with id == -1 from non-save paths (e.g.
        // Excel imports that happened earlier in the same Ctrl+U tick).
        m_sensoryPanel->inheritExistingIdsAndVersions();
    }

    // ── Save detailed-sensory sessions ──
    // Plan C C10: this block previously did not exist — closing or Ctrl+U with
    // only detailed-sensory edits persisted nothing, silently losing the work
    // even though m_detailedSensorySessionsDirty was set. Mirrors the sensory
    // block above. DetailedSensorySession carries no originalSessionName, so
    // there is no in-place-rename → force-INSERT branch here.
    // Reconciliation happens on both sides of the loop: before it,
    // inheritExistingIdsAndVersions() resolves id/version for id<=0 sessions so
    // re-imports take UPDATE, not INSERT; after it, syncSavedSessionState()
    // back-fills the written id/version (and per-image imageIds/imageVersions)
    // from the local detSessions copy into the panel's m_sessions, so the panel
    // no longer holds id == -1 and repeat saves stop re-INSERTing images. The
    // byRef tryWriteDetailedSensorySession populates those anchors on Success.
    int detSaved = 0;
    int detSkipped = 0;
    if (m_detailedSensoryPanel) {
        m_detailedSensoryPanel->inheritExistingIdsAndVersions();
        auto detSessions = m_detailedSensoryPanel->allSessions();
        for (DetailedSensorySession& sess : detSessions) {
            if (DVE::isPlaceholderSession(sess)) continue;
            if (!m_db) { ++failed; continue; }

            DVE::WriteResult r = m_db->tryWriteDetailedSensorySession(sess);

            // Name collision on a fresh INSERT: surface it and skip so the user
            // picks a different Test Title (symmetric with the sensory branch's
            // never-delete policy).
            if (r == DVE::WriteResult::UniqueViolation) {
                QMessageBox::information(
                    this,
                    tr("Detailed Sensory Session Name Taken"),
                    tr("Another detailed sensory session named \"%1\" already "
                       "exists for tester \"%2\" on %3.\n\n"
                       "Pick a different Test Title and save again — it was not "
                       "written to the database.")
                        .arg(sess.sessionName, sess.testerName, sess.date));
                ++cancelled;
                continue;
            }

            if (r == DVE::WriteResult::Success) {
                ++detSaved;
            } else if (r == DVE::WriteResult::VersionMismatch
                    || r == DVE::WriteResult::RowDeleted) {
                ++detSkipped;
                qInfo().noquote()
                    << "[onUpdateDatabase] detailed sensory session"
                    << sess.sessionName
                    << "skipped — already up to date via LiveSync (result="
                    << static_cast<int>(r) << ")";
            } else {
                ++failed;
            }
        }
        if (detSaved > 0)
            m_detailedSensorySessionsDirty = false;

        // Merge id/version (+ per-image imageIds/imageVersions) back into panel
        // state from the local `detSessions` copy that tryWriteDetailedSensorySession's
        // byRef back-fill updated — symmetric with the sensory block above. Without
        // this, allSessions() returning m_sessions by value means the panel's
        // m_sessions[i].id stays -1 after a first save, so every repeat save in the
        // same run re-INSERTs images and churns rows.
        m_detailedSensoryPanel->syncSavedSessionState(detSessions);
    }
    updateDbSyncIndicator();

    const int total   = saved + sensSaved + detSaved;
    const int skipped = filesSkipped + sensSkipped + detSkipped;
    if (total == 0 && failed == 0) {
        if (skipped > 0) {
            updateStatusBar(
                QString("Database already up to date "
                        "(%1 item%2 already live-synced).")
                    .arg(skipped).arg(skipped > 1 ? "s" : ""));
        } else {
            updateStatusBar("Database already up to date.");
        }
        return;
    }

    if (failed == 0) {
        QString msg = QString("Database updated (%1 file%2")
                          .arg(total).arg(total > 1 ? "s" : "");
        if (sensSaved > 0)
            msg += QString(", %1 sensory session%2")
                       .arg(sensSaved).arg(sensSaved > 1 ? "s" : "");
        if (detSaved > 0)
            msg += QString(", %1 detailed sensory session%2")
                       .arg(detSaved).arg(detSaved > 1 ? "s" : "");
        msg += " saved).";
        updateStatusBar(msg);
    } else {
        const QString lastError = m_db ? m_db->lastError() : QString();
        qWarning() << "Database save error:" << lastError;
        showError("Database Error",
                  QString("%1 item(s) failed to save.\n\n"
                          "Last error from the database:\n%2\n\n"
                          "Full log: %3\\dataviewer.log")
                      .arg(failed)
                      .arg(lastError.isEmpty() ? QStringLiteral("(no detail)")
                                               : lastError)
                      .arg(QStandardPaths::writableLocation(
                          QStandardPaths::AppLocalDataLocation)));
    }
}

void MainWindow::onExportToExcelTriggered()
{
    // LiveSync already persists every cell commit to the database, so Ctrl+S
    // is purely a manual flush of the debounced Excel write-back queue.
    if (m_pendingWrites.isEmpty()) {
        updateStatusBar(tr("Nothing to export"));
        return;
    }
    flushExcelWrites();
    updateStatusBar(tr("Exported to Excel"));
}

QVector<RecoveryEntry> MainWindow::captureRecoveryState() const
{
    // Plan C: snapshot ALL THREE in-memory stores every flush, regardless of the
    // active mode (both sensory panels can hold sessions simultaneously). Thin
    // glue over the already-tested serializers; runs on the UI thread and must
    // not block. Each entry carries a self-contained value payload so the flush
    // worker can persist it off-thread without touching live UI state.
    QVector<RecoveryEntry> out;
    out.reserve(m_loadedFiles.size());

    // ── TPM files: stable id = filePath (survives m_loadedFiles reshuffle) ─────
    for (const FileResult& f : m_loadedFiles) {
        RecoveryEntry e;
        e.kind        = RecoveryKind::Tpm;
        e.id          = f.filePath;
        e.displayName = f.fileName;
        e.sourcePath  = f.filePath;
        e.dirty       = m_modifiedFilePaths.contains(f.filePath);
        e.payload     = fileResultToJson(f);
        out.append(e);
    }

    // ── Sensory sessions: index-based id ("sensory_<i>"). The snapshot captures
    //    the whole set on each flush, so the prune step in the flush worker
    //    handles removals — no need for an id that survives reordering. ─────────
    if (m_sensoryPanel) {
        const QVector<SensorySession> sessions = m_sensoryPanel->allSessions();
        for (int i = 0; i < sessions.size(); ++i) {
            const SensorySession& s = sessions[i];
            // M-1: never snapshot a never-touched placeholder ("New Session"
            // with no fields/scores/samples). Otherwise the C8 reopen prompt
            // would offer an empty session as recoverable. Index i still tracks
            // the panel's position; the snapshot's prune step drops any stale
            // blob, so skipping mid-loop is safe.
            if (DVE::isPlaceholderSession(s)) continue;
            RecoveryEntry e;
            e.kind        = RecoveryKind::Sensory;
            e.id          = QStringLiteral("sensory_") + QString::number(i);
            e.displayName = s.sessionName;
            e.sourcePath  = s.sourceFilePath;
            e.dirty       = m_sensorySessionsDirty;
            e.payload     = sensorySessionToJson(s);
            out.append(e);
        }
    }

    // ── Detailed-sensory sessions: index-based id ("detailed_<i>") ─────────────
    if (m_detailedSensoryPanel) {
        const QVector<DetailedSensorySession> sessions =
            m_detailedSensoryPanel->allSessions();
        for (int i = 0; i < sessions.size(); ++i) {
            const DetailedSensorySession& s = sessions[i];
            // M-1: skip never-touched placeholders (see the sensory loop above).
            if (DVE::isPlaceholderSession(s)) continue;
            RecoveryEntry e;
            e.kind        = RecoveryKind::Detailed;
            e.id          = QStringLiteral("detailed_") + QString::number(i);
            e.displayName = s.sessionName;
            e.dirty       = m_detailedSensorySessionsDirty;
            e.payload     = detailedSensorySessionToJson(s);
            out.append(e);
        }
    }

    return out;
}

void MainWindow::maybeOfferRecovery()
{
    // Restore-last-session reopen prompt. adoptPreviousSession() (run in the
    // ctor) moved the previous session's store -- however the app last closed,
    // clean or not -- into Recovery_prev/. hasRecoverable() is false only when
    // there genuinely was no previous session (a true first run, or one the user
    // declined last time), so return silently in that case.
    if (!m_recovery || !m_recovery->hasRecoverable())
        return;

    const QVector<RecoveryEntry> items = m_recovery->recoverableItems();
    if (items.isEmpty())
        return;   // defensive: hasRecoverable() implies non-empty, but never assume.

    const int n = items.size();

    // List the item names so the user sees WHAT will reopen. Cap the visible
    // list so a large session does not produce a giant dialog; summarize the rest.
    QStringList names;
    const int kMaxShown = 12;
    for (int i = 0; i < items.size() && i < kMaxShown; ++i)
        names << (QStringLiteral("  • ") + items.at(i).displayName);
    if (items.size() > kMaxShown)
        names << tr("  …and %1 more").arg(items.size() - kMaxShown);

    const QMessageBox::StandardButton answer = QMessageBox::question(
        this, tr("Reopen Previous Session"),
        tr("You had %1 file(s)/session(s) open when you last used DataViewer:"
           "\n\n%2\n\n"
           "Reopen them and pick up where you left off?")
            .arg(n).arg(names.join(QLatin1Char('\n'))),
        QMessageBox::Yes | QMessageBox::No);

    if (answer == QMessageBox::Yes)
        restoreItems(items);
    // No (or the dialog was dismissed): do NOT reopen, but deliberately KEEP
    // Recovery_prev/ this session. An accidental "No"/Escape must never destroy
    // unsaved work -- the previous session stays retrievable via Tools->Recover
    // (C9) for the rest of this session. It is not re-offered next launch:
    // adoptPreviousSession() replaces Recovery_prev/ with the current session's
    // store at startup, so a declined session quietly ages out without nagging.
}

void MainWindow::onRecover()
{
    // Plan C C9: manual Tools->Recover entry point. Unlike the one-shot C8 prompt,
    // this works any time this session — including after the user clicked "No" on
    // startup — because Recovery_prev/ is kept until the next clean close. It also
    // lets the user pick WHICH previous-session items to reload rather than the
    // all-or-nothing reopen prompt.
    const auto items = m_recovery ? m_recovery->recoverableItems()
                                  : QVector<RecoveryEntry>{};
    if (items.isEmpty()) {
        QMessageBox::information(this, tr("Recover"),
            tr("There is no recoverable work from a previous session."));
        return;
    }

    RecoverDialog dlg(items, this);
    if (dlg.exec() == QDialog::Accepted) {
        const auto chosen = dlg.selected();
        if (!chosen.isEmpty())
            restoreItems(chosen);
    }
}

void MainWindow::restoreItems(const QVector<RecoveryEntry>& items)
{
    if (items.isEmpty())
        return;

    // The mode to switch to once everything is loaded: the FIRST restored item's
    // kind (TPM = default central view; sensory/detailed = the matching toggle).
    const RecoveryKind firstKind = items.first().kind;

    bool tpmRestored = false;
    QVector<SensorySession>         sensorySessions;
    QVector<DetailedSensorySession> detailedSessions;

    for (const RecoveryEntry& entry : items) {
        switch (entry.kind) {
        case RecoveryKind::Tpm: {
            FileResult f = fileResultFromJson(entry.payload);
            if (f.filePath.isEmpty())
                continue;   // unparseable / empty payload — skip rather than load junk.

            // Recompute the per-sheet plot series. tpmTrend/puffCounts are
            // intentionally NOT serialized (they are pure derived data), so a
            // restored FileResult would otherwise render empty plots. This
            // mirrors DatabaseManager::loadFile's rebuild-from-row-data step
            // exactly: walk every sample's rows, skipping rows with a missing
            // before/after weight (incomplete measurements), and re-derive the
            // trend + puff-count series the plot engine consumes.
            for (SheetResult& sheet : f.sheets) {
                sheet.tpmTrend.clear();
                sheet.puffCounts.clear();
                for (const SampleResult& sr : sheet.samples) {
                    for (const DataRow& dr : sr.rows) {
                        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
                        sheet.tpmTrend.append(dr.tpm);
                        sheet.puffCounts.append(dr.puffs);
                    }
                }
            }

            // Dedup against the live set by file path (mirrors the DB-load path):
            // if the same file is already open, replace it; otherwise append.
            bool alreadyLoaded = false;
            for (int i = 0; i < m_loadedFiles.size(); ++i) {
                if (m_loadedFiles[i].filePath == f.filePath) {
                    m_loadedFiles[i] = f;
                    m_currentFileIndex = i;
                    alreadyLoaded = true;
                    break;
                }
            }
            if (!alreadyLoaded) {
                m_loadedFiles.append(f);
                m_currentFileIndex = m_loadedFiles.size() - 1;
            }
            // Mark dirty so the user can re-save through the normal optimistic-
            // concurrency path. (The restored file may never have been persisted,
            // or may differ from the DB row that died with the prior session.)
            m_modifiedFilePaths.insert(f.filePath);
            tpmRestored = true;
            break;
        }
        case RecoveryKind::Sensory: {
            SensorySession s = sensorySessionFromJson(entry.payload);
            // Skip a corrupt/empty recovered blob (mirrors the TPM filePath guard
            // above) so we don't load a phantom blank session.
            if (!DVE::isPlaceholderSession(s))
                sensorySessions.append(s);
            break;
        }
        case RecoveryKind::Detailed: {
            DetailedSensorySession s = detailedSensorySessionFromJson(entry.payload);
            if (!DVE::isPlaceholderSession(s))
                detailedSessions.append(s);
            break;
        }
        }
    }

    // ── Sensory: construct the panel if a recovered session needs it before the
    //    user ever toggled the mode, then load + inherit ids/versions. ─────────
    if (!sensorySessions.isEmpty()) {
        if (!m_sensoryPanel)
            initSensoryPanel();
        m_sensoryPanel->loadSessions(sensorySessions);
        m_sensorySessionsDirty = true;
        // Match an existing DB row's id+version (by name) so the next save is an
        // UPDATE, not an INSERT — prevents spurious unique-violation / stale-row
        // dialogs for sessions that were already persisted before the crash.
        m_sensoryPanel->inheritExistingIdsAndVersions();
    }

    // ── Detailed sensory: same lazy-construct + load + inherit. ────────────────
    if (!detailedSessions.isEmpty()) {
        if (!m_detailedSensoryPanel)
            initDetailedSensoryPanel();
        m_detailedSensoryPanel->loadSessions(detailedSessions);
        m_detailedSensorySessionsDirty = true;
        m_detailedSensoryPanel->inheritExistingIdsAndVersions();
    }

    // ── Refresh the TPM UI if any file was restored (the panels self-refresh on
    //    loadSessions, so nothing extra is needed for sensory/detailed). ───────
    if (tpmRestored) {
        m_currentSheetIndex  = 0;
        m_currentSampleIndex = 0;
        populateFileTree();
        populateSheetCombo();
        displayCurrentSample();
    }

    // ── Switch to the FIRST restored item's mode so the user lands on recovered
    //    work. TPM is the default central view (toggle both sensory modes off);
    //    sensory/detailed activate the matching toggle (which builds the panel if
    //    it somehow isn't built yet and shows it). ─────────────────────────────
    switch (firstKind) {
    case RecoveryKind::Tpm:
        if (m_sensoryMode)         toggleSensoryMode(false);
        if (m_detailedSensoryMode) toggleDetailedSensoryMode(false);
        break;
    case RecoveryKind::Sensory:
        if (!m_sensoryMode)        toggleSensoryMode(true);
        break;
    case RecoveryKind::Detailed:
        if (!m_detailedSensoryMode) toggleDetailedSensoryMode(true);
        break;
    }

    // ── Re-capture the just-restored state into the NEW live store so a crash
    //    THIS session can recover it again. Gated on m_recoveryArmed (matching the
    //    other noteDirty() call sites): when recovery isn't armed we must NOT write,
    //    because the prior crash data is still stranded in the live dir and a flush
    //    here would overwrite it. ──────────────────────────────────────────────────
    if (m_recoveryArmed && m_recovery)
        m_recovery->noteDirty();

    updateStatusBar(tr("Recovered %1 item(s) from the previous session.")
                        .arg(items.size()));
}

void MainWindow::markFileModified()
{
    FileResult* f = currentFile();
    if (!f) return;
    m_modifiedFilePaths.insert(f->filePath);
    // Plan C: this is the single TPM edit chokepoint (all 7 TPM edit sites route
    // here), so one noteDirty() covers every TPM data change.
    if (m_recoveryArmed && m_recovery) m_recovery->noteDirty();
    updateDbSyncIndicator();
    // Restart debounce timer — saves 5 s after last change
    m_dbSaveTimer->start();

    // Phase 5: bump presence intent to "editing" on first edit. The
    // PresenceManager call is idempotent — repeated setIntent("editing")
    // is cheap and keeps the heartbeat fresh. TODO(Phase 5 follow-up):
    // also bump on sensory-session edits (currently scoped to TPM file
    // edits, which is the common path).
    if (m_presence && m_presence->activeIntent() != QStringLiteral("editing"))
        m_presence->setIntent(QStringLiteral("editing"));
}

void MainWindow::updateDbSyncIndicator()
{
    if (!m_db) return;
    bool isNas = m_db->currentPath().startsWith("//") ||
                 m_db->currentPath().startsWith("\\\\");
    QString prefix = isNas ? "NAS DB: " : "Local DB: ";
    bool hasTPM      = !m_modifiedFilePaths.isEmpty();
    bool hasSensory  = m_sensorySessionsDirty;
    bool hasDetailed = m_detailedSensorySessionsDirty;
    if (!hasTPM && !hasSensory && !hasDetailed) {
        setStatusDb(prefix + "Synced", DbStatusOk);
    } else {
        QStringList parts;
        if (hasTPM)
            parts << QString("%1 TPM").arg(m_modifiedFilePaths.size());
        if (hasSensory)
            parts << "sensory";
        if (hasDetailed)
            parts << "detailed sensory";
        setStatusDb(prefix + parts.join(" + ") + " modified (Ctrl+U)", DbStatusModified);
    }
}

QVector<QString> MainWindow::unsavedInventory() const
{
    // const, but intentionally non-pure: the allSessions() calls below flush the
    // active editor first (each calls saveCurrentTester()), so the panels' live
    // form state is committed into m_sessions before we inventory it. That side
    // effect reaches through the panel member pointers; it is wanted, not a leak —
    // the inventory must reflect in-progress edits, not the last-applied snapshot.
    QVector<QString> items;

    // ── Modified TPM files (by display name) ────────────────────────────────
    // Resolve each dirty path to its loaded FileResult so the user sees the
    // workbook name, not a long absolute path. A dirty path with no matching
    // loaded file (should not happen) still lists by basename so nothing is
    // silently dropped from the inventory.
    if (!m_modifiedFilePaths.isEmpty()) {
        QSet<QString> remaining = m_modifiedFilePaths;
        for (const FileResult& f : m_loadedFiles) {
            if (remaining.remove(f.filePath))
                items.append(QStringLiteral("TPM file: ") + f.fileName);
        }
        for (const QString& path : remaining)
            items.append(QStringLiteral("TPM file: ") + QFileInfo(path).fileName());
    }

    // ── Dirty Sensory sessions (placeholders excluded) ──────────────────────
    if (m_sensorySessionsDirty && m_sensoryPanel) {
        const QVector<SensorySession> sessions = m_sensoryPanel->allSessions();
        for (const SensorySession& s : sessions) {
            if (DVE::isPlaceholderSession(s)) continue;
            items.append(QStringLiteral("Sensory session: ")
                         + m_sensoryPanel->sessionLabel(s));
        }
    }

    // ── Dirty Detailed-sensory sessions (placeholders excluded) ─────────────
    if (m_detailedSensorySessionsDirty && m_detailedSensoryPanel) {
        const QVector<DetailedSensorySession> sessions =
            m_detailedSensoryPanel->allSessions();
        for (const DetailedSensorySession& s : sessions) {
            if (DVE::isPlaceholderSession(s)) continue;
            items.append(QStringLiteral("Detailed sensory session: ")
                         + m_detailedSensoryPanel->sessionLabel(s));
        }
    }

    return items;
}

bool MainWindow::promptSaveDatabase()
{
    // Plan C C10: one consolidated prompt for ALL unsaved work across the three
    // modes. The inventory already filters placeholder sessions, so a brand-new
    // empty "New Session" never triggers a prompt. Empty inventory ⇒ nothing to
    // save, close proceeds silently.
    const QVector<QString> inventory = unsavedInventory();
    if (inventory.isEmpty()) return true;

    QString body = tr("You have unsaved work:\n\n  • %1\n\n"
                      "Save your unsaved work before closing?")
                       .arg(inventory.join(QStringLiteral("\n  • ")));

    QMessageBox box(QMessageBox::Question, tr("Unsaved Changes"), body,
                    QMessageBox::NoButton, this);
    QPushButton* saveBtn    = box.addButton(tr("Save All"),
                                            QMessageBox::AcceptRole);
    QPushButton* discardBtn = box.addButton(tr("Discard"),
                                            QMessageBox::DestructiveRole);
    QPushButton* cancelBtn  = box.addButton(QMessageBox::Cancel);
    box.setDefaultButton(saveBtn);
    box.exec();

    QAbstractButton* clicked = box.clickedButton();
    // Treat the window-close ([X]) and Esc as Cancel — never lose data or
    // proceed past an ambiguous dismissal.
    if (clicked == cancelBtn || clicked == nullptr)
        return false;                       // → closeEvent vetoes the close
    if (clicked == discardBtn)
        return true;                        // → proceed, persist nothing

    // ── Save All (the only outcome left — Cancel/Discard returned above) ──────
    // For untitled (non-placeholder) sessions, route through the panel's own
    // save() so the user is offered a Save-As for the on-disk copy. The
    // "no path yet" guard is read here (before save() runs, since save() sets
    // the path); already-saved panels skip the dialog. onUpdateDatabase() is
    // the authoritative DB persist for ALL sessions (named or freshly-titled),
    // so the save() calls are purely the disk-file courtesy for untitled work.
    if (m_sensorySessionsDirty && m_sensoryPanel
        && !m_sensoryPanel->hasSavePath()) {
        m_sensoryPanel->save();
    }
    if (m_detailedSensorySessionsDirty && m_detailedSensoryPanel
        && !m_detailedSensoryPanel->hasSavePath()) {
        m_detailedSensoryPanel->save();
    }
    onUpdateDatabase();
    return true;
}

QString MainWindow::lastBrowseDir() const
{
    if (!m_lastBrowseDir.isEmpty() && QDir(m_lastBrowseDir).exists())
        return m_lastBrowseDir;

    // Fallback: user's Documents folder (via OutputPaths canonical resolver)
    return OutputPaths::documentsDir();
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
        QStringLiteral(
            "<b>DataViewer Enterprise</b><br/>Version %1<br/><br/>"
            "Professional engineering data analysis tool.<br/>"
            "Built with Qt 6 and C++17.<br/><br/>"
            "\xC2\xA9 2025 SDR. All rights reserved.")
            .arg(QApplication::applicationVersion()));
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
        if (p.endsWith(".xlsx", Qt::CaseInsensitive) ||
            p.endsWith(".xlsm", Qt::CaseInsensitive) ||
            p.endsWith(".xls",  Qt::CaseInsensitive) ||
            p.endsWith(".json", Qt::CaseInsensitive))
            routeFile(p);
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
// ─── Plan B B8: incomplete-data banner ───────────────────────────────────────
void MainWindow::updateIncompleteDataBanner()
{
    if (!m_incompleteDataBanner) return;

    // Only relevant in TPM mode; hide in sensory modes.
    if (m_sensoryMode || m_detailedSensoryMode) {
        m_incompleteDataBanner->dismiss();
        return;
    }

    const FileResult* f = currentFile();
    if (!f) {
        m_incompleteDataBanner->dismiss();
        return;
    }

    // Show when ANY sheet in the active file is flagged incomplete.
    const bool anyIncomplete = std::any_of(
        f->sheets.begin(), f->sheets.end(),
        [](const SheetResult& s) { return s.dbDataIncomplete; });

    if (anyIncomplete) {
        m_incompleteDataBanner->showForFile(f->fileName);
    } else {
        m_incompleteDataBanner->dismiss();
    }
}

// ─── Plan C T7-T9: offline-mode slots ────────────────────────────────────────
void MainWindow::onConnectionWentOffline()
{
    qInfo() << "MainWindow: ConnectionMonitor reports offline";
    if (m_db) m_db->setOnline(false);

    // Best-effort: try to open the snapshot if we don't already have one.
    if (m_snapshot && !m_snapshot->isOpen()) {
        m_snapshot->openReadOnly();
    }

    if (m_offlineBanner) {
        if (m_snapshot && m_snapshot->isOpen()) {
            m_offlineBanner->setLastSync(m_snapshot->snapshotTakenAt());
        } else {
            m_offlineBanner->setLastSync(QDateTime());  // "(No previous snapshot.)"
        }
        m_offlineBanner->setPendingCount(m_pendingEdits.size());
        m_offlineBanner->setVisible(true);
    }

    // Stop NOTIFY listening + presence heartbeat — both need a live PG
    // connection. Re-enabled in onConnectionCameOnline.
    if (m_notify)   m_notify->unsubscribe();
    if (m_presence) m_presence->deactivate();

    setStatusDb(tr("Lost database connection — working offline."), DbStatusDisconnected);
}

void MainWindow::onConnectionCameOnline()
{
    qInfo() << "MainWindow: ConnectionMonitor reports online";

    // v2.1.0: the monitor watches m_pgConn (NOTIFY/heartbeat), but writes
    // go through m_db's separate QSqlDatabase. Postgres often drops idle
    // sockets during a network blip without notifying the client; Qt's
    // QSqlDatabase doesn't poll, so the driver still reports "open" while
    // the next BEGIN / PREPARE fails with "server closed the connection
    // unexpectedly". Reopening m_db here makes the driver-level state
    // match reality before any save attempts go through.
    if (m_db && !m_db->reopen()) {
        qWarning() << "MainWindow: m_db->reopen() failed on cameOnline:"
                   << m_db->lastError() << "— staying offline";
        setStatusDb(tr("Reconnect failed — still offline."), DbStatusDisconnected);
        return;
    }
    if (m_db) m_db->setOnline(true);

    if (m_offlineBanner) {
        m_offlineBanner->setVisible(false);
    }

    // Re-subscribe to NOTIFY. Presence reactivation happens lazily — the
    // user must select a resource again (or it's set by mode switches).
    if (m_notify && !m_notify->isSubscribed()) {
        if (!m_notify->subscribe()) {
            qWarning() << "MainWindow: NOTIFY resubscribe failed:"
                       << "live updates will not resume this session";
        }
    }
    // PresenceManager doesn't have a top-level start(); it self-starts in
    // activate(). Nothing to do here — the next activate() will spin the
    // heartbeat back up.

    setStatusDb(tr("Reconnected to database."), DbStatusOk);

    // C4: drain the persistent per-cell queue FIRST. The cell-level queue
    // contains granular edits captured offline (table/row/column tuples);
    // file-level flushPendingEdits below issues full saveFile per loaded
    // file, which would race the per-cell drain if it ran first. Ordering
    // per-cell-then-file ensures every cell edit lands via the precise
    // dve_commit_cell path before any file-level UPDATE.
    if (m_liveSync) {
        const int replayed = m_liveSync->flushPending();
        if (replayed > 0) {
            qInfo() << "MainWindow: LiveSync replayed" << replayed
                    << "per-cell edits on reconnect";
        }
    }

    flushPendingEdits();
}

void MainWindow::onOfflineRetryClicked()
{
    // v2.1.0: stop/start used to leave the user waiting up to 30 s for the
    // next ping tick. forceReconnect() runs one attempt synchronously and
    // either emits cameOnline (→ m_db->reopen() in our handler) or re-arms
    // the reconnect timer with a fresh jittered delay.
    if (m_monitor) {
        setStatusDb(tr("Reconnecting…"), DbStatusDisconnected);
        m_monitor->forceReconnect();
    } else {
        // Offline-boot path: no monitor was constructed. Fall back to a
        // status message; the user must close and reopen.
        setStatusDb(tr("Offline boot — please close and reopen DataViewer to reconnect."),
                    DbStatusDisconnected);
    }
}

void MainWindow::onRefreshSnapshotTriggered()
{
    if (!m_db || !m_db->isOnline() || !m_snapshot
        || !m_pgConn || !m_pgConn->isOpen()) {
        QMessageBox::information(this, tr("Refresh Offline Snapshot"),
            tr("Snapshot can only be refreshed when connected to the database."));
        return;
    }

    QProgressDialog progress(tr("Refreshing offline snapshot..."), QString(),
                             0, 0, this);
    progress.setWindowModality(Qt::WindowModal);
    progress.setCancelButton(nullptr);
    progress.show();
    QApplication::processEvents();

    const bool ok = m_snapshot->regenerate(m_pgConn);
    progress.close();

    if (ok) {
        QMessageBox::information(this, tr("Refresh Offline Snapshot"),
            tr("Offline snapshot refreshed successfully."));
    } else {
        QMessageBox::warning(this, tr("Refresh Offline Snapshot"),
            tr("Snapshot refresh failed: %1").arg(m_snapshot->lastError()));
    }
}

void MainWindow::flushPendingEdits()
{
    if (m_pendingEdits.isEmpty()) return;

    // C4: callers must drain the per-cell LiveSync queue first (see the
    // reconnect handler above). Asserting here makes the contract
    // self-checking in debug builds so a future caller can't accidentally
    // race saveFile against an unflushed per-cell drain. The assert is
    // release-only-elided so production behavior is unchanged.
    Q_ASSERT(!m_liveSync || m_liveSync->pendingCount() == 0);

    qInfo() << "MainWindow: flushing" << m_pendingEdits.size()
            << "pending edits on reconnect";

    int succeeded = 0;
    int failed = 0;

    // H7: track which file paths' save succeeded; retain edits whose file
    // either could not be located or whose saveFile failed. Previously the
    // queue was unconditionally cleared at the end, so a transient PG outage
    // mid-retry would silently drop offline-captured cells.
    QSet<QString> succeededPaths;
    QSet<QString> failedPaths;

    // Group edits by filePath — replaying them one cell at a time would
    // issue N full saves. Group, save each FileResult once, count
    // successes per queued edit.
    QSet<QString> processedFiles;
    for (const PendingEdit& edit : m_pendingEdits) {
        if (processedFiles.contains(edit.filePath)) continue;
        processedFiles.insert(edit.filePath);

        // Locate the file by path — m_loadedFiles indices may have shifted
        // since the edit was captured.
        int fileIdx = -1;
        for (int i = 0; i < m_loadedFiles.size(); ++i) {
            if (m_loadedFiles[i].filePath == edit.filePath) {
                fileIdx = i;
                break;
            }
        }
        if (fileIdx < 0) {
            // File is no longer loaded — count every queued edit for it as
            // failed and continue.
            for (const PendingEdit& e2 : m_pendingEdits) {
                if (e2.filePath == edit.filePath) ++failed;
            }
            failedPaths.insert(edit.filePath);
            continue;
        }

        const bool savedOk = (m_db && m_db->saveFile(m_loadedFiles[fileIdx]));
        for (const PendingEdit& e2 : m_pendingEdits) {
            if (e2.filePath != edit.filePath) continue;
            if (savedOk) ++succeeded;
            else         ++failed;
        }
        if (savedOk) {
            m_modifiedFilePaths.remove(edit.filePath);
            succeededPaths.insert(edit.filePath);
        } else {
            failedPaths.insert(edit.filePath);
        }
    }

    // H7: retain only edits whose file save failed. Successful files'
    // edits are erased; failures stay queued for the next retry. Use
    // removeIf because m_pendingEdits is a QVector<PendingEdit>.
    m_pendingEdits.erase(
        std::remove_if(m_pendingEdits.begin(), m_pendingEdits.end(),
            [&](const PendingEdit& e) {
                return succeededPaths.contains(e.filePath);
            }),
        m_pendingEdits.end());

    if (m_offlineBanner) {
        m_offlineBanner->setPendingCount(m_pendingEdits.size());
        m_offlineBanner->setPendingFailureCount(failed);
    }

    if (succeeded > 0 || failed > 0) {
        setStatusDb(tr("Flushed pending edits: %1 saved, %2 failed.")
                        .arg(succeeded).arg(failed),
                    failed > 0 ? DbStatusModified : DbStatusOk);
    }
    updateDbSyncIndicator();
}

void MainWindow::closeEvent(QCloseEvent* e)
{
    // H2: drain any debounced Excel cell edits BEFORE the prompt, so the
    // workbook on disk picks up the last burst of changes. The 500 ms
    // m_excelWriteTimer may not have fired yet when the user closes.
    flushExcelWrites();
    if (!promptSaveDatabase()) { e->ignore(); return; }

    // v2.0.2 H8: wipe the session-scoped ImageCache directory. The cache
    // is regenerated lazily from BYTEA on the next launch; persisting it
    // across sessions just bloats %LOCALAPPDATA% and risks stale blobs
    // shadowing a renamed image with the same content hash.
    const QString cacheDir =
        QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation)
        + "/ImageCache";
    if (QDir(cacheDir).exists()) {
        QDir(cacheDir).removeRecursively();
    }

    // Plan C T9: refresh the offline snapshot for next launch. Synchronous —
    // full regen takes a few seconds on a large DB. Best-effort; failure is
    // logged but never blocks close. Skips when offline (the snapshot
    // already represents the last clean state).
    if (m_db && m_db->isOnline() && m_snapshot
        && m_pgConn && m_pgConn->isOpen()) {
        qInfo() << "MainWindow: regenerating offline snapshot before close";
        if (!m_snapshot->regenerate(m_pgConn)) {
            qWarning() << "Snapshot regenerate failed:" << m_snapshot->lastError();
        }
    }

    saveSettings();

    // Restore-last-session: persist the final in-memory state (open files +
    // sessions, including unsaved edits) so the NEXT launch can offer to reopen
    // this session -- regardless of how the app closed. (Previously this wiped
    // the store on a clean close, which made recovery crash-only and meant a
    // normal "Don't save" close lost the session.) flushNow(true) is a no-op
    // when recovery isn't armed (no state provider was set).
    if (m_recovery && m_recoveryArmed) m_recovery->flushNow(true);

    e->accept();
}

// ─── Accessors ────────────────────────────────────────────────────────────────

QString MainWindow::liveColumnForDataCol(int col) const
{
    if (col == DVE::Cols::RESISTANCE) {   // dual-purpose column 4
        const SheetResult* s = currentSheet();
        return (s && s->hasPerRowRegime) ? QStringLiteral("puffing_regime")
                                         : QStringLiteral("resistance");
    }
    return columnNameForDataTableColumn(col);
}

int MainWindow::dataColForLiveColumn(const QString& dbColumn) const
{
    if (dbColumn == QLatin1String("puffing_regime")) return DVE::Cols::RESISTANCE;
    return dataTableColumnForColumnName(dbColumn);
}

QStringList MainWindow::currentFileRegimes() const
{
    const FileResult* f = currentFile();
    return f ? DVE::RegimeUtils::uniqueRegimes(*f) : QStringList();
}

void MainWindow::refreshPlotRegimes()
{
    if (m_plotWidget) m_plotWidget->setAvailableRegimes(currentFileRegimes());
}

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
void MainWindow::updateStatusBar(const QString& msg)
{
    setStatusFile(msg, FileStatusOk);
}
void MainWindow::setProgress(int pct, const QString& msg)
{
    if (m_progressBar) {
        m_progressBar->setVisible(true);
        m_progressBar->setValue(pct);
        if (pct >= 100) m_progressBar->setVisible(false);
    }
    setStatusFile(msg, FileStatusOk);
    QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
}
void MainWindow::showError(const QString& t, const QString& m) { QMessageBox::critical(this, t, m); }
void MainWindow::showInfo(const QString& t, const QString& m)  { QMessageBox::information(this, t, m); }

QString MainWindow::resourcePath() const
{
    QStringList candidates = {
        QCoreApplication::applicationDirPath() + "/resources",
        QCoreApplication::applicationDirPath() + "/../resources"
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
        QCoreApplication::applicationDirPath() + "/resources/templates/" + kTemplateName
    };
    for (const QString& p : candidates)
        if (QFile::exists(p)) return p;
    return QString();
}

QString MainWindow::findPython() const
{
    if (m_pythonProbed) return m_cachedPython;
    m_pythonProbed = true;

    // Check bundled Python first (installed alongside DataViewer.exe)
    const QString bundled = QCoreApplication::applicationDirPath() + "/python/python.exe";
    if (QFile::exists(bundled)) {
        m_cachedPython = bundled;
        return m_cachedPython;
    }

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
    // v2.0.2 M2: use QTemporaryFile for the on-disk script so two
    // concurrent MainWindow instances (or two flushExcelWrites racing
    // through the 500 ms debouncer) can't clobber each other's
    // %TEMP%/dve_script.py. QTemporaryFile generates a unique name
    // (.../dve_script_XXXXXX.py) and removes the file on destruction.
    QTemporaryFile scriptTmp(QStandardPaths::writableLocation(
        QStandardPaths::TempLocation) + "/dve_script_XXXXXX.py");
    scriptTmp.setAutoRemove(true);
    if (!scriptTmp.open()) {
        errOut = "Cannot write temp script: " + scriptTmp.errorString();
        return QString();
    }
    scriptTmp.write(script.toUtf8());
    scriptTmp.flush();
    const QString scriptPath = scriptTmp.fileName();
    scriptTmp.close();   // close handle so QProcess can read it on Windows

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

bool MainWindow::writeCellsToExcel(const QString& filePath, const QString& sheetName,
                                    const QVector<CellWrite>& cells)
{
    if (cells.isEmpty()) return true;        // nothing to do is a success
    const QString python = findPython();
    if (python.isEmpty()) return false;       // no interpreter → cannot persist

    // Batch Python script: writes all cells in one load-save cycle.
    // Arguments after path+sheet are triplets: row col value row col value ...
    //
    // v2.0.2 M3: atomic save. wb.save() writes the entire OOXML zip
    // by truncating the target file first, so a crash mid-write (power
    // loss, process kill, antivirus interruption) leaves the workbook
    // truncated to zero bytes and the user's data unrecoverable. The
    // tmp + os.replace pattern is atomic on both NTFS and POSIX — the
    // original file is unchanged until the move succeeds.
    static const char* kWriteCells = R"PY(
import os
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
tmp = path + ".dve_tmp"
wb.save(tmp)
os.replace(tmp, path)
print("OK")
)PY";

    QStringList args = { filePath, sheetName };
    for (const CellWrite& cw : cells) {
        args << QString::number(cw.row) << QString::number(cw.col) << cw.value;
    }

    QString err;
    const QString out = runPython(python, kWriteCells, args, err);
    // H3: runPython returns the empty string on failure and stuffs the
    // stderr / timeout message into err. The success path prints "OK"
    // from the kWriteCells script; missing OK with empty err still
    // signals a malformed invocation we should not treat as a clean save.
    if (!err.isEmpty()) {
        qWarning() << "writeCellsToExcel failed:" << err;
        return false;
    }
    return out.contains(QLatin1String("OK"));
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

    // H3: keep pending entries when the openpyxl call fails so the next
    // tick (or the next manual save) can retry. The rate-limited warning
    // surfaces the failure to the user without spamming a dialog on every
    // 500 ms tick when the underlying problem is persistent (e.g. file
    // locked by Excel, openpyxl missing from a hand-rolled Python).
    const bool ok = writeCellsToExcel(
        m_pendingWriteFile, m_pendingWriteSheet, m_pendingWrites);
    if (ok) {
        m_pendingWrites.clear();
        m_excelWriteFailureShown = false;
        return;
    }
    if (!m_excelWriteFailureShown) {
        m_excelWriteFailureShown = true;
        QMessageBox::warning(
            this, tr("Excel save failed"),
            tr("Could not write pending changes to %1.\n\n"
               "The edits are still in memory and the database "
               "save will include them. The file on disk is unchanged. "
               "Close any other program that has the workbook open "
               "and try again.")
            .arg(m_pendingWriteFile));
    }
}

QStringList MainWindow::dataTableHeaders(bool perRowRegime)
{
    return {"Puffs","Before (g)","After (g)","Pressure",
            perRowRegime ? "Puffing Regime" : "Resistance",
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

void MainWindow::onLaunchTranslator()
{
    // 1. Resolve the exe path relative to DataViewer.exe location
    QString exePath = QCoreApplication::applicationDirPath()
                      + "/dataviewer_translator/dist/DocumentTranslator.exe";

    if (!QFile::exists(exePath)) {
        QMessageBox::information(this, "Document Translator Not Installed",
            "The Document Translator is not installed.\n\n"
            "Re-run the DataViewer installer and select the "
            "\"Document Translator\" component.");
        return;
    }

    // 2. Get API key — from file if available, otherwise prompt
    // 2. Get API key — from registry if available, otherwise prompt
    QString apiKey = loadApiKey();

    if (apiKey.isEmpty()) {
        bool ok = false;
        apiKey = QInputDialog::getText(
            this,
            "Anthropic API Key Required",
            "Enter your Anthropic API key to use the Document Translator:",
            QLineEdit::Normal, QString(), &ok);
        if (!ok || apiKey.trimmed().isEmpty())
            return;
        apiKey = apiKey.trimmed();
        saveApiKey(apiKey);
    }

    // 3. Write translator config
    if (!writeTranslatorConfig(apiKey)) {
        QMessageBox::critical(this, "Translator Config Failed",
            "Could not write the translator configuration file.\n"
            "Check that your home directory is writable.");
        return;
    }

    // 4. Launch — fire and forget
    if (!QProcess::startDetached(exePath)) {
        QMessageBox::warning(this, "Launch Failed",
            "Could not launch Document Translator.\n"
            "Path: " + exePath);
    }
}

// ── Document Translator ──────────────────────────────────────────────────────
bool MainWindow::writeTranslatorConfig(const QString& apiKey)
{
    QString configDir  = QDir::homePath() + "/.document_translator";
    QString configPath = configDir + "/config.dat";

    if (!QDir(configDir).mkpath("."))
        return false;

    // v2.0.3 — write the same format the translator's APIKeyManager
    // expects: {"api_key": "<base64-of-plaintext>"}. The previous build
    // XOR-obfuscated the key before base64-encoding it; the translator
    // has no matching XOR step, so it decoded garbage, the Anthropic
    // SDK accepted the garbage string at construction time without
    // erroring, and the FIRST API request hung at auth. End-user
    // symptom: "translator says key is configured, then hangs."
    // Persistent storage on the C++ side (QSettings via saveApiKey)
    // still uses the XOR-obfuscated form; only the translator-facing
    // file is plain base64.
    QJsonObject obj;
    obj["api_key"] = QLatin1String(apiKey.toUtf8().toBase64());
    QJsonDocument doc(obj);

    QFile file(configPath);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Truncate))
        return false;

    file.write(doc.toJson(QJsonDocument::Compact));
    return true;
}

// ── API Key Storage ─────────────────────────────────────────────────────────
QString MainWindow::loadApiKey()
{
    // Migration: check old plaintext file first
    QString keyFilePath = QCoreApplication::applicationDirPath()
                          + "/anthropic_api_key.txt";
    QFile keyFile(keyFilePath);
    if (keyFile.open(QIODevice::ReadOnly | QIODevice::Text)) {
        QTextStream in(&keyFile);
        QString key = in.readAll().trimmed();
        keyFile.close();
        if (!key.isEmpty()) {
            // Migrate: save to registry, delete plaintext file
            saveApiKey(key);
            keyFile.remove();
            return key;
        }
    }

    // Load from QSettings (stored in Windows registry under HKCU)
    QSettings settings("SDR", "DataViewer Enterprise");
    QByteArray stored = settings.value("translator/apiKey").toByteArray();
    if (stored.isEmpty()) return {};

    // De-obfuscate (XOR with fixed key)
    const QByteArray mask = QByteArrayLiteral("DVE_2026_TRANSLATOR");
    QByteArray decoded = QByteArray::fromBase64(stored);
    for (int i = 0; i < decoded.size(); ++i)
        decoded[i] = decoded[i] ^ mask[i % mask.size()];
    return QString::fromUtf8(decoded);
}

bool MainWindow::saveApiKey(const QString& key)
{
    const QByteArray raw = key.toUtf8();
    const QByteArray mask = QByteArrayLiteral("DVE_2026_TRANSLATOR");
    QByteArray obfuscated = raw;
    for (int i = 0; i < obfuscated.size(); ++i)
        obfuscated[i] = obfuscated[i] ^ mask[i % mask.size()];

    QSettings settings("SDR", "DataViewer Enterprise");
    settings.setValue("translator/apiKey", obfuscated.toBase64());
    return true;
}

} // namespace DVE
