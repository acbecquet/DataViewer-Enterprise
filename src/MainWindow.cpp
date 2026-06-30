#include "MainWindow.h"
#include "widgets/ScrollHost.h"

// DATAVIEWER-13 (MS-5) extras:
#include <QSettings>
#include <QKeySequence>
#include <QKeyEvent>
#include <QDialog>
#include "utils/AppTheme.h"
#include "utils/OutputPaths.h"
#include "utils/ResponsiveLayout.h"
#include "pipeline/SheetProcessors.h"
#include "pipeline/DataCleanup.h"
#include "ui/SensoryPanel.h"
#include "ui/DetailedSensoryPanel.h"
#include "ui/TesterRound.h"   // v2.5.0 RC5: splitTesterRound for close-dialog "what's missing"
#include "ui/SopDialog.h"
#include "ui/RecoverDialog.h"
#include "database/IdentityManager.h"
#include "database/IdentityPromptDialog.h"
#include "database/ConfigLoader.h"
#include "database/PostgresConnection.h"
#include "database/NotificationListener.h"
#include "database/PresenceManager.h"
#include "database/LiveSync.h"
#include "database/WriteOutcome.h"
#include "database/VersionLookup.h"
#include "database/OfflineSnapshot.h"
#include "database/ConnectionMonitor.h"
#include "widgets/PresenceDotsDelegate.h"
#include "widgets/PresenceAvatarBar.h"
#include "widgets/RowDeletedBanner.h"
#include "widgets/OfflineBanner.h"
#include "widgets/IncompleteDataBanner.h"
#include "widgets/NotesStoryPanel.h"
#include "pipeline/RegimeUtils.h"

#include <QApplication>
#include <QFileDialog>
#include <QMessageBox>
#include <QInputDialog>
#include <QSettings>
#include <QTimer>
#include <QThread>
#include <QEventLoop>
#include <QCloseEvent>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QMimeData>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QDialog>
#include <QDialogButtonBox>
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

// ── v2.4.16 progress-feedback helper ─────────────────────────────────────────
// Owner directive: any operation that blocks the UI must show a progress bar from
// the START -- never a frozen window with no feedback.

// Build + show a modal, cancel-less progress dialog immediately. minimumDuration 0
// makes it paint before the first heavy step (no frozen pre-bar gap). maximum==0
// gives a busy/indeterminate bar; >0 a determinate one. Caller owns the pointer.
QProgressDialog* makeBusyDialog(QWidget* parent, const QString& label, int maximum)
{
    auto* d = new QProgressDialog(label, QString(), 0, maximum, parent);
    d->setWindowModality(Qt::WindowModal);
    d->setCancelButton(nullptr);
    d->setMinimumDuration(0);
    d->setWindowTitle(QObject::tr("Please wait"));
    d->setValue(0);
    d->show();
    QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
    return d;
}

// v2.7.0: each m_centralStack page is a ScrollHost wrapping an inner widget
// (the TPM splitter / SensoryPanel / DetailedSensoryPanel). setCurrentWidget()
// needs the *page* (the ScrollHost), so resolve the inner widget up to whatever
// widget is the stack's direct child. Returns `inner` unchanged if it is not
// wrapped (defensive -- keeps old behaviour if a page is ever added unwrapped).
QWidget* stackPageFor(QStackedWidget* stack, QWidget* inner) {
    QWidget* w = inner;
    while (w && w->parentWidget() && stack->indexOf(w) < 0) {
        w = w->parentWidget();
    }
    return (w && stack->indexOf(w) >= 0) ? w : inner;
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
    // v2.7.0: a 480x360 floor lets the window corner-snap (quarter screen) and
    // half-split on common monitors. The per-region ScrollHosts guarantee that
    // anything that no longer fits scrolls into reach rather than clipping.
    setMinimumSize(480, 360);
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
        // Best-effort; absence is fine on first run. v2.4.4 R5: at boot we don't
        // yet know if PG is reachable, so a decode failure here is NOT shown as
        // a popup (it'd be a false alarm when online). We log it loudly; the
        // user-facing one-time warning fires from onConnectionWentOffline()
        // where the unreadable cache actually matters.
        if (!m_snapshot->openReadOnly() && m_snapshot->lastOpenWasDecodeFailure()) {
            qWarning().noquote()
                << "MainWindow: offline snapshot present but undecodable at boot"
                << "(MIP-encrypted?):" << m_snapshot->lastError();
        }

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

            // ── SP4.5 Stage 2a: background persist + debounced snapshot regen ──
            // Capture the cfg for the workers' own connections, then construct +
            // start both worker threads. Mirrors the LiveSync worker lifecycle
            // EXACTLY: moveToThread -> start the thread -> invokeMethod("start").
            // There is deliberately NO connect(thread, &QThread::started, ...),
            // which would double-open the worker's connection.
            m_dbConfig = cfg;
            qRegisterMetaType<DVE::PersistJob>("DVE::PersistJob");

            m_persistThread = new QThread(this);
            m_persistWorker = new DVE::PersistWorker(cfg);   // no parent (moveToThread)
            m_persistWorker->moveToThread(m_persistThread);
            connect(m_persistThread, &QThread::finished,
                    m_persistWorker, &QObject::deleteLater);
            connect(m_persistWorker, &DVE::PersistWorker::persistFinished,
                    this, &MainWindow::onPersistFinished, Qt::QueuedConnection);
            m_persistThread->start();
            QMetaObject::invokeMethod(m_persistWorker, "start", Qt::QueuedConnection);

            // The regen worker writes the SAME prod path OfflineSnapshot::path()
            // resolves to, so the close-time isCurrentVsLive() reads what the
            // worker wrote (no byte-mismatch risk).
            const QString snapPath = m_snapshot ? m_snapshot->path()
                : (QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation)
                   + QStringLiteral("/snapshot.sqlite"));
            m_regenThread = new QThread(this);
            m_regenWorker = new DVE::SnapshotRegenWorker(cfg, snapPath);  // no parent
            m_regenWorker->moveToThread(m_regenThread);
            connect(m_regenThread, &QThread::finished,
                    m_regenWorker, &QObject::deleteLater);
            connect(m_regenWorker, &DVE::SnapshotRegenWorker::regenFinished,
                    this, &MainWindow::onSnapshotRegenFinished, Qt::QueuedConnection);
            m_regenThread->start();
            QMetaObject::invokeMethod(m_regenWorker, "start", Qt::QueuedConnection);

            // Debounce: coalesce many writes into one regen 30 s after the last.
            m_snapshotRegenTimer = new QTimer(this);
            m_snapshotRegenTimer->setSingleShot(true);
            m_snapshotRegenTimer->setInterval(30000);
            connect(m_snapshotRegenTimer, &QTimer::timeout,
                    this, &MainWindow::onSnapshotRegenRequired);

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
                            // 1. Dispatch the row-deleted toast for the open
                            //    resource (T19). The TPM data-table cell
                            //    decoration (T18) is gone with the table.
                            handleRemoteRowChange(c);
                            // 2. Forward to LiveSync so it can keep its own
                            //    per-cell state in sync with remote changes.
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
                // v2.5.0 Task 4 (RC3): surface per-cell sync failures. The
                // worker now reconnect-and-retries broken connections; only
                // edits that STILL fail bump this count. Reflect it in the DB
                // sync indicator so the user knows an edit isn't live yet (it
                // is re-persisted on the next whole-session save).
                connect(m_liveSync, &DVE::LiveSync::unsyncedEditsChanged, this,
                        [this](int count) {
                            m_unsyncedEdits = count;
                            updateDbSyncIndicator();
                            // SP4.5 Stage 2a: a drain to 0 means per-cell edits
                            // landed in Postgres -> schedule a debounced regen.
                            if (count == 0 && m_snapshotRegenTimer)
                                m_snapshotRegenTimer->start();
                        });
                // v2.4.4 R5: an offline edit that couldn't even be written to
                // the local pending-edits queue (e.g. a MIP-encrypted queue
                // file the python fallback couldn't decode). The unsynced COUNT
                // is already reflected by the unsyncedEditsChanged handler above
                // (LiveSync bumps the single-owner tally on this failure), so we
                // must NOT add to m_unsyncedEdits here — doing so was the old
                // quadratic over-count that fought the sibling handler. This
                // handler's sole job is the ONE-TIME, non-blocking warning so
                // the user knows the local backup queue is degraded.
                connect(m_liveSync, &DVE::LiveSync::offlineEnqueueFailed, this,
                        [this]() {
                            if (m_offlineEnqueueWarningShown)
                                return;
                            m_offlineEnqueueWarningShown = true;
                            QMessageBox* box = new QMessageBox(
                                QMessageBox::Warning,
                                tr("Offline Edits Not Saved Locally"),
                                tr("DataViewer is offline and could not save "
                                   "your recent edit to the local backup "
                                   "queue. The local cache file may be "
                                   "encrypted at rest (Microsoft Information "
                                   "Protection).\n\n"
                                   "Your edit is kept in the open file/session "
                                   "and will be written to the database the next "
                                   "time you save (Ctrl+U) while online. Save "
                                   "soon to avoid losing work if the application "
                                   "closes unexpectedly."),
                                QMessageBox::Ok, this);
                            box->setAttribute(Qt::WA_DeleteOnClose);
                            box->setModal(false);   // non-blocking
                            box->show();
                        });
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
        const bool compactOrNarrower =
            (bp == DVE::ResponsiveLayout::Compact ||
             bp == DVE::ResponsiveLayout::VeryNarrow);
        if (m_ribbon) m_ribbon->setCompactMode(compactOrNarrower);
        setStatusBreadcrumb(m_lastBreadcrumbSegments);
        if (m_sidebarStack) {
            m_sidebarStack->setCurrentIndex(compactOrNarrower ? 1 : 0);
            // Dock min/max width: 32 px strip in compact, normal range otherwise.
            m_fileDock->setMinimumWidth(compactOrNarrower ? 32  : 220);
            // Non-compact: no max cap, so the floated dock can be re-docked
            // (see setupDockPanels).  Compact: 32px max collapses the icon strip.
            m_fileDock->setMaximumWidth(compactOrNarrower ? 32  : QWIDGETSIZE_MAX);
        }
        // VeryNarrow (<760): also fully collapse both side docks so the central
        // ScrollHost gets maximum room before it has to scroll.
        applyVeryNarrowDockState(bp == DVE::ResponsiveLayout::VeryNarrow);
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

    // Debounce timer for the Notes-panel plot re-render. A qualitative cell edit
    // writes the data + LiveSync + Excel immediately, but the heavy full plot
    // rebuild/render is COALESCED here so rapid edits / focus-out storms don't
    // fire N synchronous renders (the source of the intermittent "not responding"
    // freeze). The data model stays correct; only the redraw is deferred ~150ms.
    m_storyPlotTimer = new QTimer(this);
    m_storyPlotTimer->setSingleShot(true);
    m_storyPlotTimer->setInterval(150);
    connect(m_storyPlotTimer, &QTimer::timeout, this, [this]() {
        SheetResult* sheet = currentSheet();
        if (!sheet || sheet->samples.isEmpty()) return;
        recalculateSampleMetrics(*sheet);
        if (currentSheetHasCleanup())
            m_plotWidget->setSheetData(buildCleanedSheet(*sheet, m_currentFileIndex, m_currentSheetIndex));
        else
            m_plotWidget->setSheetData(*sheet);
        if (sheet->hasPerRowRegime && m_storyRegimeDirty) refreshPlotRegimes();
        m_storyRegimeDirty = false;
    });

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
    // This runs regardless of m_recoveryArmed: it only reads the Recovery_prev/
    // store, which is empty unless adoptPreviousSession() moved a non-empty store
    // there. (When recovery wasn't armed, the crash data is still in the *live*
    // dir, not _prev, so the read is empty and we don't double-offer.)
    QTimer::singleShot(0, this, &MainWindow::maybeOfferRecovery);

    updateStatusBar("Ready");
}

MainWindow::~MainWindow()
{
    // SP4.5 Stage 2a: stop the background workers before MainWindow's QObject
    // children (incl. the QThreads + m_db/m_pgConn) are torn down. The persist
    // worker's stop() is a fast connection-close (BlockingQueuedConnection ok);
    // the regen worker's requestRegen() is a BLOCKING slot, so we cancel() (atomic)
    // + quit()/wait() rather than BlockingQueuedConnection (which would freeze the
    // UI for the whole regen). closeEvent already drained/cancelled these on a
    // normal close; this is the backstop for any other teardown path.
    if (m_persistThread) {
        if (m_persistWorker)
            QMetaObject::invokeMethod(m_persistWorker, "stop", Qt::BlockingQueuedConnection);
        m_persistThread->quit();
        m_persistThread->wait(5000);
    }
    if (m_regenThread) {
        if (m_regenWorker) {
            m_regenWorker->cancel();
            QMetaObject::invokeMethod(m_regenWorker, "stop", Qt::QueuedConnection);
        }
        m_regenThread->quit();
        m_regenThread->wait(15000);
    }

    // SP3-T4 (R6): never let the app tear down with an Excel write abandoned on
    // a background thread, and never destroy the watcher while its future runs.
    finishExcelWritesBlocking();
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

        // DATAVIEWER-4: persist before removing, exactly like a scoped program
        // close. Only the sessions that saved are removed; ones blocked by a
        // name clash / hard error stay open so nothing is silently dropped.
        const QVector<int> failed = saveSensorySessionsBeforeClose(indices);
        QVector<int> toClose;
        for (int i : indices) if (!failed.contains(i)) toClose.append(i);
        if (toClose.isEmpty()) { updateImageButton(); return; }

        m_sensoryPanel->closeSessions(toClose);
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

        // DATAVIEWER-4: persist before removing, exactly like a scoped program
        // close. Only the sessions that saved are removed; ones blocked by a
        // name clash / hard error stay open so nothing is silently dropped.
        const QVector<int> failed = saveDetailedSensorySessionsBeforeClose(indices);
        QVector<int> toClose;
        for (int i : indices) if (!failed.contains(i)) toClose.append(i);
        if (toClose.isEmpty()) { updateImageButton(); return; }

        m_detailedSensoryPanel->closeSessions(toClose);
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

    m_undoAllCleanupBtn = m_cleanupGroup->addLargeButton("Undo All",
        AppTheme::icon("rotate-ccw"),
        "Remove ALL data exclusions across every open file");
    m_undoAllCleanupBtn->setEnabled(false);

    connect(cleanBtn,          &QToolButton::clicked, this, &MainWindow::onCleanData);
    connect(m_resetCleanupBtn, &QToolButton::clicked, this, &MainWindow::onResetCleanup);
    connect(m_undoAllCleanupBtn, &QToolButton::clicked, this, &MainWindow::onUndoAllCleanup);
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

    // DV-17: open the active item's SOURCE Excel (TPM: the loaded file; Sensory:
    // the session's sourceFilePath). Disabled in Detailed Sensory (no source
    // path); the no-source-file case is handled in onViewRawData with a message.
    auto* dataGrp = tab->addGroup("Data");
    m_viewRawDataBtn = dataGrp->addLargeButton("View Raw\nData",
        AppTheme::icon("file-text"),
        "Open the source Excel file for the current file/session");
    connect(m_viewRawDataBtn, &QToolButton::clicked, this, &MainWindow::onViewRawData);
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
        { QStringLiteral("Detailed Sensory Path"),
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

    // DATAVIEWER-13 (MS-5): rebind the Sensory stopwatch hotkey. The button label
    // stays short; the current key shows in the tooltip (updated on rebind).
    auto swTip = []() -> QString {
        QSettings s(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"));
        const int k = s.value(QStringLiteral("sensory/stopwatchHotkey"),
                              int(Qt::Key_Space)).toInt();
        return QStringLiteral("Current: %1  ·  click to rebind")
            .arg(QKeySequence(k).toString());
    };
    RibbonGroup* swGrp = tab->addGroup(QStringLiteral("Sensory"));
    QToolButton* swBtn = swGrp->addLargeButton(QStringLiteral("Stopwatch\nHotkey"),
                                               AppTheme::icon(QStringLiteral("keyboard")), swTip());
    connect(swBtn, &QToolButton::clicked, this, [this, swBtn, swTip]() {
        const int key = captureStopwatchKey();
        if (key == 0) return;   // cancelled
        QSettings(QStringLiteral("SDR"), QStringLiteral("DataViewerEnterprise"))
            .setValue(QStringLiteral("sensory/stopwatchHotkey"), key);
        swBtn->setToolTip(swTip());
        if (m_sensoryPanel) m_sensoryPanel->reloadStopwatchHotkey();
    });

    // Panel/dock master reset — un-floats and re-docks the Navigator + Notes
    // panels to their defaults. The escape hatch for a panel that floated off
    // and won't drag back, or a Navigator that restored hidden/off-screen.
    RibbonGroup* layoutGrp = tab->addGroup(QStringLiteral("Panels"));
    QToolButton* resetBtn = layoutGrp->addLargeButton(
        QStringLiteral("Reset\nPanels"),
        AppTheme::icon(QStringLiteral("rotate-ccw")),
        QStringLiteral("Restore the Navigator and Notes panels to their "
                       "default docked positions"));
    connect(resetBtn, &QToolButton::clicked, this, &MainWindow::resetPanelLayout);
}

namespace {
// DATAVIEWER-13 (MS-5): captures the live (non-modifier) key press into *out and
// echoes its name into the dialog's display label so the user can SEE which key
// registered before committing. Does NOT auto-close -- the user confirms via the
// "Set hotkey" button (mouse), which keeps a bindable Space/Enter from doubling as
// the confirm keystroke. Esc falls through so the dialog's standard reject closes.
class StopwatchKeyCatcher : public QObject {
public:
    StopwatchKeyCatcher(int* out, QLabel* display, QAbstractButton* okBtn)
        : m_out(out), m_display(display), m_ok(okBtn) {}
protected:
    bool eventFilter(QObject* o, QEvent* e) override {
        if (e->type() == QEvent::KeyPress) {
            const int k = static_cast<QKeyEvent*>(e)->key();
            if (k == Qt::Key_Shift || k == Qt::Key_Control ||
                k == Qt::Key_Alt   || k == Qt::Key_Meta) return true;  // ignore bare modifiers
            if (k == Qt::Key_Escape) return false;                     // let the dialog reject
            *m_out = k;
            m_display->setText(QStringLiteral("Selected:  %1").arg(QKeySequence(k).toString()));
            m_display->setStyleSheet(QStringLiteral(
                "font-size: 15pt; font-weight: 600; padding: 10px; color: #0066CC;"
                " border: 1px solid #0066CC; border-radius: 4px;"));
            m_ok->setEnabled(true);
            return true;
        }
        return QObject::eventFilter(o, e);
    }
private:
    int*             m_out;
    QLabel*          m_display;
    QAbstractButton* m_ok;
};
} // namespace

int MainWindow::captureStopwatchKey()
{
    QDialog dlg(this);
    dlg.setWindowTitle(QStringLiteral("Set stopwatch hotkey"));
    auto* lay = new QVBoxLayout(&dlg);
    lay->addWidget(new QLabel(QStringLiteral(
        "Press the key to use for the Sensory stopwatch start/stop,"
        "\nthen click “Set hotkey” to confirm."), &dlg));

    // Live feedback: echoes the captured key name so the user can verify it
    // registered the right key before committing.
    auto* display = new QLabel(QStringLiteral("Waiting for a key…"), &dlg);
    display->setAlignment(Qt::AlignCenter);
    display->setStyleSheet(QStringLiteral(
        "font-size: 15pt; font-weight: 600; padding: 10px; color: #888;"
        " border: 1px solid #BBB; border-radius: 4px;"));
    lay->addWidget(display);

    auto* buttons = new QDialogButtonBox(
        QDialogButtonBox::Ok | QDialogButtonBox::Cancel, &dlg);
    buttons->button(QDialogButtonBox::Ok)->setText(QStringLiteral("Set hotkey"));
    buttons->button(QDialogButtonBox::Ok)->setEnabled(false);
    // Mouse-only confirm: NoFocus keeps key events flowing to the dialog (so the
    // event filter sees every keystroke), and stops a bindable Space/Enter from
    // doubling as the button's activation key.
    buttons->button(QDialogButtonBox::Ok)->setFocusPolicy(Qt::NoFocus);
    buttons->button(QDialogButtonBox::Cancel)->setFocusPolicy(Qt::NoFocus);
    lay->addWidget(buttons);
    connect(buttons, &QDialogButtonBox::accepted, &dlg, &QDialog::accept);
    connect(buttons, &QDialogButtonBox::rejected, &dlg, &QDialog::reject);

    int captured = 0;
    StopwatchKeyCatcher catcher(&captured, display, buttons->button(QDialogButtonBox::Ok));
    dlg.installEventFilter(&catcher);
    dlg.resize(360, 150);
    return (dlg.exec() == QDialog::Accepted) ? captured : 0;
}

void MainWindow::setupCentralWidget()
{
    // Horizontal splitter: plot panel on the left, editable story panel on the right
    m_centralSplitter = new QSplitter(Qt::Horizontal, this);

    // ── Sample-nav bar ─────────────────────────────────────────────────────────
    // Nav bar: … | Prev | count | Next | … (centered)
    // Styled frame with accent-subtle background for visual hierarchy. It sits
    // atop the editable story panel in the Notes dock (built below).
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
    m_prevBtn->setMinimumSize(28, AppTheme::controlHeight());
    m_nextBtn->setMinimumSize(28, AppTheme::controlHeight());
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

    // Center the prev/count/next group with stretches on both sides.
    navHL->addStretch();
    navHL->addWidget(m_prevBtn);
    navHL->addWidget(m_sampleCountLabel);
    navHL->addWidget(m_nextBtn);
    navHL->addStretch();

    connect(m_prevBtn,      &QPushButton::clicked, this, &MainWindow::onPrevSample);
    connect(m_nextBtn,      &QPushButton::clicked, this, &MainWindow::onNextSample);

    // ── Plot panel (left) ──────────────────────────────────────────────────────
    m_plotWidget = new PlotWidget(this);
    m_plotWidget->setMinimumHeight(200);

    // Notes panel: the sample-nav bar on top of the editable story panel. The
    // plot fills the TPM center alone; the notes panel lives in a floatable
    // right dock (m_notesDock) that mirrors the left Navigator dock exactly,
    // created below.
    m_storyPanel = new NotesStoryPanel(this);
    QWidget* notesPane = new QWidget(this);
    QVBoxLayout* notesVL = new QVBoxLayout(notesPane);
    notesVL->setContentsMargins(0, 0, 0, 0);
    notesVL->setSpacing(4);
    notesVL->addWidget(m_sampleNavBar);
    notesVL->addWidget(m_storyPanel, 1);

    // m_notesDock mirrors m_fileDock exactly: same allowed areas (all four
    // edges, so the user can pin it top/left/right/bottom), same
    // Movable|Floatable features, same 220..320 width range — so it pops out
    // the same way the Navigator does. TPM-only (hidden in sensory/detailed).
    m_notesDock = new QDockWidget("Notes", this);
    m_notesDock->setObjectName(QStringLiteral("notesDock"));
    m_notesDock->setAllowedAreas(Qt::AllDockWidgetAreas);
    m_notesDock->setFeatures(QDockWidget::DockWidgetMovable | QDockWidget::DockWidgetFloatable);
    // v2.7.0: wrap the Notes content so the sample-nav bar + story panel scroll
    // (both directions) instead of clipping when the dock is narrow/short.
    m_notesScrollHost = ScrollHost::wrap(notesPane);
    m_notesDock->setWidget(m_notesScrollHost);
    // No setMaximumWidth: a max size on a QDockWidget blocks re-docking after
    // float (see setupDockPanels).  Width is controlled via resizeDocks().
    m_notesDock->setMinimumWidth(220);
    addDockWidget(Qt::RightDockWidgetArea, m_notesDock);

    // The plot fills the central splitter alone (a one-child splitter is fine;
    // it stays the index-0 page so mode-return setCurrentWidget(m_centralSplitter)
    // lines remain valid).
    m_centralSplitter->addWidget(m_plotWidget);

    connect(m_storyPanel, &NotesStoryPanel::cellEdited,    this, &MainWindow::onStoryCellEdited);
    connect(m_storyPanel, &NotesStoryPanel::noteActivated, this, &MainWindow::onStoryNoteActivated);

    // Wrap in a stacked widget (index 0 = TPM, index 1 = sensory, added lazily).
    // Each page is wrapped in a ScrollHost so it scrolls instead of clipping at
    // small window sizes (v2.7.0). m_centralSplitter stays pointing at the inner
    // splitter; the page added to the stack is its ScrollHost. setCurrentWidget
    // call sites use stackPageFor() to resolve the inner widget to its page.
    m_centralStack = new QStackedWidget(this);
    m_centralStack->addWidget(ScrollHost::wrap(m_centralSplitter));   // index 0

    // Presence avatars dock at the right end of the TPM plot's control row
    // (Plot Type / Regime / Save plot) so they share that row instead of taking
    // their own full-width band above the plot. Presence in sensory/detailed
    // modes still shows via the navigator dots. setPresence()/clear() drive its
    // visibility (hidden when no one — including self — is present).
    m_avatarBar = new DVE::PresenceAvatarBar(this);
    // Hug the right edge of the control row (the row's own stretch pushes it
    // there) instead of expanding to fill — otherwise it would split the
    // remaining width with the stretch.
    m_avatarBar->setSizePolicy(QSizePolicy::Maximum, QSizePolicy::Fixed);
    if (m_plotWidget) m_plotWidget->setHeaderTrailingWidget(m_avatarBar);

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
    // m_avatarBar is NOT added here — it's docked into the plot's top control
    // row via setHeaderTrailingWidget() above, so it no longer occupies its own
    // band between the banners and the central stack.
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
                // v2.5.0 Task-5 review (5b): route the recreate through the
                // auto-suffix wrapper. If another client recreated a row with
                // the same natural key between the delete and this re-INSERT,
                // a plain tryWrite would 23505 and the banner would silently
                // fail; the wrapper self-heals by suffixing (_1/_2/...).
                savedOk = (m_db->tryWriteSensorySessionAutoSuffix(*curr)
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
                // v2.5.0 Task-5 review (5b): twin of the sensory recreate path —
                // auto-suffix so a concurrent recreate collision self-heals
                // instead of failing the banner silently.
                savedOk = (m_db->tryWriteDetailedSensorySessionAutoSuffix(*curr)
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
    // objectName is REQUIRED for QMainWindow::saveState()/restoreState() to
    // persist this dock's geometry/float state across runs — without it Qt
    // silently skips the dock (and warns at runtime).  Mirrors m_notesDock.
    m_fileDock->setObjectName(QStringLiteral("navigatorDock"));
    // All four edges allowed so the Navigator can be pinned top/left/right/
    // bottom — the Notes dock mirrors this for symmetric pop-out behaviour.
    m_fileDock->setAllowedAreas(Qt::AllDockWidgetAreas);
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
    avgHeader->setMinimumHeight(AppTheme::controlHeight());
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
    propHeader->setMinimumHeight(AppTheme::controlHeight());
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
    imgBar->setMinimumHeight(40);

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

    // v2.7.0: wrap the whole sidebar stack so the Navigator + Test Averages +
    // Properties + image-button column scroll instead of clipping at small dock
    // heights. The 32px icon-strip page never overflows, so compact mode is
    // unaffected.
    m_navScrollHost = ScrollHost::wrap(m_sidebarStack);
    m_fileDock->setWidget(m_navScrollHost);
    addDockWidget(Qt::LeftDockWidgetArea, m_fileDock);
    // NB: do NOT cap the dock with setMaximumWidth — a maximum size on a
    // QDockWidget prevents Qt from re-docking it after it has been floated
    // (the layout can't build the drop-gap).  Width is controlled via
    // resizeDocks()/restoreState() instead.  The compact handler still uses a
    // 32px max to collapse to an icon strip, which is intentional in that mode.
    m_fileDock->setMinimumWidth(32);
    m_fileDock->setMaximumWidth(QWIDGETSIZE_MAX);
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
    // DATAVIEWER-4: pass flushPending=true via lambda. A method-pointer connect
    // would bind QAction::triggered(bool checked=false) to flushPending, so Ctrl+U
    // would NOT flush — the lambda forces the deliberate-save flush path.
    connect(dbUpdateAct, &QAction::triggered, this, [this]() { onUpdateDatabase(true); });
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

// NOTE (post-v2.0.7 cleanup backlog S001a / R-007): onEditHeaders is currently
// UNWIRED — no connect() and no caller route to it (the 'Edit Headers' action was
// never re-added after the v2.0.7 ribbon rework). It is a candidate for removal in
// a future cleanup pass; the delete-vs-keep decision is tracked there, not here.
// Until then it is hardened so that re-wiring 'Edit Headers' cannot reintroduce the
// concurrent-writer corruption SP3-T4 fixed: (1) it drains the async flush first,
// exactly like onAddRow/onRemoveRow, and (2) its kWriteHeaders script saves
// atomically (tmp + os.replace), not in place.
void MainWindow::onEditHeaders()
{
    const FileResult*  file  = currentFile();
    const SheetResult* sheet = currentSheet();
    if (!file || !sheet || sheet->samples.isEmpty()) {
        showInfo("No Sample", "Load a file and select a sample first.");
        return;
    }

    // SP3-T4 (R6 fix): single-writer invariant — same rationale as onAddRow/
    // onRemoveRow. This op ends in a SYNCHRONOUS kWriteHeaders openpyxl write
    // (below). HeaderEditDialog is modal and dlg.exec() spins the event loop, so
    // the 500 ms m_excelWriteTimer can fire and dispatch a flush worker while the
    // dialog is open. Drain any in-flight/pending async writes synchronously first
    // so this op's synchronous write is the only writer to the workbook. At the
    // OUTER op entry, NOT inside writeCellsToExcel, so the drain cannot recurse.
    finishExcelWritesBlocking();

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

    // v2.4.4 R6: atomic save (tmp + os.replace) — same crash-safety tail as
    // excelWriteCellsScript()/excelDeleteRowScript(). wb.save(path) truncates the
    // target first, so a kill mid-write would leave the source workbook torn. Even
    // though onEditHeaders now drains the async flush first (single-writer), the
    // atomic tail makes a torn workbook impossible regardless of caller.
    static const char* kWriteHeaders = R"PY(
import os, sys, json
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
tmp = d['file_path'] + ".dve_tmp"
wb.save(tmp)
os.replace(tmp, d['file_path'])
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

    if (currentSheetHasCleanup()) {
        const SheetResult cleaned = buildCleanedSheet(*sheet, m_currentFileIndex, m_currentSheetIndex);
        m_plotWidget->setSheetData(cleaned);
    } else {
        m_plotWidget->setSheetData(*sheet);
    }

    // Refresh the notes-story panel so its TPM context reflects the recalc above.
    // Safe here: the user is editing the prop table, not a panel field, so
    // re-populating the panel won't destroy an active editor.
    m_storyPanel->setSample(s, exclusionsFor(m_currentFileIndex, m_currentSheetIndex, m_currentSampleIndex), sheet->hasPerRowRegime);

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

// v2.5.0: robust "same file" test for the open TPM working set. Windows paths are
// case-insensitive and an open can arrive via QFileDialog (forward slashes) or the
// SingleInstance IPC / command line (back slashes), so a raw string compare would
// treat a re-opened file as brand new -> a duplicate in-memory entry AND a second
// DB INSERT (split data). Normalize separators + case before comparing.
static bool isSameLoadedPath(const QString& a, const QString& b)
{
    return QDir::cleanPath(QFileInfo(a).absoluteFilePath())
               .compare(QDir::cleanPath(QFileInfo(b).absoluteFilePath()),
                        Qt::CaseInsensitive) == 0;
}

void MainWindow::routeFile(const QString& path)
{
    // If this exact file is already open in the TPM working set, just switch to it.
    // Reloading would fork a duplicate working-set entry + a split DB row (and could
    // drop unsaved in-memory edits by re-reading disk). Sensory/detailed sessions
    // live in their panels (never in m_loadedFiles), so this only catches TPM files.
    for (int i = 0; i < m_loadedFiles.size(); ++i) {
        if (isSameLoadedPath(m_loadedFiles[i].filePath, path)) {
            if (m_sensoryBtn->isChecked())         m_sensoryBtn->setChecked(false);
            if (m_detailedSensoryBtn->isChecked()) m_detailedSensoryBtn->setChecked(false);
            if (i != m_currentFileIndex) {
                m_currentFileIndex   = i;
                m_currentSheetIndex  = 0;
                m_currentSampleIndex = 0;
                populateFileTree();
                populateSheetCombo();
                displayCurrentSample();
            }
            updateStatusBar("Already open: " + QFileInfo(path).fileName());
            // keep draining any queued multi-select load
            if (!m_pendingLoadPaths.isEmpty())
                routeFile(m_pendingLoadPaths.takeFirst());
            return;
        }
    }

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
    // clicks Close. SP3-T4: finish synchronously so no write is left running
    // off-thread while m_loadedFiles changes underneath it.
    finishExcelWritesBlocking();

    // SP4.5 Stage 2a: a background persist for THIS file may still be in flight.
    // Drain it (and flush the writeback) before we synchronously persist + remove
    // the slot, or the synchronous persistLoadedFile below would re-INSERT with
    // id=-1 and create a duplicate DB row (the v2.4.0 false-duplicate class).
    drainPersistWorkerBlocking(15000);

    // DATAVIEWER-4: Close == a scoped program-close. Drain LiveSync per-cell edits
    // too, then authoritatively persist this file (WriteResult-aware, OCC retry)
    // BEFORE removing it. No "save? Yes/No" prompt — Close always persists; we only
    // ask if the save actually FAILS, so the user can't silently lose data.
    if (m_liveSync && !m_liveSync->flushNowAndWait()) {
        // flushNowAndWait() returns false on EITHER a drain timeout OR a nested
        // re-entrant flush (its re-entrancy guard), so don't claim "timed out".
        qWarning() << "onCloseFile: LiveSync flush did not complete (timeout or "
                      "nested flush); proceeding with authoritative persist (pending="
                   << m_liveSync->pendingCount() << ")";
    }

    const QString fp = m_loadedFiles[m_currentFileIndex].filePath;
    if (m_modifiedFilePaths.contains(fp)) {
        persistLoadedFile(m_currentFileIndex);
        // v2.5.0 RC5: never hard-block. If the save failed/offline, offer
        // Retry (re-attempt persistLoadedFile — transient NAS blips clear),
        // Close Anyway (accept the loss), or Cancel (keep the file open). Loop
        // so a repeated failure re-shows the same options rather than trapping.
        while (m_modifiedFilePaths.contains(fp)) {   // still dirty => not saved
            QMessageBox box(QMessageBox::Warning, tr("Could Not Save File"),
                tr("'%1' could not be saved to the database.\n\n"
                   "Retry the save, close anyway and lose the unsaved database "
                   "changes, or cancel and keep the file open?")
                    .arg(QFileInfo(fp).fileName()),
                QMessageBox::NoButton, this);
            QPushButton* retryBtn  = box.addButton(tr("Retry"),
                                                   QMessageBox::AcceptRole);
            QPushButton* closeBtn  = box.addButton(tr("Close Anyway"),
                                                   QMessageBox::DestructiveRole);
            QPushButton* cancelBtn = box.addButton(QMessageBox::Cancel);
            box.setDefaultButton(retryBtn);
            box.exec();

            QAbstractButton* clicked = box.clickedButton();
            if (clicked == retryBtn) {
                persistLoadedFile(m_currentFileIndex);   // re-attempt; loop re-checks
                continue;
            }
            if (clicked == closeBtn)
                break;                                   // accept the loss, fall through
            // Cancel / [X] / Esc -> keep the file open.
            (void)cancelBtn;
            return;
        }
    }

    m_modifiedFilePaths.remove(m_loadedFiles[m_currentFileIndex].filePath);

    // GAP-B: drop this file's exclusions AND re-key the survivors. Removing the
    // file shifts every higher index down by one, so exclusions keyed by the old
    // fileIdx must shift with them or they'd point at the wrong file.
    m_excludedRows = DataCleanup::rekeyAfterClose(m_excludedRows, m_currentFileIndex);

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
        m_storyPanel->clear();
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

// DATAVIEWER-3: persist a freshly loaded/refreshed TPM file to the database and
// reflect the outcome. Unlike the old fire-and-forget saveFile(), this checks
// the WriteResult: a failed save keeps the file marked dirty (so the Ctrl+U
// batch / close-flush retries it) and surfaces the reason, instead of silently
// dropping it. Uses the mutable tryWriteFile overload so id/version are stamped
// back into m_loadedFiles for presence dots + subsequent saves.
void MainWindow::persistLoadedFile(int fileIndex)
{
    if (fileIndex < 0 || fileIndex >= m_loadedFiles.size())
        return;
    FileResult& fr = m_loadedFiles[fileIndex];

    // No database configured: nothing to persist; mirror the pre-DATAVIEWER-3
    // behavior of treating the file as not-dirty for the (no-op) indicator.
    if (!m_db) {
        m_modifiedFilePaths.remove(fr.filePath);
        updateDbSyncIndicator();
        return;
    }

    DVE::WriteResult r = m_db->tryWriteFile(fr);

    // One-shot optimistic-concurrency recovery: the DB row changed or was
    // deleted since we inherited id/version. Re-inherit from the current row
    // and retry exactly once (no loop).
    if (r == DVE::WriteResult::VersionMismatch || r == DVE::WriteResult::RowDeleted) {
        const FileResult dbRow = m_db->loadFileByPath(fr.filePath);
        if (dbRow.id > 0) {
            fr.id      = dbRow.id;
            fr.version = dbRow.version;
            r = m_db->tryWriteFile(fr);
        }
    }

    const QString name = fr.fileName.isEmpty()
                             ? QFileInfo(fr.filePath).fileName()
                             : fr.fileName;

    switch (DVE::classifyLoadSaveResult(r)) {
    case DVE::LoadSavePolicy::Saved:
        m_modifiedFilePaths.remove(fr.filePath);
        break;
    case DVE::LoadSavePolicy::RetryOffline:
        m_modifiedFilePaths.insert(fr.filePath);
        updateStatusBar(tr("Offline: '%1' was not saved to the database; "
                           "it will be retried when the connection returns.").arg(name));
        qWarning().noquote() << "[persistLoadedFile] offline, not saved:"
                             << name << "-" << m_db->lastError();
        break;
    case DVE::LoadSavePolicy::RetryConflict:
        // The file stays dirty. onUpdateDatabase() simply re-issues tryWriteFile
        // on the next tick / Ctrl+U; that wrapper now recovers internally (adopts
        // fresh versions on a near-unreachable VersionMismatch, re-INSERTs on
        // RowDeleted), so the retry actually writes the user's edits rather than
        // silently clearing the dirty flag.
        m_modifiedFilePaths.insert(fr.filePath);
        updateStatusBar(tr("'%1' was changed by another user and was not saved; "
                           "press Ctrl+U to retry.").arg(name));
        qWarning().noquote() << "[persistLoadedFile] OCC conflict, not saved:"
                             << name << "-" << m_db->lastError();
        break;
    case DVE::LoadSavePolicy::RetryError:
        m_modifiedFilePaths.insert(fr.filePath);
        updateStatusBar(tr("Failed to save '%1' to the database: %2")
                            .arg(name, m_db->lastError()));
        qWarning().noquote() << "[persistLoadedFile] save failed:"
                             << name << "-" << m_db->lastError();
        break;
    }

    updateDbSyncIndicator();
}

// ── SP4.5 Stage 2a: background persist-on-load ──────────────────────────────
// Enqueue the file at fileIndex for a background DB save. The file is already
// displayed; the save runs on the PersistWorker thread. Falls back to the
// synchronous persistLoadedFile when there is no running worker (offline / no DB
// / tests) so behavior there is unchanged.
void MainWindow::enqueuePersist(int fileIndex)
{
    if (fileIndex < 0 || fileIndex >= m_loadedFiles.size()) return;

    if (!m_persistWorker || !m_persistThread || !m_persistThread->isRunning()) {
        persistLoadedFile(fileIndex);
        return;
    }

    FileResult& fr = m_loadedFiles[fileIndex];
    DVE::PersistJob job;
    job.slotIndex  = fileIndex;
    job.filePath   = fr.filePath;
    job.generation = ++m_persistGeneration;
    job.writerUuid = m_identity
                         ? m_identity->uuid().toString(QUuid::WithoutBraces)
                         : QStringLiteral("unknown");
    job.online     = (m_db && m_db->isOnline());   // UI-thread snapshot of online state
    job.snapshot   = fr;                            // value copy at enqueue time

    m_lastEnqueuedGen[fr.filePath] = job.generation;
    m_backgroundSaveInFlight.insert(fr.filePath);   // guard onUpdateDatabase from
    m_dirtiedDuringPersist.remove(fr.filePath);     // racing this into a dup INSERT
    QMetaObject::invokeMethod(m_persistWorker, "enqueue", Qt::QueuedConnection,
                              Q_ARG(DVE::PersistJob, job));
}

// Writes the server-assigned file id/version back into m_loadedFiles (matching
// persistLoadedFile's tryWriteFile contract -- child ids are not cascaded). UI
// thread (QueuedConnection). Locates the slot by FILE PATH (the working set may
// have been reindexed since enqueue) and applies only the LATEST generation per
// file, so a superseded rapid-reload never stamps a stale id over a fresh one.
void MainWindow::onPersistFinished(DVE::PersistJob job, DVE::WriteResult wr)
{
    qInfo() << "[perf] onPersistFinished: file=" << job.filePath
            << "gen=" << job.generation << "result=" << static_cast<int>(wr);

    const quint64 latest   = m_lastEnqueuedGen.value(job.filePath, 0);
    const bool    isLatest = (job.generation >= latest);
    if (isLatest)
        m_backgroundSaveInFlight.remove(job.filePath);  // newest bg save for this file done

    int slot = -1;
    for (int i = 0; i < m_loadedFiles.size(); ++i)
        if (m_loadedFiles[i].filePath == job.filePath) { slot = i; break; }
    if (slot < 0) {
        m_dirtiedDuringPersist.remove(job.filePath);    // file no longer loaded
        qWarning() << "onPersistFinished: file no longer loaded -- discarding writeback"
                   << job.filePath;
        return;
    }
    if (!isLatest) {
        // A newer enqueue is still in flight; it will do the writeback and clear
        // the in-flight guard. Leave m_backgroundSaveInFlight set for that job.
        qWarning() << "onPersistFinished: superseded generation" << job.generation
                   << "<" << latest << "-- discarding writeback";
        return;
    }

    const bool dirtiedDuringPersist = m_dirtiedDuringPersist.contains(job.filePath);
    m_dirtiedDuringPersist.remove(job.filePath);

    FileResult& fr = m_loadedFiles[slot];
    switch (DVE::classifyLoadSaveResult(wr)) {
    case DVE::LoadSavePolicy::Saved:
        fr.id      = job.snapshot.id;
        fr.version = job.snapshot.version;
        // Only mark clean if the file was NOT edited during the background save:
        // the worker wrote the enqueue-time snapshot, so a post-enqueue edit must
        // stay dirty and be re-persisted (now with the real file id) on the next
        // whole-file save, otherwise that edit would be silently dropped.
        if (!dirtiedDuringPersist) m_modifiedFilePaths.remove(fr.filePath);
        if (m_snapshotRegenTimer) m_snapshotRegenTimer->start();  // DB changed
        break;
    case DVE::LoadSavePolicy::RetryOffline:
    case DVE::LoadSavePolicy::RetryConflict:
    case DVE::LoadSavePolicy::RetryError:
        m_modifiedFilePaths.insert(fr.filePath);
        break;
    }
    updateDbSyncIndicator();
    populateFileTree();   // presence dots now have the real file id
}

// ── SP4.5 Stage 2a: close-time persist drain guard ──────────────────────────
// Drain the persist queue (process every queued job) and then flush the queued
// onPersistFinished() slots, so m_modifiedFilePaths is authoritative before the
// save prompt / dirty check reads it. Makes load-then-close impossible to lose a
// write. The re-entrancy guard is set BEFORE the loop so a re-shown close (the
// user cancelled the save prompt, then closes again) cannot stack drains.
void MainWindow::drainPersistWorkerBlocking(int timeoutMs)
{
    if (!m_persistWorker || !m_persistThread || !m_persistThread->isRunning())
        return;
    if (m_persistDraining) return;
    m_persistDraining = true;

    qInfo() << "[perf] draining persist worker";
    QEventLoop drainLoop;
    QMetaObject::Connection c =
        connect(m_persistWorker, &DVE::PersistWorker::drainFinished,
                &drainLoop, &QEventLoop::quit, Qt::QueuedConnection);
    QTimer to;
    to.setSingleShot(true);
    connect(&to, &QTimer::timeout, &drainLoop, [&drainLoop]() {
        qWarning() << "[perf] persist drain TIMEOUT -- an in-flight write may roll "
                      "back; the affected file stays dirty for next launch";
        drainLoop.quit();
    });
    to.start(timeoutMs);
    QMetaObject::invokeMethod(m_persistWorker, "drainAndSignal", Qt::QueuedConnection);
    // ExcludeUserInputEvents so a second close-click can't re-enter closeEvent
    // during the nested loop; cross-thread drainFinished (QueuedConnection) still
    // delivers (it is not a user-input event).
    drainLoop.exec(QEventLoop::ExcludeUserInputEvents);
    disconnect(c);

    // Flush queued onPersistFinished() so m_modifiedFilePaths reflects every
    // just-committed file before the save prompt reads it.
    QCoreApplication::processEvents(QEventLoop::ExcludeUserInputEvents);

    m_persistDraining = false;
    qInfo() << "[perf] persist drain done";
}

// ── SP4.5 Stage 2a: debounced background snapshot regen ─────────────────────
void MainWindow::onSnapshotRegenRequired()
{
    if (!m_regenWorker || !m_regenThread || !m_regenThread->isRunning()) return;
    if (m_db && !m_db->isOnline()) return;     // offline: keep the snapshot as-is
    if (m_snapshotRegenInFlight) return;
    m_snapshotRegenInFlight = true;
    // Release our read-only handle on the prod snapshot so the worker's atomic
    // rename can replace it. SQLite holds the file without FILE_SHARE_DELETE on
    // Windows, so the rename would otherwise fail with a sharing violation. The
    // handle is re-opened in onSnapshotRegenFinished. We are online here, so the
    // offline read path (the only consumer of the open handle) is not in use.
    if (m_snapshot) m_snapshot->close();
    qInfo() << "[perf] dispatching background snapshot regen";
    QMetaObject::invokeMethod(m_regenWorker, "requestRegen", Qt::QueuedConnection);
}

void MainWindow::onSnapshotRegenFinished(bool ok, const QString& error)
{
    m_snapshotRegenInFlight = false;
    // Re-open the read-only view of the (now freshly promoted) prod snapshot so
    // close-time isCurrentVsLive() + any offline read see the new file.
    if (m_snapshot) m_snapshot->openReadOnly();
    if (ok) {
        m_regenFailStreak = 0;
        qInfo() << "[perf] background snapshot regen complete";
    } else {
        qWarning() << "[perf] background snapshot regen failed:" << error;
        // Bounded retry so a sustained NAS outage doesn't spin forever.
        if (++m_regenFailStreak < 3) {
            if (m_snapshotRegenTimer) m_snapshotRegenTimer->start();
        } else {
            qWarning() << "[perf] regen fail-streak hit 3 -- pausing until next write";
            m_regenFailStreak = 0;
        }
    }
}

// ── v2.5.0 RC5: never hard-block an unnamed session on close ─────────────────
// Gather the bits the close dialog needs from a sensory session.
MainWindow::SessionCloseInfo
MainWindow::sessionCloseInfoFor(const SensorySession& s) const
{
    SessionCloseInfo info;
    info.label       = m_sensoryPanel ? m_sensoryPanel->sessionLabel(s)
                                      : QStringLiteral("(unnamed)");
    info.sampleCount = s.samples.size();
    // Name what's missing so the user knows exactly what to fill in.
    const bool noTitle  = s.testTitle.trimmed().isEmpty();
    const bool noTester = DVE::splitTesterRound(s.testerName).tester.trimmed().isEmpty();
    if (noTitle && noTester)      info.missing = tr("a test name and a tester");
    else if (noTitle)             info.missing = tr("a test name");
    else                          info.missing = tr("a tester");
    // An autosaved .xlsx on disk for this panel is left in place; mention it.
    info.hasDiskFile = m_sensoryPanel && m_sensoryPanel->hasSavePath();
    return info;
}

// Twin for a detailed-sensory session.
MainWindow::SessionCloseInfo
MainWindow::sessionCloseInfoFor(const DetailedSensorySession& s) const
{
    SessionCloseInfo info;
    info.label       = m_detailedSensoryPanel ? m_detailedSensoryPanel->sessionLabel(s)
                                              : QStringLiteral("(unnamed)");
    info.sampleCount = s.samples.size();
    const bool noTitle  = s.testTitle.trimmed().isEmpty();
    const bool noTester = s.testerName.trimmed().isEmpty();
    if (noTitle && noTester)      info.missing = tr("a test name and a tester");
    else if (noTitle)             info.missing = tr("a test name");
    else                          info.missing = tr("a tester");
    info.hasDiskFile = m_detailedSensoryPanel && m_detailedSensoryPanel->hasSavePath();
    return info;
}

// One dialog per offending session: Name It Now / Discard Session / Cancel.
// Never an OK-only block — the user always has a way forward. The default is
// "Name It Now" (preserve data); Discard is destructive and must be chosen
// deliberately.
MainWindow::SessionCloseChoice
MainWindow::promptUnnamedSessionOnClose(const SessionCloseInfo& info)
{
    QString body = tr("This session is missing %1, so it can't be saved.\n\n"
                      "Session: %2\nSamples entered: %3\n\n"
                      "What would you like to do?")
                       .arg(info.missing, info.label)
                       .arg(info.sampleCount);
    if (info.hasDiskFile)
        body += tr("\n\n(An auto-saved Excel copy on disk is left untouched.)");

    QMessageBox box(QMessageBox::Question, tr("Session Has No Name"), body,
                    QMessageBox::NoButton, this);
    QPushButton* nameBtn    = box.addButton(tr("Name It Now"),
                                            QMessageBox::ActionRole);
    QPushButton* discardBtn = box.addButton(tr("Discard Session"),
                                            QMessageBox::DestructiveRole);
    QPushButton* cancelBtn  = box.addButton(QMessageBox::Cancel);
    box.setDefaultButton(nameBtn);
    box.exec();

    QAbstractButton* clicked = box.clickedButton();
    if (clicked == discardBtn) return SessionCloseChoice::Discard;
    if (clicked == nameBtn)    return SessionCloseChoice::NameIt;
    // Cancel, [X], or Esc -> keep the session open, change nothing.
    (void)cancelBtn;
    return SessionCloseChoice::Cancel;
}

// DATAVIEWER-4: persist the given sensory sessions (panel indices) before Close
// removes them. Mirrors onUpdateDatabase's per-session save EXACTLY — rename ->
// force-INSERT (old row preserved), UniqueViolation -> skip + inform, Success /
// VersionMismatch / RowDeleted treated as done, anything else kept open — but
// scoped to the closing set and quiet on success. Returns the indices that
// FAILED to save so the caller leaves those sessions open.
QVector<int> MainWindow::saveSensorySessionsBeforeClose(const QVector<int>& indices)
{
    QVector<int> failed;
    if (!m_sensoryPanel || !m_db) return failed;
    if (m_liveSync && !m_liveSync->flushNowAndWait()) {   // our scores -> DB first
        qWarning() << "saveSensorySessionsBeforeClose: LiveSync flush did not "
                      "complete (timeout or nested flush); proceeding (dirty-aware "
                      "merge keeps local edits, pending="
                   << m_liveSync->pendingCount() << ")";
    }

    // v2.5.0 RC5: indices the user chose to discard (their DB rows are removed
    // here; the caller's closeSessions() drops them from the panel because they
    // are deliberately left OUT of `failed`). Tracked so syncSavedSessionState
    // below skips them — they no longer exist as far as the user is concerned.
    QVector<int> discarded;
    bool aborted = false;   // user hit Name-It-Now / Cancel -> stop processing

    QVector<SensorySession> sessions = m_sensoryPanel->allSessions();  // flushes widgets
    for (int idx : indices) {
        if (aborted) { failed.append(idx); continue; }   // keep the rest open
        if (idx < 0 || idx >= sessions.size()) continue;
        SensorySession& sess = sessions[idx];
        if (DVE::isPlaceholderSession(sess)) continue;
        // v2.5.0 RC5: a non-empty session with no test name/tester can't be
        // keyed. Never hard-block — offer Name It Now / Discard / Cancel.
        if (!DVE::isSensorySessionSavable(sess)) {
            const SessionCloseChoice choice =
                promptUnnamedSessionOnClose(sessionCloseInfoFor(sess));
            switch (choice) {
            case SessionCloseChoice::NameIt:
                // Land the user on the Test Title field for this session and
                // abort the close so they can fill it in. Keep this session and
                // any remaining ones open.
                m_sensoryPanel->focusTitleForSession(idx);
                failed.append(idx);
                aborted = true;
                continue;
            case SessionCloseChoice::Discard:
                // Delete the DATA: remove the DB row (cascades to images) when
                // one exists; if the delete fails (offline) discard locally
                // anyway so the user is never trapped. Disk autosave files are
                // left untouched (mentioned in the dialog). Left OUT of `failed`
                // so the caller's closeSessions() removes it from the panel.
                if (sess.id > 0 && !m_db->removeSensorySession(sess.id))
                    qWarning().noquote()
                        << "[saveSensorySessionsBeforeClose] DB delete failed for"
                        << "discarded session" << sess.sessionName << "-"
                        << m_db->lastError() << "(discarding locally anyway)";
                discarded.append(idx);
                continue;
            case SessionCloseChoice::Cancel:
                // Abort the entire close; keep this and all remaining sessions.
                failed.append(idx);
                aborted = true;
                continue;
            }
            continue;
        }

        const bool isRename = sess.id > 0 && !sess.originalSessionName.isEmpty()
                              && sess.originalSessionName != sess.sessionName;
        if (isRename) { sess.id = -1; sess.version = 0; }   // preserve old row

        const QString preName = sess.sessionName;
        // v2.5.0 RC4: a name collision at close auto-suffixes (_1/_2/...) and
        // saves rather than blocking the close with a modal — never hard-block.
        DVE::WriteResult r = m_db->tryWriteSensorySessionAutoSuffix(sess);
        if (r != DVE::WriteResult::Success) {
            // RC1: VersionMismatch / RowDeleted used to be treated as "already
            // saved by LiveSync" and let the session close. The wrapper now
            // re-INSERTs on RowDeleted and adopts fresh versions, so any
            // non-Success here is a genuine failure -> keep the session open.
            failed.append(idx);
        } else {
            if (sess.sessionName != preName) {
                qInfo().noquote() << "[saveSensorySessionsBeforeClose] name taken —"
                                  << "auto-renamed" << preName << "→" << sess.sessionName;
                updateStatusBar(
                    tr("Session name was taken — saved as \"%1\"").arg(sess.sessionName));
            }
            // v2.5.0 Task 3 (RC2 review, CRITICAL 2): Success — the edits are in
            // the DB blob, so clear the dirty set on this local copy. The adopt
            // in syncSavedSessionState() then propagates it; failed sessions keep
            // their dirty cells (the adopt leaves them, the merge protects them).
            sess.dirtyCells.clear();
        }
    }
    m_sensoryPanel->syncSavedSessionState(sessions);
    refreshSensoryNavigator();   // v2.5.0 RC4: reflect any auto-suffix in the list
    updateDbSyncIndicator();
    if (!discarded.isEmpty())
        updateStatusBar(tr("Discarded %1 unnamed session%2")
                            .arg(discarded.size())
                            .arg(discarded.size() == 1 ? "" : "s"));
    return failed;
}

// DATAVIEWER-4: detailed-sensory counterpart. Symmetric with the sensory helper
// above, but DetailedSensorySession carries no originalSessionName, so there is
// no in-place-rename -> force-INSERT branch (mirrors onUpdateDatabase's detailed
// loop). Flushes LiveSync once at the top, returns the failed indices.
QVector<int> MainWindow::saveDetailedSensorySessionsBeforeClose(const QVector<int>& indices)
{
    QVector<int> failed;
    if (!m_detailedSensoryPanel || !m_db) return failed;
    if (m_liveSync && !m_liveSync->flushNowAndWait()) {   // our scores -> DB first
        qWarning() << "saveDetailedSensorySessionsBeforeClose: LiveSync flush did not "
                      "complete (timeout or nested flush); proceeding (dirty-aware "
                      "merge keeps local edits, pending="
                   << m_liveSync->pendingCount() << ")";
    }

    // DATAVIEWER-4: mirror onUpdateDatabase's detailed pre-loop reconciliation so a
    // freshly-imported (id<=0) session that already exists in the DB resolves to an
    // UPDATE instead of a spurious INSERT/UniqueViolation that would block the close.
    m_detailedSensoryPanel->inheritExistingIdsAndVersions();

    QVector<int> discarded;
    bool aborted = false;

    QVector<DetailedSensorySession> sessions = m_detailedSensoryPanel->allSessions();  // flushes widgets
    for (int idx : indices) {
        if (aborted) { failed.append(idx); continue; }   // keep the rest open
        if (idx < 0 || idx >= sessions.size()) continue;
        DetailedSensorySession& sess = sessions[idx];
        if (DVE::isPlaceholderSession(sess)) continue;
        // v2.5.0 RC5: unnamed session — offer Name It Now / Discard / Cancel
        // instead of hard-blocking the close (twin of the sensory path).
        if (!DVE::isDetailedSessionSavable(sess)) {
            const SessionCloseChoice choice =
                promptUnnamedSessionOnClose(sessionCloseInfoFor(sess));
            switch (choice) {
            case SessionCloseChoice::NameIt:
                m_detailedSensoryPanel->focusTitleForSession(idx);
                failed.append(idx);
                aborted = true;
                continue;
            case SessionCloseChoice::Discard:
                if (sess.id > 0 && !m_db->removeDetailedSensorySession(sess.id))
                    qWarning().noquote()
                        << "[saveDetailedSensorySessionsBeforeClose] DB delete"
                        << "failed for discarded session" << sess.sessionName
                        << "-" << m_db->lastError() << "(discarding locally anyway)";
                discarded.append(idx);
                continue;
            case SessionCloseChoice::Cancel:
                failed.append(idx);
                aborted = true;
                continue;
            }
            continue;
        }

        const QString preName = sess.sessionName;
        // v2.5.0 RC4: auto-suffix on collision instead of blocking the close
        // (twin of the sensory close path).
        DVE::WriteResult r = m_db->tryWriteDetailedSensorySessionAutoSuffix(sess);
        if (r != DVE::WriteResult::Success) {
            // RC1: twin of the sensory close path — any non-Success is a
            // genuine failure now (the wrapper re-INSERTs on RowDeleted and
            // adopts fresh versions), so keep the session open.
            failed.append(idx);
        } else {
            if (sess.sessionName != preName) {
                qInfo().noquote() << "[saveDetailedSensorySessionsBeforeClose] name taken —"
                                  << "auto-renamed" << preName << "→" << sess.sessionName;
                updateStatusBar(
                    tr("Session name was taken — saved as \"%1\"").arg(sess.sessionName));
            }
            // v2.5.0 Task 3 (RC2 review, CRITICAL 2): Success — clear the dirty
            // set on this local copy (twin of the sensory close path). The adopt
            // in syncSavedSessionState() propagates it; failed sessions keep
            // their protection.
            sess.dirtyCells.clear();
        }
    }
    m_detailedSensoryPanel->syncSavedSessionState(sessions);
    refreshDetailedSensoryNavigator();   // v2.5.0 RC4: reflect any auto-suffix
    updateDbSyncIndicator();
    if (!discarded.isEmpty())
        updateStatusBar(tr("Discarded %1 unnamed session%2")
                            .arg(discarded.size())
                            .arg(discarded.size() == 1 ? "" : "s"));
    return failed;
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

    // F6 (v2.5.0): a fresh load FROM DISK must mint a NEW versioned DB row,
    // stamped with the add time, so re-adding the same .xlsx later keeps the
    // prior version as history instead of overwriting it. We therefore inherit
    // id+version ONLY from an in-memory copy of the SAME open file (re-opening
    // a file already in the working set continues that exact row) and
    // deliberately DO NOT adopt any existing DB row by file_path. The composite
    // UNIQUE(file_path, added_at) index means the subsequent INSERT no longer
    // collides on the path, so id stays -1 and persistLoadedFile takes the
    // INSERT branch — see docs/superpowers/specs/2026-06-10-v24-save-sync-
    // regressions-evidence.md (identity map). In-session saves keep UPDATEing
    // this freshly-minted row by id; loads FROM the DB (DatabaseBrowserDialog)
    // adopt the loaded row's id and continue that version.
    for (const FileResult& mem : m_loadedFiles) {
        if (isSameLoadedPath(mem.filePath, result.filePath) && mem.id > 0) {
            result.id = mem.id;
            result.version = mem.version;
            break;
        }
    }

    // Replace if already loaded — preserve images/layouts/crops from the in-memory version
    for (int i = 0; i < m_loadedFiles.size(); ++i) {
        if (isSameLoadedPath(m_loadedFiles[i].filePath, result.filePath)) {
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
            enqueuePersist(i);   // SP4.5 Stage 2a: save off the UI thread
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

    // Diagnostic: a genuinely new entry joins the working set. After the normalized
    // same-path checks above, the only remaining way to get two entries for "the
    // same" file is genuinely different paths (e.g. two folders) -- log both so any
    // residual duplicate is traceable to the exact paths that didn't match.
    {
        QStringList openPaths;
        for (const FileResult& mem : m_loadedFiles) openPaths << mem.filePath;
        if (!openPaths.isEmpty())
            qInfo().noquote() << "[load] NEW working-set entry:" << result.filePath
                              << "| already open:" << openPaths.join(", ");
    }
    m_loadedFiles.append(result);
    m_currentFileIndex   = m_loadedFiles.size() - 1;
    m_currentSheetIndex  = 0;
    m_currentSampleIndex = 0;

    // DATAVIEWER-16: re-apply any exclusions persisted for this file's PATH in a
    // prior session, re-stamped onto the index it just landed at.
    restoreExclusionsForFile(m_currentFileIndex);

    enqueuePersist(m_currentFileIndex);   // SP4.5 Stage 2a: save off the UI thread
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

// ─── Editable notes-story panel (TPM) ────────────────────────────────────────
// Routes a panel qualitative edit through the same path the old data table used:
// mutate the DataRow -> recalc -> push to the plot -> mark modified -> per-cell
// LiveSync to data_rows -> debounced Excel write-back. dataRow indexes directly
// into sample.rows (the panel does no visible-row skipping).
void MainWindow::onStoryCellEdited(int dataRow, int col, const QString& text) {
    // Re-entrancy guard: an editingFinished handler may setText()/show a tooltip,
    // which can bounce focus and re-deliver editingFinished, re-entering here
    // before we return. Drop the nested call (it's a duplicate of the same edit)
    // so heavy work can never nest or loop.
    if (m_inStoryCellEdit) return;
    m_inStoryCellEdit = true;
    struct ResetGuard { bool& b; ~ResetGuard() { b = false; } } guard{m_inStoryCellEdit};

    SheetResult* sheet = currentSheet();
    FileResult*  file  = currentFile();
    if (!sheet || !file || sheet->samples.isEmpty()) return;
    if (m_currentSampleIndex >= sheet->samples.size()) return;
    SampleResult& sample = sheet->samples[m_currentSampleIndex];
    if (dataRow < 0 || dataRow >= sample.rows.size()) return;
    DataRow& dr = sample.rows[dataRow];

    switch (col) {                                   // qualitative columns only
        case Cols::RESISTANCE:
            if (!sheet->hasPerRowRegime) return;     // col 4 is regime only on per-row-regime sheets
            dr.puffingRegime = text; m_storyRegimeDirty = true; break;
        case Cols::SMELL: dr.smell = text; break;
        case Cols::CLOG:  dr.clog  = text; break;
        case Cols::NOTES: dr.notes = text; break;
        default: return;
    }

    markFileModified();

    // Per-cell LiveSync to Postgres data_rows — mirrors onDataTableItemChanged.
    // Immediate (LiveSync is throttled + off-thread, so it never blocks the UI).
    if (m_liveSync && dr.id > 0) {
        const QString column = liveColumnForDataCol(col);
        if (!column.isEmpty())
            m_liveSync->commitCell(QStringLiteral("data_rows"), dr.id, column, text);
    }

    // Debounced Excel write-back — same cell math as onTableCellChanged.
    const int excelRow = dataRow + 5;
    const int excelCol = m_currentSampleIndex * 12 + col + 1;
    queueExcelWrite(file->filePath, sheet->sheetName, excelRow, excelCol, text);

    // Coalesce the heavy recalc + full plot re-render (the per-edit synchronous
    // setSheetData was the freeze). The data is already written above, so a
    // ~150ms-deferred redraw is imperceptible and rapid edits collapse to one.
    m_storyPlotTimer->start();

    // NOTE: deliberately DO NOT call m_storyPanel->setSample() here — re-populating
    // would destroy the editor widget the user is actively typing in. The panel's
    // displayed context refreshes on the next sample switch / file reload.
}
void MainWindow::onStoryNoteActivated(int dataRow) {
    // v1 note->plot emphasis: ring + guide the clicked note's TPM-trend point,
    // keyed by its cumulative-puff count (the plot's x value for that row).
    SheetResult* sheet = currentSheet();
    if (!sheet || sheet->samples.isEmpty()) return;
    if (m_currentSampleIndex >= sheet->samples.size()) return;
    const SampleResult& sample = sheet->samples[m_currentSampleIndex];
    if (dataRow < 0 || dataRow >= sample.rows.size()) return;
    m_plotWidget->selectPuff(int(sample.rows[dataRow].puffs));
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

    // data_rows UPDATE/INSERT NOTIFYs no longer drive any UI: the TPM edit
    // surface is the NotesStoryPanel, which always re-reads the model on the
    // next sample switch / file reload, so there is no live cell decoration to
    // paint here. The row-deleted banner above (file / sensory / detailed
    // sessions) is the only remaining live-collab cue this handler drives.
}

void MainWindow::displayCurrentSample()
{
    // Update the incomplete-data banner whenever the displayed content changes.
    // This must run before the early-return paths so the banner stays accurate
    // even when switching to raw/SOP sheets or empty files.
    updateIncompleteDataBanner();

    const SheetResult* sheet = currentSheet();

    // ── Raw table (SOP / instruction sheets) ──────────────────────────────────
    // After Option B the in-app grid is gone, so a raw/SOP sheet has no in-app
    // rendering. Show a hint in the story panel directing the user to View Raw
    // Data (opens the source Excel) instead of a blank centre.
    if (sheet && sheet->isRawTable) {
        m_plotWidget->hide();
        m_sampleCountLabel->setText("SOP View");
        m_prevBtn->setEnabled(false);
        m_nextBtn->setEnabled(false);
        m_propTable->setRowCount(0);

        const auto* f = currentFile();
        setStatusBreadcrumb({ f ? f->fileName : QString(), sheet->sheetName });
        m_storyPanel->showHint(
            tr("This is a raw/SOP sheet.\n\nUse View Raw Data to open this "
               "sheet in Excel."));
        return;
    }

    // Ensure plot is visible when showing normal sheets
    m_plotWidget->show();

    if (!sheet || sheet->samples.isEmpty()) {
        m_plotWidget->clear();
        m_storyPanel->showHint(tr("No samples to show for this sheet."));
        m_propTable->setRowCount(0);
        m_sampleCountLabel->setText("0 / 0");
        m_prevBtn->setEnabled(false);
        m_nextBtn->setEnabled(false);
        return;
    }

    const bool perRowRegime = sheet && sheet->hasPerRowRegime;

    m_currentSampleIndex = qBound(0, m_currentSampleIndex, (int)sheet->samples.size() - 1);
    const SampleResult& sample = sheet->samples[m_currentSampleIndex];

    // ── Plots & Properties ────────────────────────────────────────────────────
    // When cleanup is active, pass cleaned data to the plot and property panel
    // so the stats reflect only the included rows. The notes panel below shows
    // the RAW sample with excluded rows marked (its summaries exclude them).
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

    // ── Editable notes-story panel ────────────────────────────────────────────
    // The panel takes the RAW sample + the exclusion set (it marks excluded rows
    // and excludes them from its own summaries), keyed on the current
    // (file, sheet, sample).
    const QSet<int> storyExcl =
        exclusionsFor(m_currentFileIndex, m_currentSheetIndex, m_currentSampleIndex);
    m_storyPanel->setSample(sheet->samples[m_currentSampleIndex], storyExcl, perRowRegime);

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
    const FileResult reportFile = buildCleanedFile(*file, m_currentFileIndex);
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
    picker.setMinimumSize(360, 280);
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
            files.append(buildCleanedFile(m_loadedFiles[idx], idx));   // GAP-A: per-file index
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
        m_centralStack->setCurrentWidget(stackPageFor(m_centralStack, m_sensoryPanel));
        if (m_notesDock) m_notesDock->hide();   // Notes dock is TPM-only
        m_navStack->setCurrentWidget(m_sensoryNav);
        m_navLabel->setText("Sessions:  <span style='color:gray; font-size:11px;'>select multiple to show average sensory score</span>");
        refreshSensoryNavigator();
        if (m_testAvgPanel) m_testAvgPanel->setVisible(true);
        refreshSensoryAverages();
        updateSensoryProperties();
    } else {
        m_centralStack->setCurrentWidget(stackPageFor(m_centralStack, m_centralSplitter));
        if (m_notesDock) m_notesDock->show();   // back in TPM mode
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
        m_centralStack->setCurrentWidget(stackPageFor(m_centralStack, m_detailedSensoryPanel));
        if (m_notesDock) m_notesDock->hide();   // Notes dock is TPM-only
        m_navStack->setCurrentWidget(m_detailedSensoryNav);
        m_navLabel->setText("Sessions:  <span style='color:gray; font-size:11px;'>select multiple to show average score</span>");
        refreshDetailedSensoryNavigator();
        if (m_testAvgPanel) m_testAvgPanel->setVisible(true);
        updateDetailedSensoryProperties();
    } else {
        m_centralStack->setCurrentWidget(stackPageFor(m_centralStack, m_centralSplitter));
        if (m_notesDock) m_notesDock->show();   // back in TPM mode
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
    // Wrap in a ScrollHost so the horizontal cards/chart split scrolls instead
    // of clipping at small sizes (v2.7.0). m_sensoryPanel stays the inner ptr.
    m_centralStack->addWidget(ScrollHost::wrap(m_sensoryPanel));   // index 1

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
    // Wrap in a ScrollHost (v2.7.0): the 2-column grid + 4-quadrant chart can
    // overflow at small sizes; scroll instead of clip. inner ptr unchanged.
    m_centralStack->addWidget(ScrollHost::wrap(m_detailedSensoryPanel));

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

    // DV-17: View Raw Data is available in TPM + Sensory, disabled in Detailed
    // (DetailedSensorySession has no source Excel path).
    if (m_viewRawDataBtn)
        m_viewRawDataBtn->setEnabled(!detailedSensory);
}

void MainWindow::onViewRawData()
{
    // Open the SOURCE Excel for the active item. TPM: the loaded file's path;
    // Sensory: the current session's sourceFilePath (empty for a DB-only
    // session). Detailed is disabled in updateRibbonForMode (no source path).
    QString path;
    if (m_sensoryMode && m_sensoryPanel) {
        if (SensorySession* s = m_sensoryPanel->currentSession())
            path = s->sourceFilePath;
    } else if (!m_sensoryMode && !m_detailedSensoryMode) {   // TPM
        if (m_currentFileIndex >= 0 && m_currentFileIndex < m_loadedFiles.size())
            path = m_loadedFiles[m_currentFileIndex].filePath;
    }

    if (path.isEmpty() || !QFileInfo::exists(path)) {
        QMessageBox::information(this, QStringLiteral("View Raw Data"),
            QStringLiteral("No source Excel file is available for the current %1.")
                .arg(m_sensoryMode ? QStringLiteral("session") : QStringLiteral("file")));
        return;
    }
    QDesktopServices::openUrl(QUrl::fromLocalFile(path));
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
            if (isSameLoadedPath(m_loadedFiles[i].filePath, result.filePath)) {
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
void MainWindow::onUpdateDatabase(bool flushPending)
{
    // DATAVIEWER-4: at deliberate save points, push pending debounced edits to the
    // DB BEFORE the whole-session merge runs, so a just-typed value can't be briefly
    // overwritten by the merge's view of the DB. Bounded + no-op when idle.
    if (flushPending) {
        // SP3-T4: drain Excel write-back synchronously at this deliberate save
        // point so the on-disk workbook is consistent before the merge runs.
        finishExcelWritesBlocking();
        if (m_liveSync && !m_liveSync->flushNowAndWait()) {
            qWarning() << "onUpdateDatabase: LiveSync flush did not complete (timeout "
                          "or nested flush); proceeding (dirty-aware merge keeps local "
                          "edits, pending="
                       << m_liveSync->pendingCount() << ")";
        }
    }

    int saved = 0, failed = 0;
    // v2.5.0 RC4: the old `cancelled` counter tracked UniqueViolation skips that
    // popped a modal and abandoned the save. Collisions now auto-suffix silently
    // (tryWrite*AutoSuffix), so there is nothing to cancel — counter removed.

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
    // RC1 (v2.4.0 data-loss regression, see
    // docs/superpowers/specs/2026-06-10-v24-save-sync-regressions-evidence.md):
    // the old v2.0.5 heuristic treated VersionMismatch / RowDeleted as "LiveSync
    // already wrote this file, safe to skip" and dropped the dirty flag — which
    // silently discarded the user's in-memory edits whenever any edit did NOT
    // flow through LiveSync (23 skip events in the Jun 8-10 production log).
    // The write path now adopts fresh versions (child-row OCC, Task 2 Part A)
    // and re-INSERTs on RowDeleted (Part B), so a non-Success here is a genuine
    // failure: count it and KEEP the dirty flag so the next tick retries.
    for (int i = 0; i < m_loadedFiles.size(); ++i) {
        FileResult& fr = m_loadedFiles[i];
        if (!m_modifiedFilePaths.contains(fr.filePath)) continue;
        // SP4.5 Stage 2a: a background persist for this file is in flight; skip it
        // here so we don't race the worker into a duplicate INSERT. onPersistFinished
        // keeps the file dirty if it was edited during the save, so the next tick
        // saves the edit (as an UPDATE, with the real file id).
        if (m_backgroundSaveInFlight.contains(fr.filePath)) continue;
        const QString oldPath = fr.filePath;
        if (!m_db) { ++failed; continue; }
        const DVE::WriteResult r = m_db->tryWriteFile(fr);
        if (r == DVE::WriteResult::Success) {
            m_modifiedFilePaths.remove(oldPath);
            ++saved;
        } else {
            ++failed;   // keep the dirty flag — retry on the next tick
        }
    }

    // ── Save sensory sessions ──
    // DATAVIEWER-8: names of sessions skipped for a missing test name / tester.
    // Collected across BOTH the sensory and detailed loops, but only on the
    // interactive path (flushPending) so the background 5 s auto-save stays
    // silent. Surfaced once, after both loops, as a single summary message.
    QStringList incompleteNames;
    int sensSaved = 0;
    const int failedBeforeSensory = failed;   // RC1: scope dirty-clear to this section
    if (m_sensoryPanel) {
        auto sessions = m_sensoryPanel->allSessions();
        for (SensorySession& sess : sessions) {
            if (DVE::isPlaceholderSession(sess)) continue;
            if (!DVE::isSensorySessionSavable(sess)) {
                if (flushPending)
                    incompleteNames << (sess.testTitle.trimmed().isEmpty()
                                            ? tr("(unnamed session)") : sess.testTitle);
                continue;                          // never persist an unkeyed session
            }
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

            const QString preName = sess.sessionName;
            // v2.5.0 RC4: a name collision (brand-new session OR a rename whose
            // target name another session already owns) no longer pops a modal
            // and skips — that fed the June-10 endless loop (rename detected ->
            // INSERT -> 23505 -> dialog -> skip -> stale baseline -> repeat
            // every autosave tick). tryWriteSensorySessionAutoSuffix is a strict
            // superset of tryWriteSensorySession: non-collision results are
            // identical; on collision it auto-suffixes sessionName+testTitle
            // (_1/_2/_3...) and re-baselines originalSessionName so the loop dies.
            DVE::WriteResult r = m_db->tryWriteSensorySessionAutoSuffix(sess);

            if (r == DVE::WriteResult::Success) {
                if (sess.sessionName != preName) {
                    // Resolved a collision by renaming. Non-modal notice only.
                    qInfo().noquote() << "[onUpdateDatabase] sensory name taken —"
                                      << "auto-renamed" << preName << "→" << sess.sessionName;
                    updateStatusBar(
                        tr("Session name was taken — saved as \"%1\"").arg(sess.sessionName));
                }
                ++sensSaved;
                // v2.5.0 Task 3 (RC2 review, CRITICAL 2): the write landed, so
                // the locally-edited scores are now in the DB blob — clear the
                // dirty set on THIS local copy. syncSavedSessionState() below
                // ADOPTS this set into the panel, so a failed session (which
                // never reaches here) keeps its dirty cells and stays protected
                // on the retry.
                sess.dirtyCells.clear();
            } else {
                // RC1: any non-Success is a genuine failure (the wrapper now
                // re-INSERTs on RowDeleted and adopts fresh versions, so the
                // old "skip — already up to date via LiveSync" path is gone).
                // Count it and keep m_sensorySessionsDirty set so the next tick
                // retries — never silently drop the user's edits.
                ++failed;
            }
        }
        // Only clear the dirty flag when NO sensory save failed — a single
        // failure means an edit hasn't landed yet and must be retried next
        // tick. (Scoped to this section so a TPM/detailed failure elsewhere
        // doesn't keep the sensory store dirty.)
        if (sensSaved > 0 && failed == failedBeforeSensory)
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
        // v2.5.0 RC4: an auto-suffix may have changed a session's displayed
        // title; refresh the navigator labels so the list shows "T_1" etc.
        // (refreshSensoryNavigator rebuilds from allSessions()+sessionLabel()
        // without re-marking the store dirty, unlike emitting sessionsChanged).
        refreshSensoryNavigator();
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
    const int failedBeforeDetailed = failed;   // RC1: scope dirty-clear to this section
    if (m_detailedSensoryPanel) {
        m_detailedSensoryPanel->inheritExistingIdsAndVersions();
        auto detSessions = m_detailedSensoryPanel->allSessions();
        for (DetailedSensorySession& sess : detSessions) {
            if (DVE::isPlaceholderSession(sess)) continue;
            if (!DVE::isDetailedSessionSavable(sess)) {
                if (flushPending)
                    incompleteNames << (sess.testTitle.trimmed().isEmpty()
                                            ? tr("(unnamed session)") : sess.testTitle);
                continue;                          // never persist an unkeyed session
            }
            if (!m_db) { ++failed; continue; }

            const QString preName = sess.sessionName;
            // v2.5.0 RC4: name collision auto-suffixes instead of blocking
            // (twin of the sensory branch). DetailedSensorySession carries no
            // originalSessionName, so only the plain INSERT-collision path
            // applies; the wrapper still suffixes sessionName+testTitle.
            DVE::WriteResult r = m_db->tryWriteDetailedSensorySessionAutoSuffix(sess);

            if (r == DVE::WriteResult::Success) {
                if (sess.sessionName != preName) {
                    qInfo().noquote() << "[onUpdateDatabase] detailed name taken —"
                                      << "auto-renamed" << preName << "→" << sess.sessionName;
                    updateStatusBar(
                        tr("Session name was taken — saved as \"%1\"").arg(sess.sessionName));
                }
                ++detSaved;
                // v2.5.0 Task 3 (RC2 review, CRITICAL 2): twin of the sensory
                // loop — clear the dirty set on THIS local copy only on Success.
                // syncSavedSessionState() below adopts it, so a failed session
                // keeps its dirty cells and stays protected on the retry.
                sess.dirtyCells.clear();
            } else {
                // RC1: any non-Success is a genuine failure (twin of the
                // sensory loop). Count it and keep m_detailedSensorySessionsDirty
                // set so the next tick retries — never silently drop edits.
                ++failed;
            }
        }
        if (detSaved > 0 && failed == failedBeforeDetailed)
            m_detailedSensorySessionsDirty = false;

        // Merge id/version (+ per-image imageIds/imageVersions) back into panel
        // state from the local `detSessions` copy that tryWriteDetailedSensorySession's
        // byRef back-fill updated — symmetric with the sensory block above. Without
        // this, allSessions() returning m_sessions by value means the panel's
        // m_sessions[i].id stays -1 after a first save, so every repeat save in the
        // same run re-INSERTs images and churns rows.
        m_detailedSensoryPanel->syncSavedSessionState(detSessions);
        // v2.5.0 RC4: refresh navigator labels after a possible auto-suffix
        // (twin of the sensory block above).
        refreshDetailedSensoryNavigator();
    }
    updateDbSyncIndicator();

    // DATAVIEWER-8: on a deliberate save (Ctrl+U / program-close), name the
    // sessions that were skipped because they lack a test name + tester. The
    // background auto-save (flushPending == false) collected nothing above, so
    // this never fires there. Skipped sessions are neither saved nor failed, so
    // the counters below stay accurate; the work survives via the recovery
    // snapshot until the user fills in the missing fields.
    if (flushPending && !incompleteNames.isEmpty()) {
        QMessageBox::information(this, tr("Some Sessions Not Saved"),
            tr("These sessions need a test name and tester before they can be saved:\n\n  - %1")
                .arg(incompleteNames.join(QStringLiteral("\n  - "))));
    }

    const int total = saved + sensSaved + detSaved;
    // SP4.5 Stage 2a: a Ctrl+U / auto-save that actually wrote something changed
    // the DB -> schedule a debounced background snapshot regen so close stays fast.
    if (total > 0 && m_snapshotRegenTimer) m_snapshotRegenTimer->start();
    if (total == 0 && failed == 0) {
        // Nothing was dirty (LiveSync already persisted every cell). RC1: the
        // old "N items already live-synced" path is gone — VersionMismatch /
        // RowDeleted no longer count as benign skips; they're failures now.
        updateStatusBar("Database already up to date.");
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
    if (m_pendingWrites.isEmpty() && !m_excelFlushInFlight) {
        updateStatusBar(tr("Nothing to export"));
        return;
    }
    // SP3-T4: Ctrl+S is a deliberate manual flush — finish synchronously so the
    // "Exported to Excel" status reflects a completed write (drains any in-flight
    // off-thread flush first, then any still-pending cells with the batch budget).
    finishExcelWritesBlocking();
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
    if (!m_recovery)
        return;

    // Restore-last-session reopen prompt. adoptPreviousSession() (run in the
    // ctor) moved the previous session's store -- however the app last closed,
    // clean or not -- into Recovery_prev/.
    //
    // R5 minor: read the store ONCE. The previous code called hasRecoverable()
    // (which reads the store, re-spawning the python MIP fallback) and then
    // recoverableItems() (reading it AGAIN). recoverableItems() == readAll(prev)
    // sets lastReadFailed()/lastError() as a side effect, so we branch on the
    // single read: empty + not-failed => nothing to recover (silent), empty +
    // failed => present-but-undecodable (warn), non-empty => offer reopen.
    const QVector<RecoveryEntry> items = m_recovery->recoverableItems();
    if (items.isEmpty()) {
        // The previous session's store EXISTS but could not be read -- most
        // likely a MIP/AIP-encrypted index/blob the bundled python could not
        // decrypt (the durability surface this fix hardens). Do NOT fail
        // silently: a crashed session left work behind we cannot surface. A
        // genuinely-absent store (true first run / declined-last-time) leaves
        // lastReadFailed() false and we return quietly.
        if (m_recovery->lastReadFailed()) {
            QMessageBox::warning(
                this, tr("Could Not Read Recovered Session"),
                tr("DataViewer detected unsaved work from a previous session "
                   "but could not read it back.\n\n%1\n\n"
                   "The files may be encrypted at rest (Microsoft Information "
                   "Protection). The recovery data is preserved on disk and you "
                   "can retry via Tools → Recover.")
                    .arg(m_recovery->lastError()));
        }
        return;   // nothing readable to reopen
    }

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
        // Ensure Notes dock is visible when recovering into TPM mode (e.g. last
        // session crashed while in a sensory mode and restoreState hid it).
        if (m_notesDock)
            m_notesDock->setVisible(true);
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
    // SP4.5 Stage 2a: if a background persist for this file is in flight, its job
    // snapshot predates this edit. Record the edit so onPersistFinished keeps the
    // file dirty (instead of marking it clean against the stale snapshot); the
    // next whole-file save then re-persists the edited rows with the real file id.
    if (m_backgroundSaveInFlight.contains(f->filePath))
        m_dirtiedDuringPersist.insert(f->filePath);
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

    // v2.5.0 Task 4 (RC3): per-cell live-sync failures take visual precedence.
    // The worker reconnect-and-retries broken connections, so a non-zero count
    // means an edit STILL didn't land after a retry. It is not lost — the
    // dirty-cell merge re-persists it on the next whole-session save — but the
    // user should see it isn't live yet. Warning (amber) state + explicit
    // tooltip pointing at Ctrl+U.
    if (m_unsyncedEdits > 0) {
        const QString msg = prefix
            + QString("%1 edit%2 not synced — they will be saved on the next "
                      "save (Ctrl+U)")
                  .arg(m_unsyncedEdits)
                  .arg(m_unsyncedEdits == 1 ? "" : "s");
        setStatusDb(msg, DbStatusModified);
        if (m_statusDbText)
            m_statusDbText->setToolTip(
                tr("%1 cell edit(s) could not be live-synced to the database. "
                   "Your changes are kept locally and will be written on the "
                   "next save (Ctrl+U).").arg(m_unsyncedEdits));
        return;
    }

    // v2.4.4 R5: distinct, non-blocking "queue unreadable" state. The local
    // offline pending-edits queue exists but couldn't be opened/decoded (most
    // likely MIP-encrypted at rest), so offline per-cell edits can't be
    // persisted to it. This rides on offlineQueueDegraded() (NOT pendingCount(),
    // which deliberately excludes the degraded queue so its ==0 drain invariant
    // holds). Edits are not lost — they stay in the open session and re-persist
    // on the next online save — so this is a warning, never modal/blocking.
    if (m_liveSync && m_liveSync->offlineQueueDegraded()) {
        setStatusDb(prefix + tr("Local backup queue unreadable — edits kept in "
                                "session; save (Ctrl+U) while online"),
                    DbStatusModified);
        if (m_statusDbText)
            m_statusDbText->setToolTip(
                tr("The local offline backup queue could not be read or written "
                   "(it may be encrypted at rest by Microsoft Information "
                   "Protection). Per-cell edits cannot be queued there while "
                   "offline, but your changes are kept in this session and will "
                   "be written on the next online save (Ctrl+U)."));
        return;
    }
    if (m_statusDbText) m_statusDbText->setToolTip(QString());

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

// v2.5.0 RC5: program-close counterpart to the per-session close dialog. For
// every unnamed (un-keyable) non-placeholder session in either panel, ask
// Name It Now / Discard / Cancel. Returns false to ABORT the close (the user
// wants to name a session or cancelled); true to proceed once all unnamed
// sessions are either discarded or there are none left.
bool MainWindow::resolveUnnamedSessionsForProgramClose()
{
    // Process descending so a discardSession() (which shifts later indices down)
    // can't invalidate indices we haven't visited yet. A single Name-It-Now /
    // Cancel aborts immediately — the user has signalled they want to deal with
    // it before the app exits.
    if (m_sensoryPanel) {
        const QVector<SensorySession> snap = m_sensoryPanel->allSessions();
        for (int idx = snap.size() - 1; idx >= 0; --idx) {
            const SensorySession& sess = snap[idx];
            if (DVE::isPlaceholderSession(sess)) continue;
            if (DVE::isSensorySessionSavable(sess)) continue;
            switch (promptUnnamedSessionOnClose(sessionCloseInfoFor(sess))) {
            case SessionCloseChoice::NameIt:
                m_sensoryPanel->focusTitleForSession(idx);
                return false;                 // veto the close; let them name it
            case SessionCloseChoice::Cancel:
                return false;                 // veto the close
            case SessionCloseChoice::Discard:
                if (m_db && sess.id > 0 && !m_db->removeSensorySession(sess.id))
                    qWarning().noquote()
                        << "[resolveUnnamedSessionsForProgramClose] sensory DB"
                        << "delete failed for" << sess.sessionName << "-"
                        << m_db->lastError() << "(discarding locally anyway)";
                m_sensoryPanel->discardSession(idx);
                break;
            }
        }
    }
    if (m_detailedSensoryPanel) {
        const QVector<DetailedSensorySession> snap =
            m_detailedSensoryPanel->allSessions();
        for (int idx = snap.size() - 1; idx >= 0; --idx) {
            const DetailedSensorySession& sess = snap[idx];
            if (DVE::isPlaceholderSession(sess)) continue;
            if (DVE::isDetailedSessionSavable(sess)) continue;
            switch (promptUnnamedSessionOnClose(sessionCloseInfoFor(sess))) {
            case SessionCloseChoice::NameIt:
                m_detailedSensoryPanel->focusTitleForSession(idx);
                return false;
            case SessionCloseChoice::Cancel:
                return false;
            case SessionCloseChoice::Discard:
                if (m_db && sess.id > 0
                    && !m_db->removeDetailedSensorySession(sess.id))
                    qWarning().noquote()
                        << "[resolveUnnamedSessionsForProgramClose] detailed DB"
                        << "delete failed for" << sess.sessionName << "-"
                        << m_db->lastError() << "(discarding locally anyway)";
                m_detailedSensoryPanel->discardSession(idx);
                break;
            }
        }
    }
    refreshSensoryNavigator();
    refreshDetailedSensoryNavigator();
    updateDbSyncIndicator();
    return true;
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
    // v2.5.0 RC5: before persisting, resolve any unnamed (un-keyable) sessions
    // with the same Name It Now / Discard / Cancel options used by the per-tab
    // close button — never silently drop them, and never hard-block. Name It Now
    // / Cancel veto the close (return false) so the user can finish naming; a
    // chosen Discard deletes the row + drops the session before we save the rest.
    if (!resolveUnnamedSessionsForProgramClose())
        return false;                       // → closeEvent vetoes the close

    // v2.4.16: show a progress bar from the START of the save work. Previously the
    // DB save (onUpdateDatabase) + disk saves ran with a frozen window for a couple
    // seconds, and only the snapshot regen afterwards showed a bar. This is a member
    // dialog so the snapshot regen REUSES it (one continuous bar); closeEvent tears
    // it down via closeProgressEnd(). Busy/indeterminate (max 0) for the save phase;
    // regenerateSnapshotWithProgress switches it to determinate.
    if (!m_closeProgress)
        m_closeProgress = makeBusyDialog(this, tr("Saving your work..."), 0);

    // For untitled (non-placeholder) sessions, route through the panel's own
    // save() so the user is offered a Save-As for the on-disk copy. The
    // "no path yet" guard is read here (before save() runs, since save() sets
    // the path); already-saved panels skip the dialog. onUpdateDatabase() is
    // the authoritative DB persist for ALL sessions (named or freshly-titled),
    // so the save() calls are purely the disk-file courtesy for untitled work.
    // DATAVIEWER-8: only run the disk-courtesy save() for a savable current
    // session. Otherwise save()'s hard-guard modal would fire during shutdown.
    // onUpdateDatabase(true) just below skips + summarizes any incomplete
    // session, and incomplete work survives via the recovery snapshot, so
    // gating these out loses no data.
    if (m_sensorySessionsDirty && m_sensoryPanel
        && !m_sensoryPanel->hasSavePath()
        && m_sensoryPanel->currentSessionSavable()) {
        m_sensoryPanel->save();
    }
    if (m_detailedSensorySessionsDirty && m_detailedSensoryPanel
        && !m_detailedSensoryPanel->hasSavePath()
        && m_detailedSensoryPanel->currentSessionSavable()) {
        m_detailedSensoryPanel->save();
    }
    if (m_closeProgress) {
        m_closeProgress->setLabelText(tr("Saving to database..."));
        QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
    }
    onUpdateDatabase(/*flushPending=*/true);   // DATAVIEWER-4: deliberate program-close save
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

void MainWindow::applyVeryNarrowDockState(bool veryNarrow)
{
    if (veryNarrow) {
        if (m_docksAutoCollapsed) return;   // already collapsed; idempotent
        m_docksAutoCollapsed = true;
        // Reclaim the side-dock width for the central ScrollHost. hide() keeps
        // the docks out of the layout entirely so the central page gets the
        // full client width before it must scroll.
        if (m_fileDock)  m_fileDock->hide();
        if (m_notesDock) m_notesDock->hide();
        return;
    }

    // Leaving VeryNarrow: only restore docks that WE collapsed, and only to
    // their current mode-appropriate visibility. The Navigator is always
    // reachable; the Notes dock is TPM-only, so it stays hidden in a sensory
    // mode even after the window widens.
    if (!m_docksAutoCollapsed) return;
    m_docksAutoCollapsed = false;
    if (m_fileDock)  m_fileDock->show();
    if (m_notesDock) m_notesDock->setVisible(!m_sensoryMode && !m_detailedSensoryMode);
}

// ─── Settings ─────────────────────────────────────────────────────────────────
void MainWindow::restoreSettings()
{
    QSettings s("SDR", "DataViewerEnterprise");
    restoreGeometry(s.value("geometry").toByteArray());
    restoreState(s.value("windowState").toByteArray());
    // restoreState can hide m_notesDock if the app was last closed in a sensory
    // mode.  Force-sync its visibility to the current mode so TPM always shows it.
    if (m_notesDock)
        m_notesDock->setVisible(!m_sensoryMode && !m_detailedSensoryMode);

    // Safety net: the Navigator must ALWAYS be reachable. restoreState() has no
    // visibility guard for it (unlike m_notesDock above), so a prior run that
    // closed while the Navigator was floating can restore it hidden or off the
    // screen — which reads to the user as "the navigator vanished" with no way
    // back. Deferred to the event loop so the window is shown first (isVisible()
    // and geometry are only meaningful then); a deliberate, on-screen float is
    // left untouched.
    QTimer::singleShot(0, this, [this]() {
        if (!m_fileDock) return;
        const bool offScreen = m_fileDock->isFloating()
            && !QGuiApplication::screenAt(m_fileDock->frameGeometry().center());
        if (!m_fileDock->isVisible() || offScreen) {
            m_fileDock->setFloating(false);
            addDockWidget(Qt::LeftDockWidgetArea, m_fileDock);
            if (m_sidebarStack && m_sidebarFullPanel)
                m_sidebarStack->setCurrentWidget(m_sidebarFullPanel);
            m_fileDock->show();
            m_fileDock->raise();
            resizeDocks({m_fileDock}, {280}, Qt::Horizontal);
        }
    });

    // One-time: make the Navigator and Notes docks an identical default width so
    // the two side panels are symmetric out of the box.  Keyed on a settings
    // flag so it runs once per install; afterwards the user's own dock sizing is
    // restored by restoreState() above and left untouched.  Deferred to the
    // event loop because resizeDocks() is a no-op until the docks are shown.
    // Flag bumped to _v2 because the Navigator dock gained an objectName this
    // release; older saved layouts can't match it, so re-balance once to a clean
    // symmetric default, then the user's own sizing (now persisted) takes over.
    if (!s.value("dockWidthsBalanced_v2", false).toBool()) {
        s.setValue("dockWidthsBalanced_v2", true);
        QTimer::singleShot(0, this, [this]() {
            if (m_fileDock && m_notesDock)
                resizeDocks({m_fileDock, m_notesDock}, {280, 280}, Qt::Horizontal);
        });
    }

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

void MainWindow::resetPanelLayout()
{
    // Master reset for the side panels: un-float, re-dock to defaults, show,
    // un-collapse the sidebar, re-balance widths, and persist immediately. The
    // user-facing escape hatch for "I floated a panel and can't drag it back" or
    // "the Navigator vanished". Invoked from Settings ▸ Panels ▸ Reset Panels.
    if (m_fileDock) {
        m_fileDock->setFloating(false);
        addDockWidget(Qt::LeftDockWidgetArea, m_fileDock);
        m_fileDock->show();
        m_fileDock->raise();
    }
    if (m_sidebarStack && m_sidebarFullPanel)
        m_sidebarStack->setCurrentWidget(m_sidebarFullPanel);   // un-collapse
    if (m_notesDock) {
        m_notesDock->setFloating(false);
        addDockWidget(Qt::RightDockWidgetArea, m_notesDock);
        // Notes is TPM-only; match the current mode's expected visibility.
        m_notesDock->setVisible(!m_sensoryMode && !m_detailedSensoryMode);
    }
    // Symmetric default widths, deferred until the re-dock settles.
    QTimer::singleShot(0, this, [this]() {
        if (m_fileDock && m_notesDock)
            resizeDocks({m_fileDock, m_notesDock}, {280, 280}, Qt::Horizontal);
    });
    saveSettings();   // persist now so a crash can't resurrect the broken layout
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
        if (!m_snapshot->openReadOnly()
            && m_snapshot->lastOpenWasDecodeFailure()
            && !m_offlineSnapshotDecodeWarningShown) {
            // v2.4.4 R5: the snapshot EXISTS but is MIP-encrypted and could not
            // be decoded -- we are now offline with no readable local cache. A
            // mere first-run absence does NOT warn; a decode failure does. Once
            // per session, non-blocking.
            m_offlineSnapshotDecodeWarningShown = true;
            QMessageBox* box = new QMessageBox(
                QMessageBox::Warning,
                tr("Offline Cache Unreadable"),
                tr("DataViewer is offline and the local read-only cache could "
                   "not be opened. The cache file may be encrypted at rest "
                   "(Microsoft Information Protection).\n\n"
                   "You may not be able to view previously cached data until the "
                   "database connection is restored.\n\n%1")
                    .arg(m_snapshot->lastError()),
                QMessageBox::Ok, this);
            box->setAttribute(Qt::WA_DeleteOnClose);
            box->setModal(false);
            box->show();
        }
    }

    if (m_offlineBanner) {
        if (m_snapshot && m_snapshot->isOpen()) {
            m_offlineBanner->setLastSync(m_snapshot->snapshotTakenAt());
        } else {
            m_offlineBanner->setLastSync(QDateTime());  // "(No previous snapshot.)"
        }
        m_offlineBanner->setPendingCount(0);
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

    // Re-subscribe to NOTIFY. v2.4.2 R4b: per-channel heal instead of the old
    // all-or-nothing isSubscribed() guard — if a single channel dropped (the
    // half-open / GFW case), resubscribeMissing() refills exactly the missing
    // ones without disturbing the survivors. Returns false only if some
    // channel still failed to re-attach.
    if (m_notify && !m_notify->resubscribeMissing()) {
        qWarning() << "MainWindow: NOTIFY resubscribe incomplete — "
                   << "some live-update channels did not re-attach";
    }

    // R3: re-activate local presence so peers see us again and the heartbeat
    // restarts (offline deactivated it). Lazy reactivation left the user a
    // ghost to other clients after every blip. onConnectionWentOffline()
    // called PresenceManager::deactivate(), which clears m_activeIntent to ""
    // — so activeIntent() is empty here. Fall back to the app's canonical
    // default ("viewing") rather than activating with an empty intent (which
    // would write a malformed presence row matching neither "viewing" nor
    // "editing"). A user mid-edit re-escalates to "editing" on the next edit.
    if (m_presence && !m_currentResourceType.isEmpty() && m_currentResourceId > 0) {
        QString intent = m_presence->activeIntent();
        if (intent.isEmpty()) intent = QStringLiteral("viewing");
        m_presence->activate(m_currentResourceType, m_currentResourceId, intent);
    }
    refreshAllPresence();

    setStatusDb(tr("Reconnected to database."), DbStatusOk);

    // C4: drain the persistent per-cell LiveSync queue. It holds granular
    // edits captured offline (table/row/column tuples) and replays them via the
    // precise dve_commit_cell path. This is now the only offline outbound queue
    // (the file-level TPM capture went away with the data-table edit surface).
    if (m_liveSync) {
        const int replayed = m_liveSync->flushPending();
        if (replayed > 0) {
            qInfo() << "MainWindow: LiveSync replayed" << replayed
                    << "per-cell edits on reconnect";
        }
    }

    // R3 (reset-to-5 keystone): ORDER MATTERS. The per-cell drain above lands
    // every locally-captured outbound edit in the DB first; only THEN do we pull
    // the authoritative DB state for the open resource and dirty-aware-merge it
    // back, so a remote change made during the offline window (e.g. the nightly
    // normalizer rewriting a legacy string score) is adopted without ever
    // reverting the user's own unsaved edits.
    reloadOpenResourceAfterReconnect();
}

// v2.4.2 R3 — the reset-to-5 keystone. See the header comment on the
// declaration. Best-effort + logged: any failure leaves memory untouched.
void MainWindow::reloadOpenResourceAfterReconnect()
{
    // Nothing open, or no live online connection to read truth from.
    if (m_currentResourceId <= 0 || !m_db || !m_db->isOnline())
        return;

    const QString resType = m_currentResourceType;
    const qint64  resId   = m_currentResourceId;

    if (resType == QLatin1String("file")) {
        // Find the in-memory FileResult with the matching id.
        for (int i = 0; i < m_loadedFiles.size(); ++i) {
            if (qint64(m_loadedFiles[i].id) != resId) continue;

            // TPM has no per-cell dirty merge yet (deferred). If the file has
            // local unsaved edits we must NOT clobber them — leave memory as is.
            if (m_modifiedFilePaths.contains(m_loadedFiles[i].filePath)) {
                qInfo() << "MainWindow: reconnect catch-up skipped for dirty file"
                        << m_loadedFiles[i].filePath
                        << "(local unsaved edits preserved)";
                break;
            }

            // Clean file: replace with the authoritative DB row and re-render.
            const int fileId = m_loadedFiles[i].id;
            FileResult fresh = m_db->loadFile(fileId);
            if (fresh.filePath.isEmpty()) {
                qWarning() << "MainWindow: reconnect catch-up loadFile failed for"
                           << fileId;
                break;
            }
            m_loadedFiles[i] = fresh;
            if (i == m_currentFileIndex) {
                displayCurrentSample();
            }
            qInfo() << "MainWindow: reconnect catch-up reloaded file" << fileId;
            break;
        }
    } else if (resType == QLatin1String("sensory_session") && m_sensoryPanel) {
        // currentSession() flushes the on-screen widgets back into the struct
        // first (folding in any unsaved user edits), so *curr is the freshest
        // local state before we merge.
        SensorySession* curr = m_sensoryPanel->currentSession();
        if (!curr || qint64(curr->id) != resId) {
            qInfo() << "MainWindow: reconnect catch-up skipped sensory — current "
                       "session does not match open resource id" << resId;
            return;
        }
        const SensorySession dbNow = m_db->loadSensorySession(int(resId));
        if (dbNow.id <= 0) {
            // A remote row-deletion during the offline window (dbNow.id<=0) is
            // INTENTIONALLY ignored by catch-up: local wins and the row
            // resurrects on the next save — matching the documented TPM
            // deferral, hence the asymmetry with the "file" branch above.
            qWarning() << "MainWindow: reconnect catch-up loadSensorySession "
                          "failed/empty for" << resId;
            return;
        }
        // Dirty-aware merge: DB-authoritative for non-dirty scores (adopts any
        // remote/normalizer change), in-memory wins for the user's dirty cells.
        const QJsonObject merged = mergeSensoryPreservingDbScores(
            sensorySessionToJson(*curr), sensorySessionToJson(dbNow),
            curr->dirtyCells);
        // Overlay ONLY the merged scalar scores onto the in-memory session and
        // re-render the cards from it, all inside the panel. This deliberately
        // does NOT rebuild the struct via sensorySessionFromJson (which would
        // drop images and reset id/version, breaking OCC). curr->dirtyCells is a
        // field on the struct and is left untouched, so a subsequent save still
        // treats the local edits as authoritative. The render happens BEFORE the
        // navigator refresh below, so the buildSession() that refresh triggers
        // reads the merged widgets and cannot revert the adopted remote value.
        m_sensoryPanel->applyMergedScoresToCurrentSession(merged);
        refreshSensoryNavigator();
        qInfo() << "MainWindow: reconnect catch-up merged sensory session" << resId;
    } else if (resType == QLatin1String("detailed_sensory_session")
               && m_detailedSensoryPanel) {
        DetailedSensorySession* curr = m_detailedSensoryPanel->currentSession();
        if (!curr || qint64(curr->id) != resId) {
            qInfo() << "MainWindow: reconnect catch-up skipped detailed-sensory — "
                       "current session does not match open resource id" << resId;
            return;
        }
        const DetailedSensorySession dbNow =
            m_db->loadDetailedSensorySession(int(resId));
        if (dbNow.id <= 0) {
            // As in the sensory branch: a remote row-deletion during the offline
            // window is INTENTIONALLY ignored (local wins / resurrects on next
            // save), matching the documented TPM deferral.
            qWarning() << "MainWindow: reconnect catch-up "
                          "loadDetailedSensorySession failed/empty for" << resId;
            return;
        }
        const QJsonObject merged = mergeDetailedSensoryPreservingDbScores(
            detailedSensorySessionToJson(*curr),
            detailedSensorySessionToJson(dbNow), curr->dirtyCells);
        // Same overlay-only contract as the sensory branch: scores onto the
        // in-memory struct + re-render, preserving id/version/images and the
        // dirty set, BEFORE the navigator flush.
        m_detailedSensoryPanel->applyMergedScoresToCurrentSession(merged);
        refreshDetailedSensoryNavigator();
        qInfo() << "MainWindow: reconnect catch-up merged detailed-sensory session"
                << resId;
    }
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

bool MainWindow::regenerateSnapshotWithProgress(const QString& title)
{
    // SP4.5 audit fix: run the synchronous regen on a DEDICATED PG connection,
    // never the shared m_pgConn. regenerate() opens a REPEATABLE READ READ ONLY
    // transaction and pumps the event loop to keep the progress bar painting; on
    // the shared connection a LiveSync-throttle or presence-heartbeat QTimer firing
    // during that pump would issue a write inside the read-only txn (SQLSTATE
    // 25P02) and abort the regen. A private connection lets those timers run
    // harmlessly. Falls back to the shared connection if the dedicated open fails
    // (no worse than before).
    DVE::PostgresConnection regenConn;
    const bool dedicated = regenConn.open(m_dbConfig);
    if (!dedicated)
        qWarning() << "[perf] regen: dedicated PG connection failed, using shared:"
                   << regenConn.lastError();
    DVE::PostgresConnection* const conn = dedicated ? &regenConn : m_pgConn;
    if (!conn || !conn->isOpen()) return false;

    // v2.4.16: if the close->save bar is already up (m_closeProgress), REUSE it so
    // the close shows ONE continuous bar (no flash of a second dialog). Otherwise
    // (the standalone Refresh-Snapshot menu action) own a local one. The shared
    // dialog started busy/indeterminate (max 0) during the DB save; switch it to a
    // determinate 0..100 for the regen phases. closeEvent owns the shared teardown.
    QProgressDialog* progress = m_closeProgress;
    const bool ownDialog = (progress == nullptr);
    if (ownDialog) {
        progress = makeBusyDialog(this, title, 100);
    } else {
        progress->setMaximum(100);
        progress->setLabelText(title);
        progress->setValue(0);
    }
    // ExcludeUserInputEvents keeps the bar painting without letting a second
    // close/refresh click re-enter this path mid-regen.
    QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);

    const bool ok = m_snapshot->regenerate(conn,
        [progress](int done, int total, const QString& phase) {
            progress->setLabelText(phase);
            progress->setValue(total > 0 ? (done * 100 / total) : 0);
            QApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
        });

    progress->setValue(100);
    if (ownDialog) {
        progress->close();
        delete progress;            // local dialog: tear down now
    }
    // shared m_closeProgress: leave it up; closeEvent's closeProgressEnd() ends it.
    if (dedicated) regenConn.close();
    return ok;
}

// v2.4.16: tear down the close->save->snapshot progress bar (a no-op if the user
// chose Discard / there was nothing to save). Called at the end of closeEvent.
void MainWindow::closeProgressEnd()
{
    if (!m_closeProgress) return;
    m_closeProgress->close();
    delete m_closeProgress;
    m_closeProgress = nullptr;
}

void MainWindow::onRefreshSnapshotTriggered()
{
    if (!m_db || !m_db->isOnline() || !m_snapshot
        || !m_pgConn || !m_pgConn->isOpen()) {
        QMessageBox::information(this, tr("Refresh Offline Snapshot"),
            tr("Snapshot can only be refreshed when connected to the database."));
        return;
    }

    const bool ok = regenerateSnapshotWithProgress(tr("Refreshing offline snapshot..."));

    if (ok) {
        QMessageBox::information(this, tr("Refresh Offline Snapshot"),
            tr("Offline snapshot refreshed successfully."));
    } else {
        QMessageBox::warning(this, tr("Refresh Offline Snapshot"),
            tr("Snapshot refresh failed: %1").arg(m_snapshot->lastError()));
    }
}

void MainWindow::closeEvent(QCloseEvent* e)
{
    // SP4.5: per-step close timing. The file log handler timestamps every line
    // to the ms, so these markers give a full close breakdown (esp. on the real
    // NAS DB where the snapshot regen used to cost ~9s on every close).
    qInfo() << "[perf] closeEvent: enter";

    // H2: drain any debounced Excel cell edits BEFORE the prompt, so the
    // workbook on disk picks up the last burst of changes. The 500 ms
    // m_excelWriteTimer may not have fired yet when the user closes.
    // SP3-T4: finish synchronously so the app never proceeds to teardown with
    // an Excel write still running on a background thread.
    finishExcelWritesBlocking();
    qInfo() << "[perf] closeEvent: excel flush done";

    // SP4.5 Stage 2a: drain the background persist worker (and flush its writeback
    // slots) BEFORE the save prompt, so a quick load->close commits the file write
    // and the prompt reads an authoritative dirty set.
    drainPersistWorkerBlocking(15000);
    // Stop the regen debounce and cancel any in-flight background regen. cancel()
    // is atomic + non-blocking; the worker aborts at its next checkpoint on its
    // own thread (the destructor's quit()/wait() joins it). If a regen WAS in
    // flight, the synchronous fallback below MUST be skipped -- it writes the same
    // snapshot.sqlite.tmp the aborting worker is still tearing down, and running
    // both would corrupt the snapshot. The snapshot then stays one session stale;
    // the next clean online close regenerates it.
    const bool regenWasInFlight = m_snapshotRegenInFlight;
    if (m_snapshotRegenTimer) m_snapshotRegenTimer->stop();
    if (m_regenWorker) m_regenWorker->cancel();
    if (regenWasInFlight)
        qInfo() << "[perf] closeEvent: background regen was in flight -- cancelled; "
                   "snapshot left as-is for the next clean close";

    if (!promptSaveDatabase()) { e->ignore(); return; }
    qInfo() << "[perf] closeEvent: prompt/save done";

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
    qInfo() << "[perf] closeEvent: image cache wiped";

    // Plan C T9: refresh the offline snapshot for next launch. Best-effort;
    // failure is logged but never blocks close. Skips when offline (the snapshot
    // already represents the last clean state).
    //
    // SP4.5 (perf): the full regen copies every table + all image blobs from the
    // NAS — ~9s on the real DB, and it used to run on EVERY close (the dominant
    // close-freeze). Skip it when the live DB is unchanged since the last snapshot
    // (the common case, incl. read-only sessions): a cheap COUNT + MAX(updated_at)
    // fingerprint match means there is nothing new to copy. When it HAS changed we
    // still regenerate synchronously (Stage 2 will move that off the UI thread).
    if (!regenWasInFlight && m_db && m_db->isOnline() && m_snapshot
        && m_pgConn && m_pgConn->isOpen()) {
        if (m_snapshot->isCurrentVsLive(m_pgConn)) {
            qInfo() << "[perf] closeEvent: offline snapshot already current — skipping regen";
        } else {
            qInfo() << "[perf] closeEvent: regenerating offline snapshot (DB changed)";
            // SP4.5 Stage 2b: the regen is now incremental (sub-second on a
            // data-only close; pulls only changed blobs on an image change). Show
            // a determinate "Saving..." bar so any real work reads as progress,
            // not a frozen window. processEvents() keeps the bar painting.
            const bool ok = regenerateSnapshotWithProgress(tr("Saving offline copy..."));
            if (!ok)
                qWarning() << "Snapshot regenerate failed:" << m_snapshot->lastError();
            qInfo() << "[perf] closeEvent: snapshot regen done";
        }
    }

    saveSettings();
    qInfo() << "[perf] closeEvent: settings saved";

    // Restore-last-session: persist the final in-memory state (open files +
    // sessions, including unsaved edits) so the NEXT launch can offer to reopen
    // this session -- regardless of how the app closed. (Previously this wiped
    // the store on a clean close, which made recovery crash-only and meant a
    // normal "Don't save" close lost the session.) flushNow(true) is a no-op
    // when recovery isn't armed (no state provider was set).
    if (m_recovery && m_recoveryArmed) m_recovery->flushNow(true);
    qInfo() << "[perf] closeEvent: recovery flushed — accepting close";

    closeProgressEnd();   // v2.4.16: tear down the save/snapshot progress bar
    e->accept();
}

// ─── Accessors ────────────────────────────────────────────────────────────────

QString MainWindow::liveColumnForDataCol(int col) const
{
    // Maps a DVE::Cols index to its data_rows DB column. The notes-story panel
    // only edits the qualitative columns (NOTES / SMELL / CLOG / RESISTANCE),
    // so only those need mapping for the per-cell LiveSync commit. Column 4 is
    // dual-purpose: per-row-regime sheets store the puffing regime there.
    switch (col) {
        case DVE::Cols::RESISTANCE: {
            const SheetResult* s = currentSheet();
            return (s && s->hasPerRowRegime) ? QStringLiteral("puffing_regime")
                                             : QStringLiteral("resistance");
        }
        case DVE::Cols::SMELL: return QStringLiteral("smell");
        case DVE::Cols::CLOG:  return QStringLiteral("clog");
        case DVE::Cols::NOTES: return QStringLiteral("notes");
        default:               return QString();
    }
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
                              QString& errOut,
                              int timeoutMs)
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

    if (!proc.waitForFinished(timeoutMs)) {
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

    saveExclusions();          // DATAVIEWER-16: persist across restarts
    displayCurrentSample();
}

void MainWindow::onResetCleanup()
{
    const SheetResult* sheet = currentSheet();
    if (!sheet) return;
    for (int si = 0; si < sheet->samples.size(); ++si)
        m_excludedRows.remove(cleanupKey(m_currentFileIndex, m_currentSheetIndex, si));
    saveExclusions();          // DATAVIEWER-16: persist across restarts
    displayCurrentSample();
}

void MainWindow::onUndoAllCleanup()
{
    if (m_excludedRows.isEmpty()) {
        showInfo("No Cleanup", "There are no data exclusions to undo.");
        return;
    }
    if (QMessageBox::question(
            this, tr("Undo All Cleanup"),
            tr("Remove ALL data exclusions across every open file?\n\n"
               "Plots and reports will show the full, unfiltered data again."),
            QMessageBox::Yes | QMessageBox::No, QMessageBox::No)
        != QMessageBox::Yes)
        return;

    m_excludedRows.clear();
    saveExclusions();          // clears the persisted state too (empty map)
    displayCurrentSample();    // refresh plots/tables for the current sheet
    updateCleanupButtons();
}

QString MainWindow::cleanupKey(int fileIdx, int sheetIdx, int sampleIdx) const
{
    return DataCleanup::key(fileIdx, sheetIdx, sampleIdx);
}

QSet<int> MainWindow::exclusionsFor(int fileIdx, int sheetIdx, int sampleIdx) const
{
    return DataCleanup::exclusionsFor(m_excludedRows, fileIdx, sheetIdx, sampleIdx);
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
    if (m_undoAllCleanupBtn)
        m_undoAllCleanupBtn->setEnabled(!m_excludedRows.isEmpty());
}

SampleResult MainWindow::buildCleanedSample(const SampleResult& sr,
                                             const QSet<int>& excluded) const
{
    return DataCleanup::buildCleanedSample(sr, excluded);
}

SheetResult MainWindow::buildCleanedSheet(const SheetResult& sheet,
                                          int fileIdx, int sheetIdx) const
{
    return DataCleanup::buildCleanedSheet(sheet, m_excludedRows, fileIdx, sheetIdx);
}

FileResult MainWindow::buildCleanedFile(const FileResult& file, int fileIdx) const
{
    return DataCleanup::buildCleanedFile(file, m_excludedRows, fileIdx);
}

// ─── Cleanup persistence (DATAVIEWER-16) ──────────────────────────────────────
// m_excludedRows is keyed by the VOLATILE fileIdx ("fileIdx:sheetIdx:sampleIdx").
// Across restarts a file can land at a different index (reordering, different
// open order), so persistence is keyed by the file PATH instead: each path maps
// to a list of "sheetIdx:sampleIdx=row,row,..." entries. On the next load of
// that path, restoreExclusionsForFile() re-stamps them onto the current fileIdx.
// Storage uses the same QSettings store as the rest of the app's UI state.
void MainWindow::saveExclusions() const
{
    QSettings s("SDR", "DataViewerEnterprise");
    s.beginGroup("cleanupExclusions");
    s.remove("");   // rewrite the whole group from the live map

    // path -> list of "sheetIdx:sampleIdx=r,r,r"
    QMap<QString, QStringList> byPath;
    for (auto it = m_excludedRows.constBegin(); it != m_excludedRows.constEnd(); ++it) {
        if (it.value().isEmpty()) continue;
        const QStringList parts = it.key().split(':');
        if (parts.size() != 3) continue;
        bool ok = false;
        const int fileIdx = parts[0].toInt(&ok);
        if (!ok || fileIdx < 0 || fileIdx >= m_loadedFiles.size()) continue;
        const QString path = m_loadedFiles[fileIdx].filePath;
        if (path.isEmpty()) continue;

        QList<int> rows = it.value().values();
        std::sort(rows.begin(), rows.end());
        QStringList rowStrs;
        for (int r : rows) rowStrs << QString::number(r);
        byPath[path] << QString("%1:%2=%3").arg(parts[1], parts[2], rowStrs.join(','));
    }

    int n = 0;
    for (auto it = byPath.constBegin(); it != byPath.constEnd(); ++it) {
        s.beginGroup(QString("file%1").arg(n++));
        s.setValue("path", it.key());
        s.setValue("entries", it.value());
        s.endGroup();
    }
    s.endGroup();
}

void MainWindow::restoreExclusionsForFile(int fileIdx)
{
    if (fileIdx < 0 || fileIdx >= m_loadedFiles.size()) return;
    const QString path = m_loadedFiles[fileIdx].filePath;
    if (path.isEmpty()) return;

    QSettings s("SDR", "DataViewerEnterprise");
    s.beginGroup("cleanupExclusions");
    const QStringList groups = s.childGroups();
    for (const QString& g : groups) {
        s.beginGroup(g);
        const QString storedPath = s.value("path").toString();
        // Match the same way the working set does: normalized, case-insensitive.
        const bool same = isSameLoadedPath(storedPath, path);
        if (same) {
            const QStringList entries = s.value("entries").toStringList();
            for (const QString& e : entries) {
                const int eq = e.indexOf('=');
                if (eq < 0) continue;
                const QStringList loc = e.left(eq).split(':');
                if (loc.size() != 2) continue;
                QSet<int> rows;
                const QStringList rowStrs = e.mid(eq + 1).split(',', Qt::SkipEmptyParts);
                for (const QString& r : rowStrs) {
                    bool ok = false;
                    const int v = r.toInt(&ok);
                    if (ok) rows.insert(v);
                }
                if (!rows.isEmpty())
                    m_excludedRows[cleanupKey(fileIdx, loc[0].toInt(), loc[1].toInt())] = rows;
            }
        }
        s.endGroup();
        if (same) break;
    }
    s.endGroup();
}

// ─────────────────────────────────────────────────────────────────────────────

void MainWindow::recalculateSampleMetrics(SheetResult& sheet)
{
    GenericSheetProcessor proc;
    for (SampleResult& sr : sheet.samples)
        proc.calculateMetrics(sr);
    proc.computeSheetAggregates(sheet);
}

void MainWindow::writeCellToExcel(const QString& filePath, const QString& sheetName,
                                   int excelRow1, int excelCol1, const QString& value)
{
    writeCellsToExcel(filePath, sheetName, { { excelRow1, excelCol1, value } });
}

bool MainWindow::writeCellsToExcel(const QString& filePath, const QString& sheetName,
                                    const QVector<CellWrite>& cells, int timeoutMs)
{
    if (cells.isEmpty()) return true;        // nothing to do is a success
    const QString python = findPython();
    if (python.isEmpty()) return false;       // no interpreter → cannot persist

    // SP3-T4 (R6): the openpyxl source + argv come from the shared builders in
    // ExcelWritePayload (single source of truth, pinned by the equivalence test).
    // This is the SYNCHRONOUS path — used by bulk row ops and by
    // finishExcelWritesBlocking on the close paths. The debounced single-cell
    // flush goes off-thread through runExcelWriteWorker, which calls runPython
    // with the SAME script + args, so the two paths write byte-identically.
    const QStringList args = DVE::buildWriteCellsArgs(filePath, sheetName, cells);
    const ExcelWriteResult r =
        runExcelWriteWorker(python, QString::fromUtf8(DVE::excelWriteCellsScript()),
                            args, timeoutMs);
    if (!r.ok && !r.error.isEmpty()) {
        qWarning() << "writeCellsToExcel failed:" << r.error;
    }
    return r.ok;
}

MainWindow::ExcelWriteResult
MainWindow::runExcelWriteWorker(const QString& python,
                                const QString& script,
                                const QStringList& args,
                                int timeoutMs)
{
    // Pure, self-contained: runs on a QtConcurrent worker thread for the
    // debounced flush, and inline on the UI thread for the synchronous path.
    // It touches NO MainWindow member and NO QWidget — only the value copies
    // it was handed. runPython is a pure static (temp script file + QProcess),
    // safe to call off the UI thread.
    QString err;
    const QString out = runPython(python, script, args, err, timeoutMs);
    // H3: runPython returns the empty string on failure and stuffs the
    // stderr / timeout message into err. The success path prints "OK" from
    // the kWriteCells script; missing OK with empty err still signals a
    // malformed invocation we should not treat as a clean save.
    ExcelWriteResult result;
    if (!err.isEmpty()) {
        result.ok    = false;
        result.error = err;
        return result;
    }
    result.ok = out.contains(QLatin1String("OK"));
    if (!result.ok)
        result.error = QStringLiteral("openpyxl did not report OK");
    return result;
}

void MainWindow::queueExcelWrite(const QString& filePath, const QString& sheetName,
                                  int excelRow1, int excelCol1, const QString& value)
{
    // SP3-T4 (R6): if the target file/sheet changed from what is pending OR what a
    // worker is currently writing, FULLY DRAIN synchronously before re-labelling.
    // Otherwise the cleared-but-in-flight async batch (or accumulated pending
    // cells for the old sheet) would get mixed with the new file/sheet labels —
    // the old synchronous code never had this risk because it drained inline on
    // every switch. File switches are user-driven and infrequent, so a synchronous
    // drain here is cheap and preserves the original "flush old before switching"
    // semantics exactly.
    const bool oldPendingMismatch =
        !m_pendingWrites.isEmpty() &&
        (m_pendingWriteFile != filePath || m_pendingWriteSheet != sheetName);
    const bool inFlightMismatch =
        m_excelFlushInFlight &&
        (m_inFlightWriteFile != filePath || m_inFlightWriteSheet != sheetName);
    if (oldPendingMismatch || inFlightMismatch) {
        finishExcelWritesBlocking();
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

    // SP3-T4 (R6): dispatch the openpyxl save to a QtConcurrent worker so the
    // UI thread never blocks on the subprocess (was an up-to-30 s freeze).
    //
    // Re-entrancy guard (mirrors RecoveryManager::dispatchFlush): only ONE
    // python writer to the workbook at a time. If a worker is already running,
    // record that another flush is wanted and return — the finished slot will
    // re-fire flushExcelWrites() with whatever has accumulated by then. This
    // bounds us to a single concurrent writer, so two writes can't race the
    // same file.
    if (m_excelFlushInFlight) {
        m_excelFlushPending = true;
        return;
    }

    const QString python = findPython();
    if (python.isEmpty()) {
        // No interpreter → cannot persist. Mirror the old failure behaviour:
        // keep m_pendingWrites for retry and surface the rate-limited warning.
        if (!m_excelWriteFailureShown) {
            m_excelWriteFailureShown = true;
            QMessageBox::warning(
                this, tr("Excel save failed"),
                tr("Could not write pending changes to %1.\n\n"
                   "No bundled Python interpreter was found. The edits are "
                   "still in memory and the database save will include them. "
                   "The file on disk is unchanged.")
                .arg(m_pendingWriteFile));
        }
        return;
    }

    // Snapshot pending state into the in-flight copy and clear the queue
    // optimistically. New edits arriving during the worker land in a fresh
    // m_pendingWrites and trigger a follow-up flush via the finished slot.
    m_inFlightWriteFile  = m_pendingWriteFile;
    m_inFlightWriteSheet = m_pendingWriteSheet;
    m_inFlightWrites     = m_pendingWrites;       // value copy
    m_pendingWrites.clear();

    // Tiered timeout: the debounced flush is interactive (single/few cells),
    // so use the shorter interactive budget; bulk multi-cell row ops go through
    // the synchronous writeCellsToExcel with the batch budget instead.
    const QString    script = QString::fromUtf8(DVE::excelWriteCellsScript());
    const QStringList args  = DVE::buildWriteCellsArgs(
        m_inFlightWriteFile, m_inFlightWriteSheet, m_inFlightWrites);
    const int timeoutMs = kExcelInteractiveTimeoutMs;

    if (!m_excelFlushWatcher) {
        m_excelFlushWatcher = new QFutureWatcher<ExcelWriteResult>(this);
        connect(m_excelFlushWatcher, &QFutureWatcher<ExcelWriteResult>::finished,
                this, &MainWindow::onExcelFlushFinished);
    }

    m_excelFlushInFlight = true;
    // Capture only value copies; the worker touches no member and no QWidget.
    m_excelFlushWatcher->setFuture(QtConcurrent::run(
        &MainWindow::runExcelWriteWorker, python, script, args, timeoutMs));
}

void MainWindow::onExcelFlushFinished()
{
    // UI thread. Idempotent: finishExcelWritesBlocking() may have already
    // consumed this completion synchronously (after waitForFinished), so a
    // later-delivered queued `finished` signal must be a no-op rather than
    // double-processing the result or spawning a spurious flush.
    if (!m_excelFlushInFlight)
        return;

    // Read the worker's result and clear the in-flight state.
    const ExcelWriteResult r = m_excelFlushWatcher->result();
    m_excelFlushInFlight = false;

    if (r.ok) {
        m_excelWriteFailureShown = false;
        m_inFlightWrites.clear();
    } else {
        // H3: keep the failed writes for retry. Re-queue the in-flight cells
        // back into m_pendingWrites so a later flush (the next user edit's debounce
        // tick, or finishExcelWritesBlocking on close/Ctrl+S) retries them. The
        // merge SEEDS the older in-flight cells then OVERLAYS the newer pending
        // ones (newest wins per cell) — see mergePendingWithInFlight. Only merge
        // when the file/sheet still match what's pending; if the user switched
        // files mid-flight, persist the stale batch on its own so it isn't lost.
        if (m_pendingWrites.isEmpty()) {
            m_pendingWriteFile  = m_inFlightWriteFile;
            m_pendingWriteSheet = m_inFlightWriteSheet;
            m_pendingWrites     = m_inFlightWrites;
        } else if (m_pendingWriteFile == m_inFlightWriteFile
                   && m_pendingWriteSheet == m_inFlightWriteSheet) {
            m_pendingWrites =
                DVE::mergePendingWithInFlight(m_inFlightWrites, m_pendingWrites);
        } else {
            // File/sheet switched mid-flight and new edits are pending for the
            // OTHER file. Persist the stale batch synchronously now so it is not
            // dropped (rare; deterministic, no nested async dispatch).
            writeCellsToExcel(m_inFlightWriteFile, m_inFlightWriteSheet,
                              m_inFlightWrites, kExcelInteractiveTimeoutMs);
        }
        m_inFlightWrites.clear();

        if (!m_excelWriteFailureShown) {
            m_excelWriteFailureShown = true;
            QMessageBox::warning(
                this, tr("Excel save failed"),
                tr("Could not write pending changes to %1.\n\n"
                   "The edits are still in memory and the database "
                   "save will include them. The file on disk is unchanged. "
                   "Close any other program that has the workbook open "
                   "and try again.")
                .arg(m_inFlightWriteFile));
        }
    }

    // Re-fire ONLY on an explicit coalesced request (m_excelFlushPending), exactly
    // like RecoveryManager. We do NOT auto-re-dispatch just because m_pendingWrites
    // is non-empty: a failed batch re-queued above would otherwise spin a tight
    // worker-retry loop against a persistently locked file. The debounce timer is
    // not autonomous — it only runs after queueExcelWrite (re)starts it — so a
    // re-queued failure stays on disk until the NEXT USER EDIT (whose queueExcelWrite
    // restarts the timer, whose tick flushes pending and so retries the failure) or
    // a CLOSE / Ctrl+S drain (finishExcelWritesBlocking). m_excelFlushPending is set
    // only when something needed a flush *now* while one was in flight (e.g. a
    // file-switch flush or a close path).
    if (m_excelFlushPending) {
        m_excelFlushPending = false;
        flushExcelWrites();
    }
}

void MainWindow::finishExcelWritesBlocking()
{
    // SP3-T4 (R6): used by the close paths (5 inline-flush sites + ~MainWindow)
    // in place of the async flushExcelWrites(). We must NOT exit with a write
    // abandoned on a background thread, NOT crash on a still-running future, and
    // NOT spawn a fresh async worker while tearing down. So this is entirely
    // synchronous and self-contained: it does NOT call onExcelFlushFinished()
    // (which re-dispatches async) — it folds the in-flight batch into the
    // pending set inline, then writes synchronously with the batch budget.
    //
    // 1. Wait out any in-flight worker, then consume its result directly. Setting
    //    m_excelFlushInFlight=false makes the worker's queued `finished` signal a
    //    no-op (the guard at the top of onExcelFlushFinished), so it can never
    //    double-process or re-dispatch after we've taken over synchronously.
    if (m_excelFlushInFlight && m_excelFlushWatcher) {
        m_excelFlushWatcher->waitForFinished();
        const ExcelWriteResult r = m_excelFlushWatcher->result();
        m_excelFlushInFlight = false;
        m_excelFlushPending  = false;
        if (r.ok) {
            m_excelWriteFailureShown = false;
        } else {
            // Fold the failed in-flight cells back into pending so step 2 retries
            // them synchronously. Newer same-cell edits that landed mid-flight win.
            if (m_pendingWrites.isEmpty()) {
                m_pendingWriteFile  = m_inFlightWriteFile;
                m_pendingWriteSheet = m_inFlightWriteSheet;
                m_pendingWrites     = m_inFlightWrites;
            } else if (m_pendingWriteFile == m_inFlightWriteFile
                       && m_pendingWriteSheet == m_inFlightWriteSheet) {
                m_pendingWrites =
                    DVE::mergePendingWithInFlight(m_inFlightWrites, m_pendingWrites);
            } else {
                // Stale batch belongs to a different file than what's pending;
                // persist it synchronously now so it's not dropped.
                writeCellsToExcel(m_inFlightWriteFile, m_inFlightWriteSheet,
                                  m_inFlightWrites, kExcelBatchTimeoutMs);
            }
        }
        m_inFlightWrites.clear();
    }

    // 2. Run any still-pending writes synchronously with the batch budget, so
    //    the on-disk workbook reflects the last burst before the app exits.
    if (!m_pendingWrites.isEmpty()) {
        const bool ok = writeCellsToExcel(
            m_pendingWriteFile, m_pendingWriteSheet, m_pendingWrites,
            kExcelBatchTimeoutMs);
        if (ok) {
            m_pendingWrites.clear();
            m_excelWriteFailureShown = false;
        }
        // On failure the edits remain in m_pendingWrites (in memory) and the DB
        // save still includes them; we do not loop or block close further.
    }
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
