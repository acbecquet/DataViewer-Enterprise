#pragma once

#include <QWidget>
#include <QScrollArea>
#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QGroupBox>
#include <QLineEdit>
#include <QDoubleSpinBox>
#include <QComboBox>
#include <QTextEdit>
#include <QLabel>
#include <QSplitter>
#include <QTimer>
#include <QVector>
#include <QPushButton>
#include <QStackedWidget>
#include <QTableWidget>
#include <QDialogButtonBox>
#include <QListWidget>

#include "pipeline/DetailedSensoryData.h"
#include "ui/RadarChartWidget.h"
#include "database/DatabaseManager.h"

namespace QXlsx { class Document; }

namespace DVE {

class LiveSync;

class DetailedSensoryPanel : public QWidget
{
    Q_OBJECT

public:
    explicit DetailedSensoryPanel(DatabaseManager* db, QWidget* parent = nullptr);

    // ── LiveSync routing (optional; injected from MainWindow) ────────────────
    // When set, per-cell commits on the active sample's question form are
    // forwarded to LiveSync, and remote cell-changed signals are applied
    // back to the in-memory session (and the visible form if it's showing
    // the affected sample) with signals blocked.
    void setLiveSync(LiveSync* sync);

    void loadSessions(const QVector<DetailedSensorySession>& sessions);
    void selectSession(int index);
    void showAveragedChart(const QVector<int>& sessionIndices);
    void newSession();
    void closeSessions(const QVector<int>& indices);
    void renameSession(int index, const QString& newLabel);

    // v2.5.0 RC5 (close options): drop a single session from the panel — twin of
    // SensoryPanel::discardSession. Removes it from m_sessions + fixes the
    // current-index bookkeeping. Disk autosave files are intentionally left in
    // place (the caller has already removed the DB row when one existed).
    void discardSession(int index);

    // v2.5.0 RC5 (close options): make `index` current and focus the Test Title
    // field — drives the "Name It Now" close option (twin of the sensory panel).
    void focusTitleForSession(int index);

    void save();
    void loadFile(const QString& path);
    void loadFiles();
    void loadFromDatabase();

    void generateFullReport();

    QVector<DetailedSensorySession> allSessions();

    // v2.0.5: symmetric to SensoryPanel::inheritExistingIdsAndVersions.
    // For every loaded session with id <= 0, inherit (id, version) from
    // any existing detailed_sensory_sessions row matching the natural
    // key. MainWindow calls this before the file-level save loop in
    // onUpdateDatabase so re-imports take the UPDATE branch instead of
    // INSERT-failing the UNIQUE constraint idx_detailed_sensory_sessions_key.
    void inheritExistingIdsAndVersions();

    // Plan C C10: after MainWindow's onUpdateDatabase save loop runs over the
    // by-value vector from allSessions(), merge the server-assigned id/version
    // (and the parallel imageIds/imageVersions that tryWriteDetailedSensorySession
    // back-fills for newly-inserted image rows) back into this panel's m_sessions.
    // Without it the panel's id stays -1 after a first save, so every repeat save
    // in the same run re-INSERTs images and churns rows. Mirrors
    // SensoryPanel::syncSavedSessionState (minus originalSessionName, which
    // DetailedSensorySession does not carry). Matched by index when counts agree,
    // else by natural key (sessionName/testerName/date). Persistence anchors only —
    // never overwrites user-editable content.
    void syncSavedSessionState(const QVector<DetailedSensorySession>& saved);

    // DV-28: per-session dirty marking for MainWindow sites that mutate a session
    // without going through the panel widgets (prop-table, image attach) - twin of
    // SensoryPanel. markAllSessionsDirty is for crash-recovery restores. Consumed
    // by the save-loop gate (detailedSessionNeedsSave).
    void markSessionDirty(int idx);
    void markAllSessionsDirty();

    int  currentSessionIndex() const { return m_currentTesterIdx; }
    QString sessionLabel(const DetailedSensorySession& s) const;
    DetailedSensorySession* currentSession();

    // v2.4.2 SP2-T4 (reset-to-5 keystone, Defect-1/2 fix): overlay ONLY the
    // per-metric kDetailedAllMetrics scalar scores from a pre-computed merged
    // blob (mergeDetailedSensoryPreservingDbScores output) onto the CURRENT
    // in-memory session, then re-render the visible form from it. Twin of
    // SensoryPanel::applyMergedScoresToCurrentSession and the blessed overlay in
    // dbAuthoritativeSessions: it touches nothing but scores, so
    // id/version/imagePaths/imageLayouts/imageCrops/imageIds/imageVersions and
    // every other field stay in-memory-authoritative (no fromJson round-trip,
    // which would drop images and reset the persistence anchors). Re-rendering
    // here — BEFORE any caller-side navigator refresh / allSessions() flush —
    // means a subsequent buildSession() reads the merged widgets and cannot
    // revert the adopted remote value. No-op if no current session.
    void applyMergedScoresToCurrentSession(const QJsonObject& mergedSession);

    // Plan C C10: true once this panel's sessions have been saved to a disk
    // file at least this run. The consolidated close prompt's "Save All" uses
    // this to decide whether untitled work still needs a Save-As (the DB save
    // itself is handled separately by MainWindow::onUpdateDatabase).
    bool hasSavePath() const { return !m_savePath.isEmpty(); }

    // DATAVIEWER-8: true iff the CURRENTLY-displayed session has the non-empty
    // test name + tester needed to be persisted. MainWindow gates the program-
    // close disk-courtesy save() on this so app shutdown never trips save()'s
    // hard-guard modal for an unnamed current session.
    bool currentSessionSavable() const;

    void showAveragedTable(const QStringList& deviceNames,
                           const QVector<QMap<QString, double>>& deviceAvgs);
    void showNormalView();

    static bool generateCombinedPptx(const QVector<DetailedSensorySession>& sessions,
                                      const QString& filePath,
                                      QString& errorOut);

signals:
    // Structural changes only (new/close/rename/load/save, add/remove sample).
    // Per-FIELD value edits do NOT emit this.
    void sessionsChanged();

    // Plan C (C6 fix): fires on EVERY per-field edit that mutates the
    // in-memory session — sample name/comments/spins/combos (via
    // saveCurrentSampleToSession) and session-level header + oil-smell/clog/
    // mouthpiece fields (via commitSessionField). MainWindow connects this to
    // RecoveryManager::noteDirty() so routine detailed-sensory data entry is
    // captured by the crash snapshot. Distinct from sessionsChanged so the
    // structural consumers are not re-run on every keystroke.
    void dataEdited();

private slots:
    // v2.0.1: applied when LiveSync receives a remote per-cell change.
    void onRemoteCellChanged(const QString& table, qint64 rowId,
                             const QString& column, const QVariant& newValue);

private:
    void buildHeaderRow(QWidget* container);
    void buildQuestionForm();
    void buildSampleNavBar();

    // DV-28: mark the CURRENT session dirty (self-connected to dataEdited). A
    // no-op while m_applyingSession is true (twin of SensoryPanel).
    void markCurrentSessionDirty();

    // Returns the persisted id of the currently-selected session, or -1
    // when out of range or not yet persisted. Gate for LiveSync commits.
    int activeSessionId() const;

    // v2.0.1: forward a per-cell commit to LiveSync. commitSessionField
    // targets a top-level column (e.g. "test_title"); commitSampleField
    // targets a path on the active sample, prefixed with
    // "samples[currentSampleIdx].". Both are no-ops when LiveSync isn't
    // wired or the session hasn't been persisted yet.
    void commitSessionField(const QString& fieldPath, const QVariant& value);
    void commitSampleField (const QString& fieldPath, const QVariant& value);

    // Patch the in-memory sample with a JSON-path field update from LiveSync.
    void applyRemoteFieldToSample(DetailedSensorySample& sample,
                                  const QString& fieldPath,
                                  const QVariant& value);
    // Patch a session-level field (oil_smell_liking, clog, …) and refresh UI.
    void applyRemoteSessionField(DetailedSensorySession& sess,
                                 const QString& fieldPath,
                                 const QVariant& value);
    // Re-bind the form widgets from m_sessions[m_currentTesterIdx].samples[i]
    // with signals blocked.
    void rebindCurrentSampleFromMemory();

    void displayCurrentSample();
    void saveCurrentSampleToSession();
    void onPrevSample();
    void onNextSample();
    void onAddSample();
    void onRemoveSample();
    void updateSampleNav();

    DetailedSensorySession buildSession() const;

    // DATAVIEWER-4: DB-authoritative SCORES for export. Flushes LiveSync, then for
    // each saved session overlays the DB row's kDetailedAllMetrics scores onto a
    // COPY of the in-memory session (images/anchors/structure preserved -- NOT a
    // lossy fromJson round-trip). id<=0 or no-DB pass through unchanged.
    QVector<DetailedSensorySession> dbAuthoritativeSessions(const QVector<DetailedSensorySession>& inMem);

    void           applySession(const DetailedSensorySession& session);
    void           saveCurrentTester();
    bool           isDefaultState() const;

    void scheduleChartRefresh();
    void onRefreshChart();

    void saveToExcel(const QString& path, const DetailedSensorySession& sess);

    // Header row
    QLineEdit*        m_testTitleEdit;
    QLineEdit*        m_assessorEdit;
    QLineEdit*        m_testerEdit;
    QLineEdit*        m_mediaEdit;
    QLabel*           m_dateLabel;

    // Sample navigation
    QPushButton*      m_prevBtn;
    QPushButton*      m_nextBtn;
    QLabel*           m_sampleCountLabel;
    QPushButton*      m_addSampleBtn;
    QPushButton*      m_removeSampleBtn;

    // Question form
    QWidget*          m_questionForm;
    QScrollArea*      m_questionScroll;
    QMap<QString, QDoubleSpinBox*> m_spinBoxes;
    QMap<QString, QComboBox*>      m_comboBoxes;
    QLineEdit*        m_sampleNameEdit;
    QTextEdit*        m_commentsEdit;
    QDoubleSpinBox*   m_oilSmellSpin = nullptr;
    QComboBox*        m_clogCombo = nullptr;
    QComboBox*        m_mouthpieceCombo = nullptr;

    // Dual radar charts
    RadarChartWidget* m_vaporQualityChart;
    RadarChartWidget* m_consistencyChart;

    // Averaged table overlay
    QStackedWidget*   m_topStack = nullptr;
    QTableWidget*     m_avgOverlayTable = nullptr;

    // Session data
    QVector<DetailedSensorySession> m_sessions;
    int  m_currentTesterIdx  = -1;
    int  m_currentSampleIdx  = 0;
    // DV-28: true only while applySession() programmatically repopulates the
    // header/cards, so markCurrentSessionDirty() ignores the resulting signals.
    bool m_applyingSession   = false;

    QTimer* m_refreshTimer;
    // v2.0.1: debounce comments LiveSync commits — QTextEdit fires textChanged
    // on every keystroke; coalesce bursts into a single commit on timeout.
    QTimer* m_commentsCommitTimer = nullptr;
    QString m_savePath;
    DatabaseManager* m_db;

    // LiveSync (optional; owned by MainWindow; null in tests/offline).
    LiveSync* m_liveSync = nullptr;

    QString m_lastBrowseDir;
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
};

} // namespace DVE
