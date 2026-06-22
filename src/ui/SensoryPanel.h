#pragma once

#include <QWidget>
#include <QScrollArea>
#include <QVBoxLayout>
#include <QGroupBox>
#include <QLineEdit>
#include <QSpinBox>
#include <QTextEdit>
#include <QLabel>
#include <QSplitter>
#include <QTimer>
#include <QVector>
#include <QComboBox>
#include <QStackedWidget>
#include <QTableWidget>
#include <QDoubleSpinBox>
#include <QPushButton>
#include <functional>

#include "pipeline/SensoryData.h"
#include "ui/RadarChartWidget.h"
#include "database/DatabaseManager.h"

namespace QXlsx { class Document; }

namespace DVE {

class LiveSync;

// ─── Per-sample card widget ────────────────────────────────────────────────────
class SampleCard : public QGroupBox
{
    Q_OBJECT

public:
    explicit SampleCard(int index, QWidget* parent = nullptr);

    SensorySample toSample() const;
    void          fromSample(const SensorySample& s);

    // v2.0.4: hook the sample-name field up to a "Saved values" dropdown.
    // The provider is invoked on every dropdown click so it always
    // reflects the latest DB state — no caching, no NOTIFY plumbing.
    // Calling this more than once is safe; only the most recent
    // provider is used (older dropdown actions stay but their captured
    // provider is benign).
    void attachNamePresetDropdown(std::function<QStringList()> provider,
                                  std::function<void(const QString&)> onDelete = {});

signals:
    void changed();
    void removeRequested(SampleCard* card);
    // v2.0.1: per-widget commit event for LiveSync routing. Fires alongside
    // changed() with a JSON sub-path (e.g. "name", "voltage", "Burnt Taste")
    // identifying which field the new value belongs to inside the sample.
    void cellCommitted(const QString& jsonPath, const QVariant& value);

private:
    QLineEdit* m_nameEdit;
    QMap<QString, QDoubleSpinBox*> m_spinBoxes;  // all entries are NoWheelDoubleSpinBox instances (see SensoryPanel.cpp)
    QTextEdit* m_commentsEdit;

    // Per-sample device properties
    QLineEdit* m_voltageEdit;
    QLineEdit* m_resistanceEdit;
    QComboBox* m_heatingTechCombo;
    QLabel*    m_powerLabel;

    // #7: per-sample test conditions
    QComboBox*      m_powerTypeCombo;
    QDoubleSpinBox* m_puffLengthSpin;  // NoWheelDoubleSpinBox instance

    // v2.0.1: debounce comments LiveSync commits — QTextEdit fires textChanged
    // on every keystroke; coalesce bursts into a single commit on timeout.
    QTimer* m_commentsCommitTimer = nullptr;

    void recalcPower();
};

// ─── Flow layout (cards wrap left→right, then down) ─────────────────────────
class FlowLayout : public QLayout
{
public:
    explicit FlowLayout(QWidget* parent = nullptr, int margin = -1,
                        int hSpacing = -1, int vSpacing = -1);
    ~FlowLayout() override;

    void addItem(QLayoutItem* item) override;
    int  horizontalSpacing() const;
    int  verticalSpacing()  const;
    Qt::Orientations expandingDirections() const override;
    bool hasHeightForWidth() const override;
    int  heightForWidth(int) const override;
    int  count() const override;
    QLayoutItem* itemAt(int index) const override;
    QSize minimumSize() const override;
    void  setGeometry(const QRect& rect) override;
    QSize sizeHint() const override;
    QLayoutItem* takeAt(int index) override;

private:
    int doLayout(const QRect& rect, bool testOnly) const;

    QList<QLayoutItem*> m_items;
    int m_hSpace;
    int m_vSpace;
};

// ─── Main sensory panel (embeds in MainWindow central area) ─────────────────
class SensoryPanel : public QWidget
{
    Q_OBJECT

public:
    explicit SensoryPanel(DatabaseManager* db, QWidget* parent = nullptr);

    // ── LiveSync routing (optional; injected from MainWindow) ────────────────
    // When set, per-cell commits from SampleCards are forwarded to LiveSync,
    // and remote cell-changed signals are applied back to the visible card
    // with signals blocked. Null in standalone/test usage.
    void setLiveSync(LiveSync* sync);

    // v2.0.5: for each session loaded from disk (id <= 0), look up the
    // natural key (session_name, tester_name, date) in the database and
    // inherit id+version if it already exists. Public so MainWindow can
    // call it before the file-level save loop in onUpdateDatabase —
    // without this prep, re-importing the same Excel/JSON file fails
    // with UNIQUE-violation against the prior import's DB row.
    void inheritExistingIdsAndVersions();

    // v2.1.0+: after MainWindow's save loop runs, pull id/version/
    // originalSessionName updates that the byRef tryWriteSensorySession()
    // back-fill wrote into the local copy back into m_sessions. Indexed in
    // parallel (saved[i] mirrors m_sessions[i] since allSessions() returns a
    // vector copy and we don't add/remove sessions during the save loop).
    // Without this, a Test Title rename routed through INSERT lands in DB
    // with a new id, but m_sessions[i].id remains the pre-rename row's id,
    // so the next Ctrl+U tick treats the same edits as another rename and
    // duplicates the row.
    void syncSavedSessionState(const QVector<SensorySession>& saved);

    // ── Session management (called by MainWindow) ────────────────────────────
    void loadSessions(const QVector<SensorySession>& sessions);
    void selectSession(int index);
    void showAveragedChart(const QVector<int>& sessionIndices);
    void newSession();
    void closeSessions(const QVector<int>& indices);
    void renameSession(int index, const QString& newLabel);

    // v2.5.0 RC5 (close options): drop a single session from the panel — used
    // by MainWindow's "Discard Session" close option when the user opts to
    // delete an unnamed session's data rather than name it. Removes it from
    // m_sessions and fixes up the current-index bookkeeping (clamp/clear),
    // re-applying whatever session remains current. Disk autosave files are
    // intentionally NOT touched (the user is told the .xlsx is left in place);
    // only the in-memory session + DB row (removed by the caller) go away.
    void discardSession(int index);

    // v2.5.0 RC5 (close options): make `index` the current session and put the
    // keyboard focus on the Test Title field — drives the "Name It Now" close
    // option so the user lands directly on the field that needs filling in.
    void focusTitleForSession(int index);

    // ── File operations (called from MainWindow ribbon) ──────────────────────
    void save();
    void loadFile(const QString& path);
    void loadFiles();
    void loadFromDatabase();

    // ── Report generation (called from MainWindow ribbon) ────────────────────
    void generateFullReport();
    void generateStats();

    // ── Session access ───────────────────────────────────────────────────────
    QVector<SensorySession> allSessions();
    int  currentSessionIndex() const { return m_currentTesterIdx; }
    QString sessionLabel(const SensorySession& s) const;
    SensorySession* currentSession();

    // v2.4.2 SP2-T4 (reset-to-5 keystone, Defect-1/2 fix): overlay ONLY the
    // per-metric kSensoryMetrics scalar scores from a pre-computed merged blob
    // (mergeSensoryPreservingDbScores output) onto the CURRENT in-memory
    // session, then re-render the visible cards from it. Mirrors the blessed
    // overlay in dbAuthoritativeSessions: it touches nothing but scores, so
    // id/version/imagePaths/imageLayouts/imageCrops/imageIds/imageVersions and
    // every other field stay in-memory-authoritative (no fromJson round-trip,
    // which would drop images and reset the persistence anchors). Re-rendering
    // here — BEFORE any caller-side navigator refresh / allSessions() flush —
    // means a subsequent buildSession() reads the merged widgets and cannot
    // revert the adopted remote value. No-op if no current session.
    void applyMergedScoresToCurrentSession(const QJsonObject& mergedSession);

    // Plan C C10: true once the panel's sessions have been saved to a disk
    // file at least this run. The consolidated close prompt's "Save All" uses
    // this to decide whether untitled work still needs a Save-As (the DB save
    // itself is handled separately by MainWindow::onUpdateDatabase).
    bool hasSavePath() const { return !m_savePath.isEmpty(); }

    // DATAVIEWER-8: true iff the CURRENTLY-displayed session has the non-empty
    // test name + tester needed to be persisted. MainWindow gates the program-
    // close disk-courtesy save() on this so app shutdown never trips save()'s
    // hard-guard modal for an unnamed current session.
    bool currentSessionSavable() const;

    // ── Averaged table overlay (called by MainWindow when test avg selected) ─
    void showAveragedTable(const QStringList& deviceNames,
                           const QVector<QMap<QString, double>>& deviceAvgs);
    void showNormalView();

    // ── Static combined PPTX ─────────────────────────────────────────────────
    static bool generateCombinedPptx(const QVector<SensorySession>& sessions,
                                      const QString& filePath,
                                      QString& errorOut);

signals:
    // Structural changes only: new/close/rename/load/save and add/remove
    // sample. Per-FIELD value edits (scores, comments, header text) do NOT
    // emit this — they go through scheduleChartRefresh and never reach
    // MainWindow. Use dataEdited() for those.
    void sessionsChanged();

    // Plan C (C6 fix): fires on EVERY per-field edit that mutates the
    // in-memory session — score sliders, comments, sample names, and the
    // header fields (title/assessor/tester/media/round). MainWindow connects
    // this to RecoveryManager::noteDirty() so routine data entry is captured
    // by the crash snapshot (sessionsChanged alone misses all of it). Distinct
    // from sessionsChanged so the structural consumers (navigator refresh,
    // averages, properties) are not re-run on every keystroke.
    void dataEdited();

private slots:
    // v2.0.1: applied when LiveSync receives a remote per-cell change.
    void onRemoteCellChanged(const QString& table, qint64 rowId,
                             const QString& column, const QVariant& newValue);

private:
    void buildHeaderRow(QWidget* container);

    // Patch the in-memory sample with a JSON-path field update from LiveSync.
    void applyRemoteFieldToSample(SensorySample& s,
                                  const QString& fieldPath,
                                  const QVariant& value);

    void addSampleCard(const SensorySample& sample = SensorySample{});
    void clearAllCards();

    // Returns the persisted id of the currently-selected session, or -1 if
    // the selection is out of range or the session hasn't been persisted
    // yet. Used to gate LiveSync commits — placeholder rows must not
    // broadcast.
    int activeSessionId() const;

    SensorySession buildSession() const;

    // DATAVIEWER-4: return DB-authoritative copies of `inMem` for export. Flushes
    // LiveSync first so the DB holds this client's latest per-cell scores, then
    // overlays DB scores onto in-memory metadata via mergeSensoryPreservingDbScores.
    // Sessions with id <= 0 (never saved) or when no DB is available pass through.
    QVector<SensorySession> dbAuthoritativeSessions(const QVector<SensorySession>& inMem);

    void           applySession(const SensorySession& session);
    void           saveCurrentTester();
    bool           isDefaultState() const;

    // inheritExistingIdsAndVersions() declared public above (v2.0.5).

    void scheduleChartRefresh();
    void onRefreshChart();
    void onAddSample();
    void onRemoveCard(SampleCard* card);
    void onSaveChart();

    void writeStatsCsv(const QString& path);

    void saveToJson(const QString& path, const SensorySession& sess);
    void saveToExcel(const QString& path, const SensorySession& sess);

    int loadExcelSavedFormat(QXlsx::Document& xlsx, const QStringList& excelMetrics,
                             const QString& testTitle, const QString& filePath);
    int loadExcelStandardFormat(QXlsx::Document& xlsx, const QStringList& excelMetrics,
                                const QString& testTitle);

    // ── UI elements ──────────────────────────────────────────────────────────
    QLineEdit*        m_testTitleEdit;
    QLineEdit*        m_assessorEdit;
    QLineEdit*        m_testerEdit;
    QComboBox*        m_roundCombo = nullptr;   // Bug 2: double-blind round (1/2/N/A)
    QLineEdit*        m_mediaEdit;
    QLabel*           m_dateLabel;
    QPushButton*      m_saveHeadersBtn = nullptr;
    QPushButton*      m_addSampleBtn   = nullptr;
    QSplitter*        m_splitter;
    QScrollArea*      m_scrollArea;
    FlowLayout*       m_flowLayout;
    QWidget*          m_flowContainer;
    RadarChartWidget* m_chart;
    QVector<SampleCard*> m_cards;

    // ── Averaged table overlay (replaces cards when test avg selected) ──
    QStackedWidget*   m_leftStack       = nullptr;
    QTableWidget*     m_avgOverlayTable = nullptr;

    // ── Session data ─────────────────────────────────────────────────────────
    QVector<SensorySession> m_sessions;
    int                     m_currentTesterIdx = -1;

    // ── Debounce timer ───────────────────────────────────────────────────────
    QTimer* m_refreshTimer;

    // ── Save path (set once on first save, reused on Ctrl+S) ─────────────────
    QString m_savePath;

    // ── Database ─────────────────────────────────────────────────────────────
    DatabaseManager* m_db;

    // ── LiveSync (optional; owned by MainWindow; null in tests/offline) ──
    LiveSync* m_liveSync = nullptr;

    // ── Last browse directory ────────────────────────────────────────────────
    QString m_lastBrowseDir;
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
};

} // namespace DVE
