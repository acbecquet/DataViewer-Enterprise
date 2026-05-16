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

class SaveCoordinator;
class LiveSync;

class DetailedSensoryPanel : public QWidget
{
    Q_OBJECT

public:
    explicit DetailedSensoryPanel(DatabaseManager* db, QWidget* parent = nullptr);

    // ── Save routing ─────────────────────────────────────────────────────────
    // Optional. When set, save call sites route through the coordinator so
    // conflicts (VersionMismatch / RowDeleted / UniqueViolation) surface
    // dialogs to the user. When null (standalone / test usage) the panel
    // falls back to the legacy bool saveDetailedSensorySession path.
    void setSaveCoordinator(SaveCoordinator* coord) { m_saveCoord = coord; }

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

    void save();
    void loadFile(const QString& path);
    void loadFiles();
    void loadFromDatabase();

    void generateFullReport();

    QVector<DetailedSensorySession> allSessions();
    int  currentSessionIndex() const { return m_currentTesterIdx; }
    QString sessionLabel(const DetailedSensorySession& s) const;
    DetailedSensorySession* currentSession();

    void showAveragedTable(const QStringList& deviceNames,
                           const QVector<QMap<QString, double>>& deviceAvgs);
    void showNormalView();

    static bool generateCombinedPptx(const QVector<DetailedSensorySession>& sessions,
                                      const QString& filePath,
                                      QString& errorOut);

signals:
    void sessionsChanged();

private slots:
    // v2.0.1: applied when LiveSync receives a remote per-cell change.
    void onRemoteCellChanged(const QString& table, qint64 rowId,
                             const QString& column, const QVariant& newValue);

private:
    void buildHeaderRow(QWidget* container);
    void buildQuestionForm();
    void buildSampleNavBar();

    // v2.0.1: forward a per-cell commit on the active sample to LiveSync.
    // sampleField=true paths get prefixed with "samples[currentSampleIdx].";
    // sampleField=false paths are top-level session fields. No-op when
    // LiveSync isn't wired or the session hasn't been persisted yet.
    void emitCellCommit(const QString& fieldPath, const QVariant& value,
                        bool sampleField);

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

    QTimer* m_refreshTimer;
    QString m_savePath;
    DatabaseManager* m_db;

    // Save coordinator (optional; nullptr falls back to bool save).
    // Plain pointer — see comment in SensoryPanel.h.
    SaveCoordinator* m_saveCoord = nullptr;

    // LiveSync (optional; owned by MainWindow; null in tests/offline).
    LiveSync* m_liveSync = nullptr;

    QString m_lastBrowseDir;
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
};

} // namespace DVE
