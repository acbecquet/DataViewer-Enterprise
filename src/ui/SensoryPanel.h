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

#include "pipeline/SensoryData.h"
#include "ui/RadarChartWidget.h"
#include "database/DatabaseManager.h"

namespace QXlsx { class Document; }

namespace DVE {

// ─── Per-sample card widget ────────────────────────────────────────────────────
class SampleCard : public QGroupBox
{
    Q_OBJECT

public:
    explicit SampleCard(int index, QWidget* parent = nullptr);

    SensorySample toSample() const;
    void          fromSample(const SensorySample& s);

signals:
    void changed();
    void removeRequested(SampleCard* card);

private:
    QLineEdit* m_nameEdit;
    QMap<QString, QDoubleSpinBox*> m_spinBoxes;  // all entries are NoWheelDoubleSpinBox instances (see SensoryPanel.cpp)
    QTextEdit* m_commentsEdit;

    // Per-sample device properties
    QLineEdit* m_voltageEdit;
    QLineEdit* m_resistanceEdit;
    QComboBox* m_heatingTechCombo;
    QLabel*    m_powerLabel;

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

    // ── Session management (called by MainWindow) ────────────────────────────
    void loadSessions(const QVector<SensorySession>& sessions);
    void selectSession(int index);
    void showAveragedChart(const QVector<int>& sessionIndices);
    void newSession();
    void closeSessions(const QVector<int>& indices);
    void renameSession(int index, const QString& newLabel);

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

    // ── Averaged table overlay (called by MainWindow when test avg selected) ─
    void showAveragedTable(const QStringList& deviceNames,
                           const QVector<QMap<QString, double>>& deviceAvgs);
    void showNormalView();

    // ── Static combined PPTX ─────────────────────────────────────────────────
    static bool generateCombinedPptx(const QVector<SensorySession>& sessions,
                                      const QString& filePath,
                                      QString& errorOut);

signals:
    void sessionsChanged();

private:
    void buildHeaderRow(QWidget* container);

    void addSampleCard(const SensorySample& sample = SensorySample{});
    void clearAllCards();

    SensorySession buildSession() const;
    void           applySession(const SensorySession& session);
    void           saveCurrentTester();
    QString        resolveTestName();
    bool           isDefaultState() const;

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
    QLineEdit*        m_mediaEdit;
    QLabel*           m_dateLabel;
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

    // ── Last browse directory ────────────────────────────────────────────────
    QString m_lastBrowseDir;
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
};

} // namespace DVE
