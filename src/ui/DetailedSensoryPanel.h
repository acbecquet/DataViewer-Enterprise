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

class DetailedSensoryPanel : public QWidget
{
    Q_OBJECT

public:
    explicit DetailedSensoryPanel(DatabaseManager* db, QWidget* parent = nullptr);

    void loadSessions(const QVector<DetailedSensorySession>& sessions);
    void selectSession(int index);
    void showAveragedChart(const QVector<int>& sessionIndices);
    void newSession();
    void closeSessions(const QVector<int>& indices);
    void renameSession(int index, const QString& newLabel);

    void save();
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

private:
    void buildHeaderRow(QWidget* container);
    void buildQuestionForm();
    void buildSampleNavBar();

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
    QLineEdit*        m_mouthpieceEdit = nullptr;

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

    QString m_lastBrowseDir;
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
};

} // namespace DVE
