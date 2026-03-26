#pragma once

#include <QDialog>
#include <QScrollArea>
#include <QVBoxLayout>
#include <QGroupBox>
#include <QLineEdit>
#include <QSpinBox>
#include <QTextEdit>
#include <QLabel>
#include <QSplitter>
#include <QToolBar>
#include <QMenu>
#include <QTimer>
#include <QShortcut>
#include <QVector>
#include <QComboBox>
#include <QListWidget>
#include <QTreeWidget>

#include "pipeline/SensoryData.h"
#include "ui/RadarChartWidget.h"
#include "database/DatabaseManager.h"

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
    QMap<QString, QSpinBox*> m_spinBoxes;
    QTextEdit* m_commentsEdit;
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

// ─── Main sensory dialog ───────────────────────────────────────────────────────
class SensoryDialog : public QDialog
{
    Q_OBJECT

public:
    explicit SensoryDialog(DatabaseManager* db, QWidget* parent = nullptr);
    explicit SensoryDialog(DatabaseManager* db, QWidget* parent,
                           const SensorySession& initialSession);
    explicit SensoryDialog(DatabaseManager* db, QWidget* parent,
                           const QVector<SensorySession>& sessions);

    // Generate a combined PPTX report with one slide per session.
    static bool generateCombinedReport(const QVector<SensorySession>& sessions,
                                       const QString& filePath,
                                       QString& errorOut);

private slots:
    void onSave();
    void onNewSession();
    void onLoadExcel();
    void onLoadFromDatabase();
    void onRefreshChart();
    void onAddSample();
    void onRemoveCard(SampleCard* card);
    void onGeneratePptx();
    void onGenerateStats();
    void onTesterChanged(int index);
    void onNavigatorClicked(int row);

private:
    void init();
    void buildToolBar();
    void buildHeaderRow(QWidget* container);

    void addSampleCard(const SensorySample& sample = SensorySample{});
    void clearAllCards();

    SensorySession buildSession() const;
    void           applySession(const SensorySession& session);
    void           saveCurrentTester();
    void           populateTesterCombo();
    void           refreshNavigator();
    QString        resolveTestName();
    QString        sessionLabel(const SensorySession& s) const;
    bool           isDefaultState() const;

    void scheduleChartRefresh();

    QByteArray renderChartToImage(int width, int height) const;
    void writeStatsCsv(const QString& path);

    void saveToJson(const QString& path, const SensorySession& sess);
    void saveToExcel(const QString& path, const SensorySession& sess);

    // ── UI elements ──────────────────────────────────────────────────────────
    QToolBar*         m_toolbar;
    QLineEdit*        m_testTitleEdit;
    QLineEdit*        m_assessorEdit;
    QLineEdit*        m_testerEdit;
    QLineEdit*        m_mediaEdit;
    QLabel*           m_dateLabel;
    QComboBox*        m_testerCombo;
    QListWidget*      m_navigator;
    QWidget*          m_navPanel;
    QSplitter*        m_splitter;
    QScrollArea*      m_scrollArea;
    FlowLayout*       m_flowLayout;
    QWidget*          m_flowContainer;
    RadarChartWidget* m_chart;
    QVector<SampleCard*> m_cards;

    // ── Multi-tester sessions ────────────────────────────────────────────────
    QVector<SensorySession> m_sessions;
    int                     m_currentTesterIdx = -1;

    // ── Debounce timer ───────────────────────────────────────────────────────
    QTimer* m_refreshTimer;

    // ── Save path (set once on first save, reused on Ctrl+S) ─────────────────
    QString m_savePath;  // base path without extension

    // ── Database ─────────────────────────────────────────────────────────────
    DatabaseManager* m_db;

    // ── Last browse directory ────────────────────────────────────────────────
    QString m_lastBrowseDir;
    QString lastBrowseDir() const;
    void    setLastBrowseDir(const QString& filePath);
};

} // namespace DVE
