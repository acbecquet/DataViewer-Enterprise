#pragma once

#include <QDialog>
#include <QTreeWidget>
#include <QTabWidget>
#include <QLabel>
#include <QPushButton>
#include <QLineEdit>
#include <QComboBox>
#include "../database/DatabaseManager.h"
#include "../pipeline/DetailedSensoryData.h"

namespace DVE {

class DatabaseBrowserDialog : public QDialog
{
    Q_OBJECT

public:
    explicit DatabaseBrowserDialog(DatabaseManager* db, QWidget* parent = nullptr);

    // Returns the file IDs the user chose to load (empty if cancelled).
    QVector<int> selectedFileIds() const { return m_selectedIds; }

    // Returns the sensory session IDs selected (empty if none).
    QVector<int> selectedSensoryIds() const { return m_selectedSensoryIds; }

    // Returns true if the user selected from the sensory tab.
    bool isSensorySelection() const { return m_sensorySelection; }

    QVector<int> selectedDetailedSensoryIds() const { return m_selectedDetSensIds; }
    bool isDetailedSensorySelection() const { return m_detSensSelection; }

private slots:
    void onRefresh();
    void onFilterChanged(const QString& text);
    void onLoad();
    void onLoadAll();
    void onDelete();
    void onCleanup();
    void onSelectionChanged();
    void onItemDoubleClicked(QTreeWidgetItem* item, int column);
    void onItemExpanded(QTreeWidgetItem* item);
    void onTabChanged(int index);

    // Sensory tab slots
    void onSensorySelectionChanged();
    void onSensoryDoubleClicked(QTreeWidgetItem* item, int column);
    void onSensoryDelete();
    void onSensoryGenerateReport();

private:
    void populateTree(const QString& filter = QString());
    void populateSensoryTree(const QString& filter = QString());
    int  idFromItem(QTreeWidgetItem* item) const;

    DatabaseManager*   m_db;
    QTabWidget*        m_tabWidget;

    // TPM Data tab
    QTreeWidget*       m_tree;
    QLineEdit*         m_filterEdit;
    QLabel*            m_statusLabel;
    QPushButton*       m_loadBtn;
    QPushButton*       m_loadAllBtn;
    QPushButton*       m_deleteBtn;

    // Sensory Data tab
    QTreeWidget*       m_sensoryTree;
    QLineEdit*         m_sensoryFilterEdit;
    QLabel*            m_sensoryStatusLabel;
    QPushButton*       m_sensoryLoadBtn;
    QPushButton*       m_sensoryDeleteBtn;
    QPushButton*       m_sensoryReportBtn;

    QVector<FileRecord>   m_allRecords;
    QVector<SensoryRecord> m_sensoryRecords;
    QVector<int>          m_selectedIds;
    QVector<int>          m_selectedSensoryIds;
    bool                  m_sensorySelection  = false;

    // Detailed Sensory Data tab
    QTreeWidget*       m_detSensTree = nullptr;
    QLineEdit*         m_detSensFilterEdit = nullptr;
    QLabel*            m_detSensStatusLabel = nullptr;
    QPushButton*       m_detSensLoadBtn = nullptr;
    QPushButton*       m_detSensDeleteBtn = nullptr;
    QPushButton*       m_detSensReportBtn = nullptr;
    QVector<DetailedSensoryRecord> m_detSensRecords;
    QVector<int>       m_selectedDetSensIds;
    bool               m_detSensSelection = false;

    void populateDetailedSensoryTree(const QString& filter = {});
    void onDetailedSensoryDelete();
    void onDetailedSensoryGenerateReport();
    void onDetailedSensorySelectionChanged();
};

} // namespace DVE
