#pragma once

#include <QDialog>
#include <QTreeWidget>
#include <QLabel>
#include <QPushButton>
#include <QLineEdit>
#include <QComboBox>
#include "../database/DatabaseManager.h"

namespace DVE {

class DatabaseBrowserDialog : public QDialog
{
    Q_OBJECT

public:
    explicit DatabaseBrowserDialog(DatabaseManager* db, QWidget* parent = nullptr);

    // Returns the file IDs the user chose to load (empty if cancelled).
    QVector<int> selectedFileIds() const { return m_selectedIds; }

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

private:
    void populateTree(const QString& filter = QString());
    int  idFromItem(QTreeWidgetItem* item) const;

    DatabaseManager*   m_db;
    QTreeWidget*       m_tree;
    QLineEdit*         m_filterEdit;
    QLabel*            m_statusLabel;
    QPushButton*       m_loadBtn;
    QPushButton*       m_loadAllBtn;
    QPushButton*       m_deleteBtn;

    QVector<FileRecord> m_allRecords;
    QVector<int>        m_selectedIds;
};

} // namespace DVE
