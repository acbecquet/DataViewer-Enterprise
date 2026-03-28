#pragma once

#include <QDialog>
#include <QListWidget>
#include <QTextBrowser>
#include <QVector>

namespace DVE {

struct SopEntry {
    QString test;
    QString sop;
    QString objective;
    QString passCriteria;
    QString equipment;
    QString quantity;
    QString estDuration1mL;
    QString estDuration2mL;
    QString note;
};

class SopDialog : public QDialog
{
    Q_OBJECT

public:
    explicit SopDialog(const QString& xlsxPath, QWidget* parent = nullptr);

private:
    void loadFromExcel(const QString& xlsxPath);
    void onTestSelected(int row);

    QListWidget*       m_testList;
    QTextBrowser*      m_detail;
    QVector<SopEntry>  m_entries;
};

} // namespace DVE
