#pragma once

#include <QDialog>
#include <QVariantMap>

class QTableWidget;
class QButtonGroup;

namespace DVE {

class VersionMismatchDialog : public QDialog {
    Q_OBJECT
public:
    enum BulkChoice { Mine, Theirs, Selection };

    VersionMismatchDialog(const QString& tableLabel,
                           const QString& whoChanged,
                           const QString& howLongAgo,
                           const QVariantMap& mine,
                           const QVariantMap& theirs,
                           QWidget* parent = nullptr);

    BulkChoice  bulkChoice() const { return m_choice; }
    QVariantMap selectedValues() const;

private slots:
    void onKeepMine();
    void onTakeTheirs();
    void onApplySelection();

private:
    QTableWidget*           m_table;
    QVariantMap             m_mine;
    QVariantMap             m_theirs;
    BulkChoice              m_choice = Selection;
    // For each field, m_radios[field] holds a QButtonGroup whose checked button
    // has data 0=mine, 1=theirs.
    QHash<QString, QButtonGroup*> m_radios;
};

} // namespace DVE
