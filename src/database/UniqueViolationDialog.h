#pragma once

#include <QDialog>
#include <QString>

class QLineEdit;

namespace DVE {

class UniqueViolationDialog : public QDialog {
    Q_OBJECT
public:
    enum Choice { OpenExisting, RenameMine, CancelChoice };

    UniqueViolationDialog(const QString& description,
                          const QString& suggestedRename,
                          QWidget* parent = nullptr);

    Choice  choice() const { return m_choice; }
    QString renamedName() const;

private slots:
    void onOpenExisting() { m_choice = OpenExisting; accept(); }
    void onRename();
    void onCancel()       { m_choice = CancelChoice; reject(); }

private:
    QLineEdit* m_renameEdit;
    Choice     m_choice = CancelChoice;
};

} // namespace DVE
