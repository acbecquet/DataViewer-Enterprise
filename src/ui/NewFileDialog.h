#pragma once

#include <QDialog>
#include <QListWidget>
#include <QStringList>

namespace DVE {

// ---------------------------------------------------------------------------
// NewFileDialog
//
// Shows a scrollable list of test sheets from the template with checkboxes.
// The user selects which tests to include, then confirms. The caller is
// responsible for getting a save path and running the creation logic.
// ---------------------------------------------------------------------------
class NewFileDialog : public QDialog
{
    Q_OBJECT
public:
    explicit NewFileDialog(const QStringList& availableTests,
                           QWidget* parent = nullptr);

    // Returns the subset of tests that were checked on Accept.
    QStringList selectedTests() const;

private slots:
    void onSelectAll();
    void onDeselectAll();

private:
    QListWidget* m_list;
};

} // namespace DVE
