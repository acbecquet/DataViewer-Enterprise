#pragma once

#include <QDialog>
#include <QVariantMap>

namespace DVE {

class RowDeletedDialog : public QDialog {
    Q_OBJECT
public:
    enum Choice { Recreate, Discard };

    RowDeletedDialog(const QString& tableLabel,
                     const QString& whoDeleted,
                     const QVariantMap& myUnsavedEdits,
                     QWidget* parent = nullptr);

    Choice choice() const { return m_choice; }

private slots:
    void onRecreate() { m_choice = Recreate; accept(); }
    void onDiscard()  { m_choice = Discard;  accept(); }

private:
    Choice m_choice = Discard;
};

} // namespace DVE
