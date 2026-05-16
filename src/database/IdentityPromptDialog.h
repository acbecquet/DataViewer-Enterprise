#pragma once

#include <QDialog>
#include <QHash>

class QLineEdit;
class QButtonGroup;

namespace DVE {

class IdentityManager;
class PostgresConnection;

class IdentityPromptDialog : public QDialog {
    Q_OBJECT
public:
    // `conn` is optional; when provided the dialog queries the presence
    // table for currently-active user_color values and visually marks
    // those swatches as "taken" (the user can still pick them).
    explicit IdentityPromptDialog(IdentityManager* mgr,
                                  PostgresConnection* conn = nullptr,
                                  QWidget* parent = nullptr);

private slots:
    void onAccept();

private:
    IdentityManager*    m_mgr;
    PostgresConnection* m_conn = nullptr;
    QLineEdit*          m_nameEdit;
    QButtonGroup*       m_colorGroup;
    QString             m_selectedColor;

    // Returns lowercased color hex -> comma-joined names of active users
    // holding that color. Empty if no connection or query failed.
    QHash<QString, QString> queryTakenColors() const;
};

} // namespace DVE
