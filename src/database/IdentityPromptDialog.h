#pragma once

#include <QDialog>

class QLineEdit;
class QButtonGroup;

namespace DVE {

class IdentityManager;

class IdentityPromptDialog : public QDialog {
    Q_OBJECT
public:
    explicit IdentityPromptDialog(IdentityManager* mgr, QWidget* parent = nullptr);

private slots:
    void onAccept();

private:
    IdentityManager* m_mgr;
    QLineEdit*       m_nameEdit;
    QButtonGroup*    m_colorGroup;
    QString          m_selectedColor;
};

} // namespace DVE
