#include "IdentityPromptDialog.h"
#include "IdentityManager.h"
#include "PostgresConnection.h"

#include <QVBoxLayout>
#include <QHBoxLayout>
#include <QGridLayout>
#include <QLabel>
#include <QLineEdit>
#include <QPushButton>
#include <QButtonGroup>
#include <QToolButton>
#include <QMessageBox>
#include <QHash>
#include <QSqlQuery>
#include <QSqlError>
#include <QVariant>
#include <QDebug>

namespace DVE {

IdentityPromptDialog::IdentityPromptDialog(IdentityManager* mgr,
                                           PostgresConnection* conn,
                                           QWidget* parent)
    : QDialog(parent), m_mgr(mgr), m_conn(conn) {
    setWindowTitle(tr("Welcome to DataViewer"));
    setModal(true);
    // Intentional UX choice: dismissing this dialog without filling both
    // fields (Esc/X) leaves firstLaunchPending() true so MainWindow will
    // re-prompt on next launch. The spec described this as "shown once" —
    // the broader contract ("show until identity is set") is the safer
    // fail-loud version of the same intent.

    auto* root = new QVBoxLayout(this);

    root->addWidget(new QLabel(
        tr("Other users will see your name and color when you have files "
           "or sessions open. Pick a name and a color now — you can change "
           "either later from Tools → Identity.")));

    auto* nameRow = new QHBoxLayout;
    nameRow->addWidget(new QLabel(tr("Display name:")));
    m_nameEdit = new QLineEdit;
    m_nameEdit->setPlaceholderText(tr("e.g. Sarah Chen"));
    m_nameEdit->setMaxLength(40);
    nameRow->addWidget(m_nameEdit);
    root->addLayout(nameRow);

    root->addWidget(new QLabel(tr("Pick a color:")));
    auto* grid = new QGridLayout;
    m_colorGroup = new QButtonGroup(this);
    const QStringList palette = IdentityManager::defaultColorPalette();
    const QHash<QString, QString> taken = queryTakenColors();
    for (int i = 0; i < palette.size(); ++i) {
        auto* btn = new QToolButton;
        btn->setCheckable(true);
        btn->setMinimumSize(36, 36);
        const QString key = palette[i].toLower();
        const auto it = taken.constFind(key);
        if (it != taken.constEnd()) {
            btn->setStyleSheet(QString(
                "QToolButton { background:%1; border:3px solid #888888; "
                "border-radius:18px; } "
                "QToolButton:checked { border:3px solid #000; }"
            ).arg(palette[i]));
            btn->setToolTip(tr("Taken by %1").arg(it.value()));
        } else {
            btn->setStyleSheet(QString(
                "QToolButton { background:%1; border:2px solid #00000022; "
                "border-radius:18px; } "
                "QToolButton:checked { border:3px solid #000; }"
            ).arg(palette[i]));
        }
        btn->setProperty("dve_color_hex", palette[i]);
        m_colorGroup->addButton(btn, i);
        grid->addWidget(btn, i / 6, i % 6);
    }
    root->addLayout(grid);

    auto* buttons = new QHBoxLayout;
    buttons->addStretch();
    auto* okBtn = new QPushButton(tr("Save"));
    okBtn->setDefault(true);
    buttons->addWidget(okBtn);
    root->addLayout(buttons);

    connect(okBtn, &QPushButton::clicked,
            this, &IdentityPromptDialog::onAccept);
}

void IdentityPromptDialog::onAccept() {
    const QString name = m_nameEdit->text().trimmed();
    if (name.isEmpty()) {
        QMessageBox::warning(this, tr("Name required"),
                             tr("Please enter a display name."));
        return;
    }
    auto* picked = m_colorGroup->checkedButton();
    if (!picked) {
        QMessageBox::warning(this, tr("Color required"),
                             tr("Please pick a color."));
        return;
    }
    m_selectedColor = picked->property("dve_color_hex").toString();
    m_mgr->setDisplayName(name);
    m_mgr->setColor(m_selectedColor);
    accept();
}

QHash<QString, QString> IdentityPromptDialog::queryTakenColors() const
{
    QHash<QString, QString> taken;
    if (!m_conn || !m_conn->isOpen()) return taken;

    QSqlQuery q(m_conn->queryDb());
    // matches presence heartbeat freshness window (see init.sql)
    if (!q.exec("SELECT user_color, string_agg(DISTINCT user_name, ', ') "
                "FROM presence "
                "WHERE last_heartbeat > now() - interval '30 seconds' "
                "GROUP BY user_color")) {
        qDebug() << "IdentityPromptDialog: queryTakenColors failed:"
                 << q.lastError().text();
        return taken;
    }
    while (q.next()) {
        const QString color = q.value(0).toString().trimmed().toLower();
        const QString names = q.value(1).toString().trimmed();
        if (!color.isEmpty()) taken.insert(color, names);
    }
    return taken;
}

} // namespace DVE
