#include "RegimeComboDelegate.h"
#include "../pipeline/RegimeUtils.h"
#include <QComboBox>

namespace DVE {

static const QStringList kPresetRegimes = {
    "60mL/3s/30s", "200mL/10s/60s", "100mL/2.5s/15s", "60mL/2s/5s", "200mL/3s/30s"
};

RegimeComboDelegate::RegimeComboDelegate(QObject* parent)
    : CellFocusDelegate(parent) {}

QWidget* RegimeComboDelegate::createEditor(QWidget* parent, const QStyleOptionViewItem& opt,
                                           const QModelIndex& idx) const
{
    if (!m_active)
        return CellFocusDelegate::createEditor(parent, opt, idx);  // plain editor for old template
    QComboBox* cb = new QComboBox(parent);
    cb->setEditable(true);
    QStringList items = m_regimes;
    for (const QString& p : kPresetRegimes)
        if (!items.contains(p)) items << p;
    cb->addItems(items);
    return cb;
}

void RegimeComboDelegate::setEditorData(QWidget* editor, const QModelIndex& idx) const
{
    if (auto* cb = qobject_cast<QComboBox*>(editor)) {
        cb->setCurrentText(idx.data(Qt::EditRole).toString());
        return;
    }
    CellFocusDelegate::setEditorData(editor, idx);
}

void RegimeComboDelegate::setModelData(QWidget* editor, QAbstractItemModel* model,
                                       const QModelIndex& idx) const
{
    if (auto* cb = qobject_cast<QComboBox*>(editor)) {
        // Enforce the standard like the Notes-panel regime field: canonicalise
        // (CORESTA -> 60mL/3s/30s) and write only if it matches the parametric
        // format; otherwise leave the model's existing value (the cell reverts).
        const QString canon = RegimeUtils::canonicalRegime(cb->currentText());
        if (RegimeUtils::isStandardRegimeFormat(canon))
            model->setData(idx, canon, Qt::EditRole);
        return;
    }
    CellFocusDelegate::setModelData(editor, model, idx);
}

} // namespace DVE
