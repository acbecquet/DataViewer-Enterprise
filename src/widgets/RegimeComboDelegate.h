#pragma once
#include "CellFocusDelegate.h"
#include <QStringList>

namespace DVE {

// Editable-combo editor for the per-row Puffing Regime column. Subclasses
// CellFocusDelegate so the remote-focus border / flash painting is preserved.
// When inactive (old-template sheet), falls back to the base (plain) editor.
class RegimeComboDelegate : public CellFocusDelegate {
    Q_OBJECT
public:
    explicit RegimeComboDelegate(QObject* parent = nullptr);
    void setActive(bool a) { m_active = a; }
    void setRegimes(const QStringList& r) { m_regimes = r; }

    QWidget* createEditor(QWidget* parent, const QStyleOptionViewItem& opt,
                          const QModelIndex& idx) const override;
    void setEditorData(QWidget* editor, const QModelIndex& idx) const override;
    void setModelData(QWidget* editor, QAbstractItemModel* model,
                      const QModelIndex& idx) const override;
private:
    bool        m_active = false;
    QStringList m_regimes;
};

} // namespace DVE
