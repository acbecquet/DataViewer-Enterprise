#include "RemoteCellHelpers.h"

#include <QTableWidget>
#include <QTableWidgetItem>

namespace DVE {

bool applyRemoteValueToCell(QTableWidget* table, int row, int col,
                            const QString& newValue)
{
    if (!table) return false;
    if (row < 0 || row >= table->rowCount()) return false;
    if (col < 0 || col >= table->columnCount()) return false;
    QTableWidgetItem* it = table->item(row, col);
    if (!it) return false;

    // C6: never clobber a cell the user is actively editing. The yellow
    // remote-conflict decoration that handleRemoteRowChange already painted
    // surfaces the conflict; onDataTableItemClicked lets the user accept.
    if (it->data(Qt::UserRole + 3).toBool()) return false;

    // Block itemChanged so the remote-applied text isn't echoed back
    // via commitCell. Treat the new value as the baseline so the cell
    // shows clean afterward.
    const bool wasBlocked = table->blockSignals(true);
    it->setText(newValue);
    it->setData(Qt::UserRole + 2, newValue);
    it->setData(Qt::UserRole + 3, false);
    table->blockSignals(wasBlocked);
    return true;
}

} // namespace DVE
