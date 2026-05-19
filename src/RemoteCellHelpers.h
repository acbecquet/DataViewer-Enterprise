#pragma once

#include <QString>

class QTableWidget;

namespace DVE {

// Apply a remote-arrived value to a cell, honoring the dirty-cell guard
// (UserRole+3). Used by MainWindow::onRemoteCellChanged so a remote update
// to a cell the user is actively editing does NOT overwrite their
// in-progress edit — handleRemoteRowChange has already painted the yellow
// remote-conflict decoration on the dirty cell, and the click-to-accept
// path (onDataTableItemClicked) lets the user take the remote value.
//
// Returns true if the cell text was changed, false if skipped (dirty
// cell, missing item, or out-of-range row/col). The function blocks
// itemChanged signals while it writes so onDataTableItemChanged does not
// echo the write back through LiveSync.
bool applyRemoteValueToCell(QTableWidget* table, int row, int col,
                            const QString& newValue);

} // namespace DVE
