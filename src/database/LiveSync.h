#pragma once

#include <QObject>
#include <QPointer>
#include <QString>
#include <QVariant>

namespace DVE {

class PostgresConnection;
class IdentityManager;
struct RowChange;
struct CellFocusChange;

// LiveSync is the single chokepoint for per-cell writes and remote
// applies in v2.0.1. Every editable widget commits through
// commitCell(); every NOTIFY-driven cell change emits cellChanged().
//
// JSONB columns are addressed with column names prefixed "json_path:"
// e.g. column = "json_path:samples[2].scores.smoothness". LiveSync
// translates these into jsonb_set() UPDATEs at the database layer so
// concurrent edits to different paths in the same sample don't
// clobber each other.
class LiveSync : public QObject {
    Q_OBJECT
public:
    LiveSync(PostgresConnection* conn, IdentityManager* identity,
             QObject* parent = nullptr);

    // Scalar-column UPDATE on a known table+row+column. Returns true on
    // success, false on connection failure or DB error. Bumps version
    // and stamps updated_by. Sets the dve.live_column / dve.live_value
    // session vars first so the trigger emits a column-aware payload.
    bool commitCell(const QString& table, qint64 rowId,
                    const QString& column, const QVariant& value);

    // Upsert a cell_focus row for the current user. The user can only
    // hold one focus at a time; calling focusCell again first deletes
    // any previous focus row owned by this user.
    bool focusCell(const QString& table, qint64 rowId, const QString& column);

    // Delete the current user's focus row (if any).
    bool blurCell();

signals:
    // Emitted when a remote cell change arrives (after self-UUID filter).
    void cellChanged(const QString& table, qint64 rowId,
                     const QString& column, const QVariant& newValue);

    // Emitted when a remote cell-focus change arrives.
    void cellFocused(const QString& table, qint64 rowId,
                     const QString& column,
                     const QString& userName, const QString& userColor);
    void cellBlurred(const QString& table, qint64 rowId,
                     const QString& column);

public slots:
    // Called by MainWindow when NotificationListener emits its signals,
    // after the own-UUID echo filter. Splits row/focus dispatch.
    void onRowChanged(const RowChange& change);
    void onCellFocusChanged(const CellFocusChange& change);

private:
    bool runScalarUpdate(const QString& table, qint64 rowId,
                         const QString& column, const QVariant& value);
    bool runJsonPathUpdate(const QString& table, qint64 rowId,
                           const QString& jsonPath, const QVariant& value);

    QPointer<PostgresConnection> m_conn;
    QPointer<IdentityManager>    m_identity;
};

} // namespace DVE
