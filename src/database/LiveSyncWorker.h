#pragma once

#include <QObject>
#include <QSqlDatabase>
#include <QString>
#include <QVariant>
#include "ConfigLoader.h"

namespace DVE {

// LiveSyncWorker runs on its own QThread. The UI thread fires fire-and-forget
// queued-connection signals; the worker owns its own QSqlDatabase connection
// so writes never block the UI on Postgres round trips.
//
// Lifecycle:
//   ctor builds the QObject (still on the caller's thread).
//   start() must be called AFTER moveToThread() — it runs on the worker
//   thread and opens the dedicated connection.
//   stop()  closes the connection on the worker thread.
//
// Signals back to the UI (commitFailed, commitSucceeded) are queued so the
// caller sees them on the UI thread and can update offline-snapshot state.
class LiveSyncWorker : public QObject {
    Q_OBJECT
public:
    explicit LiveSyncWorker(const DbConfig& cfg, QObject* parent = nullptr);
    ~LiveSyncWorker() override;

public slots:
    // Open the dedicated connection on the worker thread.
    void start();
    // Close it on the worker thread.
    void stop();

    // Scalar UPDATE via dve_commit_cell stored function.
    // expectedVersion < 0 disables OCC (legacy callers/tests). When >= 0,
    // the stored function's WHERE clause matches AND version = expected;
    // a row with a different version triggers commitConflict, not commitFailed.
    void commitScalar(QString table, qint64 rowId,
                      QString column, QVariant value, QString uuid,
                      qint64 expectedVersion);

    // JSONB UPDATE via dve_commit_cell_json stored function. expectedVersion
    // semantics match commitScalar.
    // jsonPath is the unparsed path (e.g. "samples[2].name"); the worker
    // parses it into a TEXT[] array client-side and sends both forms.
    void commitJson(QString table, qint64 rowId,
                    QString jsonPath, QVariant value, QString uuid,
                    qint64 expectedVersion);

    void focusCell(QString uuid, QString table, qint64 rowId,
                   QString column, QString userName, QString userColor);
    void blurCell(QString uuid);

signals:
    // Emitted on driver-level commit failure (network drop, syntax error,
    // permission denial). LiveSync (UI thread) enqueues the lost edit to
    // the offline snapshot for replay.
    void commitFailed(QString table, qint64 rowId,
                      QString column, QVariant value);

    // Emitted when the stored function returns FALSE — the caller passed
    // an expectedVersion that no longer matches the row's current version
    // (optimistic-concurrency miss) OR the row was deleted. Distinct from
    // commitFailed so LiveSync DOES NOT enqueue to the offline snapshot
    // (that would re-apply a stale write).
    void commitConflict(QString table, qint64 rowId,
                        QString column, QVariant attemptedValue,
                        qint64 expectedVersion);

private:
    bool        openConnection();
    QString     parsePathToPgArray(const QString& jsonPath) const;

    DbConfig     m_cfg;
    QSqlDatabase m_db;
    QString      m_connName;
    bool         m_open = false;
};

} // namespace DVE
