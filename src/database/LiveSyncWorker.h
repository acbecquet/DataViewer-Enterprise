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
    void commitScalar(QString table, qint64 rowId,
                      QString column, QVariant value, QString uuid);

    // JSONB UPDATE via dve_commit_cell_json stored function.
    // jsonPath is the unparsed path (e.g. "samples[2].name"); the worker
    // parses it into a TEXT[] array client-side and sends both forms.
    void commitJson(QString table, qint64 rowId,
                    QString jsonPath, QVariant value, QString uuid);

    void focusCell(QString uuid, QString table, qint64 rowId,
                   QString column, QString userName, QString userColor);
    void blurCell(QString uuid);

signals:
    // Emitted on commit failure so LiveSync (on UI thread) can enqueue
    // the lost edit to the offline snapshot for replay.
    void commitFailed(QString table, qint64 rowId,
                      QString column, QVariant value);

private:
    bool        openConnection();
    QString     parsePathToPgArray(const QString& jsonPath) const;

    DbConfig     m_cfg;
    QSqlDatabase m_db;
    QString      m_connName;
    bool         m_open = false;
};

} // namespace DVE
