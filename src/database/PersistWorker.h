// PersistWorker.h -- SP4.5 Stage 2a. Background thread that serialises whole-file
// FileResult saves to Postgres so loading a file no longer blocks the UI on a
// multi-second NAS transaction. Constructed with no parent; the caller
// moveToThread()s it, wires signals, starts the thread, then invokeMethod
// ("start") -- mirrors LiveSync's worker lifecycle EXACTLY (there is NO
// QThread::started connect, which would double-open the connection).
//
// Owns ONE uniquely-named QPSQL connection, opened on the worker thread in
// start() and closed in stop(). The actual write goes through the thread-safe
// free function DVE::persistFileOnConnection (DatabaseOps.cpp) -- the worker
// never touches a DatabaseManager member, so there is no cross-thread access of
// the UI-thread connection (Qt asserts on that).
#pragma once

#include <QObject>
#include <QQueue>
#include <QSqlDatabase>
#include <QString>
#include "../pipeline/ReportData.h"   // DVE::FileResult
#include "DatabaseManager.h"          // DVE::WriteResult
#include "ConfigLoader.h"             // DVE::DbConfig

namespace DVE {

struct PersistJob {
    int        slotIndex  = -1;   // index into MainWindow's working set at enqueue time
    QString    filePath;          // stable key; the writeback re-finds the slot by path
    quint64    generation = 0;    // reload guard: only the latest gen per file writes back
    QString    writerUuid;        // computed on the UI thread at enqueue time
    bool       online     = true; // UI-thread snapshot of DatabaseManager::isOnline()
    FileResult snapshot;          // value copy of the FileResult at enqueue time
};

class PersistWorker : public QObject
{
    Q_OBJECT
public:
    explicit PersistWorker(DbConfig cfg, QObject* parent = nullptr);
    ~PersistWorker() override = default;

public slots:
    void start();                       // open own connection (after moveToThread)
    void stop();                        // close own connection
    void enqueue(DVE::PersistJob job);  // UI thread posts jobs via QueuedConnection
    void drainAndSignal();              // process the queue then emit drainFinished()

signals:
    void persistFinished(DVE::PersistJob job, DVE::WriteResult wr);
    void drainFinished();

private:
    void processQueue();
    void openConnection();
    void closeConnection();

    DbConfig           m_cfg;
    QSqlDatabase       m_pgDb;
    QString            m_connName;
    bool               m_open = false;
    QQueue<PersistJob> m_queue;
};

} // namespace DVE

Q_DECLARE_METATYPE(DVE::PersistJob)
