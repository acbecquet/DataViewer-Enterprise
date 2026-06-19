// SnapshotRegenWorker.h -- SP4.5 Stage 2a. Background thread that runs
// OfflineSnapshot::regenToPath() against its OWN PostgresConnection (a separate
// socket, so the REPEATABLE READ single-connection consistency invariant is
// preserved). Debounced during the session by MainWindow so the close-time
// isCurrentVsLive() check usually skips the regen entirely.
//
// Lifecycle mirrors LiveSync's worker EXACTLY: constructed with no parent,
// moveToThread, start the thread, then invokeMethod("start"). requestRegen() is
// a BLOCKING slot (the regen takes seconds) -- so teardown/close use cancel()
// (an atomic flag polled inside regenToPath) + quit()/wait(), NEVER a
// BlockingQueuedConnection on the blocking slot (that would freeze the UI for
// the whole regen, defeating the point of Stage 2a).
#pragma once

#include <QObject>
#include <QString>
#include <atomic>
#include "ConfigLoader.h"   // DVE::DbConfig

namespace DVE {

class PostgresConnection;

class SnapshotRegenWorker : public QObject
{
    Q_OBJECT
public:
    explicit SnapshotRegenWorker(DbConfig cfg, QString snapshotProductionPath,
                                 QObject* parent = nullptr);
    ~SnapshotRegenWorker() override;

public slots:
    void start();         // open own PostgresConnection (after moveToThread)
    void stop();          // close + delete it
    void requestRegen();  // run a regen; if one is already in flight, coalesce
    void cancel();        // set the atomic cancel flag (cross-thread safe)

signals:
    void regenFinished(bool ok, QString error);

private:
    DbConfig            m_cfg;
    QString             m_productionPath;
    PostgresConnection* m_pg      = nullptr;  // owned; created on the worker thread
    bool                m_running = false;    // worker-thread-only
    bool                m_pending = false;    // worker-thread-only: request arrived mid-regen
    std::atomic<bool>   m_cancel{false};
};

} // namespace DVE
