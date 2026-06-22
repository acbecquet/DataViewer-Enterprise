#include "SnapshotRegenWorker.h"
#include "PostgresConnection.h"
#include "OfflineSnapshot.h"

#include <QDateTime>
#include <QDebug>

namespace DVE {

SnapshotRegenWorker::SnapshotRegenWorker(DbConfig cfg, QString snapshotProductionPath,
                                         QObject* parent)
    : QObject(parent)
    , m_cfg(std::move(cfg))
    , m_productionPath(std::move(snapshotProductionPath))
{}

SnapshotRegenWorker::~SnapshotRegenWorker()
{
    delete m_pg;   // null-safe; stop() normally clears it first on the worker thread
}

void SnapshotRegenWorker::start()
{
    m_pg = new PostgresConnection;          // no parent: owned by this worker
    if (!m_pg->open(m_cfg))
        qWarning() << "[SnapshotRegenWorker] PG open failed:" << m_pg->lastError();
    else
        qInfo() << "[SnapshotRegenWorker] PG connection open";
}

void SnapshotRegenWorker::stop()
{
    if (m_pg) {
        m_pg->close();
        delete m_pg;
        m_pg = nullptr;
    }
}

void SnapshotRegenWorker::cancel()
{
    m_cancel.store(true);   // polled inside regenToPath between blob tables
}

void SnapshotRegenWorker::requestRegen()
{
    if (m_running) {
        // A request arrived while a regen is in flight: remember it and coalesce
        // into exactly one follow-up (the do/while below), never a nested regen.
        m_pending = true;
        qInfo() << "[SnapshotRegenWorker] regen in flight -- coalescing follow-up";
        return;
    }
    if (!m_pg || !m_pg->isOpen()) {
        emit regenFinished(false, QStringLiteral("PG connection not open"));
        return;
    }

    do {
        m_pending = false;
        m_running = true;
        m_cancel.store(false);
        qInfo() << "[SnapshotRegenWorker] starting background regen";

        QString fp, err;
        OfflineSnapshot::RegenStats stats;
        const bool ok = DVE::OfflineSnapshot::regenToPath(
            m_pg, m_productionPath, &m_cancel, &fp, /*outServerTimeUtc*/nullptr, &err,
            [this](int d, int t, const QString& ph){ emit regenProgress(d, t, ph); },
            &stats);

        m_running = false;
        if (ok) qInfo()    << "[SnapshotRegenWorker] regen complete, fp:" << fp
                           << "incremental:" << stats.wasIncremental
                           << "blobsPulled:" << stats.imageRowsPulledFromPg;
        else    qWarning() << "[SnapshotRegenWorker] regen failed:" << err;
        emit regenFinished(ok, err);
        // If a request coalesced while we were running, run exactly one more.
        // Terminates: each pass clears m_pending up front; a cancel breaks out.
    } while (m_pending && !m_cancel.load());
}

} // namespace DVE
