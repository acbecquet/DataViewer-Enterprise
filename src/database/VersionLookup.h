#pragma once

#include <QString>
#include <QVector>

namespace DVE {

struct FileResult;
struct SensorySession;
struct DetailedSensorySession;

// Resolve the caller's last-known row version for an OCC-protected table.
// Returns -1 when the row isn't found in the in-memory cache, or when the
// cached version is 0 (fresh INSERT that hasn't round-tripped yet) — both
// cases opt out of OCC so the first commit isn't gated on a meaningless
// anchor. This is the pure-function backbone of MainWindow's
// LiveSync::setVersionLookup callback.
qint64 versionForRow(const QVector<FileResult>& files,
                     const QVector<SensorySession>& sensorySessions,
                     const QVector<DetailedSensorySession>& detailedSessions,
                     const QString& table,
                     qint64 rowId);

} // namespace DVE
