#include "VersionLookup.h"

#include "../pipeline/ReportData.h"
#include "../pipeline/SensoryData.h"
#include "../pipeline/DetailedSensoryData.h"

#include <QStringList>

namespace DVE {

namespace {

// A row with id > 0 but version == 0 is fresh from disk and hasn't been
// round-tripped through the DB yet. Treat it like "not found" so OCC opts
// out — the first commit must not be gated on a meaningless anchor.
inline qint64 asAnchor(int version)
{
    return version > 0 ? static_cast<qint64>(version) : qint64(-1);
}

} // anonymous

qint64 versionForRow(const QVector<FileResult>& files,
                     const QVector<SensorySession>& sensorySessions,
                     const QVector<DetailedSensorySession>& detailedSessions,
                     const QString& table,
                     qint64 rowId)
{
    if (table == QLatin1String("files")) {
        for (const FileResult& f : files) {
            if (f.id == rowId) return asAnchor(f.version);
        }
        return -1;
    }
    if (table == QLatin1String("tests")) {
        for (const FileResult& f : files) {
            for (const SheetResult& s : f.sheets) {
                if (s.id == rowId) return asAnchor(s.version);
            }
        }
        return -1;
    }
    if (table == QLatin1String("samples")) {
        for (const FileResult& f : files) {
            for (const SheetResult& s : f.sheets) {
                for (const SampleResult& sm : s.samples) {
                    if (sm.id == rowId) return asAnchor(sm.version);
                }
            }
        }
        return -1;
    }
    if (table == QLatin1String("data_rows")) {
        for (const FileResult& f : files) {
            for (const SheetResult& s : f.sheets) {
                for (const SampleResult& sm : s.samples) {
                    for (const DataRow& r : sm.rows) {
                        if (r.id == rowId) return asAnchor(r.version);
                    }
                }
            }
        }
        return -1;
    }
    if (table == QLatin1String("sensory_sessions")) {
        for (const SensorySession& ss : sensorySessions) {
            if (static_cast<qint64>(ss.id) == rowId) return asAnchor(ss.version);
        }
        return -1;
    }
    if (table == QLatin1String("detailed_sensory_sessions")) {
        for (const DetailedSensorySession& d : detailedSessions) {
            if (static_cast<qint64>(d.id) == rowId) return asAnchor(d.version);
        }
        return -1;
    }
    return -1;
}

} // namespace DVE
