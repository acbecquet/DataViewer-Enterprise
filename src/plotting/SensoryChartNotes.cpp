#include "plotting/SensoryChartNotes.h"
#include "utils/AppTheme.h"
#include "utils/SampleColorMap.h"

namespace DVE {

QVector<SampleNote> collectSensoryNotes(const QVector<SensorySession>& sessions,
                                        const QSet<int>&               hidden)
{
    // DV-26: resolve colors for ALL samples in draw order (same order/keying as
    // the radar), so swatches match the name-pinned polygons.
    QVector<QString> names;
    for (const SensorySession& s : sessions)
        for (const SensorySample& sm : s.samples) names.append(sm.name);
    const QVector<QColor> cols = SampleColorMap::instance().colorsForPlot(names);

    QVector<SampleNote> out;
    int gi = 0;
    for (const SensorySession& s : sessions) {
        for (const SensorySample& sm : s.samples) {
            const int idx = gi++;                 // count every sample, then filter
            if (hidden.contains(idx)) continue;
            const QString body = sm.comments.trimmed();
            if (body.isEmpty()) continue;
            SampleNote n;
            n.title  = sm.name.isEmpty() ? QStringLiteral("Sample %1").arg(idx + 1) : sm.name;
            n.body   = body;
            n.swatch = cols.value(idx, AppTheme::seriesColor(idx));
            out.append(n);
        }
    }
    return out;
}

} // namespace DVE
