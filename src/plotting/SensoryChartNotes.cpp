#include "plotting/SensoryChartNotes.h"
#include "utils/AppTheme.h"

namespace DVE {

QVector<SampleNote> collectSensoryNotes(const QVector<SensorySession>& sessions,
                                        const QSet<int>&               hidden)
{
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
            n.swatch = AppTheme::seriesColor(idx);
            out.append(n);
        }
    }
    return out;
}

} // namespace DVE
