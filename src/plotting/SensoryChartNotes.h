#pragma once
#include <QVector>
#include <QSet>
#include "pipeline/SensoryData.h"
#include "plotting/PlotEngine.h"   // SampleNote

namespace DVE {

// DV-27: derive the per-sample note blocks for the sensory annotated export
// from the radar's own sample source. Iterates sessions then samples, counting
// a global index for EVERY sample (matching RadarChartWidget's colorIdx), skips
// hidden indices and samples whose comment is empty/whitespace, and colors each
// swatch to match the radar polygon. Pure and widget-free so it is unit-testable
// without a QApplication.
QVector<SampleNote> collectSensoryNotes(const QVector<SensorySession>& sessions,
                                        const QSet<int>&               hidden);

} // namespace DVE
