#pragma once
#include "PlotEngine.h"
#include <QPoint>
#include <QVector>

// DV-18 (note<->plot linking, v2): a pure, widget-free hit-test for mapping a
// click in unzoomed plot-pixmap pixels to the nearest note-bearing TPM-trend
// point. Lives in its own header (rather than on PlotWidget) so it can be
// unit-tested without linking the heavy PlotWidget.cpp — only PlotEngine.h's
// PlotTransform is needed.

namespace DVE {

// A note-bearing plotted point, resolved back to its sample+row.
struct NotePoint {
    int    sampleIndex   = -1;
    int    dataRowIndex  = -1;
    double puffs         = 0.0;
    double tpm           = 0.0;
};

// Return the note point nearest clickPx (in unzoomed pixmap pixels) within
// tolPx, or false if none / the transform is invalid / pts is empty.
// On success, outSample/outRow are set to the winning point's indices; on
// failure they are set to -1.
inline bool nearestNotePoint(const QPoint& clickPx, const PlotTransform& tf,
                              const QVector<NotePoint>& pts, int tolPx,
                              int& outSample, int& outRow)
{
    outSample = -1;
    outRow    = -1;
    if (!tf.valid) return false;
    double best = double(tolPx) * tolPx;   // compare squared distances
    for (const NotePoint& np : pts) {
        const QPointF p  = tf.dataToPixel(np.puffs, np.tpm);
        const double  dx = p.x() - clickPx.x();
        const double  dy = p.y() - clickPx.y();
        const double  d2 = dx * dx + dy * dy;
        if (d2 <= best) {
            best      = d2;
            outSample = np.sampleIndex;
            outRow    = np.dataRowIndex;
        }
    }
    return outSample >= 0;
}

} // namespace DVE
