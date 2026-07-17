#pragma once
#include <QString>
#include <QColor>
#include <QVector>
#include <QHash>
#include <QMutex>

namespace DVE {

// DV-26: stable sample-name -> palette-index registry so the SAME name renders
// in the SAME color everywhere it appears (sensory radar, TPM plots, single-
// file report plots). First-seen order, never reassigned. Session/working-set
// lifetime - NOT persisted; clear() on a fresh working-set load. Thread-safe
// because report builds can run off the GUI thread.
class SampleColorMap {
public:
    static SampleColorMap& instance();

    // Stable index for a non-empty (trimmed) name; assigns the next free index
    // on first sight. Blank/whitespace -> -1 (caller supplies a local fallback).
    int indexFor(const QString& name);

    // One color per name for a single plot. Named samples take their global
    // pinned index; blank names take the lowest palette index unused in THIS
    // plot, guaranteeing in-plot distinctness while pinning names globally.
    QVector<QColor> colorsForPlot(const QVector<QString>& names);

    void clear();          // drop all mappings (fresh working set)
    int  size() const;     // registered-name count (tests/diagnostics)

private:
    SampleColorMap() = default;
    mutable QMutex      m_mutex;
    QHash<QString, int> m_index;
    int                 m_next = 0;
};

} // namespace DVE
