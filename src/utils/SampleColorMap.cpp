#include "utils/SampleColorMap.h"
#include "utils/AppTheme.h"
#include <QSet>

namespace DVE {

SampleColorMap& SampleColorMap::instance()
{
    static SampleColorMap inst;
    return inst;
}

int SampleColorMap::indexFor(const QString& name)
{
    const QString key = name.trimmed();
    if (key.isEmpty()) return -1;
    QMutexLocker lock(&m_mutex);
    const auto it = m_index.constFind(key);
    if (it != m_index.constEnd()) return it.value();
    const int idx = m_next++;
    m_index.insert(key, idx);
    return idx;
}

QVector<QColor> SampleColorMap::colorsForPlot(const QVector<QString>& names)
{
    const int n = names.size();
    QVector<int> idxOf(n, -1);
    QSet<int>    used;

    // Pass 1: named samples take their global pinned index.
    for (int i = 0; i < n; ++i) {
        const QString key = names[i].trimmed();
        if (key.isEmpty()) continue;
        const int gi = indexFor(key);      // locks internally; not held here
        idxOf[i] = gi;
        used.insert(gi);
    }
    // Pass 2: blank names take the lowest palette index unused in THIS plot.
    int probe = 0;
    for (int i = 0; i < n; ++i) {
        if (idxOf[i] >= 0) continue;
        while (used.contains(probe)) ++probe;
        idxOf[i] = probe;
        used.insert(probe);
    }

    QVector<QColor> out(n);
    for (int i = 0; i < n; ++i)
        out[i] = AppTheme::seriesColor(idxOf[i]);
    return out;
}

void SampleColorMap::clear()
{
    QMutexLocker lock(&m_mutex);
    m_index.clear();
    m_next = 0;
}

int SampleColorMap::size() const
{
    QMutexLocker lock(&m_mutex);
    return m_index.size();
}

} // namespace DVE
