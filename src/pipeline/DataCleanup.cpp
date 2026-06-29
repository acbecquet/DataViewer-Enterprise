#include "DataCleanup.h"

#include "SheetProcessors.h"

#include <QStringList>

namespace DVE {
namespace DataCleanup {

QString key(int fileIdx, int sheetIdx, int sampleIdx)
{
    return QString("%1:%2:%3").arg(fileIdx).arg(sheetIdx).arg(sampleIdx);
}

QSet<int> exclusionsFor(const QMap<QString, QSet<int>>& all,
                        int fileIdx, int sheetIdx, int sampleIdx)
{
    return all.value(key(fileIdx, sheetIdx, sampleIdx));
}

QMap<QString, QSet<int>> rekeyAfterClose(const QMap<QString, QSet<int>>& all,
                                         int closedIdx)
{
    QMap<QString, QSet<int>> out;
    for (auto it = all.constBegin(); it != all.constEnd(); ++it) {
        // key format is "fileIdx:sheetIdx:sampleIdx"
        const QStringList parts = it.key().split(':');
        if (parts.size() != 3) {
            // Unexpected key shape — preserve it untouched rather than dropping.
            out.insert(it.key(), it.value());
            continue;
        }
        bool ok = false;
        const int fileIdx = parts[0].toInt(&ok);
        if (!ok) {
            out.insert(it.key(), it.value());
            continue;
        }
        if (fileIdx == closedIdx)
            continue;                       // belonged to the closed file — drop
        const int newFileIdx = (fileIdx > closedIdx) ? fileIdx - 1 : fileIdx;
        out.insert(key(newFileIdx, parts[1].toInt(), parts[2].toInt()), it.value());
    }
    return out;
}

SampleResult buildCleanedSample(const SampleResult& sr, const QSet<int>& excluded)
{
    if (excluded.isEmpty()) return sr;

    SampleResult cleaned = sr;

    // Build a cleanup note listing every excluded data point
    QStringList parts;
    int rowNum = 0;
    for (int i = 0; i < sr.rows.size(); ++i) {
        const DataRow& dr = sr.rows[i];
        if (dr.beforeWeight == 0.0 || dr.afterWeight == 0.0) continue;
        ++rowNum;
        if (excluded.contains(i))
            parts << QString("Puff %1 (TPM=%2)")
                     .arg(dr.puffs, 0, 'f', 0)
                     .arg(dr.tpm,   0, 'f', 3);
    }
    if (!parts.isEmpty())
        cleaned.extra["cleanupNote"] =
            QString("Data cleanup: %1 row(s) excluded [%2]")
            .arg(excluded.size())
            .arg(parts.join(", "));

    // Remove excluded rows from the copy
    QVector<DataRow> kept;
    kept.reserve(sr.rows.size() - excluded.size());
    for (int i = 0; i < sr.rows.size(); ++i)
        if (!excluded.contains(i))
            kept.append(sr.rows[i]);
    cleaned.rows = kept;

    // Recalculate derived metrics from the surviving rows
    GenericSheetProcessor proc;
    proc.calculateMetrics(cleaned);
    return cleaned;
}

SheetResult buildCleanedSheet(const SheetResult& sheet,
                              const QMap<QString, QSet<int>>& all,
                              int fileIdx, int sheetIdx)
{
    SheetResult cleaned = sheet;
    for (int si = 0; si < sheet.samples.size(); ++si) {
        const QSet<int> ex = exclusionsFor(all, fileIdx, sheetIdx, si);
        if (!ex.isEmpty())
            cleaned.samples[si] = buildCleanedSample(sheet.samples[si], ex);
    }
    GenericSheetProcessor proc;
    proc.computeSheetAggregates(cleaned);
    return cleaned;
}

FileResult buildCleanedFile(const FileResult& file,
                            const QMap<QString, QSet<int>>& all,
                            int fileIdx)
{
    if (fileIdx < 0) return file;
    FileResult cleaned = file;
    for (int si = 0; si < file.sheets.size(); ++si) {
        SheetResult& cs = cleaned.sheets[si];
        bool anyChanged = false;
        for (int sampleIdx = 0; sampleIdx < file.sheets[si].samples.size(); ++sampleIdx) {
            const QSet<int> ex = exclusionsFor(all, fileIdx, si, sampleIdx);
            if (!ex.isEmpty()) {
                cs.samples[sampleIdx] =
                    buildCleanedSample(file.sheets[si].samples[sampleIdx], ex);
                anyChanged = true;
            }
        }
        if (anyChanged && cs.hasSamples()) {
            GenericSheetProcessor proc;
            proc.computeSheetAggregates(cs);
        }
    }
    return cleaned;
}

} // namespace DataCleanup
} // namespace DVE
