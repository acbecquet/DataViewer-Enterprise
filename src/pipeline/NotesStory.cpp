#include "NotesStory.h"

namespace DVE {

static bool isVisible(const DataRow& r) {
    return r.beforeWeight != 0.0 && r.afterWeight != 0.0;
}

QVector<StorySegment> NotesStory::build(const SampleResult& sample, const QSet<int>& excluded) {
    QVector<StorySegment> out;

    StorySummary run;
    int    runIncluded = 0;
    double sumTpm = 0.0, sumVar = 0.0;
    bool   runOpen = false;

    auto flushRun = [&]() {
        if (!runOpen) return;
        run.count  = runIncluded;
        run.avgTpm = runIncluded ? sumTpm / runIncluded : 0.0;
        run.varTpm = runIncluded ? sumVar / runIncluded : 0.0;
        StorySegment seg; seg.kind = StorySegment::Summary; seg.summary = run;
        out.append(seg);
        run = StorySummary{}; runIncluded = 0; sumTpm = sumVar = 0.0; runOpen = false;
    };

    for (int i = 0; i < sample.rows.size(); ++i) {
        const DataRow& r = sample.rows[i];
        if (!isVisible(r)) continue;

        if (!r.notes.trimmed().isEmpty()) {
            flushRun();
            StorySegment seg;
            seg.kind = StorySegment::Note;
            seg.rowIndex = i;
            seg.excluded = excluded.contains(i);
            out.append(seg);
            continue;
        }

        // puffStart/puffEnd span ALL visible rows in the run, including excluded ones, so the label shows the real puff range; count/avgTpm/varTpm below cover INCLUDED rows only. An all-excluded run therefore yields count==0 with avg/var==0 but a real range — consumers must guard count==0 before dividing or labelling an average.
        if (!runOpen) { run.puffStart = int(r.puffs); runOpen = true; }
        run.puffEnd = int(r.puffs);
        if (!excluded.contains(i)) {
            ++runIncluded; sumTpm += r.tpm; sumVar += r.variationTPM;
        }
    }
    flushRun();
    return out;
}

} // namespace DVE
