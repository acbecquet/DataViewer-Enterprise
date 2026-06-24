#include <QtTest>
#include "../../src/pipeline/NotesStory.h"
using namespace DVE;

static DataRow row(double puffs, double tpm, double var, const QString& notes,
                   double before = 1.0, double after = 0.9) {
    DataRow r; r.puffs = puffs; r.tpm = tpm; r.variationTPM = var; r.notes = notes;
    r.beforeWeight = before; r.afterWeight = after; return r;
}

class TstNotesStory : public QObject { Q_OBJECT
private slots:
    void segmentsSplitOnNoteRows();
    void summaryAggregatesOverIncludedRowsOnly();
    void zeroWeightRowsAreSkipped();
    void allExcludedRunYieldsZeroAggregatesButRealRange();
};

void TstNotesStory::segmentsSplitOnNoteRows() {
    SampleResult s;
    s.rows = { row(1,3.6,0.1,""), row(2,3.5,0.1,""), row(3,3.4,0.2,"burnt onset"),
               row(4,3.3,0.2,""), row(5,3.1,0.2,"clog") };
    const QVector<StorySegment> segs = NotesStory::build(s, /*excluded=*/{});
    QCOMPARE(segs.size(), 4);
    QCOMPARE(segs[0].kind, StorySegment::Summary);      // puffs 1-2
    QCOMPARE(segs[0].summary.count, 2);
    QCOMPARE(segs[1].kind, StorySegment::Note);         // puff 3
    QCOMPARE(segs[1].rowIndex, 2);
    QCOMPARE(segs[2].kind, StorySegment::Summary);      // puff 4
    QCOMPARE(segs[3].kind, StorySegment::Note);         // puff 5
}

void TstNotesStory::summaryAggregatesOverIncludedRowsOnly() {
    SampleResult s;
    s.rows = { row(1,4.0,0.0,""), row(2,2.0,0.0,""), row(3,3.0,0.0,"") };
    const QVector<StorySegment> segs = NotesStory::build(s, /*excluded=*/{1}); // drop puff 2
    QCOMPARE(segs.size(), 1);
    QCOMPARE(segs[0].kind, StorySegment::Summary);
    QCOMPARE(segs[0].summary.count, 2);                 // 2 included rows
    QCOMPARE(segs[0].summary.avgTpm, 3.5);              // (4.0 + 3.0) / 2
}

void TstNotesStory::zeroWeightRowsAreSkipped() {
    SampleResult s;
    s.rows = { row(1,3.0,0.1,""), row(2,0,0,"", 0.0, 0.0), row(3,3.0,0.1,"note") };
    const QVector<StorySegment> segs = NotesStory::build(s, {});
    QCOMPARE(segs.size(), 2);                            // empty row dropped: 1 summary + 1 note
    QCOMPARE(segs[0].summary.count, 1);
    QCOMPARE(segs[1].kind, StorySegment::Note);
    QCOMPARE(segs[1].rowIndex, 2);
}

void TstNotesStory::allExcludedRunYieldsZeroAggregatesButRealRange() {
    SampleResult s;
    s.rows = { row(10,4.0,0.1,""), row(11,3.0,0.2,""), row(12,2.0,0.3,"") };
    const QVector<StorySegment> segs = NotesStory::build(s, /*excluded=*/{0,1,2});
    QCOMPARE(segs.size(), 1);
    QCOMPARE(segs[0].kind, StorySegment::Summary);
    QCOMPARE(segs[0].summary.count, 0);          // no included rows
    QCOMPARE(segs[0].summary.avgTpm, 0.0);
    QCOMPARE(segs[0].summary.varTpm, 0.0);
    QCOMPARE(segs[0].summary.puffStart, 10);     // range still spans the excluded run
    QCOMPARE(segs[0].summary.puffEnd, 12);
}

QTEST_MAIN(TstNotesStory)
#include "tst_notesstory.moc"
