#pragma once
#include "ReportData.h"
#include <QSet>
#include <QVector>

namespace DVE {

struct StorySummary {
    int    puffStart = 0;
    int    puffEnd   = 0;
    int    count     = 0;
    double avgTpm    = 0.0;
    double varTpm    = 0.0;
    QVector<int> rowIndices; // indices into SampleResult::rows of all visible note-less rows in this run (incl. excluded)
};

struct StorySegment {
    enum Kind { Note, Summary } kind = Summary;
    int          rowIndex = -1;     // index into SampleResult::rows (Note only)
    bool         excluded = false;  // Note only
    StorySummary summary;           // Summary only
};

namespace NotesStory {
    QVector<StorySegment> build(const SampleResult& sample, const QSet<int>& excluded);
}

} // namespace DVE
