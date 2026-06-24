#pragma once

// ─── DataCleanup ──────────────────────────────────────────────────────────────
// Pure, GUI-free helpers for the non-destructive data-cleanup feature.
//
// Cleanup exclusions are stored as a flat map keyed by the string
// "fileIdx:sheetIdx:sampleIdx" → set of excluded row indices. This module owns
// the keying convention and the transforms that operate on it, so the logic is
// unit-testable without spinning up MainWindow.
//
// MainWindow holds the live QMap<QString,QSet<int>> and delegates here.
// ──────────────────────────────────────────────────────────────────────────────

#include "ReportData.h"

#include <QMap>
#include <QSet>
#include <QString>

namespace DVE {
namespace DataCleanup {

// Canonical key for one (file, sheet, sample) triple.
QString key(int fileIdx, int sheetIdx, int sampleIdx);

// The set of excluded row indices for one (file, sheet, sample) triple.
QSet<int> exclusionsFor(const QMap<QString, QSet<int>>& all,
                        int fileIdx, int sheetIdx, int sampleIdx);

// GAP-B: when the file at `closedIdx` is removed from the working set, every
// file after it shifts down by one. Return a re-keyed copy of `all` that drops
// the closed file's keys and decrements fileIdx for every higher-indexed file.
QMap<QString, QSet<int>> rekeyAfterClose(const QMap<QString, QSet<int>>& all,
                                         int closedIdx);

// Build a cleaned copy of one sample with `excluded` rows removed and derived
// metrics recomputed. A "cleanupNote" describing the dropped rows is stored in
// SampleResult::extra. Returns `sr` unchanged when `excluded` is empty.
SampleResult buildCleanedSample(const SampleResult& sr, const QSet<int>& excluded);

// Build a cleaned copy of one sheet, applying the exclusions recorded for
// (fileIdx, sheetIdx, *) in `all`. Sheet-level aggregates are recomputed.
SheetResult buildCleanedSheet(const SheetResult& sheet,
                              const QMap<QString, QSet<int>>& all,
                              int fileIdx, int sheetIdx);

// GAP-A: build a cleaned copy of one file, applying the exclusions recorded for
// `fileIdx` in `all`. `fileIdx` must be the file's ACTUAL index in the working
// set — not the currently-selected file — so multi-file / Combined reports
// apply each file's own exclusions.
FileResult buildCleanedFile(const FileResult& file,
                            const QMap<QString, QSet<int>>& all,
                            int fileIdx);

} // namespace DataCleanup
} // namespace DVE
