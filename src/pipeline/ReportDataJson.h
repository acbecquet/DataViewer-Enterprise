#pragma once

#include <QJsonObject>

#include "pipeline/ReportData.h"

namespace DVE {

// Canonical JSON contract for the TPM FileResult tree. Used by the Plan C
// RecoveryManager to snapshot every open TPM file to disk so a crash or
// auto-update can restore it losslessly. The round-trip
// (fileResultToJson -> fileResultFromJson) must preserve every data-bearing
// field at every level, including server-assigned id/version anchors.
//
// Intentionally OMITTED (re-derivable or transient, never persisted here):
//   * SheetResult::tpmTrend, SheetResult::puffCounts  (recomputed on restore)
//   * SheetResult::images                              (embedded Excel bytes)
//   * SheetResult::dbDataIncomplete                    (transient DB-load flag)
//
// Any new data-bearing field on DataRow / SampleResult / SheetResult /
// FileResult MUST be added to BOTH functions and to the round-trip test in
// tst_reportdatajson, or recovery will silently drop it.
QJsonObject fileResultToJson(const FileResult& f);
FileResult  fileResultFromJson(const QJsonObject& obj);

} // namespace DVE
