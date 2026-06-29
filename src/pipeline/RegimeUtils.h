#pragma once

#include "ReportData.h"
#include <QString>
#include <QStringList>

namespace DVE {
namespace RegimeUtils {

QString     unspecifiedLabel();
QString     regimeKey(const DataRow& row);
bool        isRegimeHeader(const QString& colEHeader);
QStringList uniqueRegimes(const SheetResult& sheet);     // non-blank, for the picker
QStringList uniqueRegimes(const FileResult& file);
QStringList uniqueRegimeKeys(const SheetResult& sheet);  // incl. "(unspecified)", for report fan-out
bool        sheetHasRegimeData(const SheetResult& sheet);
SheetResult filterByRegime(const SheetResult& sheet, const QString& regime);

// ── Editable-regime format / alias helpers (shared by the Notes panel field
//    and the TPM-table RegimeComboDelegate, so both accept/refuse identically) ──
//
// isStandardRegimeFormat: true when `text` is the parametric standard
//   "<vol>mL/<puff>s/<interval>s" (the mL prefix and surrounding whitespace are
//   optional; decimals allowed, e.g. 100mL/2.5s/15s) — or empty. This drives
//   both the advisory red flag and the accept/refuse decision.
// canonicalRegime: trims `text` and maps known aliases to the canonical
//   parametric label — currently "CORESTA" -> "60mL/3s/30s" (the standard IS
//   the CORESTA standard but is always labeled parametrically, never as the
//   word). Any other text is returned trimmed, unchanged.
bool        isStandardRegimeFormat(const QString& text);
QString     canonicalRegime(const QString& text);

} // namespace RegimeUtils
} // namespace DVE
