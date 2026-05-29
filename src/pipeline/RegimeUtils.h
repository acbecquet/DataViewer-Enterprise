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

} // namespace RegimeUtils
} // namespace DVE
