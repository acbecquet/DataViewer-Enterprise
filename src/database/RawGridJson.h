#pragma once
#include <QString>
#include <QStringList>
#include <QVector>
namespace DVE {
// Serialize a raw-table grid to compact JSON:
// {"headers":["h1",...],"rows":[["a","b",...],...]}.
// Returns empty QString when both headers AND rows are empty.
QString rawGridToJson(const QStringList& headers, const QVector<QStringList>& rows);
// Parse back; empty or invalid input clears both outputs.
void rawGridFromJson(const QString& json, QStringList& headers, QVector<QStringList>& rows);
} // namespace DVE
