#pragma once

#include <QString>
#include <QVector>

namespace DVE {

struct SopEntry {
    QString test;
    QString sop;
    QString objective;
    QString passCriteria;
    QString equipment;
    QString quantity;
    QString estDuration1mL;
    QString estDuration2mL;
    QString note;
};

class SopLoader {
public:
    /// Load all SOP rows from the standardized template's "Test SOPs" sheet.
    /// Returns an empty vector and emits a qWarning if the file is missing,
    /// unreadable, or has no parseable rows. Never throws.
    static QVector<SopEntry> load(const QString& xlsxPath);
};

} // namespace DVE
