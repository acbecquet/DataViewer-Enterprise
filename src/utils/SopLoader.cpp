#include "SopLoader.h"

#include <QDebug>
#include <QFile>
#include <xlsxdocument.h>

namespace DVE {

QVector<SopEntry> SopLoader::load(const QString& xlsxPath)
{
    QVector<SopEntry> result;

    if (!QFile::exists(xlsxPath)) {
        qWarning() << "SopLoader: file not found:" << xlsxPath;
        return result;
    }

    QXlsx::Document xlsx(xlsxPath);
    if (!xlsx.load()) {
        qWarning() << "SopLoader: QXlsx failed to load:" << xlsxPath;
        return result;
    }

    if (!xlsx.selectSheet(QStringLiteral("Test SOPs"))) {
        qWarning() << "SopLoader: 'Test SOPs' sheet not found in" << xlsxPath;
        return result;
    }

    // Row 1 is headers, data starts at row 2.
    // Columns: A=Test, B=SOP, C=Objective, D=Pass Criteria, E=Equipment,
    //          F=Quantity, G=Est Duration (1mL), H=Est Duration (2mL), I=Note
    for (int row = 2; row <= 200; ++row) {
        const QString test = xlsx.read(row, 1).toString().trimmed();
        if (test.isEmpty()) break;

        SopEntry e;
        e.test           = test;
        e.sop            = xlsx.read(row, 2).toString().trimmed();
        e.objective      = xlsx.read(row, 3).toString().trimmed();
        e.passCriteria   = xlsx.read(row, 4).toString().trimmed();
        e.equipment      = xlsx.read(row, 5).toString().trimmed();
        e.quantity       = xlsx.read(row, 6).toString().trimmed();
        e.estDuration1mL = xlsx.read(row, 7).toString().trimmed();
        e.estDuration2mL = xlsx.read(row, 8).toString().trimmed();
        e.note           = xlsx.read(row, 9).toString().trimmed();
        result.append(e);
    }

    return result;
}

} // namespace DVE
