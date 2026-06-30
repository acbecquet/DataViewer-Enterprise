#include <QtTest/QtTest>
#include <QApplication>
#include <QDoubleSpinBox>
#include "ui/DataCleanupDialog.h"
#include "pipeline/ReportData.h"

using namespace DVE;

class TstSizingSweep : public QObject
{
    Q_OBJECT
private slots:
    void cleanupDialogFloorLowered();
    void cleanupDialogControlsNotVerticallyFixed();
};

void TstSizingSweep::cleanupDialogFloorLowered()
{
    SheetResult sheet;
    sheet.sheetName = QStringLiteral("S1");
    SampleResult sample;
    sample.sampleName = QStringLiteral("A");
    DataRow row; row.puffs = 1; row.beforeWeight = 1.0; row.afterWeight = 0.9;
    sample.rows.append(row);
    sheet.samples.append(sample);

    QMap<int, QSet<int>> none;
    DataCleanupDialog dlg(sheet, none);
    QVERIFY2(dlg.minimumHeight() <= 400,
             qPrintable(QStringLiteral("minimumHeight=%1").arg(dlg.minimumHeight())));
    QVERIFY2(dlg.minimumWidth() <= 560,
             qPrintable(QStringLiteral("minimumWidth=%1").arg(dlg.minimumWidth())));
}

void TstSizingSweep::cleanupDialogControlsNotVerticallyFixed()
{
    SheetResult sheet;
    sheet.sheetName = QStringLiteral("S1");
    SampleResult sample;
    sample.sampleName = QStringLiteral("A");
    sheet.samples.append(sample);

    QMap<int, QSet<int>> none;
    DataCleanupDialog dlg(sheet, none);
    const auto spins = dlg.findChildren<QDoubleSpinBox*>();
    QVERIFY(!spins.isEmpty());
    for (auto* sp : spins) {
        const bool fixedBand =
            sp->sizePolicy().horizontalPolicy() == QSizePolicy::Fixed &&
            sp->minimumWidth() == sp->maximumWidth();
        QVERIFY2(!fixedBand, "spin box must be growable, not a fixed width band");
    }
}

int main(int argc, char** argv)
{
    QApplication app(argc, argv);
    TstSizingSweep tc;
    return QTest::qExec(&tc, argc, argv);
}

#include "tst_sizingsweep.moc"
