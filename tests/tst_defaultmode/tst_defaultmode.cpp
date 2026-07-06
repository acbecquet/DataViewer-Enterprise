// Tests for DefaultModeDialog (v2.7.0): the "Pick a default mode:" picker shown
// once on first run and reopened from Settings > Set Default Mode, plus its
// QSettings-backed persistence. The dialog is standalone (no MainWindow), so
// this suite stays light per the responsive-UI plan's no-MainWindow rule.
#include <QtTest>
#include <QLabel>
#include <QPushButton>
#include <QSettings>

#include "ui/DefaultModeDialog.h"

using namespace DVE;

namespace {
const QString kOrg = QStringLiteral("SDR");
const QString kApp = QStringLiteral("DataViewerEnterprise");
const QString kKey = QStringLiteral("ui/defaultMode");
}

class TestDefaultMode : public QObject
{
    Q_OBJECT

private slots:
    void initTestCase();
    void cleanupTestCase();
    void init();

    void unsetMeansNoSavedModeAndTpmDefault();
    void setSavedRoundTripsEveryMode();
    void invalidStoredValueFallsBackToTpm();
    void clickingTpmSavesAndAccepts();
    void clickingSensorySavesAndAccepts();
    void clickingDetailedSavesAndAccepts();
    void rejectingLeavesSettingUnset();
    void buttonsAreStackedVerticallyTpmSensoryDetailed();

private:
    // The developer machine's real registry value, restored afterwards so the
    // suite never clobbers the user's actual default-mode choice.
    QVariant m_priorValue;
    bool m_hadPrior = false;
};

void TestDefaultMode::initTestCase()
{
    QSettings s(kOrg, kApp);
    m_hadPrior = s.contains(kKey);
    m_priorValue = s.value(kKey);
}

void TestDefaultMode::cleanupTestCase()
{
    QSettings s(kOrg, kApp);
    if (m_hadPrior) s.setValue(kKey, m_priorValue);
    else            s.remove(kKey);
}

void TestDefaultMode::init()
{
    QSettings(kOrg, kApp).remove(kKey);
}

void TestDefaultMode::unsetMeansNoSavedModeAndTpmDefault()
{
    QVERIFY(!DefaultModeDialog::hasSavedDefaultMode());
    QCOMPARE(DefaultModeDialog::savedDefaultMode(), ReportMode::Tpm);
}

void TestDefaultMode::setSavedRoundTripsEveryMode()
{
    const ReportMode modes[] = { ReportMode::Tpm, ReportMode::Sensory,
                                 ReportMode::DetailedSensory };
    for (ReportMode m : modes) {
        DefaultModeDialog::setSavedDefaultMode(m);
        QVERIFY(DefaultModeDialog::hasSavedDefaultMode());
        QCOMPARE(DefaultModeDialog::savedDefaultMode(), m);
    }
}

void TestDefaultMode::invalidStoredValueFallsBackToTpm()
{
    QSettings(kOrg, kApp).setValue(kKey, QStringLiteral("banana"));
    QCOMPARE(DefaultModeDialog::savedDefaultMode(), ReportMode::Tpm);
    QSettings(kOrg, kApp).setValue(kKey, 99);
    QCOMPARE(DefaultModeDialog::savedDefaultMode(), ReportMode::Tpm);
}

void TestDefaultMode::clickingTpmSavesAndAccepts()
{
    DefaultModeDialog dlg;
    QPushButton* b = dlg.findChild<QPushButton*>("tpmButton");
    QVERIFY2(b, "tpmButton missing");
    b->click();
    QCOMPARE(dlg.result(), int(QDialog::Accepted));
    QCOMPARE(dlg.selectedMode(), ReportMode::Tpm);
    QVERIFY(DefaultModeDialog::hasSavedDefaultMode());
    QCOMPARE(DefaultModeDialog::savedDefaultMode(), ReportMode::Tpm);
}

void TestDefaultMode::clickingSensorySavesAndAccepts()
{
    DefaultModeDialog dlg;
    QPushButton* b = dlg.findChild<QPushButton*>("sensoryButton");
    QVERIFY2(b, "sensoryButton missing");
    b->click();
    QCOMPARE(dlg.result(), int(QDialog::Accepted));
    QCOMPARE(dlg.selectedMode(), ReportMode::Sensory);
    QCOMPARE(DefaultModeDialog::savedDefaultMode(), ReportMode::Sensory);
}

void TestDefaultMode::clickingDetailedSavesAndAccepts()
{
    DefaultModeDialog dlg;
    QPushButton* b = dlg.findChild<QPushButton*>("detailedButton");
    QVERIFY2(b, "detailedButton missing");
    b->click();
    QCOMPARE(dlg.result(), int(QDialog::Accepted));
    QCOMPARE(dlg.selectedMode(), ReportMode::DetailedSensory);
    QCOMPARE(DefaultModeDialog::savedDefaultMode(), ReportMode::DetailedSensory);
}

void TestDefaultMode::rejectingLeavesSettingUnset()
{
    DefaultModeDialog dlg;
    dlg.reject();
    QVERIFY(!DefaultModeDialog::hasSavedDefaultMode());
    QCOMPARE(dlg.selectedMode(), ReportMode::Tpm);
}

void TestDefaultMode::buttonsAreStackedVerticallyTpmSensoryDetailed()
{
    DefaultModeDialog dlg;
    dlg.show();   // realize the layout so geometry is meaningful

    QLabel* prompt = nullptr;
    for (QLabel* l : dlg.findChildren<QLabel*>()) {
        if (l->text() == QStringLiteral("Pick a default mode:")) { prompt = l; break; }
    }
    QVERIFY2(prompt, "prompt label 'Pick a default mode:' missing");

    QPushButton* tpm  = dlg.findChild<QPushButton*>("tpmButton");
    QPushButton* sens = dlg.findChild<QPushButton*>("sensoryButton");
    QPushButton* det  = dlg.findChild<QPushButton*>("detailedButton");
    QVERIFY(tpm && sens && det);
    QCOMPARE(tpm->text(),  QStringLiteral("TPM"));
    QCOMPARE(sens->text(), QStringLiteral("Sensory"));
    QCOMPARE(det->text(),  QStringLiteral("Detailed Sensory"));

    // Spec: buttons stacked vertically, TPM top / Sensory middle / Detailed bottom.
    QVERIFY2(tpm->y() < sens->y(),  "TPM must sit above Sensory");
    QVERIFY2(sens->y() < det->y(),  "Sensory must sit above Detailed Sensory");
    QCOMPARE(tpm->x(), sens->x());
    QCOMPARE(sens->x(), det->x());
}

QTEST_MAIN(TestDefaultMode)
#include "tst_defaultmode.moc"
