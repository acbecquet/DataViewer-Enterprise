#include <QtTest>
#include "ui/TesterRound.h"

using namespace DVE;

class tst_TesterRound : public QObject
{
    Q_OBJECT
private slots:
    void splitParsesR1() {
        const TesterRound tr = splitTesterRound("Charlie R1");
        QCOMPARE(tr.tester, QString("Charlie"));
        QCOMPARE(tr.round,  QString("1"));
    }
    void splitParsesR2WithSpaceInName() {
        const TesterRound tr = splitTesterRound("Mary Jane R2");
        QCOMPARE(tr.tester, QString("Mary Jane"));
        QCOMPARE(tr.round,  QString("2"));
    }
    void splitPlainNameYieldsNA() {
        const TesterRound tr = splitTesterRound("Charlie");
        QCOMPARE(tr.tester, QString("Charlie"));
        QCOMPARE(tr.round,  QString("N/A"));
    }
    void splitUnknownRoundIsNotParsed() {
        const TesterRound tr = splitTesterRound("Charlie R3");
        QCOMPARE(tr.tester, QString("Charlie R3"));
        QCOMPARE(tr.round,  QString("N/A"));
    }
    void splitEmptyYieldsNA() {
        const TesterRound tr = splitTesterRound("");
        QCOMPARE(tr.tester, QString(""));
        QCOMPARE(tr.round,  QString("N/A"));
    }
    void combineRound1AppendsSuffix() {
        QCOMPARE(combineTesterRound("Charlie", "1"), QString("Charlie R1"));
    }
    void combineRound2AppendsSuffix() {
        QCOMPARE(combineTesterRound("Charlie", "2"), QString("Charlie R2"));
    }
    void combineNAAppendsNothing() {
        QCOMPARE(combineTesterRound("Charlie", "N/A"), QString("Charlie"));
    }
    void combineTrimsAndKeepsEmptyEmpty() {
        QCOMPARE(combineTesterRound("  Charlie  ", "1"), QString("Charlie R1"));
        QCOMPARE(combineTesterRound("", "1"), QString(""));
        QCOMPARE(combineTesterRound("   ", "2"), QString(""));
    }
    void roundTripAllRounds() {
        const QStringList rounds{ "1", "2", "N/A" };
        for (const QString& r : rounds) {
            const QString combined = combineTesterRound("Charlie", r);
            const TesterRound tr = splitTesterRound(combined);
            QCOMPARE(tr.tester, QString("Charlie"));
            QCOMPARE(tr.round,  r);
        }
    }
};

QTEST_APPLESS_MAIN(tst_TesterRound)
#include "tst_testerround.moc"
