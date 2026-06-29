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
    void splitParsesR3() {
        const TesterRound tr = splitTesterRound("Charlie R3");
        QCOMPARE(tr.tester, QString("Charlie"));
        QCOMPARE(tr.round,  QString("3"));
    }
    void splitParsesR4WithSpaceInName() {
        const TesterRound tr = splitTesterRound("Mary Jane R4");
        QCOMPARE(tr.tester, QString("Mary Jane"));
        QCOMPARE(tr.round,  QString("4"));
    }
    void splitUnknownRoundIsNotParsed() {
        // R5 is outside the supported 1-4 range (v2.5.13 extended R1/R2 -> R1-R4),
        // so it stays part of the name.
        const TesterRound tr = splitTesterRound("Charlie R5");
        QCOMPARE(tr.tester, QString("Charlie R5"));
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
    void combineRound3And4AppendSuffix() {
        QCOMPARE(combineTesterRound("Charlie", "3"), QString("Charlie R3"));
        QCOMPARE(combineTesterRound("Charlie", "4"), QString("Charlie R4"));
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
        const QStringList rounds{ "1", "2", "3", "4", "N/A" };
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
