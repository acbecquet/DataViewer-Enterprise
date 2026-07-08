#include <QtTest>
#include "pipeline/HeatingTech.h"

using namespace DVE;

class TstHeatingTech : public QObject
{
    Q_OBJECT
private slots:
    void offset_t58gFamilyIs078()
    {
        QCOMPARE(heatingTechResistanceOffset("T58G"), 0.78);
        QCOMPARE(heatingTechResistanceOffset("CCELL3.0"), 0.78);
        QCOMPARE(heatingTechResistanceOffset("CCELL 3.0"), 0.78);
        // New T58G variants share the family offset.
        QCOMPARE(heatingTechResistanceOffset("S17B"), 0.78);
        QCOMPARE(heatingTechResistanceOffset("S25B"), 0.78);
        QCOMPARE(heatingTechResistanceOffset("S25B1"), 0.78);
    }

    void offset_t51Is025_othersZero()
    {
        QCOMPARE(heatingTechResistanceOffset("T51"), 0.25);
        QCOMPARE(heatingTechResistanceOffset("EVO"), 0.0);
        QCOMPARE(heatingTechResistanceOffset(""), 0.0);
        QCOMPARE(heatingTechResistanceOffset("Competitor"), 0.0);
    }

    void offset_trimsAndUppercases()
    {
        QCOMPARE(heatingTechResistanceOffset("  t58g  "), 0.78);
        QCOMPARE(heatingTechResistanceOffset("s25b1"), 0.78);
    }

    void defaultResistance_variantsSeed()
    {
        double r = -1.0;
        QVERIFY(heatingTechDefaultResistance("S17B", r));  QCOMPARE(r, 1.3);
        QVERIFY(heatingTechDefaultResistance("S25B", r));  QCOMPARE(r, 1.1);
        QVERIFY(heatingTechDefaultResistance("s25b1", r)); QCOMPARE(r, 1.0);
    }

    void defaultResistance_nonVariantsReturnFalse()
    {
        double r = 42.0;
        QVERIFY(!heatingTechDefaultResistance("T58G", r));
        QVERIFY(!heatingTechDefaultResistance("EVO", r));
        QVERIFY(!heatingTechDefaultResistance("", r));
        QCOMPARE(r, 42.0);  // out is left untouched on false
    }
};

QTEST_MAIN(TstHeatingTech)
#include "tst_heatingtech.moc"
