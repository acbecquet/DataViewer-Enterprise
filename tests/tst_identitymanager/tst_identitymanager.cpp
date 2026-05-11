#include <QtTest/QtTest>
#include <QCoreApplication>
#include <QSettings>
#include "../../src/database/IdentityManager.h"

using DVE::IdentityManager;

class TstIdentityManager : public QObject {
    Q_OBJECT
private slots:
    void initTestCase() {
        QCoreApplication::setOrganizationName("DataViewerTest");
        QCoreApplication::setApplicationName("tst_identitymanager");
    }

    // Per-slot setup: clear identity state so each test starts from a known
    // baseline regardless of slot order. (Qt may emit harmless RegDeleteKey
    // warnings on Windows when the key is already absent — they are not
    // test failures.)
    void init() {
        QSettings s;
        s.clear();
    }

    void uuid_isGenerated_onFirstAccess() {
        IdentityManager m;
        QVERIFY(!m.uuid().isNull());
        QVERIFY(!m.uuid().toString().isEmpty());
    }

    void uuid_persistsAcrossInstances() {
        IdentityManager m1;
        const QString u1 = m1.uuid().toString();
        IdentityManager m2;
        const QString u2 = m2.uuid().toString();
        QCOMPARE(u1, u2);
    }

    void uuid_regeneratesWhenStoredValueCorrupt() {
        QSettings s;
        s.setValue("identity/uuid", "not-a-uuid");
        IdentityManager m;
        QVERIFY(!m.uuid().isNull());
        QVERIFY(m.uuid().toString(QUuid::WithoutBraces) != QStringLiteral("not-a-uuid"));
    }

    void displayName_defaultsToEmpty_untilSet() {
        IdentityManager m;
        QVERIFY(m.displayName().isEmpty());
    }

    void displayName_persistsAfterSet() {
        IdentityManager m;
        m.setDisplayName("Charlie B.");
        IdentityManager m2;
        QCOMPARE(m2.displayName(), QString("Charlie B."));
    }

    void color_defaultsToEmpty_untilSet() {
        IdentityManager m;
        QVERIFY(m.color().isEmpty());
    }

    void color_persistsAfterSet() {
        IdentityManager m;
        m.setColor("#3b82f6");
        IdentityManager m2;
        QCOMPARE(m2.color(), QString("#3b82f6"));
    }

    void firstLaunchPending_trueWhenBothEmpty() {
        IdentityManager m;
        QVERIFY(m.firstLaunchPending());
    }

    void firstLaunchPending_trueWhenOnlyNameMissing() {
        IdentityManager m;
        m.setColor("#10b981");
        // displayName still empty
        QVERIFY(m.firstLaunchPending());
    }

    void firstLaunchPending_trueWhenOnlyColorMissing() {
        IdentityManager m;
        m.setDisplayName("Bob");
        // color still empty
        QVERIFY(m.firstLaunchPending());
    }

    void firstLaunchPending_falseAfterNameAndColorSet() {
        IdentityManager m;
        m.setDisplayName("Sarah");
        m.setColor("#10b981");
        QVERIFY(!m.firstLaunchPending());
    }

    void defaultColorPalette_has12Hexes() {
        const QStringList palette = IdentityManager::defaultColorPalette();
        QCOMPARE(palette.size(), 12);
        for (const QString& hex : palette) {
            QVERIFY2(hex.length() == 7, qPrintable("bad length: " + hex));
            QVERIFY2(hex.startsWith('#'), qPrintable("missing #: " + hex));
        }
    }
};

QTEST_MAIN(TstIdentityManager)
#include "tst_identitymanager.moc"
