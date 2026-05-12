#include <QtTest/QtTest>
#include "../../src/database/NotificationListener.h"
#include "../../src/database/PostgresConnection.h"

using namespace DVE;

class TstNotificationListener : public QObject {
    Q_OBJECT
private slots:
    void constructor_createsValidObject() {
        PostgresConnection pg;
        NotificationListener nl(&pg);
        QVERIFY(!nl.isSubscribed());
    }

    void structs_canBeInstantiated() {
        RowChange rc;
        rc.table = "test_table";
        rc.op = "INSERT";
        rc.id = 123;
        rc.updatedBy = "user123";

        QCOMPARE(rc.table, QString("test_table"));
        QCOMPARE(rc.op, QString("INSERT"));
        QCOMPARE(rc.id, qint64(123));
    }

    void presenceChangeStruct_canBeInstantiated() {
        PresenceChange pc;
        pc.op = "INSERT";
        pc.userUuid = QUuid::createUuid();
        pc.resourceType = "file";
        pc.resourceId = 42;
        pc.intent = "viewing";

        QCOMPARE(pc.op, QString("INSERT"));
        QCOMPARE(pc.resourceType, QString("file"));
        QCOMPARE(pc.resourceId, qint64(42));
    }

    void metatype_rowChangeCanBeRegistered() {
        qRegisterMetaType<DVE::RowChange>();
        int typeId = QMetaType::type("DVE::RowChange");
        QVERIFY(typeId != QMetaType::UnknownType);
    }

    void metatype_presenceChangeCanBeRegistered() {
        qRegisterMetaType<DVE::PresenceChange>();
        int typeId = QMetaType::type("DVE::PresenceChange");
        QVERIFY(typeId != QMetaType::UnknownType);
    }
};

QTEST_MAIN(TstNotificationListener)
#include "tst_notificationlistener.moc"
