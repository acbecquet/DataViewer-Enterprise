#include <QtTest/QtTest>
#include <QTemporaryFile>
#include "../../src/database/ConfigLoader.h"

using DVE::ConfigLoader;
using DVE::DbConfig;

class TstConfigLoader : public QObject {
    Q_OBJECT
private slots:
    void encryptDecrypt_roundTrip_returnsOriginal() {
        const QString plain = QStringLiteral("s3cr3t-password!@#");
        const QString encrypted = ConfigLoader::encryptPassword(plain);
        QVERIFY(encrypted != plain);
        QCOMPARE(ConfigLoader::decryptPassword(encrypted), plain);
    }

    void parseConfig_validIni_returnsPopulatedStruct() {
        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        const QByteArray content =
            "[postgres]\n"
            "host=db.example.local\n"
            "port=5432\n"
            "database=dataviewer\n"
            "user=dataviewer_app\n"
            "password_encrypted=" + ConfigLoader::encryptPassword("hunter2").toUtf8() + "\n";
        tmp.write(content);
        tmp.flush();
        DbConfig cfg;
        QVERIFY(ConfigLoader::load(tmp.fileName(), cfg));
        QCOMPARE(cfg.host,     QString("db.example.local"));
        QCOMPARE(cfg.port,     5432);
        QCOMPARE(cfg.database, QString("dataviewer"));
        QCOMPARE(cfg.user,     QString("dataviewer_app"));
        QCOMPARE(cfg.password, QString("hunter2"));
    }

    void parseConfig_missingFile_returnsFalse() {
        DbConfig cfg;
        QVERIFY(!ConfigLoader::load("C:/no/such/file.conf", cfg));
    }

    void parseConfig_missingPassword_returnsFalseWithError() {
        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        tmp.write("[postgres]\nhost=h\nport=5432\ndatabase=d\nuser=u\n");
        tmp.flush();
        DbConfig cfg;
        QString err;
        QVERIFY(!ConfigLoader::load(tmp.fileName(), cfg, &err));
        QVERIFY(err.contains("password", Qt::CaseInsensitive));
    }

    void parseConfig_missingHost_returnsFalseWithError() {
        QTemporaryFile tmp;
        QVERIFY(tmp.open());
        tmp.write("[postgres]\nport=5432\ndatabase=d\nuser=u\npassword_encrypted=x\n");
        tmp.flush();
        DbConfig cfg;
        QString err;
        QVERIFY(!ConfigLoader::load(tmp.fileName(), cfg, &err));
        QVERIFY(err.contains("host", Qt::CaseInsensitive));
    }
};

QTEST_MAIN(TstConfigLoader)
#include "tst_configloader.moc"
