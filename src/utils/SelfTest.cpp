#include "SelfTest.h"
#include "UpdateChecker.h"
#include "../database/ConfigLoader.h"
#include "../database/PostgresConnection.h"

#include <QApplication>
#include <QCoreApplication>
#include <QDateTime>
#include <QDir>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QProcess>
#include <QSettings>
#include <QSqlDatabase>
#include <QSqlError>
#include <QSqlQuery>
#include <QStandardPaths>
#include <QString>
#include <QTextStream>
#include <QVersionNumber>

namespace DVE {

namespace {

struct TestResult {
    QString name;
    bool    pass = false;
    QString detail;
};

QString bundledPython()
{
    return QCoreApplication::applicationDirPath() + "/python/python.exe";
}

TestResult testRegistry()
{
    TestResult r{"registry", false, ""};
    QSettings s("SDR", "DataViewerEnterprise");
    const QString key   = QStringLiteral("selftest/probe");
    const QString value = QString::number(QDateTime::currentMSecsSinceEpoch());

    s.setValue(key, value);
    s.sync();
    if (s.status() != QSettings::NoError) {
        r.detail = QStringLiteral("QSettings status=%1 (write)").arg(s.status());
        return r;
    }

    const QString got = s.value(key).toString();
    s.remove(key);
    s.sync();

    if (got != value) {
        r.detail = QStringLiteral("readback mismatch");
        return r;
    }

    r.pass = true;
    return r;
}

TestResult testSqlite()
{
    TestResult r{"sqlite", false, ""};
    const QString tmpDb = QDir::tempPath() + "/dataviewer_selftest.db";
    QFile::remove(tmpDb);

    {
        QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE", "selftest_conn");
        db.setDatabaseName(tmpDb);
        if (!db.open()) {
            r.detail = QStringLiteral("open failed: %1").arg(db.lastError().text());
            QSqlDatabase::removeDatabase("selftest_conn");
            return r;
        }

        QSqlQuery q(db);
        const bool ok =
            q.exec("CREATE TABLE t (i INTEGER)") &&
            q.exec("INSERT INTO t VALUES (42)") &&
            q.exec("SELECT i FROM t");
        if (!ok) {
            r.detail = QStringLiteral("query failed: %1").arg(q.lastError().text());
            db.close();
            QSqlDatabase::removeDatabase("selftest_conn");
            return r;
        }

        if (!q.next() || q.value(0).toInt() != 42) {
            r.detail = QStringLiteral("readback failed");
            db.close();
            QSqlDatabase::removeDatabase("selftest_conn");
            return r;
        }
        db.close();
    }
    QSqlDatabase::removeDatabase("selftest_conn");
    QFile::remove(tmpDb);

    r.pass   = true;
    r.detail = tmpDb;
    return r;
}

TestResult testPythonBundle()
{
    TestResult r{"python_bundle", false, ""};
    const QString python = bundledPython();
    if (!QFile::exists(python)) {
        r.detail = QStringLiteral("not found at %1").arg(python);
        return r;
    }

    QProcess p;
    p.start(python, {QStringLiteral("-c"),
                     QStringLiteral("import sys, openpyxl; "
                                    "print(sys.version.split()[0] + '|' + openpyxl.__version__)")});
    if (!p.waitForFinished(15000)) {
        r.detail = QStringLiteral("timed out (15s)");
        return r;
    }

    const QString out = QString::fromLocal8Bit(p.readAllStandardOutput()).trimmed();
    const QString err = QString::fromLocal8Bit(p.readAllStandardError()).trimmed();
    if (p.exitCode() != 0) {
        r.detail = QStringLiteral("exit=%1 stderr=%2").arg(p.exitCode()).arg(err);
        return r;
    }
    r.pass   = true;
    r.detail = out;
    return r;
}

TestResult testExcelRoundtrip()
{
    TestResult r{"excel_roundtrip", false, ""};
    const QString python = bundledPython();
    if (!QFile::exists(python)) {
        r.detail = QStringLiteral("python not found");
        return r;
    }

    const QString tmpXlsx = QDir::tempPath() + "/dataviewer_selftest.xlsx";
    QFile::remove(tmpXlsx);

    const QString script = QStringLiteral(
        "import sys\n"
        "from openpyxl import Workbook, load_workbook\n"
        "path = sys.argv[1]\n"
        "wb = Workbook(); ws = wb.active; ws['A1'] = 'hello'; wb.save(path)\n"
        "wb2 = load_workbook(path)\n"
        "print(wb2.active['A1'].value)\n");

    QProcess p;
    p.start(python, {QStringLiteral("-c"), script, tmpXlsx});
    if (!p.waitForFinished(15000)) {
        r.detail = QStringLiteral("timed out (15s)");
        QFile::remove(tmpXlsx);
        return r;
    }

    const QString out = QString::fromLocal8Bit(p.readAllStandardOutput()).trimmed();
    const QString err = QString::fromLocal8Bit(p.readAllStandardError()).trimmed();
    QFile::remove(tmpXlsx);

    if (p.exitCode() != 0 || out != QStringLiteral("hello")) {
        r.detail = QStringLiteral("exit=%1 out='%2' stderr=%3")
                       .arg(p.exitCode()).arg(out, err);
        return r;
    }
    r.pass = true;
    return r;
}

TestResult testSynologyRoot()
{
    TestResult r{"synology_root", false, ""};
    const UpdateChecker::ProbeResult probe = UpdateChecker::probe();
    r.detail = probe.updateRoot;
    if (!probe.rootExists) {
        r.detail += QStringLiteral(" (does not exist)");
        return r;
    }
    r.pass = true;
    return r;
}

TestResult testLatestVersion()
{
    TestResult r{"latest_version", false, ""};
    const UpdateChecker::ProbeResult probe = UpdateChecker::probe();
    if (!probe.rootExists) {
        r.detail = QStringLiteral("update root missing — see synology_root");
        return r;
    }
    if (probe.latest.isNull() || probe.installerPath.isEmpty()) {
        r.detail = QStringLiteral("no version subdir contained DataViewer-setup.exe");
        return r;
    }
    r.pass   = true;
    r.detail = QStringLiteral("%1 @ %2").arg(probe.latest.toString(), probe.installerPath);
    return r;
}

TestResult testPostgresConnection()
{
    TestResult r{"postgres_connection", false, ""};

    const QString confPath = QString::fromLocal8Bit(qgetenv("PROGRAMDATA"))
                              + "/DataViewer/db.conf";
    DVE::DbConfig cfg;
    QString err;
    if (!DVE::ConfigLoader::load(confPath, cfg, &err)) {
        r.detail = QStringLiteral("Could not read %1: %2").arg(confPath, err);
        return r;
    }

    DVE::PostgresConnection pg;
    if (!pg.open(cfg)) {
        r.detail = QStringLiteral("Connect to %1:%2 failed: %3")
                       .arg(cfg.host).arg(cfg.port).arg(pg.lastError());
        return r;
    }
    if (!pg.ping()) {
        r.detail = QStringLiteral("Ping failed after connect");
        pg.close();
        return r;
    }
    pg.close();
    r.pass   = true;
    r.detail = QStringLiteral("Connected to %1 successfully").arg(cfg.host);
    return r;
}

TestResult testAppDataWritable()
{
    TestResult r{"appdata_writable", false, ""};
    const QString dir = QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
    QDir().mkpath(dir);
    const QString probe = dir + "/.selftest_probe";
    QFile f(probe);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        r.detail = QStringLiteral("cannot write %1: %2").arg(probe, f.errorString());
        return r;
    }
    f.write("ok");
    f.close();
    QFile::remove(probe);
    r.pass   = true;
    r.detail = dir;
    return r;
}

} // anonymous namespace

int runSelfTest(const QString& outputPath)
{
    const QList<TestResult> results = {
        testAppDataWritable(),
        testRegistry(),
        testSqlite(),
        testPythonBundle(),
        testExcelRoundtrip(),
        testSynologyRoot(),
        testLatestVersion(),
        testPostgresConnection(),
    };

    QJsonArray arr;
    bool allPass = true;
    for (const TestResult& r : results) {
        QJsonObject o;
        o["name"]   = r.name;
        o["pass"]   = r.pass;
        o["detail"] = r.detail;
        arr.append(o);
        if (!r.pass) allPass = false;
    }

    QJsonObject root;
    root["version"]  = QApplication::applicationVersion();
    root["all_pass"] = allPass;
    root["tests"]    = arr;

    const QByteArray json = QJsonDocument(root).toJson(QJsonDocument::Indented);

    if (!outputPath.isEmpty()) {
        QFile f(outputPath);
        if (f.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
            f.write(json);
            f.close();
        }
    }

    QTextStream(stderr) << json << Qt::endl;

    return allPass ? 0 : 1;
}

} // namespace DVE
