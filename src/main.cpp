#include <QApplication>
#include <QDateTime>
#include <QDebug>
#include <QDir>
#include <QFile>
#include <QFont>
#include <QIcon>
#include <QCommandLineParser>
#include <QMutex>
#include <QStandardPaths>
#include <QTextStream>

#ifdef Q_OS_WIN
#  include <windows.h>
#endif

#include "MainWindow.h"
#include "utils/AppTheme.h"
#include "utils/SelfTest.h"
#include "utils/SingleInstance.h"
#include "database/MigrationTool.h"
#include "database/ConfigLoader.h"

static const QString SERVER_KEY = QStringLiteral("DataViewerEnterprise_SingleInstance");

// v2.0.5 file logger. Routes every qDebug/qInfo/qWarning/qCritical to
// %LOCALAPPDATA%/DataViewer/dataviewer.log in append mode, so a deployed
// (windows-subsystem) binary has somewhere to surface "Check the log
// for details" — the previous build wrote only to OutputDebugString,
// which was invisible without DebugView attached.
//
// Rotates once at 5 MB (truncates and starts over) so the file never
// grows unbounded across many sessions. Mutex-protected because Qt
// may call the handler from worker threads (LiveSyncWorker, etc.).
static QFile*  g_logFile = nullptr;
static QMutex  g_logMutex;

static void dveMessageHandler(QtMsgType type,
                              const QMessageLogContext& /*ctx*/,
                              const QString& msg)
{
    QMutexLocker lock(&g_logMutex);
    if (!g_logFile || !g_logFile->isOpen()) return;

    if (g_logFile->size() > 5 * 1024 * 1024) {
        g_logFile->close();
        g_logFile->open(QIODevice::WriteOnly | QIODevice::Truncate);
    }

    const char* lvl = "INFO";
    switch (type) {
        case QtDebugMsg:    lvl = "DEBUG"; break;
        case QtInfoMsg:     lvl = "INFO";  break;
        case QtWarningMsg:  lvl = "WARN";  break;
        case QtCriticalMsg: lvl = "ERROR"; break;
        case QtFatalMsg:    lvl = "FATAL"; break;
    }

    QTextStream ts(g_logFile);
    ts.setEncoding(QStringConverter::Utf8);
    ts << QDateTime::currentDateTime().toString(Qt::ISODateWithMs)
       << " [" << lvl << "] " << msg << '\n';
    ts.flush();
    g_logFile->flush();
}

static void installFileLogger()
{
    const QString logDir =
        QStandardPaths::writableLocation(QStandardPaths::AppLocalDataLocation);
    if (logDir.isEmpty() || !QDir().mkpath(logDir)) return;
    const QString logPath = logDir + "/dataviewer.log";
    g_logFile = new QFile(logPath);
    if (!g_logFile->open(QIODevice::WriteOnly | QIODevice::Append)) {
        delete g_logFile;
        g_logFile = nullptr;
        return;
    }
    qInstallMessageHandler(dveMessageHandler);
}

int main(int argc, char* argv[])
{
    QApplication app(argc, argv);
    app.setApplicationName("DataViewer Enterprise");
    app.setApplicationVersion(QStringLiteral(DVE_APP_VERSION));
    app.setOrganizationName("SDR");
    app.setOrganizationDomain("sdr.com");

    // v2.0.5: install file logger BEFORE anything else so we capture
    // even early startup diagnostics. Has to be after setApplicationName
    // / setOrganizationName because writableLocation uses those.
    installFileLogger();
    qInfo().noquote() << "──── DataViewer Enterprise v" DVE_APP_VERSION
                      << "started at"
                      << QDateTime::currentDateTime().toString(Qt::ISODate)
                      << "────";

    // Named kernel mutex so Inno Setup's AppMutex= can detect that we're
    // running and offer to close the app before installing an update.
    // Released automatically when the process exits.
#ifdef Q_OS_WIN
    HANDLE updateMutex = ::CreateMutexW(nullptr, FALSE,
        L"DataViewerEnterprise_SingleInstance");
    Q_UNUSED(updateMutex)
#endif

    // --- Argument scan ---
    // Use QCommandLineParser to handle multiple modes:
    //   DataViewer.exe <file>                  open file in existing instance
    //   --self-test [--self-test-out PATH]     run deployment diagnostics and exit
    //   --migrate-from-sqlite=PATH --to-postgres="..." [--force] [--migration-report-out PATH]
    //   --encrypt-password=PLAIN --encrypted-out=PATH
    QCommandLineParser parser;
    QCommandLineOption optSelfTest("self-test", "Run deployment diagnostics and exit");
    QCommandLineOption optSelfTestOut("self-test-out", "Write JSON report to PATH (with --self-test)", "path");
    QCommandLineOption optMigrate("migrate-from-sqlite", "Migrate a SQLite database to Postgres (headless)", "path");
    QCommandLineOption optTo("to-postgres", "Target Postgres connection string (key=value form, space-separated)", "conn");
    QCommandLineOption optForce("force", "Wipe existing Postgres data before migrating");
    QCommandLineOption optReportOut("migration-report-out", "Path to write migration JSON report", "path");
    QCommandLineOption optEncrypt("encrypt-password", "Encrypt a password and exit", "plain");
    QCommandLineOption optEncryptOut("encrypted-out", "Output file path for --encrypt-password", "path");

    parser.addOption(optSelfTest);
    parser.addOption(optSelfTestOut);
    parser.addOption(optMigrate);
    parser.addOption(optTo);
    parser.addOption(optForce);
    parser.addOption(optReportOut);
    parser.addOption(optEncrypt);
    parser.addOption(optEncryptOut);
    parser.parse(QCoreApplication::arguments());

    // Handle password encryption (used by installer in Task 24)
    if (parser.isSet(optEncrypt) && parser.isSet(optEncryptOut)) {
        const QString enc = DVE::ConfigLoader::encryptPassword(parser.value(optEncrypt));
        QFile out(parser.value(optEncryptOut));
        if (!out.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
            qCritical() << "Could not open output file:" << parser.value(optEncryptOut);
            return 3;
        }
        out.write(enc.toUtf8());
        out.close();
        return 0;
    }

    // Handle migration (headless mode)
    if (parser.isSet(optMigrate)) {
        const QString src     = parser.value(optMigrate);
        const QString connStr = parser.value(optTo);
        const bool    force   = parser.isSet(optForce);
        const QString reportPath = parser.isSet(optReportOut)
            ? parser.value(optReportOut)
            : (QDir::tempPath() + "/dataviewer_migration.json");

        // Parse connection string into DbConfig
        DVE::DbConfig cfg;
        for (const QString& part : connStr.split(' ', Qt::SkipEmptyParts)) {
            const QStringList kv = part.split('=');
            if (kv.size() != 2) continue;
            const QString k = kv[0], v = kv[1];
            if      (k == "host")     cfg.host = v;
            else if (k == "port")     cfg.port = v.toInt();
            else if (k == "dbname")   cfg.database = v;
            else if (k == "user")     cfg.user = v;
            else if (k == "password") cfg.password = v;
        }

        // Run migration
        DVE::MigrationTool m;
        if (!m.open(src, cfg)) {
            qCritical().noquote() << "Migration setup failed:" << m.lastError();
            m.report().writeJson(reportPath);
            return 2;
        }
        const bool ok = m.run(force);
        m.report().writeJson(reportPath);
        qInfo().noquote() << "Migration:" << m.report().summary();
        qInfo().noquote() << "Report written to:" << reportPath;
        return ok ? 0 : 1;
    }

    // Handle self-test
    if (parser.isSet(optSelfTest)) {
        return DVE::runSelfTest(parser.value(optSelfTestOut));
    }

    // Normal mode: extract file argument from positional arguments
    QString fileArg;
    const QStringList positional = parser.positionalArguments();
    for (const QString& a : positional) {
        if (fileArg.isEmpty() && QFile::exists(a)) {
            fileArg = a;
            break;
        }
    }

    // --- Single-instance check ---
    SingleInstance instance(SERVER_KEY);

    // If another instance is running, send it the file path and exit
    if (instance.sendToRunning(fileArg.isEmpty() ? QStringLiteral("__RAISE__") : fileArg)) {
        qDebug() << "Another instance is running; sent message and exiting.";
        return 0;
    }

    // We are the primary instance — start the server
    instance.startServer();

    // Load app icon from filesystem (Qt 6.10 rcc doesn't generate C++ source)
    QStringList iconCandidates = {
        QCoreApplication::applicationDirPath() + "/resources/images/ccell_icon.png",
        QCoreApplication::applicationDirPath() + "/../resources/images/ccell_icon.png",
    };
    for (const QString& p : iconCandidates) {
        if (QFile::exists(p)) { app.setWindowIcon(QIcon(p)); break; }
    }

    // Apply professional engineering theme
    AppTheme::apply();

    qDebug() << "DataViewer Enterprise starting | Qt" << QT_VERSION_STR;

    DVE::MainWindow window;
    window.show();

    // Load file from command-line argument
    if (!fileArg.isEmpty()) {
        window.loadFile(fileArg);
    }

    // When another instance sends a message, load the file and bring window to front
    QObject::connect(&instance, &SingleInstance::messageReceived, &window,
        [&window](const QString& message) {
            if (!message.isEmpty() && message != "__RAISE__" && QFile::exists(message)) {
                window.loadFile(message);
            }
            // Bring window to front
            window.setWindowState((window.windowState() & ~Qt::WindowMinimized) | Qt::WindowActive);
            window.raise();
            window.activateWindow();
        });

    return app.exec();
}
