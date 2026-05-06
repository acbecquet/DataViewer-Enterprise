#include <QApplication>
#include <QDebug>
#include <QDir>
#include <QFont>
#include <QIcon>

#ifdef Q_OS_WIN
#  include <windows.h>
#endif

#include "MainWindow.h"
#include "utils/AppTheme.h"
#include "utils/SelfTest.h"
#include "utils/SingleInstance.h"

static const QString SERVER_KEY = QStringLiteral("DataViewerEnterprise_SingleInstance");

int main(int argc, char* argv[])
{
    QApplication app(argc, argv);
    app.setApplicationName("DataViewer Enterprise");
    app.setApplicationVersion(QStringLiteral(DVE_APP_VERSION));
    app.setOrganizationName("SDR");
    app.setOrganizationDomain("sdr.com");

    // Named kernel mutex so Inno Setup's AppMutex= can detect that we're
    // running and offer to close the app before installing an update.
    // Released automatically when the process exits.
#ifdef Q_OS_WIN
    HANDLE updateMutex = ::CreateMutexW(nullptr, FALSE,
        L"DataViewerEnterprise_SingleInstance");
    Q_UNUSED(updateMutex)
#endif

    // --- Argument scan ---
    // Two flags besides a file path are recognised:
    //   --self-test               run deployment diagnostics and exit
    //   --self-test-out PATH      write JSON report to PATH (with --self-test)
    QString fileArg;
    QString selfTestOut;
    bool    selfTest = false;
    const QStringList args = app.arguments();
    for (int i = 1; i < args.size(); ++i) {
        const QString& a = args[i];
        if (a == QLatin1String("--self-test")) {
            selfTest = true;
        } else if (a == QLatin1String("--self-test-out") && i + 1 < args.size()) {
            selfTestOut = args[++i];
        } else if (fileArg.isEmpty() && QFile::exists(a)) {
            fileArg = a;
        }
    }

    if (selfTest) {
        return DVE::runSelfTest(selfTestOut);
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
