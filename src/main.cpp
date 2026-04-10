#include <QApplication>
#include <QDebug>
#include <QDir>
#include <QFont>
#include <QIcon>

#include "MainWindow.h"
#include "utils/AppTheme.h"
#include "utils/SingleInstance.h"

static const QString SERVER_KEY = QStringLiteral("DataViewerEnterprise_SingleInstance");

int main(int argc, char* argv[])
{
    QApplication app(argc, argv);
    app.setApplicationName("DataViewer Enterprise");
    app.setApplicationVersion("0.9.1");
    app.setOrganizationName("SDR");
    app.setOrganizationDomain("sdr.com");

    // --- Single-instance check ---
    // Find file argument (if any) before deciding what to do
    QString fileArg;
    const QStringList args = app.arguments();
    for (int i = 1; i < args.size(); ++i) {
        if (QFile::exists(args[i])) {
            fileArg = args[i];
            break;
        }
    }

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
