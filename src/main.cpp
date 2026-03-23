#include <QApplication>
#include <QDebug>
#include <QDir>
#include <QFont>
#include <QIcon>

#include "MainWindow.h"
#include "utils/AppTheme.h"

int main(int argc, char* argv[])
{
    QApplication app(argc, argv);
    app.setApplicationName("DataViewer Enterprise");
    app.setApplicationVersion("1.0.0");
    app.setOrganizationName("SDR");
    app.setOrganizationDomain("sdr.com");

    // Load app icon from filesystem (Qt 6.10 rcc doesn't generate C++ source)
    QStringList iconCandidates = {
        QCoreApplication::applicationDirPath() + "/resources/images/ccell_icon.png",
        QCoreApplication::applicationDirPath() + "/../resources/images/ccell_icon.png",
        "C:/Users/S1134987/Documents/Python/DataViewer Dev/DataViewer-Enterprise/resources/images/ccell_icon.png"
    };
    for (const QString& p : iconCandidates) {
        if (QFile::exists(p)) { app.setWindowIcon(QIcon(p)); break; }
    }

    // Apply professional engineering theme
    AppTheme::apply();

    qDebug() << "DataViewer Enterprise starting | Qt" << QT_VERSION_STR;

    DVE::MainWindow window;
    window.show();

    return app.exec();
}
