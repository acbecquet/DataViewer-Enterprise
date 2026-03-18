#include <QApplication>
#include <QDebug>
#include <QFont>

#include "MainWindow.h"
#include "utils/AppTheme.h"

int main(int argc, char* argv[])
{
    QApplication app(argc, argv);
    app.setApplicationName("DataViewer Enterprise");
    app.setApplicationVersion("1.0.0");
    app.setOrganizationName("SDR");
    app.setOrganizationDomain("sdr.com");

    // Apply professional engineering theme
    AppTheme::apply();

    qDebug() << "DataViewer Enterprise starting | Qt" << QT_VERSION_STR;

    DVE::MainWindow window;
    window.show();

    return app.exec();
}
