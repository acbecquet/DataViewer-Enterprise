#include "MainWindow.h"
#include <QApplication>
#include <QDebug>

int main(int argc, char* argv[])
{
	qDebug() << "DEBUG: Application starting..";

	QApplication app(argc, argv);
	app.setApplicationName("DataViewer Enterprise");
	app.setOrganizationName("SDR");

	qDebug() << "DEBUG: Creating MainWindow...";
	MainWindow window;
	window.show();

	qDebug() << "DEBUG: Entering main event loop...";
	return app.exec();
}