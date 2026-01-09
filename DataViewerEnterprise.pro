QT += core gui widgets sql

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

TARGET = DataViewerEnterprise
TEMPLATE = app

SOURCES += \
	src/main.cpp \
	src/MainWindow.cpp

HEADERS += \
	src/MainWindow.h

INCLUDEPATH += src

DEFINES += QT_DEPRECATED_WARNINGS

qnx: target.path = /tmp/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target