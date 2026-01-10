QT += core gui widgets sql

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17

TARGET = DataViewerEnterprise
TEMPLATE = app

# Include QXlsx source files directly
include(external/QXlsx/QXlsx/QXlsx.pri)

INCLUDEPATH += $$PWD/external/QXlsx/header
#LIBS += -L$$PWD/external/QXlsx/lib -lQXlsx

SOURCES += \
	src/main.cpp \
	src/MainWindow.cpp \
	src/ExcelReader.cpp

HEADERS += \
        src/MainWindow.h \
	src/ExcelReader.h

INCLUDEPATH += src

DEFINES += QT_DEPRECATED_WARNINGS

qnx: target.path = /tmp/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target
