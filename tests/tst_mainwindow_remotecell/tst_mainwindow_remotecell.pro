QT       += core gui widgets testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src

SOURCES += tst_mainwindow_remotecell.cpp \
           ../../src/RemoteCellHelpers.cpp

HEADERS += ../../src/RemoteCellHelpers.h
