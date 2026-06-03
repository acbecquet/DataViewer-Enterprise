QT       += core testlib
QT       -= gui
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app
TARGET    = tst_outputpaths

INCLUDEPATH += ../../src ../../src/utils

SOURCES += tst_outputpaths.cpp            ../../src/utils/OutputPaths.cpp
HEADERS += ../../src/utils/OutputPaths.h
