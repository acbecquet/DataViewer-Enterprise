QT       += core testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database

SOURCES += tst_configloader.cpp \
           ../../src/database/ConfigLoader.cpp

HEADERS += ../../src/database/ConfigLoader.h
