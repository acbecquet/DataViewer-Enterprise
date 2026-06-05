QT       += core testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/pipeline

SOURCES += tst_detailedsensoryjson.cpp \
           ../../src/pipeline/DetailedSensoryData.cpp

HEADERS += ../../src/pipeline/DetailedSensoryData.h
