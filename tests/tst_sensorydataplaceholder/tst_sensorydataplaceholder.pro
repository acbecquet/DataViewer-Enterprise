QT       += core testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/pipeline

SOURCES += tst_sensorydataplaceholder.cpp

HEADERS += ../../src/pipeline/SensoryData.h
