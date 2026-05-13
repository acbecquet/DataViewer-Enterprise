QT += core sql testlib network
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting                ../../src/reporting ../../src/database ../common

SOURCES += tst_connectionmonitor.cpp            ../../src/database/ConnectionMonitor.cpp            ../../src/database/PostgresConnection.cpp            ../../src/database/ConfigLoader.cpp

HEADERS += ../../src/database/ConnectionMonitor.h            ../../src/database/PostgresConnection.h            ../../src/database/ConfigLoader.h
