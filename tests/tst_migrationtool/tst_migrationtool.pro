QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database ../common

SOURCES += tst_migrationtool.cpp            ../../src/database/MigrationTool.cpp            ../../src/database/MigrationReport.cpp            ../../src/database/ConfigLoader.cpp

HEADERS += ../../src/database/MigrationTool.h            ../../src/database/MigrationReport.h            ../../src/database/ConfigLoader.h
