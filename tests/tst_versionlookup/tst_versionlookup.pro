QT       += core testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database ../../src/pipeline

SOURCES += tst_versionlookup.cpp            ../../src/database/VersionLookup.cpp            ../../src/pipeline/SensoryData.cpp

HEADERS += ../../src/database/VersionLookup.h            ../../src/pipeline/ReportData.h            ../../src/pipeline/SensoryData.h            ../../src/pipeline/DetailedSensoryData.h
