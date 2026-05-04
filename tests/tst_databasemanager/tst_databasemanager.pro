QT += core sql testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

SOURCES += tst_databasemanager.cpp \
           ../../src/database/DatabaseManager.cpp

HEADERS += ../../src/database/DatabaseManager.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/SensoryData.h \
           ../../src/pipeline/DetailedSensoryData.h \
           ../common/TestHelpers.h
