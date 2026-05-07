QT += core gui sql network testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/reporting ../../src/pipeline ../../src/database ../../src/utils

SOURCES += tst_sensoryreportsource.cpp \
           ../../src/reporting/SensoryReportSource.cpp \
           ../../src/reporting/ReportLayout.cpp \
           ../../src/database/DatabaseManager.cpp

HEADERS += ../../src/reporting/SensoryReportSource.h \
           ../../src/reporting/IReportSource.h \
           ../../src/reporting/ReportLayout.h \
           ../../src/database/DatabaseManager.h \
           ../../src/pipeline/SensoryData.h
