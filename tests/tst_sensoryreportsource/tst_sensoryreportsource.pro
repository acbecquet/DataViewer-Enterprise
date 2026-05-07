QT += core gui testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/reporting ../../src/pipeline ../../src/database

SOURCES += tst_sensoryreportsource.cpp            ../../src/reporting/SensoryReportSource.cpp            ../../src/reporting/ReportLayout.cpp

HEADERS += ../../src/reporting/SensoryReportSource.h            ../../src/reporting/IReportSource.h            ../../src/reporting/ReportLayout.h            ../../src/pipeline/SensoryData.h
