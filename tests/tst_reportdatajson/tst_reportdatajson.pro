QT += core testlib
QT -= gui
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_reportdatajson

INCLUDEPATH += ../../src ../../src/pipeline

SOURCES += tst_reportdatajson.cpp ../../src/pipeline/ReportDataJson.cpp
HEADERS += ../../src/pipeline/ReportDataJson.h ../../src/pipeline/ReportData.h
