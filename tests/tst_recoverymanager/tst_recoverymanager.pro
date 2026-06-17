QT += core testlib concurrent
QT -= gui
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_recoverymanager

INCLUDEPATH += ../../src ../../src/utils ../../src/pipeline

SOURCES += tst_recoverymanager.cpp ../../src/utils/RecoveryManager.cpp ../../src/utils/MipFallback.cpp ../../src/pipeline/ReportDataJson.cpp
HEADERS += ../../src/utils/RecoveryManager.h ../../src/utils/MipFallback.h ../../src/pipeline/ReportDataJson.h ../../src/pipeline/ReportData.h
