QT += core gui widgets testlib
CONFIG += c++17 console
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_sizingsweep

INCLUDEPATH += ../../src ../../src/ui ../../src/utils ../../src/widgets

SOURCES += tst_sizingsweep.cpp            ../../src/ui/DataCleanupDialog.cpp            ../../src/utils/AppTheme.cpp            ../../src/widgets/ScrollHost.cpp

HEADERS += ../../src/ui/DataCleanupDialog.h            ../../src/utils/AppTheme.h            ../../src/widgets/ScrollHost.h            ../../src/pipeline/ReportData.h
