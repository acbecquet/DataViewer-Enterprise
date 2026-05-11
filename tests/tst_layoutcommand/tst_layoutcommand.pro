QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/reporting

SOURCES += tst_layoutcommand.cpp            ../../src/reporting/ReportLayout.cpp            ../../src/reporting/LayoutCommand.cpp

HEADERS += ../../src/reporting/ReportLayout.h            ../../src/reporting/LayoutCommand.h
