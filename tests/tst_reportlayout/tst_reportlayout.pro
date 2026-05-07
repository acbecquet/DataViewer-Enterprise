QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/reporting

SOURCES += tst_reportlayout.cpp \
           ../../src/reporting/ReportLayout.cpp

HEADERS += ../../src/reporting/ReportLayout.h
