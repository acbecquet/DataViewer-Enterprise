QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
INCLUDEPATH += ../../src ../../src/pipeline
SOURCES += tst_v3roundtrip.cpp \
           ../../src/pipeline/CellAddress.cpp
HEADERS += ../../src/pipeline/CellAddress.h \
           ../../src/pipeline/ReportData.h
