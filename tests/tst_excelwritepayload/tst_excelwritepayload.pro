QT       += core testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src

SOURCES += tst_excelwritepayload.cpp            ../../src/utils/ExcelWritePayload.cpp

HEADERS += ../../src/utils/ExcelWritePayload.h
