QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/utils ../common

SOURCES += tst_excelsurgery.cpp \
           ../../src/utils/ExcelWritePayload.cpp \
           ../../src/ExcelReader.cpp

HEADERS += ../../src/utils/ExcelWritePayload.h \
           ../../src/ExcelReader.h
