QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

include(../../external/QXlsx/QXlsx/QXlsx.pri)

TARGET = tst_regimeutils

SOURCES += tst_regimeutils.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/ExcelReader.cpp

HEADERS += ../../src/pipeline/RegimeUtils.h
