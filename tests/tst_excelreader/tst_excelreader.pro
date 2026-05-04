QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

include(../../external/QXlsx/QXlsx/QXlsx.pri)

SOURCES += tst_excelreader.cpp \
           ../../src/ExcelReader.cpp

HEADERS += ../../src/ExcelReader.h \
           ../common/TestHelpers.h
