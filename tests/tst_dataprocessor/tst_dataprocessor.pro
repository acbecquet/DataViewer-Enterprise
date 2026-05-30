QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

include(../../external/QXlsx/QXlsx/QXlsx.pri)

SOURCES += tst_dataprocessor.cpp \
           ../../src/pipeline/DataProcessor.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/ExcelReader.cpp

HEADERS += ../../src/pipeline/DataProcessor.h \
           ../../src/pipeline/SheetProcessors.h \
           ../../src/pipeline/TpmCalculator.h \
           ../../src/pipeline/ReportData.h \
           ../../src/ExcelReader.h \
           ../common/TestHelpers.h
