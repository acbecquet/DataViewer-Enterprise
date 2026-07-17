QT += core gui widgets sql testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

include(../../external/QXlsx/QXlsx/QXlsx.pri)

SOURCES += tst_reportgenerator.cpp \
           ../../src/reporting/ReportGenerator.cpp \
           ../../src/reporting/PptxWriter.cpp \
           ../../src/pipeline/DataProcessor.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/DataCleanup.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/ExcelReader.cpp \
           ../../src/plotting/PlotEngine.cpp \
           ../../src/utils/ZipWriter.cpp \
           ../../src/utils/XmlBuilder.cpp \
           ../../src/utils/ImageUtils.cpp \
           ../../src/utils/SopLoader.cpp \
           ../../src/utils/AppTheme.cpp \
           ../../src/utils/SampleColorMap.cpp

HEADERS += ../../src/reporting/ReportGenerator.h \
           ../../src/reporting/PptxWriter.h \
           ../../src/pipeline/DataProcessor.h \
           ../../src/pipeline/SheetProcessors.h \
           ../../src/pipeline/DataCleanup.h \
           ../../src/pipeline/TpmCalculator.h \
           ../../src/pipeline/ReportData.h \
           ../../src/ExcelReader.h \
           ../../src/plotting/PlotEngine.h \
           ../../src/utils/ZipWriter.h \
           ../../src/utils/XmlBuilder.h \
           ../../src/utils/ImageUtils.h \
           ../../src/utils/SopLoader.h \
           ../../src/utils/AppTheme.h \
           ../common/TestHelpers.h

LIBS += -lz
