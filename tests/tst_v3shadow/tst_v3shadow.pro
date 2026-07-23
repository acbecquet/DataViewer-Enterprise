QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
DEFINES += SRCDIR=\\\"$$shell_quote($$PWD)\\\"

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

include(../../external/QXlsx/QXlsx/QXlsx.pri)

SOURCES += tst_v3shadow.cpp \
           ../../src/pipeline/DataProcessor.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/ExcelReader.cpp \
           ../../src/pipeline/ReportDataJson.cpp \
           ../../src/model/StandardSchema.cpp \
           ../../src/model/SchemaDrivenReader.cpp \
           ../../src/model/LegacyAdapter.cpp \
           ../common/CorpusUtils.cpp \
           ../common/JsonDiff.cpp

HEADERS += ../../src/pipeline/DataProcessor.h \
           ../../src/pipeline/SheetProcessors.h \
           ../../src/pipeline/TpmCalculator.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/ReportDataJson.h \
           ../../src/ExcelReader.h \
           ../../src/model/MetricDef.h \
           ../../src/model/MetricSample.h \
           ../../src/model/TemplateSchema.h \
           ../../src/model/StandardSchema.h \
           ../../src/model/SchemaDrivenReader.h \
           ../../src/model/LegacyAdapter.h \
           ../common/CorpusUtils.h \
           ../common/JsonDiff.h
