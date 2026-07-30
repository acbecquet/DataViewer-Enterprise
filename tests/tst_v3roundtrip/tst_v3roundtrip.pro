QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
DEFINES += SRCDIR=\\\"$$shell_quote($$PWD)\\\"

INCLUDEPATH += ../../src ../../src/pipeline ../common

include(../../external/QXlsx/QXlsx/QXlsx.pri)

# Task 1's CellAddress unit tests plus the Task 4 round-trip harness, which
# parses every fixture/corpus workbook through the PRODUCTION path
# (DataProcessor::processFile - openpyxl subprocess) and identity-checks the
# derived write addresses against the original grids - hence the full
# pipeline + ExcelReader + CorpusUtils link (mirrors tst_v3shadow.pro).
SOURCES += tst_v3roundtrip.cpp \
           ../../src/pipeline/CellAddress.cpp \
           ../../src/pipeline/DataProcessor.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/ExcelReader.cpp \
           ../../src/model/TemplateSchema.cpp \
           ../../src/model/StandardSchema.cpp \
           ../../src/model/SchemaDrivenReader.cpp \
           ../../src/model/SchemaInference.cpp \
           ../../src/model/Manifest.cpp \
           ../../src/model/SchemaResolver.cpp \
           ../../src/model/LegacyAdapter.cpp \
           ../../src/model/MetricRegistry.cpp \
           ../common/CorpusUtils.cpp

HEADERS += ../../src/pipeline/CellAddress.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/DataProcessor.h \
           ../../src/pipeline/SheetProcessors.h \
           ../../src/pipeline/TpmCalculator.h \
           ../../src/ExcelReader.h \
           ../../src/model/MetricDef.h \
           ../../src/model/MetricSample.h \
           ../../src/model/TemplateSchema.h \
           ../../src/model/StandardSchema.h \
           ../../src/model/SchemaDrivenReader.h \
           ../../src/model/SchemaInference.h \
           ../../src/model/Manifest.h \
           ../../src/model/SchemaResolver.h \
           ../../src/model/LegacyAdapter.h \
           ../../src/model/MetricRegistry.h \
           ../common/CorpusUtils.h
