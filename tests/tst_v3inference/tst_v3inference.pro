QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
DEFINES += SRCDIR=\\\"$$shell_quote($$PWD)\\\"

INCLUDEPATH += ../../src ../../src/pipeline ../common

include(../../external/QXlsx/QXlsx/QXlsx.pri)

# Pure-model unit tests (inferBlockCols / standardFits / inferSchema) plus the
# end-to-end value tests that run the real DataProcessor::processFile pipeline
# (openpyxl subprocess) over the pv13 / usersim8 fixtures and the optional real
# corpus - hence the full pipeline + ExcelReader + CorpusUtils link.
SOURCES += tst_v3inference.cpp \
           ../../src/pipeline/DataProcessor.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/pipeline/CellAddress.cpp \
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

HEADERS += ../../src/pipeline/DataProcessor.h \
           ../../src/pipeline/SheetProcessors.h \
           ../../src/pipeline/TpmCalculator.h \
           ../../src/pipeline/CellAddress.h \
           ../../src/pipeline/ReportData.h \
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
           ../common/CorpusUtils.h \
           ../common/TestHelpers.h
