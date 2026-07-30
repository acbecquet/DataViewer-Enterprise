QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
INCLUDEPATH += ../../src
SOURCES += tst_v3model.cpp \
           ../../src/model/TemplateSchema.cpp \
           ../../src/model/StandardSchema.cpp \
           ../../src/model/SchemaDrivenReader.cpp \
           ../../src/model/LegacyAdapter.cpp \
           ../../src/model/MetricRegistry.cpp \
           ../../src/model/RegimeParser.cpp \
           ../../src/model/Manifest.cpp \
           ../../src/model/SchemaResolver.cpp \
           ../../src/model/SchemaInference.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/ExcelReader.cpp
HEADERS += ../../src/model/MetricDef.h \
           ../../src/model/MetricSample.h \
           ../../src/model/TemplateSchema.h \
           ../../src/model/StandardSchema.h \
           ../../src/model/SchemaDrivenReader.h \
           ../../src/model/LegacyAdapter.h \
           ../../src/model/MetricRegistry.h \
           ../../src/model/RegimeParser.h \
           ../../src/model/Manifest.h \
           ../../src/model/SchemaResolver.h \
           ../../src/model/SchemaInference.h \
           ../../src/pipeline/SheetProcessors.h \
           ../../src/pipeline/TpmCalculator.h \
           ../../src/pipeline/RegimeUtils.h \
           ../../src/pipeline/ReportData.h \
           ../../src/ExcelReader.h \
           ../../src/pipeline/HeatingTech.h
