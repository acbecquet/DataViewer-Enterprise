QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
INCLUDEPATH += ../../src
SOURCES += tst_v3model.cpp \
           ../../src/model/TemplateSchema.cpp \
           ../../src/model/StandardSchema.cpp \
           ../../src/model/SchemaDrivenReader.cpp
HEADERS += ../../src/model/MetricDef.h \
           ../../src/model/MetricSample.h \
           ../../src/model/TemplateSchema.h \
           ../../src/model/StandardSchema.h \
           ../../src/model/SchemaDrivenReader.h
