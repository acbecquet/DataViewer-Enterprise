QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

SOURCES += tst_pptxwriter.cpp \
           ../../src/reporting/PptxWriter.cpp \
           ../../src/utils/ZipWriter.cpp \
           ../../src/utils/XmlBuilder.cpp

HEADERS += ../../src/reporting/PptxWriter.h \
           ../../src/utils/ZipWriter.h \
           ../../src/utils/XmlBuilder.h \
           ../../src/pipeline/ReportData.h \
           ../common/TestHelpers.h

LIBS += -lz
