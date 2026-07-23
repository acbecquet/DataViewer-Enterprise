QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
DEFINES += SRCDIR=\\\"$$shell_quote($$PWD)\\\"
INCLUDEPATH += ../../src ../../src/pipeline ../common
SOURCES += tst_v3harness.cpp ../common/CorpusUtils.cpp ../common/JsonDiff.cpp
HEADERS += ../common/CorpusUtils.h ../common/JsonDiff.h
