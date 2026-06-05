QT += core sql testlib network
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils \
               ../../src/reporting ../../src/plotting ../../src/database ../common

TARGET = tst_writeoutcome

SOURCES += tst_writeoutcome.cpp \
           ../../src/database/WriteOutcome.cpp

HEADERS += ../../src/database/WriteOutcome.h
