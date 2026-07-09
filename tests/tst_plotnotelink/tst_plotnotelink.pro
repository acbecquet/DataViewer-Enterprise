QT += testlib core gui
CONFIG += c++17 console testcase
CONFIG -= app_bundle
TARGET = tst_plotnotelink
INCLUDEPATH += ../../src
SOURCES += tst_plotnotelink.cpp
HEADERS += ../../src/plotting/PlotHitTest.h \
           ../../src/plotting/PlotEngine.h
