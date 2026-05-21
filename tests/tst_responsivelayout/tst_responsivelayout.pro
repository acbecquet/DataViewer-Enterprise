QT += testlib widgets
CONFIG += console
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_responsivelayout

INCLUDEPATH += ../../src

SOURCES += tst_responsivelayout.cpp     ../../src/utils/ResponsiveLayout.cpp

HEADERS += ../../src/utils/ResponsiveLayout.h
