QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle

TEMPLATE = app
TARGET = tst_defaultmode

INCLUDEPATH += ../../src ../common

SOURCES += tst_defaultmode.cpp \
           ../../src/ui/DefaultModeDialog.cpp

HEADERS += ../../src/ui/DefaultModeDialog.h
