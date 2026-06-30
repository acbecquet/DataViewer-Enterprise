QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_scrollhost

INCLUDEPATH += ../../src ../../src/widgets

HEADERS += ../../src/widgets/ScrollHost.h

SOURCES += tst_scrollhost.cpp            ../../src/widgets/ScrollHost.cpp
