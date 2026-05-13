QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
TARGET = tst_offlinebanner

INCLUDEPATH += ../../src ../../src/widgets

HEADERS += ../../src/widgets/OfflineBanner.h

SOURCES += tst_offlinebanner.cpp            ../../src/widgets/OfflineBanner.cpp
