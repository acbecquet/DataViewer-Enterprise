QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

INCLUDEPATH += ../../src ../../src/ui

SOURCES += tst_slidecanvasitems.cpp            ../../src/ui/SlideCanvasItems.cpp

HEADERS += ../../src/ui/SlideCanvasItems.h
