QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle

TEMPLATE = app
TARGET = tst_ribbonlayout

INCLUDEPATH += ../../src ../../src/utils ../../src/widgets ../common

SOURCES += tst_ribbonlayout.cpp            ../../src/widgets/RibbonWidget.cpp            ../../src/widgets/ScrollHost.cpp            ../../src/utils/AppTheme.cpp

HEADERS += ../../src/widgets/RibbonWidget.h            ../../src/widgets/ScrollHost.h            ../../src/utils/AppTheme.h
