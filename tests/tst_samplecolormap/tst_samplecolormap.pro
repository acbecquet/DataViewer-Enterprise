QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
INCLUDEPATH += ../../src ../../src/utils ../common
SOURCES += tst_samplecolormap.cpp
SOURCES += ../../src/utils/SampleColorMap.cpp
SOURCES += ../../src/utils/AppTheme.cpp
