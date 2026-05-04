QT += core gui widgets testlib
CONFIG += console c++17
CONFIG -= app_bundle

TEMPLATE = app

# Test data directory

# Include paths
INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

SOURCES += tst_imageutils.cpp
SOURCES += ../../src/utils/ImageUtils.cpp
