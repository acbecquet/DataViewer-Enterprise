QT += core testlib
CONFIG += console c++17
CONFIG -= app_bundle

TEMPLATE = app

# Test data directory

# Include paths
INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting ../../src/reporting ../../src/database ../common

SOURCES += tst_tpmcalculator.cpp
SOURCES += ../../src/pipeline/TpmCalculator.cpp
