QT += testlib gui widgets
CONFIG += console c++17
TEMPLATE = app

INCLUDEPATH += ../../src

SOURCES += tst_cellfocusdelegate.cpp     ../../src/widgets/CellFocusDelegate.cpp

HEADERS += ../../src/widgets/CellFocusDelegate.h
