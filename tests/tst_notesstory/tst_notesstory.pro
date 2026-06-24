QT += testlib
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app
SOURCES += tst_notesstory.cpp ../../src/pipeline/NotesStory.cpp
INCLUDEPATH += ../../src
