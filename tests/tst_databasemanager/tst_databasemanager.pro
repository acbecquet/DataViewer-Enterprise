QT += core sql testlib network
CONFIG += console c++17
CONFIG -= app_bundle
TEMPLATE = app

# DbRepair links DataProcessor + ExcelReader (+ QXlsx for ExcelReader) and
# SheetProcessors / TpmCalculator / RegimeUtils.  Include the QXlsx .pri so
# the linker can resolve the xlsx symbols pulled in by ExcelReader.cpp.
include(../../external/QXlsx/QXlsx/QXlsx.pri)

INCLUDEPATH += ../../src ../../src/pipeline ../../src/utils ../../src/plotting \
               ../../src/reporting ../../src/database ../common

SOURCES += tst_databasemanager.cpp \
           ../../src/database/DatabaseManager.cpp \
           ../../src/database/DatabaseOps.cpp \
           ../../src/database/MetricDefCache.cpp \
           ../../src/database/PersistWorker.cpp \
           ../../src/utils/OutputPaths.cpp \
           ../../src/database/DbRepair.cpp \
           ../../src/database/RawGridJson.cpp \
           ../../src/database/OfflineSnapshot.cpp \
           ../../src/utils/MipFallback.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/pipeline/SensoryData.cpp \
           ../../src/pipeline/DetailedSensoryData.cpp \
           ../../src/pipeline/DataProcessor.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/model/TemplateSchema.cpp \
           ../../src/model/StandardSchema.cpp \
           ../../src/model/SchemaDrivenReader.cpp \
           ../../src/model/LegacyAdapter.cpp \
           ../../src/model/SchemaInference.cpp \
           ../../src/model/Manifest.cpp \
           ../../src/model/SchemaResolver.cpp \
           ../../src/model/MetricRegistry.cpp \
           ../../src/ExcelReader.cpp

HEADERS += ../../src/database/DatabaseManager.h \
           ../../src/database/DatabaseOps.h \
           ../../src/database/MetricDefCache.h \
           ../../src/database/PersistWorker.h \
           ../../src/utils/OutputPaths.h \
           ../../src/database/DbRepair.h \
           ../../src/database/OfflineSnapshot.h \
           ../../src/utils/MipFallback.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/ConfigLoader.h \
           ../../src/database/IdentityManager.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/SensoryData.h \
           ../../src/pipeline/DetailedSensoryData.h \
           ../../src/pipeline/DataProcessor.h \
           ../../src/ExcelReader.h \
           ../common/TestHelpers.h
