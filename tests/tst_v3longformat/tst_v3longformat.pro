QT       += core sql testlib
CONFIG   += console c++17
CONFIG   -= app_bundle
TEMPLATE  = app

INCLUDEPATH += ../../src ../../src/database ../../src/pipeline ../../src/model \
               ../../src/utils ../common

# This suite is now TWO harnesses sharing one Postgres fixture.
#
# The 3b half links NO application source: what is under test there is
# server-side SQL (dve_migrate_to_long_format() plus the data_rows_v /
# samples_v compat views). Every value comparison is executed INSIDE Postgres
# and only counts and booleans cross the wire, so it cannot be fooled by Qt's
# double<->text formatting on the way in or out. Nothing below changes that -
# those slots still touch no C++ but MetricDefCache.
#
# The 3c half (plan Task 5) is the phase's END-TO-END GATE, and it is what
# pulls the rest of this list in. Its claim is that the manifest demo
# workbook's custom `coil_temp` column survives a real save and a real reload,
# so it has to drive the REAL chain and cannot mock any link of it:
#
#   ExcelReader (openpyxl subprocess) -> DataProcessor -> SchemaResolver ->
#   SchemaDrivenReader -> LegacyAdapter        ... the pipeline half
#   DatabaseManager / DatabaseOps / MetricDefCache / OfflineSnapshot
#                                              ... the persistence half
#
# The two blocks below are the transitive closure of exactly those two entry
# points - the union of what tst_v3inference.pro and tst_saveintegrity_e2e.pro
# already link. A missing member of this closure does NOT show up as a compile
# error; it shows up as an undefined reference at the final ld step, which is
# why the list is spelled out rather than trimmed by inspection.
#
# QXlsx is deliberately absent: ExcelReader shells out to python + openpyxl and
# nothing in src/pipeline or src/model includes a QXlsx header, so the sibling
# suites' QXlsx.pri include is inherited history rather than a dependency here.

SOURCES += tst_v3longformat.cpp \
           ../../src/database/MetricDefCache.cpp \
           ../../src/database/DatabaseManager.cpp \
           ../../src/database/DatabaseOps.cpp \
           ../../src/database/OfflineSnapshot.cpp \
           ../../src/database/PostgresConnection.cpp \
           ../../src/database/NotificationListener.cpp \
           ../../src/database/PresenceManager.cpp \
           ../../src/database/IdentityManager.cpp \
           ../../src/database/ConfigLoader.cpp \
           ../../src/database/LiveSync.cpp \
           ../../src/database/LiveSyncWorker.cpp \
           ../../src/database/RawGridJson.cpp \
           ../../src/utils/MipFallback.cpp \
           ../../src/utils/OutputPaths.cpp \
           ../../src/pipeline/SensoryData.cpp \
           ../../src/pipeline/DetailedSensoryData.cpp \
           ../../src/pipeline/DataProcessor.cpp \
           ../../src/pipeline/SheetProcessors.cpp \
           ../../src/pipeline/TpmCalculator.cpp \
           ../../src/pipeline/RegimeUtils.cpp \
           ../../src/ExcelReader.cpp \
           ../../src/model/TemplateSchema.cpp \
           ../../src/model/StandardSchema.cpp \
           ../../src/model/SchemaDrivenReader.cpp \
           ../../src/model/SchemaInference.cpp \
           ../../src/model/Manifest.cpp \
           ../../src/model/SchemaResolver.cpp \
           ../../src/model/LegacyAdapter.cpp \
           ../../src/model/MetricRegistry.cpp

HEADERS += ../../src/database/MetricDefCache.h \
           ../../src/database/DatabaseManager.h \
           ../../src/database/DatabaseOps.h \
           ../../src/database/OfflineSnapshot.h \
           ../../src/database/PostgresConnection.h \
           ../../src/database/NotificationListener.h \
           ../../src/database/PresenceManager.h \
           ../../src/database/IdentityManager.h \
           ../../src/database/ConfigLoader.h \
           ../../src/database/LiveSync.h \
           ../../src/database/LiveSyncWorker.h \
           ../../src/database/RawGridJson.h \
           ../../src/utils/MipFallback.h \
           ../../src/utils/OutputPaths.h \
           ../../src/pipeline/ReportData.h \
           ../../src/pipeline/SensoryData.h \
           ../../src/pipeline/DetailedSensoryData.h \
           ../../src/pipeline/DataProcessor.h \
           ../../src/pipeline/SheetProcessors.h \
           ../../src/pipeline/TpmCalculator.h \
           ../../src/ExcelReader.h \
           ../../src/model/MetricDef.h \
           ../../src/model/MetricSample.h \
           ../../src/model/TemplateSchema.h \
           ../../src/model/StandardSchema.h \
           ../../src/model/SchemaDrivenReader.h \
           ../../src/model/SchemaInference.h \
           ../../src/model/Manifest.h \
           ../../src/model/SchemaResolver.h \
           ../../src/model/LegacyAdapter.h \
           ../../src/model/MetricRegistry.h \
           ../common/TestHelpers.h
