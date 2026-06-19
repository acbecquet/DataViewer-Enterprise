// DatabaseOps.h -- thread-agnostic whole-file write, shared by the UI-thread
// DatabaseManager and the background PersistWorker (SP4.5 Stage 2a). None of
// these functions touch any DatabaseManager member; everything is passed in.
// That is what lets PersistWorker run the exact same OCC algorithm on its OWN
// QSqlDatabase connection on a worker thread (Qt asserts when a QSqlDatabase is
// used off its owning thread).
#pragma once

#include <QSqlDatabase>
#include <QString>
#include "DatabaseManager.h"          // DVE::WriteResult
#include "../pipeline/ReportData.h"   // DVE::FileResult
#include "IdentityManager.h"          // DVE::IdentityManager (writerUuid arg)

namespace DVE {

// Promoted from file-scope statics in DatabaseManager.cpp (now external so both
// DatabaseManager.cpp and DatabaseOps.cpp link one copy; existing unqualified
// call sites in DatabaseManager.cpp resolve here via this header).
QString     writerUuid(IdentityManager* id);
WriteResult classifyMissingUpdate(QSqlDatabase& db, const QString& table,
                                  qint64 id, QString* outDetail);
int         freshChildVersion(QSqlDatabase& db, const QString& table,
                              qint64 id, int inMemoryFallback);
void        resetFileIdsForReinsert(FileResult& result, bool resetFileRow);
bool        fileRowExistsOnDb(QSqlDatabase& db, qint64 id);

// Whole-file write core (was DatabaseManager::tryWriteFileCore). Uses ONLY the
// supplied connection + writer UUID. On Success the file-row id/version are
// written back into result (child ids are intentionally not cascaded, same as
// the original). Errors go to *outError (never thrown).
WriteResult persistFileCore(QSqlDatabase& db, const QString& who,
                            FileResult& result, QString* outError);

// RowDeleted/VersionMismatch retry wrapper (was DatabaseManager::tryWriteFile
// (FileResult&)). online/dbOpen are snapshots the caller takes (the worker
// passes its UI-thread-captured online flag + its own connection state).
WriteResult persistFileOnConnection(QSqlDatabase& db, const QString& who,
                                    bool online, bool dbOpen,
                                    FileResult& result, QString* outError);

} // namespace DVE
