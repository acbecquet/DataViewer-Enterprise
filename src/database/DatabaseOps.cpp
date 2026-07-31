#include "DatabaseOps.h"
#include "IdentityManager.h"
#include "MetricDefCache.h"
#include "RawGridJson.h"

#include <QDebug>
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlRecord>
#include <QVariant>
#include <QUuid>
#include <QDateTime>
#include <QFile>
#include <QFileInfo>
#include <QRectF>
#include <QByteArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>

namespace DVE {

namespace {
// Postgres SQLSTATE for unique_violation (mirrors the anon-namespace constant in
// DatabaseManager.cpp; the core body maps it to WriteResult::UniqueViolation).
constexpr const char* kSqlStateUniqueViolation = "23505";
}

// File-local stand-in for DatabaseManager::logDebug so the verbatim body's
// logDebug(...) calls compile unchanged.
static void logDebug(const QString& msg) {
    qDebug().noquote() << "[DatabaseOps]" << msg;
}

// --- helper: who is making this change? -------------------------------------
// Postgres requires a non-null updated_by on every write. IdentityManager
// always supplies one once open() has been called with a real identity. If
// m_identity is somehow null (shouldn't happen post-3a) we fall back to a
// generic marker rather than crash.
QString writerUuid(IdentityManager* id) {
    if (id) return id->uuid().toString(QUuid::WithoutBraces);
    return QStringLiteral("unknown");
}

// --- helper: post-UPDATE rowcount-zero diagnostic ---------------------------
// Called when an optimistic UPDATE returned numRowsAffected() == 0. Issues a
// SELECT on the same id to distinguish "row exists with newer version"
// (VersionMismatch) from "row no longer exists" (RowDeleted). Any SQL error
// in the diagnostic itself collapses to OtherError so we never silently
// upgrade a conflict to success.
//
// v2.0.2 M9 — transactional race documentation. The diagnostic SELECT
// runs in the same transaction as the failed UPDATE (callers haven't
// rolled back yet) and therefore sees the same snapshot the UPDATE saw.
// A concurrent DELETE that committed AFTER our snapshot was taken but
// BEFORE we issue this SELECT is invisible to us — the row still appears
// to exist, so we report VersionMismatch when the truth is RowDeleted.
// Conversely, a concurrent INSERT of the same id (only possible during
// natural-key recreate flows) is also invisible, so a RowDeleted return
// might shadow a row that was actually re-created.
//
// This racy classification is benign: every caller (tryWriteFile,
// tryWriteSensorySession, tryWriteDetailedSensorySession) maps both
// VersionMismatch and RowDeleted to the same conflict-dialog surface, so
// the user is prompted either way and resolves with current truth on
// re-read. Tightening the classification would require running the SELECT
// in a new SERIALIZABLE transaction, which is more cost than the dialog
// disambiguation saves.
WriteResult classifyMissingUpdate(QSqlDatabase& db,
                                         const QString& table,
                                         qint64 id,
                                         QString* outDetail)
{
    QSqlQuery q(db);
    q.prepare(QString("SELECT 1 FROM %1 WHERE id = ?").arg(table));
    q.addBindValue(static_cast<qlonglong>(id));
    if (!q.exec()) {
        if (outDetail) *outDetail = q.lastError().text();
        return WriteResult::OtherError;
    }
    return q.next() ? WriteResult::VersionMismatch : WriteResult::RowDeleted;
}

// --- helper: fresh-version read for a child UPDATE --------------------------
// RC1 (v2.4.0 data-loss regression), CHILD-row twin of the file-row fix in
// commits 0c21100/377f827. The per-child UPDATEs in tryWriteFile bound the
// in-memory sheet/sample/data_row/image version in WHERE id=? AND version=?.
// Those versions routinely go stale (a LiveSync per-cell commit or this
// client's own prior save bumps the child row's version), so a routine
// whole-file save failed on the first stale child with VersionMismatch
// ("tryWriteFile(UPDATE test id=142): version mismatch" in the production
// log) — which callers then treated as "already synced" and silently dropped.
//
// DESIGN (v2.5.0 decision, see
// docs/superpowers/specs/2026-06-10-v24-save-sync-regressions-evidence.md):
// a whole-file save deliberately adopts each child row's CURRENT committed
// version, so the file's child rows resolve as ROW-LEVEL last-writer-wins by
// design — identical semantics to the file row itself. Cross-client cell-level
// protection lives in the LiveSync per-cell stream. The SELECT takes FOR
// UPDATE inside tryWriteFile's transaction, so a concurrent whole-file saver
// serializes behind it instead of interleaving with this read; the
// in-transaction SELECT->UPDATE race is therefore closed and the `AND version
// = ?` clause is a defensive invariant whose mismatch is unreachable. When the
// row is gone the SELECT returns no row and we keep the in-memory fallback, so
// the guarded UPDATE still classifies RowDeleted exactly as before.
int freshChildVersion(QSqlDatabase& db, const QString& table,
                             qint64 id, int inMemoryFallback)
{
    QSqlQuery sel(db);
    sel.prepare(QString("SELECT version FROM %1 WHERE id = ? FOR UPDATE").arg(table));
    sel.addBindValue(static_cast<qlonglong>(id));
    if (!sel.exec()) {
        // Audit fix: a SQL error here was previously indistinguishable from a
        // missing row (both silently fall back). The guarded UPDATE still
        // classifies the outcome correctly, but log so a real error is visible.
        qWarning() << "freshChildVersion: version read failed on" << table
                   << "id" << id << "--" << sel.lastError().text()
                   << "(using in-memory version)";
        return inMemoryFallback;
    }
    if (sel.next())
        return sel.value(0).toInt();
    return inMemoryFallback;
}

// --- helper: clear ids/versions in a FileResult for a fresh re-INSERT -------
// Used by tryWriteFile's RowDeleted recovery. Two cases, distinguished by
// `resetFileRow`:
//   * file row itself was deleted (resetFileRow=true): zero EVERYTHING so the
//     retry re-INSERTs the whole tree as a brand-new file.
//   * a CHILD row was deleted but the file row survives (resetFileRow=false):
//     keep the file id/version (the core's fresh-version OCC adopts the current
//     file version) and zero only the children, so the deleted child is
//     re-INSERTed without colliding on the files.file_path UNIQUE key. The
//     child fresh-version OCC (Part A) means child rows otherwise never raise a
//     conflict, so a child UPDATE that still misses can only be a true delete.
void resetFileIdsForReinsert(FileResult& result, bool resetFileRow)
{
    if (resetFileRow) {
        result.id = -1;
        result.version = 0;
    }
    for (SheetResult& sheet : result.sheets) {
        sheet.id = -1;
        sheet.version = 0;
        for (SampleResult& sr : sheet.samples) {
            sr.id = -1;
            sr.version = 0;
            for (DataRow& dr : sr.rows) {
                dr.id = -1;
                dr.version = 0;
            }
            for (qint64& imgId : sr.imageIds)   imgId = -1;
            for (int& imgVer : sr.imageVersions) imgVer = 0;
        }
    }
}

bool fileRowExistsOnDb(QSqlDatabase& db, qint64 id) {
    QSqlQuery q(db);
    q.prepare("SELECT 1 FROM files WHERE id = ? LIMIT 1");
    q.addBindValue(static_cast<qlonglong>(id));
    return q.exec() && q.next();
}

WriteResult persistFileCore(QSqlDatabase& db, const QString& who,
                            FileResult& result, QString* outError) {
    // Body lifted VERBATIM from DatabaseManager::tryWriteFileCore (v2.4.5).
    // The only edits vs the original: the offline/!isOpen guards + the two
    // local 'db'/'who' decls are dropped (now parameters), and the error
    // member m_lastError is redirected to the caller out-param via a reference.
    // DatabaseOps NEVER touches a DatabaseManager member -> safe off-thread.
    QString _discard;
    QString& lastError = outError ? *outError : _discard;
    lastError.clear();

    int totalSamples = 0;
    for (const auto& sheet : result.sheets)
        totalSamples += sheet.samples.size();

    // ── v3 Phase 3c: open metrics (DataRow::extra / SampleResult::extra) ─────
    // These are the metric maps Phase 2 already fills for inference and manifest
    // sheets. They have no wide column at all, which is exactly why they move
    // first (Phase 3 index, SEQUENCING CORRECTION): they cannot go stale against
    // a wide copy, so read and write can move together without touching the
    // standard 13-column path at all. Everything below that mentions
    // measurements / sample_headers is ADDITIVE - no existing statement changes.
    //
    // The probe runs BEFORE the transaction on purpose. A failed statement
    // poisons a Postgres transaction, so probing from inside one would turn
    // "this database has no long tables yet" into "this file cannot be saved",
    // and the standard path would go down with it. ensureSchema creates all
    // three tables on every connect (index D8), so !longFormat means something
    // is genuinely wrong; log it and save the wide data anyway.
    MetricDefCache metricDefs(db, who);
    QString metricDefsError;
    const bool longFormat = metricDefs.load(&metricDefsError);
    if (!longFormat) {
        logDebug(QStringLiteral("persistFileCore: long-format tables unreachable, "
                                "open metrics NOT persisted this save -- %1")
                     .arg(metricDefsError));
    }

    if (!db.transaction()) {
        lastError = QStringLiteral("tryWriteFile(begin transaction): ")
                      + db.lastError().text();
        logDebug(lastError);
        return WriteResult::OtherError;
    }

    // -- Upsert the file row by file_path with optimistic concurrency. There
    //    are two branches:
    //      (a) result.id != -1 && result.version > 0 — caller has a server-
    //          loaded struct and wants to UPDATE that exact row. We require
    //          WHERE id = ? AND version = ? to refuse stale writes.
    //      (b) result.id == -1 — fresh struct, INSERT. UniqueViolation on
    //          file_path collision; caller is expected to recover (e.g.,
    //          load-then-merge) and re-issue.
    qint64 fileId = -1;
    int    newVer = 0;  // server-assigned version, captured via RETURNING
    if (result.id != -1 && result.version > 0) {
        // Whole-file save: adopt the row's CURRENT committed version, read
        // inside this transaction, rather than the in-memory result.version.
        // files.version routinely outruns the struct (LiveSync per-cell
        // commits; this client's own prior saves), so binding result.version
        // made routine whole-file saves fail with VersionMismatch — which
        // callers treated as "already synced" and silently dropped.
        //
        // DESIGN (v2.5.0 decision, see
        // docs/superpowers/specs/2026-06-10-v24-save-sync-regressions-evidence.md):
        // a whole-file save deliberately adopts the current version, so two
        // clients each saving the whole file resolve as ROW-LEVEL
        // last-writer-wins. That is intentional — cross-client cell-level
        // protection lives elsewhere: the LiveSync per-cell stream (and, for
        // sensory/detailed, the DB-score-preserving merge; dirty-aware in
        // plan Task 3). The SELECT below takes FOR UPDATE, so any concurrent
        // whole-file saver serializes behind this transaction instead of
        // interleaving with the read; the in-transaction SELECT→UPDATE race is
        // therefore closed. The `AND version = ?` clause is now a defensive
        // invariant: with the lock held and the fresh version bound, a version
        // mismatch is unreachable. result.version stays as the fallback when
        // the row is missing, so the guarded UPDATE still classifies
        // RowDeleted exactly as before.
        int expectedVersion = result.version;
        {
            QSqlQuery sel(db);
            sel.prepare("SELECT version FROM files WHERE id = ? FOR UPDATE");
            sel.addBindValue(static_cast<qlonglong>(result.id));
            if (sel.exec() && sel.next())
                expectedVersion = sel.value(0).toInt();
        }

        QSqlQuery q(db);
        q.prepare(
            "UPDATE files SET "
            "  file_path = ?, "
            "  file_name = ?, "
            "  loaded_at = ?, "
            "  template_version = ?, "
            "  sheet_count = ?, "
            "  sample_count = ?, "
            "  updated_by = ? "
            "WHERE id = ? AND version = ? "
            "RETURNING version");
        q.addBindValue(result.filePath);
        q.addBindValue(result.fileName);
        q.addBindValue(QDateTime::currentDateTime().toString(Qt::ISODate));
        q.addBindValue(result.templateVersion);
        q.addBindValue(result.sheets.size());
        q.addBindValue(totalSamples);
        q.addBindValue(who);
        q.addBindValue(static_cast<qlonglong>(result.id));
        q.addBindValue(expectedVersion);
        if (!q.exec()) {
            const QString code = q.lastError().nativeErrorCode();
            lastError = QStringLiteral("tryWriteFile(UPDATE files): ")
                          + q.lastError().text();
            db.rollback();
            logDebug(lastError);
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        // RETURNING means a Success row is available via q.next(); absence
        // signals the optimistic-concurrency miss (the WHERE didn't match).
        if (!q.next()) {
            QString detail;
            const WriteResult cls = classifyMissingUpdate(
                db, QStringLiteral("files"), result.id, &detail);
            db.rollback();
            if (cls == WriteResult::VersionMismatch) {
                lastError = QStringLiteral(
                    "tryWriteFile(UPDATE files): version mismatch (id=%1, "
                    "expected version=%2)").arg(result.id).arg(expectedVersion);
            } else if (cls == WriteResult::RowDeleted) {
                lastError = QStringLiteral(
                    "tryWriteFile(UPDATE files): row deleted (id=%1)").arg(result.id);
            } else {
                lastError = QStringLiteral(
                    "tryWriteFile(UPDATE files): classify failed: ") + detail;
            }
            logDebug(lastError);
            return cls;
        }
        fileId = result.id;
        newVer = q.value(0).toInt();
    } else {
        QSqlQuery q(db);
        q.prepare(
            "INSERT INTO files (file_path, file_name, loaded_at, template_version, "
            "sheet_count, sample_count, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?) RETURNING id, version");
        q.addBindValue(result.filePath);
        q.addBindValue(result.fileName);
        q.addBindValue(QDateTime::currentDateTime().toString(Qt::ISODate));
        q.addBindValue(result.templateVersion);
        q.addBindValue(result.sheets.size());
        q.addBindValue(totalSamples);
        q.addBindValue(who);
        if (!q.exec() || !q.next()) {
            const QString code = q.lastError().nativeErrorCode();
            lastError = QStringLiteral("tryWriteFile(INSERT files): ")
                          + q.lastError().text();
            db.rollback();
            logDebug(lastError);
            if (code == QString::fromLatin1(kSqlStateUniqueViolation))
                return WriteResult::UniqueViolation;
            return WriteResult::OtherError;
        }
        fileId = q.value(0).toLongLong();
        newVer = q.value(1).toInt();
    }

    // -- C3 id-aware upsert ------------------------------------------------
    // Replaces the legacy DELETE-cascade-rebuild (which destroyed concurrent
    // users' work) with a three-phase algorithm:
    //   (A) capture the pre-image set of child ids under this file
    //   (B) for each in-memory child: UPDATE existing rows by id+version
    //       (aborting whole save on VersionMismatch via classifyMissingUpdate)
    //       or INSERT new rows (back-filling the struct's id/version)
    //   (C) DELETE rows present pre-save but absent post-save (user deletes)
    //
    // Concurrent users' rows that aren't in A's in-memory FileResult are
    // still wiped in phase C — A reloading before saving would pull them in.
    // C5 (drainPendingEdits replayed_at sentinel) protects offline edits;
    // OCC protects same-row clobbers; this loop protects against the
    // catastrophic full-subtree wipe that v2.0.1 had.

    // ── v3 Phase 3c: the wide-column catalog, read ONCE (hazard H1) ─────────
    // The single source of truth for BOTH halves of the H1 guard: the prune
    // pre-image in phase A, and the extras writer in phase B. A key that has a
    // same-named column in the wide table IS a standard metric - the writer must
    // not put it in the long tables and the prune must not be able to delete it.
    // The two disagreeing is not a wash: the writer's half wrote a row that the
    // prune's half could never remove, the read has no key filter (H25) so it
    // came back into `extra` on every load, and deleting that column from the
    // workbook could never delete its data.
    //
    // information_schema.columns is the obvious source and the wrong one twice
    // over: it is PRIVILEGE-FILTERED (a role with no privilege on data_rows sees
    // no rows for it at all) and its table_schema has to be hardcoded, so it can
    // silently disagree with the search_path the surrounding statements resolve
    // through. Either failure yields an EMPTY column set - and `key NOT IN
    // (<empty>)` is true for every key, so the guard evaporates exactly when it
    // is needed and every measurement becomes prunable. That is the H1 failure
    // itself. Every other catalog probe in this codebase is search-path aware
    // (DatabaseManager::ensureSchema uses regclass / pg_table_is_visible) and so
    // is this one.
    //
    // to_regclass() rather than CAST(? AS regclass): the cast RAISES for an
    // absent relation, and we are inside the save transaction, where one failed
    // statement poisons the whole thing. A NULL attrelid simply matches nothing,
    // which the emptiness check below then handles deliberately.
    QSet<QString> dataRowWideCols, sampleWideCols;
    if (longFormat) {
        auto readWideCols = [&](const QString& table, QSet<QString>* out) -> bool {
            QSqlQuery c(db);
            if (!c.prepare(QStringLiteral(
                    "SELECT a.attname FROM pg_attribute a "
                    "WHERE a.attrelid = to_regclass(?) AND a.attnum > 0 "
                    "  AND NOT a.attisdropped"))) {
                lastError = QStringLiteral("tryWriteFile(wide-column catalog %1): ")
                                .arg(table) + c.lastError().text();
                return false;
            }
            c.addBindValue(table);
            if (!c.exec()) {
                lastError = QStringLiteral("tryWriteFile(wide-column catalog %1): ")
                                .arg(table) + c.lastError().text();
                return false;
            }
            while (c.next()) out->insert(c.value(0).toString());
            return true;
        };
        if (!readWideCols(QStringLiteral("data_rows"), &dataRowWideCols) ||
            !readWideCols(QStringLiteral("samples"),   &sampleWideCols)) {
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
    }

    // A wide table with zero visible columns means the guard CANNOT be
    // evaluated. Both halves then decline: nothing is written (phase B) and
    // nothing is prunable (phase A leaves the pre-image empty), so a standard
    // metric can never be deleted by a guard that is not actually running.
    // Stale open-metric rows may be left behind; that is recoverable, deleting
    // real measurements is not.
    const bool guardRows    = longFormat && !dataRowWideCols.isEmpty();
    const bool guardHeaders = longFormat && !sampleWideCols.isEmpty();
    if (longFormat && (!guardRows || !guardHeaders)) {
        logDebug(QStringLiteral(
            "persistFileCore: wide-column catalog empty for %1 -- the H1 guard "
            "cannot be evaluated, so open metrics are neither written nor pruned "
            "this save")
                .arg(!guardRows && !guardHeaders
                         ? QStringLiteral("data_rows and samples")
                         : (!guardRows ? QStringLiteral("data_rows")
                                       : QStringLiteral("samples"))));
    }

    // Phase A: pre-image. Four scoped SELECTs ride the same QSqlQuery (six with
    // the v3 Phase 3c long tables).
    QSet<qint64> preTestIds, preSampleIds, preDataRowIds, preImageIds;
    QSet<qint64> preMeasurementIds, preSampleHeaderIds;
    {
        QSqlQuery q(db);
        q.prepare("SELECT id FROM tests WHERE file_id = ?");
        q.addBindValue(fileId);
        if (!q.exec()) {
            lastError = QStringLiteral("tryWriteFile(preImage tests): ")
                          + q.lastError().text();
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
        while (q.next()) preTestIds.insert(q.value(0).toLongLong());

        q.prepare("SELECT s.id FROM samples s "
                  "JOIN tests t ON s.test_id = t.id WHERE t.file_id = ?");
        q.addBindValue(fileId);
        if (!q.exec()) {
            lastError = QStringLiteral("tryWriteFile(preImage samples): ")
                          + q.lastError().text();
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
        while (q.next()) preSampleIds.insert(q.value(0).toLongLong());

        q.prepare("SELECT dr.id FROM data_rows dr "
                  "JOIN samples s ON dr.sample_id = s.id "
                  "JOIN tests t ON s.test_id = t.id WHERE t.file_id = ?");
        q.addBindValue(fileId);
        if (!q.exec()) {
            lastError = QStringLiteral("tryWriteFile(preImage data_rows): ")
                          + q.lastError().text();
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
        while (q.next()) preDataRowIds.insert(q.value(0).toLongLong());

        q.prepare("SELECT im.id FROM images im "
                  "JOIN samples s ON im.sample_id = s.id "
                  "JOIN tests t ON s.test_id = t.id WHERE t.file_id = ?");
        q.addBindValue(fileId);
        if (!q.exec()) {
            lastError = QStringLiteral("tryWriteFile(preImage images): ")
                          + q.lastError().text();
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
        while (q.next()) preImageIds.insert(q.value(0).toLongLong());

        // ── v3 Phase 3c: the long tables' pre-image ──────────────────────
        //
        // ┌── READ THIS BEFORE TOUCHING EITHER QUERY (hazard H1) ─────────────┐
        // │ These two SELECTs decide what phase C is allowed to DELETE. The   │
        // │ prune's contract is "anything the in-memory model did not         │
        // │ reproduce is gone", and in 3c the in-memory model reproduces      │
        // │ EXCLUSIVELY the open metrics - DataRow::extra and                 │
        // │ SampleResult::extra. The 13 data_rows and 22 samples value        │
        // │ columns still live in the WIDE tables and are written by nobody   │
        // │ here, so if a standard metric ever appeared in `measurements`     │
        // │ this prune would delete it on the next save. That is H1, the top  │
        // │ data-loss risk of the whole phase.                                │
        // │                                                                   │
        // │ The `md.key NOT IN (<wide columns>)` clause is what makes that    │
        // │ structurally impossible rather than merely unlikely: a metric key │
        // │ that has a same-named column in the wide table is a STANDARD      │
        // │ metric, and this prune cannot see it, let alone delete it.        │
        // │ Driving it off the live catalog rather than a hand-written key    │
        // │ list means it cannot drift - and the 3b harness already complains │
        // │ that three hand copies of those lists exist. The SAME catalog     │
        // │ read (dataRowWideCols / sampleWideCols, above) also governs the   │
        // │ WRITER in phase B, so the two halves cannot disagree.             │
        // │                                                                   │
        // │ 3d - the standard-metric cutover - is where this changes. When    │
        // │ the migration has run and persistFileCore writes standard metrics │
        // │ as measurements, the guard must be lifted DELIBERATELY, together  │
        // │ with a write path that reproduces every one of them. Lifting it   │
        // │ before that is the H1 failure.                                    │
        // │                                                                   │
        // │ It fails safe in the direction that matters - a protected row is  │
        // │ never deleted, only ever left behind - PROVIDED the guard is      │
        // │ actually evaluated. An empty column set would protect nothing at  │
        // │ all, which is why guardRows / guardHeaders skip the pre-image     │
        // │ entirely rather than run an unguarded query.                      │
        // └───────────────────────────────────────────────────────────────────┘
        //
        // Scoped to this file's own samples, exactly like the four above.
        // `wideCols` is bound, never interpolated - the values are catalog
        // identifiers, but they still go through binds like everything else.
        auto preImageExcluding = [&](const QString& label, const QString& sql,
                                     const QSet<QString>& wideCols,
                                     QSet<qint64>* out) -> bool {
            QStringList marks;
            marks.reserve(wideCols.size());
            for (int i = 0; i < wideCols.size(); ++i) marks << QStringLiteral("?");
            if (!q.prepare(sql.arg(marks.join(QLatin1String(", "))))) {
                lastError = QStringLiteral("tryWriteFile(preImage %1): ").arg(label)
                              + q.lastError().text();
                return false;
            }
            q.addBindValue(fileId);
            for (const QString& c : wideCols) q.addBindValue(c);
            if (!q.exec()) {
                lastError = QStringLiteral("tryWriteFile(preImage %1): ").arg(label)
                              + q.lastError().text();
                return false;
            }
            while (q.next()) out->insert(q.value(0).toLongLong());
            return true;
        };

        if (guardRows &&
            !preImageExcluding(QStringLiteral("measurements"),
                               QStringLiteral(
                                   "SELECT m.id FROM measurements m "
                                   "JOIN samples s ON s.id = m.sample_id "
                                   "JOIN tests t ON s.test_id = t.id "
                                   "JOIN metric_defs md ON md.id = m.metric_id "
                                   "WHERE t.file_id = ? AND md.key NOT IN (%1)"),
                               dataRowWideCols, &preMeasurementIds)) {
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }

        if (guardHeaders &&
            !preImageExcluding(QStringLiteral("sample_headers"),
                               QStringLiteral(
                                   "SELECT sh.id FROM sample_headers sh "
                                   "JOIN samples s ON s.id = sh.sample_id "
                                   "JOIN tests t ON s.test_id = t.id "
                                   "JOIN metric_defs md ON md.id = sh.field_id "
                                   "WHERE t.file_id = ? AND md.key NOT IN (%1)"),
                               sampleWideCols, &preSampleHeaderIds)) {
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
    }

    // Prepare the eight UPDATE/INSERT statements once.
    QSqlQuery updateTest(db), insertTest(db);
    if (!updateTest.prepare(
            "UPDATE tests SET file_id = ?, sheet_name = ?, template_version = ?, "
            "overall_avg_tpm = ?, overall_stddev_tpm = ?, is_raw_table = ?, "
            "from_inferred_schema = ?, sort_order = ?, raw_grid = ?, updated_by = ? "
            "WHERE id = ? AND version = ? RETURNING version") ||
        !insertTest.prepare(
            "INSERT INTO tests (file_id, sheet_name, template_version, "
            "overall_avg_tpm, overall_stddev_tpm, is_raw_table, from_inferred_schema, "
            "sort_order, raw_grid, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id, version")) {
        lastError = QStringLiteral("tryWriteFile(prepare tests): ")
                      + (updateTest.lastError().isValid()
                            ? updateTest.lastError().text() : insertTest.lastError().text());
        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
    }
    QSqlQuery updateSample(db), insertSample(db);
    if (!updateSample.prepare(
            "UPDATE samples SET test_id = ?, sort_order = ?, sample_name = ?, sample_id = ?, "
            "date = ?, tester = ?, media = ?, viscosity = ?, resistance = ?, voltage = ?, "
            "power = ?, heating_technology = ?, puffing_regime = ?, initial_oil_mass = ?, "
            "average_tpm = ?, stddev_tpm = ?, avg_power_density = ?, efficiency_percent = ?, "
            "total_oil_consumed = ?, total_puffs = ?, normalized_tpm = ?, burn_status = ?, "
            "clog_status = ?, leak_status = ?, updated_by = ? "
            "WHERE id = ? AND version = ? RETURNING version") ||
        !insertSample.prepare(
            "INSERT INTO samples (test_id, sort_order, sample_name, sample_id, date, tester, "
            "media, viscosity, resistance, voltage, power, heating_technology, puffing_regime, "
            "initial_oil_mass, average_tpm, stddev_tpm, avg_power_density, efficiency_percent, "
            "total_oil_consumed, total_puffs, normalized_tpm, burn_status, clog_status, leak_status, "
            "updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
            "RETURNING id, version")) {
        lastError = QStringLiteral("tryWriteFile(prepare samples): ")
                      + (updateSample.lastError().isValid()
                            ? updateSample.lastError().text() : insertSample.lastError().text());
        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
    }
    QSqlQuery updateRow(db), insertRow(db);
    if (!updateRow.prepare(
            "UPDATE data_rows SET sample_id = ?, sort_order = ?, puffs = ?, "
            "before_weight = ?, after_weight = ?, draw_pressure = ?, resistance = ?, "
            "smell = ?, clog = ?, notes = ?, tpm = ?, tpm_power_density = ?, "
            "variation_tpm = ?, oil_consumed = ?, puffing_regime = ?, updated_by = ? "
            "WHERE id = ? AND version = ? RETURNING version") ||
        !insertRow.prepare(
            "INSERT INTO data_rows (sample_id, sort_order, puffs, before_weight, after_weight, "
            "draw_pressure, resistance, smell, clog, notes, tpm, tpm_power_density, "
            "variation_tpm, oil_consumed, puffing_regime, updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id, version")) {
        lastError = QStringLiteral("tryWriteFile(prepare data_rows): ")
                      + (updateRow.lastError().isValid()
                            ? updateRow.lastError().text() : insertRow.lastError().text());
        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
    }
    QSqlQuery updateImage(db), insertImage(db);
    if (!updateImage.prepare(
            "UPDATE images SET sample_id = ?, sort_order = ?, file_name = ?, image_data = ?, "
            "layout_x = ?, layout_y = ?, layout_w = ?, layout_h = ?, "
            "crop_x = ?, crop_y = ?, crop_w = ?, crop_h = ?, updated_by = ? "
            "WHERE id = ? AND version = ? RETURNING version") ||
        !insertImage.prepare(
            "INSERT INTO images (sample_id, sort_order, file_name, image_data, "
            "layout_x, layout_y, layout_w, layout_h, crop_x, crop_y, crop_w, crop_h, "
            "updated_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) RETURNING id, version")) {
        lastError = QStringLiteral("tryWriteFile(prepare images): ")
                      + (updateImage.lastError().isValid()
                            ? updateImage.lastError().text() : insertImage.lastError().text());
        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
    }

    // Phase B: upsert. Track post-image ids so phase C can identify deletions.
    QSet<qint64> postTestIds, postSampleIds, postDataRowIds, postImageIds;
    QSet<qint64> postMeasurementIds, postSampleHeaderIds;

    // ── v3 Phase 3c: batched open-metric writers ─────────────────────────────
    // Statement count here is already O(children) (hazard H8) and one extras row
    // per metric per data row multiplies it, so these go out as ONE multi-row
    // INSERT per sample rather than a round trip per value. A sample's keys are
    // unique by construction (QMap) and sort_order is the row index, so no two
    // tuples in a batch can hit the same ON CONFLICT target - which would
    // otherwise raise "cannot affect row a second time".
    //
    // The batch is capped so a pathological sample cannot exceed Postgres'
    // 65535-bind-parameter ceiling - 6 params per measurement row, 5 per
    // sample_headers row. BOTH batches honour the cap; a header batch is
    // normally tiny, but "normally" is not what a ceiling is for.
    struct ExtraRow { qint64 defId; int sortOrder; QVariant num; QVariant text; };
    constexpr int kExtrasBatchRows = 500;

    auto flushExtras = [&](const QString& table, qint64 sampleId,
                           QVector<ExtraRow>& batch, QSet<qint64>& postIds) -> bool {
        if (batch.isEmpty()) return true;
        const bool isMeasurement = (table == QLatin1String("measurements"));

        QStringList tuples;
        tuples.reserve(batch.size());
        for (int i = 0; i < batch.size(); ++i)
            tuples << (isMeasurement ? QStringLiteral("(?, ?, ?, ?, ?, ?)")
                                     : QStringLiteral("(?, ?, ?, ?, ?)"));

        const QString sql = isMeasurement
            ? QStringLiteral(
                  "INSERT INTO measurements (sample_id, metric_id, sort_order, "
                  "value_num, value_text, updated_by) VALUES %1 "
                  "ON CONFLICT (sample_id, metric_id, sort_order) DO UPDATE SET "
                  "  value_num = EXCLUDED.value_num, value_text = EXCLUDED.value_text, "
                  "  updated_by = EXCLUDED.updated_by "
                  "RETURNING id").arg(tuples.join(QLatin1String(", ")))
            : QStringLiteral(
                  "INSERT INTO sample_headers (sample_id, field_id, "
                  "value_num, value_text, updated_by) VALUES %1 "
                  "ON CONFLICT (sample_id, field_id) DO UPDATE SET "
                  "  value_num = EXCLUDED.value_num, value_text = EXCLUDED.value_text, "
                  "  updated_by = EXCLUDED.updated_by "
                  "RETURNING id").arg(tuples.join(QLatin1String(", ")));

        QSqlQuery q(db);
        if (!q.prepare(sql)) {
            lastError = QStringLiteral("tryWriteFile(prepare %1): ").arg(table)
                          + q.lastError().text();
            return false;
        }
        for (const ExtraRow& e : batch) {
            q.addBindValue(static_cast<qlonglong>(sampleId));
            q.addBindValue(static_cast<qlonglong>(e.defId));
            if (isMeasurement) q.addBindValue(e.sortOrder);
            q.addBindValue(e.num);
            q.addBindValue(e.text);
            q.addBindValue(who);
        }
        if (!q.exec()) {
            lastError = QStringLiteral("tryWriteFile(upsert %1, sample id=%2): ")
                              .arg(table).arg(sampleId) + q.lastError().text();
            return false;
        }
        // RETURNING yields one row per tuple, inserted or updated alike, so this
        // is the complete post-image for the prune.
        while (q.next()) postIds.insert(q.value(0).toLongLong());
        batch.clear();
        return true;
    };

    // Resolves one extra into a batch entry. Returns false only on a real SQL
    // failure; an absent value is skipped (sparse rule, index D2 - a measurement
    // exists exactly where there is a value, and a numeric 0 IS a value).
    //
    // A key with a same-named WIDE column is skipped too, and that is the WRITE
    // half of the H1 guard whose read half is the prune pre-image above - same
    // catalog read, so the two cannot drift. Writing one of these would create a
    // row that is unprunable BY CONSTRUCTION: the pre-image excludes exactly
    // these keys, so it could never enter the delete set, the read has no key
    // filter (H25) so it would come back into `extra` on every load, and it
    // would be rewritten on every save. Deleting the column from the workbook
    // could then never delete its data. The wide column stays authoritative.
    QSet<QString> loggedCollisions;
    auto appendExtra = [&](const QString& kind, const QString& key,
                           const QVariant& value, int sortOrder,
                           const QSet<QString>& wideCols,
                           QVector<ExtraRow>& batch) -> bool {
        if (!value.isValid() || value.isNull()) return true;
        if (wideCols.contains(key)) {
            // Once per key, not once per row: a 1000-row sheet would otherwise
            // emit 1000 identical lines.
            const QString ck = kind + QLatin1Char('|') + key;
            if (!loggedCollisions.contains(ck)) {
                loggedCollisions.insert(ck);
                logDebug(QStringLiteral(
                    "persistFileCore: open metric %1 has a same-named wide column, "
                    "so it is NOT written to the long tables (H1) - the wide "
                    "column remains authoritative").arg(ck));
            }
            return true;
        }
        ExtraRow e;
        e.defId = metricDefs.ensureMetric(kind, key, value, &lastError);
        if (e.defId < 0) return false;
        e.sortOrder = sortOrder;
        // Routed by the RESOLVED metric_defs.value_type, never by the C++ type:
        // the read path routes on the same column, so the registry's ratified
        // vocabulary is the single authority on what a metric is.
        MetricDefCache::encodeValue(metricDefs.valueType(e.defId), value,
                                    e.num, e.text);
        batch.append(e);
        return true;
    };

    for (int si = 0; si < result.sheets.size(); ++si) {
        SheetResult& sheet = result.sheets[si];

        qint64 testId = -1;
        if (sheet.id != -1 && sheet.version > 0) {
            updateTest.bindValue(0, fileId);
            updateTest.bindValue(1, sheet.sheetName);
            updateTest.bindValue(2, sheet.templateVersion);
            updateTest.bindValue(3, sheet.overallAvgTPM);
            updateTest.bindValue(4, sheet.overallStdDevTPM);
            updateTest.bindValue(5, sheet.isRawTable ? 1 : 0);
            updateTest.bindValue(6, sheet.fromInferredSchema ? 1 : 0);
            updateTest.bindValue(7, si);
            updateTest.bindValue(8, sheet.isRawTable
                ? QVariant(rawGridToJson(sheet.rawHeaders, sheet.rawRows))
                : QVariant());
            updateTest.bindValue(9, who);
            updateTest.bindValue(10, static_cast<qlonglong>(sheet.id));
            // RC1 child-row fresh-version OCC (see freshChildVersion): adopt
            // the row's current committed version, not the routinely-stale
            // sheet.version. Row-level last-writer-wins by design.
            updateTest.bindValue(11, freshChildVersion(db, QStringLiteral("tests"),
                                                        sheet.id, sheet.version));
            if (!updateTest.exec()) {
                lastError = QStringLiteral("tryWriteFile(UPDATE test id=%1): ")
                                  .arg(sheet.id) + updateTest.lastError().text();
                db.rollback(); logDebug(lastError); return WriteResult::OtherError;
            }
            if (!updateTest.next()) {
                QString detail;
                const WriteResult cls = classifyMissingUpdate(
                    db, QStringLiteral("tests"), sheet.id, &detail);
                lastError = QStringLiteral("tryWriteFile(UPDATE test id=%1): %2")
                                  .arg(sheet.id).arg(cls == WriteResult::VersionMismatch
                                      ? "version mismatch"
                                      : (cls == WriteResult::RowDeleted ? "row deleted" : detail));
                db.rollback(); logDebug(lastError); return cls;
            }
            testId = sheet.id;
            sheet.version = updateTest.value(0).toInt();
        } else {
            insertTest.bindValue(0, fileId);
            insertTest.bindValue(1, sheet.sheetName);
            insertTest.bindValue(2, sheet.templateVersion);
            insertTest.bindValue(3, sheet.overallAvgTPM);
            insertTest.bindValue(4, sheet.overallStdDevTPM);
            insertTest.bindValue(5, sheet.isRawTable ? 1 : 0);
            insertTest.bindValue(6, sheet.fromInferredSchema ? 1 : 0);
            insertTest.bindValue(7, si);
            insertTest.bindValue(8, sheet.isRawTable
                ? QVariant(rawGridToJson(sheet.rawHeaders, sheet.rawRows))
                : QVariant());
            insertTest.bindValue(9, who);
            if (!insertTest.exec() || !insertTest.next()) {
                lastError = QStringLiteral("tryWriteFile(INSERT test): ")
                              + insertTest.lastError().text();
                db.rollback(); logDebug(lastError); return WriteResult::OtherError;
            }
            testId = insertTest.value(0).toLongLong();
            sheet.id      = testId;
            sheet.version = insertTest.value(1).toInt();
        }
        postTestIds.insert(testId);

        for (int sj = 0; sj < sheet.samples.size(); ++sj) {
            SampleResult& sr = sheet.samples[sj];

            qint64 sampleId = -1;
            if (sr.id != -1 && sr.version > 0) {
                updateSample.bindValue(0,  static_cast<qlonglong>(testId));
                updateSample.bindValue(1,  sj);
                updateSample.bindValue(2,  sr.sampleName);
                updateSample.bindValue(3,  sr.sampleID);
                updateSample.bindValue(4,  sr.date);
                updateSample.bindValue(5,  sr.tester);
                updateSample.bindValue(6,  sr.media);
                updateSample.bindValue(7,  sr.viscosity);
                updateSample.bindValue(8,  sr.resistance);
                updateSample.bindValue(9,  sr.voltage);
                updateSample.bindValue(10, sr.power);
                updateSample.bindValue(11, sr.heatingTechnology);
                updateSample.bindValue(12, sr.puffingRegime);
                updateSample.bindValue(13, sr.initialOilMass);
                updateSample.bindValue(14, sr.averageTPM);
                updateSample.bindValue(15, sr.stdDevTPM);
                updateSample.bindValue(16, sr.averagePowerDensity);
                updateSample.bindValue(17, sr.efficiencyPercent);
                updateSample.bindValue(18, sr.totalOilConsumed);
                updateSample.bindValue(19, sr.totalPuffs);
                updateSample.bindValue(20, sr.normalizedTPM);
                updateSample.bindValue(21, sr.burnStatus);
                updateSample.bindValue(22, sr.clogStatus);
                updateSample.bindValue(23, sr.leakStatus);
                updateSample.bindValue(24, who);
                updateSample.bindValue(25, static_cast<qlonglong>(sr.id));
                // RC1 child-row fresh-version OCC (see freshChildVersion).
                updateSample.bindValue(26, freshChildVersion(db, QStringLiteral("samples"),
                                                              sr.id, sr.version));
                if (!updateSample.exec()) {
                    lastError = QStringLiteral("tryWriteFile(UPDATE sample id=%1): ")
                                      .arg(sr.id) + updateSample.lastError().text();
                    db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                }
                if (!updateSample.next()) {
                    QString detail;
                    const WriteResult cls = classifyMissingUpdate(
                        db, QStringLiteral("samples"), sr.id, &detail);
                    lastError = QStringLiteral("tryWriteFile(UPDATE sample id=%1): %2")
                                      .arg(sr.id).arg(cls == WriteResult::VersionMismatch
                                          ? "version mismatch"
                                          : (cls == WriteResult::RowDeleted ? "row deleted" : detail));
                    db.rollback(); logDebug(lastError); return cls;
                }
                sampleId = sr.id;
                sr.version = updateSample.value(0).toInt();
            } else {
                insertSample.bindValue(0,  static_cast<qlonglong>(testId));
                insertSample.bindValue(1,  sj);
                insertSample.bindValue(2,  sr.sampleName);
                insertSample.bindValue(3,  sr.sampleID);
                insertSample.bindValue(4,  sr.date);
                insertSample.bindValue(5,  sr.tester);
                insertSample.bindValue(6,  sr.media);
                insertSample.bindValue(7,  sr.viscosity);
                insertSample.bindValue(8,  sr.resistance);
                insertSample.bindValue(9,  sr.voltage);
                insertSample.bindValue(10, sr.power);
                insertSample.bindValue(11, sr.heatingTechnology);
                insertSample.bindValue(12, sr.puffingRegime);
                insertSample.bindValue(13, sr.initialOilMass);
                insertSample.bindValue(14, sr.averageTPM);
                insertSample.bindValue(15, sr.stdDevTPM);
                insertSample.bindValue(16, sr.averagePowerDensity);
                insertSample.bindValue(17, sr.efficiencyPercent);
                insertSample.bindValue(18, sr.totalOilConsumed);
                insertSample.bindValue(19, sr.totalPuffs);
                insertSample.bindValue(20, sr.normalizedTPM);
                insertSample.bindValue(21, sr.burnStatus);
                insertSample.bindValue(22, sr.clogStatus);
                insertSample.bindValue(23, sr.leakStatus);
                insertSample.bindValue(24, who);
                if (!insertSample.exec() || !insertSample.next()) {
                    lastError = QStringLiteral("tryWriteFile(INSERT sample): ")
                                  + insertSample.lastError().text();
                    db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                }
                sampleId = insertSample.value(0).toLongLong();
                sr.id      = sampleId;
                sr.version = insertSample.value(1).toInt();
            }
            postSampleIds.insert(sampleId);

            // -- data rows ------------------------------------------------
            for (int ri = 0; ri < sr.rows.size(); ++ri) {
                DataRow& dr = sr.rows[ri];
                if (dr.id != -1 && dr.version > 0) {
                    updateRow.bindValue(0,  static_cast<qlonglong>(sampleId));
                    updateRow.bindValue(1,  ri);
                    updateRow.bindValue(2,  dr.puffs);
                    updateRow.bindValue(3,  dr.beforeWeight);
                    updateRow.bindValue(4,  dr.afterWeight);
                    updateRow.bindValue(5,  dr.drawPressure);
                    updateRow.bindValue(6,  dr.resistance);
                    updateRow.bindValue(7,  dr.smell);
                    updateRow.bindValue(8,  dr.clog);
                    updateRow.bindValue(9,  dr.notes);
                    updateRow.bindValue(10, dr.tpm);
                    updateRow.bindValue(11, dr.tpmPowerDensity);
                    updateRow.bindValue(12, dr.variationTPM);
                    updateRow.bindValue(13, dr.oilConsumed);
                    updateRow.bindValue(14, sheet.hasPerRowRegime ? QVariant(dr.puffingRegime) : QVariant());
                    updateRow.bindValue(15, who);
                    updateRow.bindValue(16, static_cast<qlonglong>(dr.id));
                    // RC1 child-row fresh-version OCC (see freshChildVersion).
                    updateRow.bindValue(17, freshChildVersion(db, QStringLiteral("data_rows"),
                                                              dr.id, dr.version));
                    if (!updateRow.exec()) {
                        lastError = QStringLiteral("tryWriteFile(UPDATE data_row id=%1): ")
                                          .arg(dr.id) + updateRow.lastError().text();
                        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                    }
                    if (!updateRow.next()) {
                        QString detail;
                        const WriteResult cls = classifyMissingUpdate(
                            db, QStringLiteral("data_rows"), dr.id, &detail);
                        lastError = QStringLiteral("tryWriteFile(UPDATE data_row id=%1): %2")
                                          .arg(dr.id).arg(cls == WriteResult::VersionMismatch
                                              ? "version mismatch"
                                              : (cls == WriteResult::RowDeleted ? "row deleted" : detail));
                        db.rollback(); logDebug(lastError); return cls;
                    }
                    dr.version = updateRow.value(0).toInt();
                } else {
                    insertRow.bindValue(0,  static_cast<qlonglong>(sampleId));
                    insertRow.bindValue(1,  ri);
                    insertRow.bindValue(2,  dr.puffs);
                    insertRow.bindValue(3,  dr.beforeWeight);
                    insertRow.bindValue(4,  dr.afterWeight);
                    insertRow.bindValue(5,  dr.drawPressure);
                    insertRow.bindValue(6,  dr.resistance);
                    insertRow.bindValue(7,  dr.smell);
                    insertRow.bindValue(8,  dr.clog);
                    insertRow.bindValue(9,  dr.notes);
                    insertRow.bindValue(10, dr.tpm);
                    insertRow.bindValue(11, dr.tpmPowerDensity);
                    insertRow.bindValue(12, dr.variationTPM);
                    insertRow.bindValue(13, dr.oilConsumed);
                    insertRow.bindValue(14, sheet.hasPerRowRegime ? QVariant(dr.puffingRegime) : QVariant());
                    insertRow.bindValue(15, who);
                    if (!insertRow.exec() || !insertRow.next()) {
                        lastError = QStringLiteral("tryWriteFile(INSERT data_row): ")
                                      + insertRow.lastError().text();
                        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                    }
                    dr.id      = insertRow.value(0).toLongLong();
                    dr.version = insertRow.value(1).toInt();
                }
                postDataRowIds.insert(dr.id);
            }

            // -- open metrics (v3 Phase 3c) -------------------------------
            // A separate pass over the same rows on purpose: the standard
            // 13-column loop above is a stated non-goal of 3c and stays
            // byte-for-byte what it was. measurements are keyed
            // (sample_id, metric_id, sort_order = ri) - the row's ORDINAL, not
            // its id, because measurements has no link to data_rows at all
            // (hazard H22; see the prune in phase C).
            if (guardRows) {
                QVector<ExtraRow> rowBatch;
                for (int ri = 0; ri < sr.rows.size(); ++ri) {
                    const DataRow& dr = sr.rows[ri];
                    for (auto it = dr.extra.constBegin(); it != dr.extra.constEnd(); ++it) {
                        if (!appendExtra(QStringLiteral("metric"), it.key(),
                                         it.value(), ri, dataRowWideCols, rowBatch)) {
                            db.rollback(); logDebug(lastError);
                            return WriteResult::OtherError;
                        }
                        if (rowBatch.size() >= kExtrasBatchRows
                            && !flushExtras(QStringLiteral("measurements"), sampleId,
                                            rowBatch, postMeasurementIds)) {
                            db.rollback(); logDebug(lastError);
                            return WriteResult::OtherError;
                        }
                    }
                }
                if (!flushExtras(QStringLiteral("measurements"), sampleId,
                                 rowBatch, postMeasurementIds)) {
                    db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                }
            }

            // SampleResult::extra -> sample_headers, kind 'header'. Keyed
            // (sample_id, field_id): a header field is per-sample, so there
            // is no sort_order dimension.
            if (guardHeaders) {
                QVector<ExtraRow> hdrBatch;
                for (auto it = sr.extra.constBegin(); it != sr.extra.constEnd(); ++it) {
                    if (!appendExtra(QStringLiteral("header"), it.key(),
                                     it.value(), 0, sampleWideCols, hdrBatch)) {
                        db.rollback(); logDebug(lastError);
                        return WriteResult::OtherError;
                    }
                    if (hdrBatch.size() >= kExtrasBatchRows
                        && !flushExtras(QStringLiteral("sample_headers"), sampleId,
                                        hdrBatch, postSampleHeaderIds)) {
                        db.rollback(); logDebug(lastError);
                        return WriteResult::OtherError;
                    }
                }
                if (!flushExtras(QStringLiteral("sample_headers"), sampleId,
                                 hdrBatch, postSampleHeaderIds)) {
                    db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                }
            }

            // -- images (per-sample) --------------------------------------
            // Reads each on-disk image into a BYTEA blob and stores the
            // layout/crop rectangles. Skips files we can't open or that
            // exceed 100 MB. Parallel imageIds/imageVersions vectors drive
            // the UPDATE-vs-INSERT split.
            const int imageN = sr.imagePaths.size();
            while (sr.imageIds.size() < imageN)      sr.imageIds.append(-1);
            while (sr.imageVersions.size() < imageN) sr.imageVersions.append(0);
            for (int ii = 0; ii < imageN; ++ii) {
                QByteArray imgData;
                QFile imgFile(sr.imagePaths[ii]);
                if (imgFile.open(QIODevice::ReadOnly)) {
                    constexpr qint64 kMaxImageSize = 100 * 1024 * 1024;
                    if (imgFile.size() <= kMaxImageSize)
                        imgData = imgFile.readAll();
                    else
                        qWarning() << "Skipping oversized image:" << imgFile.fileName();
                }
                const QRectF layout = (ii < sr.imageLayouts.size()) ? sr.imageLayouts[ii] : QRectF();
                const QRectF crop   = (ii < sr.imageCrops.size())   ? sr.imageCrops[ii]   : QRectF(0,0,1,1);
                const QString fname = QFileInfo(sr.imagePaths[ii]).fileName();

                const qint64 imgId  = sr.imageIds[ii];
                const int    imgVer = sr.imageVersions[ii];

                if (imgId != -1 && imgVer > 0) {
                    updateImage.bindValue(0,  static_cast<qlonglong>(sampleId));
                    updateImage.bindValue(1,  ii);
                    updateImage.bindValue(2,  fname);
                    updateImage.bindValue(3,  imgData);
                    updateImage.bindValue(4,  layout.x());
                    updateImage.bindValue(5,  layout.y());
                    updateImage.bindValue(6,  layout.width());
                    updateImage.bindValue(7,  layout.height());
                    updateImage.bindValue(8,  crop.x());
                    updateImage.bindValue(9,  crop.y());
                    updateImage.bindValue(10, crop.width());
                    updateImage.bindValue(11, crop.height());
                    updateImage.bindValue(12, who);
                    updateImage.bindValue(13, static_cast<qlonglong>(imgId));
                    // RC1 child-row fresh-version OCC (see freshChildVersion).
                    updateImage.bindValue(14, freshChildVersion(db, QStringLiteral("images"),
                                                                imgId, imgVer));
                    if (!updateImage.exec()) {
                        lastError = QStringLiteral("tryWriteFile(UPDATE image id=%1): ")
                                          .arg(imgId) + updateImage.lastError().text();
                        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                    }
                    if (!updateImage.next()) {
                        QString detail;
                        const WriteResult cls = classifyMissingUpdate(
                            db, QStringLiteral("images"), imgId, &detail);
                        lastError = QStringLiteral("tryWriteFile(UPDATE image id=%1): %2")
                                          .arg(imgId).arg(cls == WriteResult::VersionMismatch
                                              ? "version mismatch"
                                              : (cls == WriteResult::RowDeleted ? "row deleted" : detail));
                        db.rollback(); logDebug(lastError); return cls;
                    }
                    sr.imageVersions[ii] = updateImage.value(0).toInt();
                    postImageIds.insert(imgId);
                } else {
                    insertImage.bindValue(0,  static_cast<qlonglong>(sampleId));
                    insertImage.bindValue(1,  ii);
                    insertImage.bindValue(2,  fname);
                    insertImage.bindValue(3,  imgData);
                    insertImage.bindValue(4,  layout.x());
                    insertImage.bindValue(5,  layout.y());
                    insertImage.bindValue(6,  layout.width());
                    insertImage.bindValue(7,  layout.height());
                    insertImage.bindValue(8,  crop.x());
                    insertImage.bindValue(9,  crop.y());
                    insertImage.bindValue(10, crop.width());
                    insertImage.bindValue(11, crop.height());
                    insertImage.bindValue(12, who);
                    if (!insertImage.exec() || !insertImage.next()) {
                        lastError = QStringLiteral("tryWriteFile(INSERT image): ")
                                      + insertImage.lastError().text();
                        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                    }
                    sr.imageIds[ii]      = insertImage.value(0).toLongLong();
                    sr.imageVersions[ii] = insertImage.value(1).toInt();
                    postImageIds.insert(sr.imageIds[ii]);
                }
            }
        }
    }

    // Phase C: post-prune. Build an id list of rows that existed pre-save but
    // were not touched this pass. Delete in child-first order so FKs don't
    // CASCADE-wipe descendants we just upserted.
    auto pruneOrphans = [&](const QString& table,
                             const QSet<qint64>& pre,
                             const QSet<qint64>& post) -> WriteResult {
        QStringList orphanCsv;
        orphanCsv.reserve(pre.size());
        for (qint64 id : pre) {
            if (!post.contains(id)) orphanCsv.append(QString::number(id));
        }
        if (orphanCsv.isEmpty()) return WriteResult::Success;
        QSqlQuery q(db);
        // Safe to interpolate: orphanCsv elements are all qint64-from-DB ids
        // formatted as base-10 ints — no SQL injection surface.
        if (!q.exec(QString("DELETE FROM %1 WHERE id IN (%2)")
                        .arg(table, orphanCsv.join(","))) ) {
            lastError = QStringLiteral("tryWriteFile(prune %1): ").arg(table)
                          + q.lastError().text();
            return WriteResult::OtherError;
        }
        return WriteResult::Success;
    };
    // v3 Phase 3c, DECISION D9 / HAZARD H22 - why measurements has to be pruned
    // here and cannot be left to a foreign key.
    //
    // measurements has NO referential link to data_rows. Its only lifecycle FK
    // is sample_id -> samples(id) ON DELETE CASCADE, which fires when a SAMPLE
    // is deleted and does nothing whatsoever when a data ROW is. So the
    // `DELETE FROM data_rows` below leaves the deleted row's measurements
    // behind: delete row 5 of 10 and save, the survivors renumber their
    // sort_order, and the stranded measurements re-bind to the WRONG rows on the
    // next load. Silent corruption, not a crash.
    //
    // D9 ruled the prune owns it. The tempting fix -
    // `measurements.data_row_id REFERENCES data_rows(id) ON DELETE CASCADE` -
    // cannot survive 3d, where data_rows is renamed aside and a view takes its
    // name; an FK onto a view is impossible, so it would have to be dropped at
    // exactly the moment the data matters most.
    //
    // It needs no special case: the surviving-row set the prune already computes
    // IS the answer. A deleted data row contributes nothing to phase B, so its
    // measurements are in the pre-image and not in the post-image, and the same
    // pre-minus-post that removes the data row removes them. The same mechanism
    // also removes a custom column the user deleted. Both are guarded by
    // tst_v3longformat::orphanTripwire_noMeasurementWithoutItsDataRow and by
    // tst_saveintegrity_e2e scenarios 20 and 21.
    //
    // Long tables go FIRST - child-most, before the data_rows and samples they
    // hang off. See the H1 banner on the phase A pre-image for the constraint
    // that makes this delete safe in 3c and unsafe the moment 3d changes it.
    if ((guardRows
         && pruneOrphans("measurements",   preMeasurementIds,  postMeasurementIds)  != WriteResult::Success) ||
        (guardHeaders
         && pruneOrphans("sample_headers", preSampleHeaderIds, postSampleHeaderIds) != WriteResult::Success) ||
        pruneOrphans("images",    preImageIds,   postImageIds)   != WriteResult::Success ||
        pruneOrphans("data_rows", preDataRowIds, postDataRowIds) != WriteResult::Success ||
        pruneOrphans("samples",   preSampleIds,  postSampleIds)  != WriteResult::Success ||
        pruneOrphans("tests",     preTestIds,    postTestIds)    != WriteResult::Success) {
        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
    }

    if (!db.commit()) {
        lastError = QStringLiteral("tryWriteFile(commit): ") + db.lastError().text();
        db.rollback();
        logDebug(lastError);
        return WriteResult::OtherError;
    }
    logDebug(QString("Saved file '%1' (fileId=%2, version=%3, %4 sheets, %5 samples)")
                 .arg(result.fileName).arg(fileId).arg(newVer)
                 .arg(result.sheets.size()).arg(totalSamples));
    // Writeback: parent file id + version only. Child ids (sample/data_row/
    // image) are intentionally NOT cascaded back — callers that need fresh
    // child ids should issue a follow-up loadFile(result.id) after this
    // returns. This keeps the writeback footprint small while letting the
    // recreate handler in MainWindow round-trip the parent id correctly.
    result.id      = fileId;
    result.version = newVer;
    return WriteResult::Success;
}

WriteResult persistFileOnConnection(QSqlDatabase& db, const QString& who,
                                    bool online, bool dbOpen,
                                    FileResult& result, QString* outError) {
    if (!online) {
        if (outError) *outError = QStringLiteral("DatabaseManager is offline (read-only mode)");
        return WriteResult::OfflineReadOnly;
    }
    if (!dbOpen) {
        if (outError) *outError = QStringLiteral("persistFile: database not open");
        return WriteResult::OtherError;
    }
    WriteResult r = persistFileCore(db, who, result, outError);
    if (r == WriteResult::RowDeleted) {
        const bool fileRowGone = (result.id != -1) && !fileRowExistsOnDb(db, result.id);
        logDebug(QStringLiteral("persistFileOnConnection: row deleted out-of-band "
                                "(fileId=%1, fileRowGone=%2) -- re-INSERTing as fresh rows")
                     .arg(result.id).arg(fileRowGone));
        resetFileIdsForReinsert(result, fileRowGone);
        r = persistFileCore(db, who, result, outError);
    } else if (r == WriteResult::VersionMismatch) {
        logDebug(QStringLiteral("persistFileOnConnection: version mismatch (fileId=%1) "
                                "-- retrying core once").arg(result.id));
        r = persistFileCore(db, who, result, outError);
    }
    return r;
}

} // namespace DVE
