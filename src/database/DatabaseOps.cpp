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

    // ── v3 Phase 3d: EVERY metric is a long-format write now ─────────────────
    // The standard 13 data-row metrics and 22 sample headers write to
    // measurements / sample_headers exactly like the open metrics 3c moved
    // first; post-cutover data_rows is a read-only name-holder view and the
    // wide statements this function used to prepare could only fail.
    //
    // The probe runs BEFORE the transaction on purpose. A failed statement
    // poisons a Postgres transaction, so probing from inside one would turn
    // "this database has no long tables yet" into a poisoned save. But unlike
    // 3c - where the wide path was still there to fall back on - there is NO
    // fallback any more: a save that cannot reach the long tables has nowhere
    // to put the standard metrics, and a half-save is data loss. Abort.
    MetricDefCache metricDefs(db, who);
    QString metricDefsError;
    if (!metricDefs.load(&metricDefsError)) {
        lastError = QStringLiteral(
            "tryWriteFile(long-format probe): metric_defs/measurements/"
            "sample_headers unreachable and there is no wide write path any "
            "more (v3 Phase 3d) -- ") + metricDefsError;
        logDebug(lastError);
        return WriteResult::OtherError;
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

    // ── the wide-column catalog, read ONCE ───────────────────────────────────
    // v3 Phase 3d: with the prune exclusion lifted, ONE consumer remains -
    // appendExtra's skip. A key with a same-named column in the wide relation
    // (post-cutover: the name-holder view) is a STANDARD metric, written from
    // its member field by appendStandard below; a stale pre-3d recovery-JSON
    // `extra` map carrying the same key must not double-write it. The catalog
    // is also what loadFile's H25 read filter consults (as a SQL subquery), so
    // writer and reader agree by construction on what "standard" means.
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
    {
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

    // v3 Phase 3d: an empty catalog set is no longer a degrade-gracefully
    // case. The set defines which keys are STANDARD (the appendExtra
    // double-write skip below reads it), and post-cutover `data_rows` /
    // `samples` are the name-holder views whose column sets are guaranteed by
    // the cutover migration. Zero visible columns means the relation is
    // missing or the role cannot see it - either way this connection cannot
    // be trusted with a save.
    if (dataRowWideCols.isEmpty() || sampleWideCols.isEmpty()) {
        lastError = QStringLiteral(
            "tryWriteFile(wide-column catalog): zero visible columns for %1 - "
            "refusing to save against an unrecognizable schema")
                .arg(dataRowWideCols.isEmpty() && sampleWideCols.isEmpty()
                         ? QStringLiteral("data_rows and samples")
                         : (dataRowWideCols.isEmpty() ? QStringLiteral("data_rows")
                                                      : QStringLiteral("samples")));
        db.rollback(); logDebug(lastError); return WriteResult::OtherError;
    }

    // Phase A: pre-image. Five scoped SELECTs ride the same QSqlQuery.
    // v3 Phase 3d: there is no data_rows pre-image any more - a data row IS
    // its measurements, so the measurements pre-image below carries the
    // deleted-row story (D9: the prune owns the row-measurement lifecycle).
    // samples_core replaces samples as the identity table; the name `samples`
    // is the read-only name-holder view.
    QSet<qint64> preTestIds, preSampleIds, preImageIds;
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

        q.prepare("SELECT s.id FROM samples_core s "
                  "JOIN tests t ON s.test_id = t.id WHERE t.file_id = ?");
        q.addBindValue(fileId);
        if (!q.exec()) {
            lastError = QStringLiteral("tryWriteFile(preImage samples_core): ")
                          + q.lastError().text();
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
        while (q.next()) preSampleIds.insert(q.value(0).toLongLong());

        q.prepare("SELECT im.id FROM images im "
                  "JOIN samples_core s ON im.sample_id = s.id "
                  "JOIN tests t ON s.test_id = t.id WHERE t.file_id = ?");
        q.addBindValue(fileId);
        if (!q.exec()) {
            lastError = QStringLiteral("tryWriteFile(preImage images): ")
                          + q.lastError().text();
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }
        while (q.next()) preImageIds.insert(q.value(0).toLongLong());

        // ── the long tables' pre-image ───────────────────────────────────
        //
        // v3 Phase 3d: the H1 guard's `md.key NOT IN (<wide columns>)`
        // exclusion is LIFTED - deliberately, together with the phase-B
        // writer below that now reproduces EVERY standard metric of every
        // surviving row, which is exactly the condition the 3c banner set
        // for lifting it. The full pre-image is what makes deleted-row
        // cleanup work at all under the long shape: delete row 5 of 10 and
        // save, and the only record that row ever existed is its
        // measurements - they are in this pre-image, nothing re-writes them
        // in phase B, and pre-minus-post removes them (D9/H22).
        //
        // What SURVIVES of the H1 guard is the write half: appendExtra still
        // skips a key with a same-named wide column, so a standard key that
        // sneaks in via a stale pre-3d recovery-JSON `extra` map cannot
        // double-write against the member-field batch.
        auto preImageIdsOf = [&](const QString& label, const QString& sql,
                                 QSet<qint64>* out) -> bool {
            if (!q.prepare(sql)) {
                lastError = QStringLiteral("tryWriteFile(preImage %1): ").arg(label)
                              + q.lastError().text();
                return false;
            }
            q.addBindValue(fileId);
            if (!q.exec()) {
                lastError = QStringLiteral("tryWriteFile(preImage %1): ").arg(label)
                              + q.lastError().text();
                return false;
            }
            while (q.next()) out->insert(q.value(0).toLongLong());
            return true;
        };

        if (!preImageIdsOf(QStringLiteral("measurements"),
                           QStringLiteral(
                               "SELECT m.id FROM measurements m "
                               "JOIN samples_core s ON s.id = m.sample_id "
                               "JOIN tests t ON s.test_id = t.id "
                               "WHERE t.file_id = ?"),
                           &preMeasurementIds)) {
            db.rollback(); logDebug(lastError); return WriteResult::OtherError;
        }

        if (!preImageIdsOf(QStringLiteral("sample_headers"),
                           QStringLiteral(
                               "SELECT sh.id FROM sample_headers sh "
                               "JOIN samples_core s ON s.id = sh.sample_id "
                               "JOIN tests t ON s.test_id = t.id "
                               "WHERE t.file_id = ?"),
                           &preSampleHeaderIds)) {
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
    // v3 Phase 3d: samples_core is the surviving NARROW identity table -
    // id/test_id/sort_order/audit only. The 22 value columns ride the
    // sample_headers batch below, and the 13 data-row metrics ride the
    // measurements batch; there are no wide UPDATE/INSERT statements left.
    QSqlQuery updateSample(db), insertSample(db);
    if (!updateSample.prepare(
            "UPDATE samples_core SET test_id = ?, sort_order = ?, updated_by = ? "
            "WHERE id = ? AND version = ? RETURNING version") ||
        !insertSample.prepare(
            "INSERT INTO samples_core (test_id, sort_order, updated_by) "
            "VALUES (?, ?, ?) RETURNING id, version")) {
        lastError = QStringLiteral("tryWriteFile(prepare samples_core): ")
                      + (updateSample.lastError().isValid()
                            ? updateSample.lastError().text() : insertSample.lastError().text());
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
    QSet<qint64> postTestIds, postSampleIds, postImageIds;
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

    // The sparse rule's "no value" test. Qt 6 gotcha: QVariant::isNull() no
    // longer delegates to the CONTAINED value, so a QVariant holding a null
    // QString - which is exactly what an untouched text member or a wide NULL
    // read back through loadFile produces - reports NOT null. Without the
    // string leg, every null-string member minted a measurement row with BOTH
    // value columns NULL: a D2 violation encodeValue's contract explicitly
    // rules out, caught by scenario23's both-null probe on the first
    // post-cutover run. An EMPTY (non-null) string is a real value and writes.
    auto isAbsentValue = [](const QVariant& v) -> bool {
        return !v.isValid() || v.isNull()
            || (v.typeId() == QMetaType::QString && v.toString().isNull());
    };

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
        if (isAbsentValue(value)) return true;   // sparse (see isAbsentValue)
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

    // ── v3 Phase 3d: the STANDARD metrics join the same batches ──────────────
    // Same ExtraRow/flush machinery, two differences from appendExtra:
    // resolution is LOOKUP-ONLY (a standard key missing from metric_defs means
    // the seed/ensureSchema contract is broken - abort loudly, never
    // auto-register a guessed type), and there is no wide-column skip (these
    // ARE the wide columns).
    //
    // The sparse rule at write time is "a measurement exists exactly where the
    // old wide bind was non-NULL": the only nullable wide binds were the text
    // columns' null QStrings and the !hasPerRowRegime regime - both arrive
    // here as null QVariants and are skipped, exactly like appendExtra.
    auto appendStandard = [&](const QString& kind, const char* key,
                              const QVariant& value, int sortOrder,
                              QVector<ExtraRow>& batch) -> bool {
        if (isAbsentValue(value)) return true;   // sparse (see isAbsentValue)
        ExtraRow e;
        e.defId = metricDefs.lookup(kind, QLatin1String(key));
        if (e.defId < 0) {
            lastError = QStringLiteral(
                "tryWriteFile(standard metric %1/%2): no metric_defs row - the "
                "seed/ensureSchema contract is broken; aborting the save")
                    .arg(kind, QLatin1String(key));
            return false;
        }
        e.sortOrder = sortOrder;
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
                updateSample.bindValue(0, static_cast<qlonglong>(testId));
                updateSample.bindValue(1, sj);
                updateSample.bindValue(2, who);
                updateSample.bindValue(3, static_cast<qlonglong>(sr.id));
                // RC1 child-row fresh-version OCC (see freshChildVersion).
                updateSample.bindValue(4, freshChildVersion(db, QStringLiteral("samples_core"),
                                                             sr.id, sr.version));
                if (!updateSample.exec()) {
                    lastError = QStringLiteral("tryWriteFile(UPDATE sample id=%1): ")
                                      .arg(sr.id) + updateSample.lastError().text();
                    db.rollback(); logDebug(lastError); return WriteResult::OtherError;
                }
                if (!updateSample.next()) {
                    QString detail;
                    const WriteResult cls = classifyMissingUpdate(
                        db, QStringLiteral("samples_core"), sr.id, &detail);
                    lastError = QStringLiteral("tryWriteFile(UPDATE sample id=%1): %2")
                                      .arg(sr.id).arg(cls == WriteResult::VersionMismatch
                                          ? "version mismatch"
                                          : (cls == WriteResult::RowDeleted ? "row deleted" : detail));
                    db.rollback(); logDebug(lastError); return cls;
                }
                sampleId = sr.id;
                sr.version = updateSample.value(0).toInt();
            } else {
                insertSample.bindValue(0, static_cast<qlonglong>(testId));
                insertSample.bindValue(1, sj);
                insertSample.bindValue(2, who);
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

            // -- data rows: 13 standard metrics + open extras, ONE batch ---
            // v3 Phase 3d: a data row IS its measurements. Keyed
            // (sample_id, metric_id, sort_order = ri) - the row's ORDINAL.
            // dr.id / dr.version are deliberately NOT read or written here:
            // post-cutover loadFile fills them from the name-holder view,
            // where id is a synthetic MIN(measurement id) surrogate banned
            // from write-back (H2 closes - keyed upserts need no id anchors,
            // which also dissolves the stale-child-id classes H3/H13 for
            // rows). Row-level OCC goes with them: the wide path already
            // adopted the fresh version (row-level last-writer-wins by
            // design, RC1), and the keyed upsert has identical semantics
            // with per-measurement no-op suppression (D4/H23) on top.
            //
            // The sparse rule mirrors the old binds exactly: every numeric
            // member always carried a value and always writes; a null QString
            // was a NULL bind and is skipped; puffing_regime is gated on
            // hasPerRowRegime just as its bind was.
            {
                QVector<ExtraRow> rowBatch;
                for (int ri = 0; ri < sr.rows.size(); ++ri) {
                    const DataRow& dr = sr.rows[ri];
                    const struct { const char* key; QVariant value; } stdMetrics[] = {
                        { "puffs",             QVariant(dr.puffs) },
                        { "before_weight",     QVariant(dr.beforeWeight) },
                        { "after_weight",      QVariant(dr.afterWeight) },
                        { "draw_pressure",     QVariant(dr.drawPressure) },
                        { "resistance",        QVariant(dr.resistance) },
                        { "smell",             QVariant(dr.smell) },
                        { "clog",              QVariant(dr.clog) },
                        { "notes",             QVariant(dr.notes) },
                        { "tpm",               QVariant(dr.tpm) },
                        { "tpm_power_density", QVariant(dr.tpmPowerDensity) },
                        { "variation_tpm",     QVariant(dr.variationTPM) },
                        { "oil_consumed",      QVariant(dr.oilConsumed) },
                        { "puffing_regime",    sheet.hasPerRowRegime
                                                   ? QVariant(dr.puffingRegime)
                                                   : QVariant() },
                    };
                    for (const auto& m : stdMetrics) {
                        if (!appendStandard(QStringLiteral("metric"), m.key,
                                            m.value, ri, rowBatch)) {
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

            // -- sample headers: 22 standard fields + open extras ----------
            // Keyed (sample_id, field_id): a header field is per-sample, so
            // there is no sort_order dimension. Same sparse mirror as the
            // rows: doubles/ints always write, null QStrings skip.
            {
                QVector<ExtraRow> hdrBatch;
                const struct { const char* key; QVariant value; } stdHeaders[] = {
                    { "sample_name",        QVariant(sr.sampleName) },
                    { "sample_id",          QVariant(sr.sampleID) },
                    { "date",               QVariant(sr.date) },
                    { "tester",             QVariant(sr.tester) },
                    { "media",              QVariant(sr.media) },
                    { "viscosity",          QVariant(sr.viscosity) },
                    { "resistance",         QVariant(sr.resistance) },
                    { "voltage",            QVariant(sr.voltage) },
                    { "power",              QVariant(sr.power) },
                    { "heating_technology", QVariant(sr.heatingTechnology) },
                    { "puffing_regime",     QVariant(sr.puffingRegime) },
                    { "initial_oil_mass",   QVariant(sr.initialOilMass) },
                    { "average_tpm",        QVariant(sr.averageTPM) },
                    { "stddev_tpm",         QVariant(sr.stdDevTPM) },
                    { "avg_power_density",  QVariant(sr.averagePowerDensity) },
                    { "efficiency_percent", QVariant(sr.efficiencyPercent) },
                    { "total_oil_consumed", QVariant(sr.totalOilConsumed) },
                    { "total_puffs",        QVariant(sr.totalPuffs) },
                    { "normalized_tpm",     QVariant(sr.normalizedTPM) },
                    { "burn_status",        QVariant(sr.burnStatus) },
                    { "clog_status",        QVariant(sr.clogStatus) },
                    { "leak_status",        QVariant(sr.leakStatus) },
                };
                for (const auto& h : stdHeaders) {
                    if (!appendStandard(QStringLiteral("header"), h.key,
                                        h.value, 0, hdrBatch)) {
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
    // v3 Phase 3d, DECISION D9 / HAZARD H22 - the prune owns the
    // row-measurement lifecycle, and under the long shape it IS row deletion.
    //
    // measurements has no referential link to any data-row identity (there is
    // no data_rows table any more; the name is a read-only view). Its only
    // lifecycle FK is sample_id -> samples_core(id) ON DELETE CASCADE, which
    // fires when a SAMPLE is deleted and does nothing when a data ROW is. A
    // deleted row contributes nothing to phase B, so its measurements are in
    // the pre-image and not in the post-image, and pre-minus-post removes
    // them - which is exactly what "the row was deleted" MEANS now. Without
    // it, survivors renumber their sort_order and the stranded measurements
    // re-bind to the WRONG rows on the next load. Guarded by
    // tst_v3longformat::orphanTripwire_noMeasurementWithoutItsDataRow (the
    // pre-cutover rehearsal) and tst_saveintegrity_e2e scenarios 20, 21 and
    // 25.
    //
    // With the phase-A exclusion lifted (see the pre-image banner), the same
    // mechanism now covers STANDARD metrics too - safe precisely because
    // phase B reproduces every standard metric of every surviving row.
    //
    // Long tables go FIRST - child-most. samples_core's DELETE cascades any
    // long rows a pruned sample still owns.
    if (pruneOrphans("measurements",   preMeasurementIds,  postMeasurementIds)  != WriteResult::Success ||
        pruneOrphans("sample_headers", preSampleHeaderIds, postSampleHeaderIds) != WriteResult::Success ||
        pruneOrphans("images",       preImageIds,  postImageIds)  != WriteResult::Success ||
        pruneOrphans("samples_core", preSampleIds, postSampleIds) != WriteResult::Success ||
        pruneOrphans("tests",        preTestIds,   postTestIds)   != WriteResult::Success) {
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
