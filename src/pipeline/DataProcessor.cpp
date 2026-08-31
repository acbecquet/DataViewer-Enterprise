#include "DataProcessor.h"
#include "SheetProcessors.h"
#include "RegimeUtils.h"

#include "../model/StandardSchema.h"
#include "../model/SchemaDrivenReader.h"
#include "../model/SchemaResolver.h"
#include "../model/LegacyAdapter.h"

#include <QDebug>
#include <QFileInfo>
#include <memory>

namespace DVE {

// ===========================================================================
// Constructor / Destructor
// ===========================================================================

DataProcessor::DataProcessor() = default;
DataProcessor::~DataProcessor() = default;

// ===========================================================================
// Private helpers
// ===========================================================================

void DataProcessor::logDebug(const QString& msg) const
{
#ifndef QT_NO_DEBUG
    qDebug() << "[DataProcessor]" << msg;
#else
    Q_UNUSED(msg);
#endif
}

void DataProcessor::setError(const QString& msg)
{
    m_lastError = msg;
    qWarning() << "[DataProcessor] ERROR:" << msg;
}

namespace {

// The workbook's _dve_schema manifest sheet name as the reader reported it,
// or empty. Case-insensitive - Excel sheet names are case-insensitively
// unique, so at most one can match. Shared by BOTH parsers (production and
// the legacy shadow referee) so the exclusion stays symmetric.
QString manifestSheetIn(const QStringList& sheetNames)
{
    for (const QString& n : sheetNames)
        if (n.trimmed().compare(model::Manifest::kSheetName, Qt::CaseInsensitive) == 0)
            return n;
    return QString();
}

} // namespace

// ===========================================================================
// processFileLegacy — legacy positional read path (shadow-harness referee).
// Not the production path; see processFile below.
// ===========================================================================

FileResult DataProcessor::processFileLegacy(
    const QString& filePath,
    std::function<void(int, const QString&)> progressCallback)
{
    m_lastError.clear();
    // Legacy positional path never infers and never parses manifests - keep
    // both flags false so a caller that diffs legacy-vs-production can read
    // them consistently on either processor.
    m_lastUsedInference = false;
    m_lastHadManifest = false;

    FileResult result;
    result.filePath = filePath;
    result.fileName = QFileInfo(filePath).fileName();

    // -----------------------------------------------------------------------
    // 1. Open the file
    // -----------------------------------------------------------------------
    auto notify = [&](int pct, const QString& msg) {
        logDebug(QString("(%1%) %2").arg(pct).arg(msg));
        if (progressCallback)
            progressCallback(pct, msg);
    };

    notify(0, QStringLiteral("Opening file: ") + result.fileName);

    ExcelReader reader;
    if (!reader.loadFile(filePath)) {
        setError(QStringLiteral("Failed to open file: ") + reader.getLastError());
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();  // signal failure to onFileLoadFinished
        return result;
    }

    // -----------------------------------------------------------------------
    // 2. Detect template version
    // -----------------------------------------------------------------------
    QString templateVersion = reader.detectTemplateVersion();
    result.templateVersion  = templateVersion;
    logDebug(QStringLiteral("Template version: ") + templateVersion);

    // -----------------------------------------------------------------------
    // 3. Enumerate sheets
    // -----------------------------------------------------------------------
    QStringList sheetNames = reader.getSheetNames();

    // _dve_schema (the workbook's veryHidden manifest sheet) must never
    // display - excluded SYMMETRICALLY with the production parser (the
    // shadow harness diffs the two). Legacy only SKIPS it; it never parses
    // manifests.
    const QString manifestSheet = manifestSheetIn(sheetNames);
    if (!manifestSheet.isEmpty())
        sheetNames.removeOne(manifestSheet);

    result.sheetNames = sheetNames;

    if (sheetNames.isEmpty()) {
        setError(QStringLiteral("No sheets found in file: ") + filePath);
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();
        return result;
    }

    notify(5, QString("Found %1 sheet(s)").arg(sheetNames.size()));

    // -----------------------------------------------------------------------
    // 4. Process each sheet
    // -----------------------------------------------------------------------
    const int nSheets = sheetNames.size();

    for (int idx = 0; idx < nSheets; ++idx) {
        const QString& sheetName = sheetNames[idx];

        int progressBase = 5 + static_cast<int>(90.0 * idx / nSheets);
        notify(progressBase, QStringLiteral("Processing sheet: ") + sheetName);

        // Select the sheet in the reader.
        if (!reader.selectSheet(sheetName)) {
            qWarning() << "[DataProcessor] Could not select sheet:" << sheetName
                       << "-" << reader.getLastError();
            // Continue with remaining sheets rather than aborting.
            continue;
        }

        SheetResult sheetResult = processSheetLegacy(reader, sheetName);
        sheetResult.templateVersion = templateVersion;

        // Capture column headers (populated after selectSheet).
        sheetResult.columnHeaders = reader.getColumnHeaders();

        result.sheets.append(sheetResult);
    }

    notify(100, QStringLiteral("Done. Processed ")
               + QString::number(result.sheets.size())
               + QStringLiteral(" sheet(s)."));

    return result;
}

// ===========================================================================
// processSheetLegacy — legacy positional single-sheet read (shadow referee).
// ===========================================================================

SheetResult DataProcessor::processSheetLegacy(ExcelReader& reader, const QString& sheetName)
{
    logDebug(QStringLiteral("processSheetLegacy: ") + sheetName);

    SheetResult empty;
    empty.sheetName = sheetName;

    // Retrieve all samples from the currently-selected sheet.
    QVector<ExcelReader::SampleData> rawSamples = reader.getAllSamples();

    // ── Diagnostic dump (debug builds only) ─────────────────────────────────
#ifndef QT_NO_DEBUG
    qDebug() << "[DVE DIAG] Sheet:" << sheetName
             << " | rawSamples:" << rawSamples.size();
    for (int si = 0; si < rawSamples.size(); ++si) {
        const auto& rs = rawSamples[si];
        qDebug() << "  Sample" << si
                 << "| startCol:" << rs.startColumn
                 << "| dataRows:" << rs.dataRows.size()
                 << "| sampleID:" << rs.metadata.sampleID
                 << "| testName:" << rs.metadata.testName
                 << "| voltage:"  << rs.metadata.voltage
                 << "| resistance:" << rs.metadata.resistance;
        for (int ri = 0; ri < qMin(3, rs.dataRows.size()); ++ri) {
            QString rowStr;
            for (int ci = 0; ci < rs.dataRows[ri].size(); ++ci) {
                const QVariant& v = rs.dataRows[ri][ci];
                rowStr += QString("  [%1] type=%2 val=%3")
                    .arg(ci)
                    .arg(static_cast<int>(v.typeId()))
                    .arg(v.toString());
            }
            qDebug() << "    Row" << ri << ":" << rowStr;
        }
    }
#endif
    // ────────────────────────────────────────────────────────────────────────

    // ── SOP / instruction sheet — show raw table, no plot ───────────────
    if (sheetName.contains(QStringLiteral("SOP"), Qt::CaseInsensitive))
        return processSopSheet(reader, sheetName);

    if (rawSamples.isEmpty()) {
        logDebug(QStringLiteral("No samples found in sheet: ") + sheetName);
        // Return the empty result rather than an error — some sheets may be
        // intentionally blank (e.g. summary / chart sheets).
        return empty;
    }

    logDebug(QString("Sheet '%1': %2 sample(s)").arg(sheetName).arg(rawSamples.size()));

    // Select the appropriate processor and run it.
    std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));

    // Column index 4's header decides per-row Resistance vs Puffing Regime.
    const QStringList hdrs = reader.getColumnHeaders();
    const bool perRowRegime =
        (hdrs.size() > 4) && RegimeUtils::isRegimeHeader(hdrs.at(4));
    processor->setPerRowRegime(perRowRegime);

    // detectTemplateVersion() is a const method on the reader; we call it here
    // so each processSheet call carries the version even when called standalone.
    QString templateVersion = reader.detectTemplateVersion();

    SheetResult result;
    try {
        result = processor->process(rawSamples, sheetName, templateVersion);
    } catch (const std::exception& ex) {
        setError(QString("Exception processing sheet '%1': %2")
                     .arg(sheetName, QString::fromLatin1(ex.what())));
        return empty;
    } catch (...) {
        setError(QString("Unknown exception processing sheet '%1'").arg(sheetName));
        return empty;
    }

    logDebug(QString("Sheet '%1' done: %2 sample(s), overallAvgTPM=%3")
                 .arg(sheetName)
                 .arg(result.samples.size())
                 .arg(result.overallAvgTPM, 0, 'f', 4));

    return result;
}

// ===========================================================================
// processSopSheet — SOP / instruction sheet: raw table, no plot.
// Shared by processSheet (production) and processSheetLegacy; behavior unchanged.
// ===========================================================================

SheetResult DataProcessor::processSopSheet(ExcelReader& reader, const QString& sheetName)
{
    SheetResult sopResult;
    sopResult.sheetName       = sheetName;
    sopResult.templateVersion = reader.detectTemplateVersion();
    sopResult.isRawTable      = true;

    QVector<QStringList> rows = reader.getSheetRows();
    if (!rows.isEmpty()) {
        // First non-empty row becomes the column headers.
        int headerIdx = 0;
        for (int i = 0; i < rows.size(); ++i) {
            bool hasContent = false;
            for (const QString& cell : rows[i]) {
                if (!cell.trimmed().isEmpty()) { hasContent = true; break; }
            }
            if (hasContent) { headerIdx = i; break; }
        }
        sopResult.rawHeaders = rows[headerIdx];
        for (int i = headerIdx + 1; i < rows.size(); ++i)
            sopResult.rawRows.append(rows[i]);
    }
    return sopResult;
}

// ===========================================================================
// processFile — production read path (schema-driven)
//
// Load / sheet-iteration / FileResult assembly identical to processFileLegacy,
// with processSheet doing schema-driven extraction (SchemaDrivenReader +
// LegacyAdapter). The Phase-1 shadow gate (tst_v3shadow) proved this
// byte-identical to processFileLegacy over the whole fixture corpus.
// ===========================================================================

FileResult DataProcessor::processFile(
    const QString& filePath,
    std::function<void(int, const QString&)> progressCallback)
{
    m_lastError.clear();
    // Reset before the sheet loop; set true by any sheet that takes the
    // inference path (see below). Exposed via lastFileUsedInference().
    m_lastUsedInference = false;
    // Reset alongside; set true when a _dve_schema manifest sheet is found
    // and parsed below. Exposed via lastFileHadManifest().
    m_lastHadManifest = false;

    FileResult result;
    result.filePath = filePath;
    result.fileName = QFileInfo(filePath).fileName();

    auto notify = [&](int pct, const QString& msg) {
        logDebug(QString("(%1%) %2").arg(pct).arg(msg));
        if (progressCallback)
            progressCallback(pct, msg);
    };

    notify(0, QStringLiteral("Opening file: ") + result.fileName);

    ExcelReader reader;
    if (!reader.loadFile(filePath)) {
        setError(QStringLiteral("Failed to open file: ") + reader.getLastError());
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();  // signal failure to onFileLoadFinished
        return result;
    }

    QString templateVersion = reader.detectTemplateVersion();
    result.templateVersion  = templateVersion;
    logDebug(QStringLiteral("Template version: ") + templateVersion);

    QStringList sheetNames = reader.getSheetNames();

    // ── Manifest (_dve_schema): a veryHidden internal sheet declaring the
    //    template. Parse it BEFORE the sheet loop (it feeds the resolver's
    //    rung 1) and exclude it from sheetNames - it must never appear in the
    //    navigator or as a data sheet. Warnings are poka-yoke (registry rule
    //    5): log and proceed, never abort. ──
    model::Manifest::ParseResult manifest;
    bool hasManifest = false;
    const QString manifestSheet = manifestSheetIn(sheetNames);
    if (!manifestSheet.isEmpty()) {
        sheetNames.removeOne(manifestSheet);
        if (reader.selectSheet(manifestSheet)) {
            manifest    = model::Manifest::parse(reader.currentSheetCells());
            hasManifest = true;
            m_lastHadManifest = true;
            for (const QString& w : manifest.warnings)
                qWarning() << "[DataProcessor] manifest:" << w;
        } else {
            qWarning() << "[DataProcessor] Could not select manifest sheet:"
                       << manifestSheet << "-" << reader.getLastError();
        }
    }

    result.sheetNames = sheetNames;

    if (sheetNames.isEmpty()) {
        setError(QStringLiteral("No sheets found in file: ") + filePath);
        notify(100, QStringLiteral("Error: ") + m_lastError);
        result.filePath.clear();
        return result;
    }

    notify(5, QString("Found %1 sheet(s)").arg(sheetNames.size()));

    // True when the resolver's rung 1 will route this sheet (mirrors
    // processSheet: a manifest block hit that is not the SOP raw-table
    // branch, which keeps precedence over the resolver).
    auto manifestRoutes = [&](const QString& name) {
        return hasManifest
            && !name.contains(QStringLiteral("SOP"), Qt::CaseInsensitive)
            && model::Manifest::blockForSheet(manifest, name) != nullptr;
    };

    const int nSheets = sheetNames.size();

    for (int idx = 0; idx < nSheets; ++idx) {
        const QString& sheetName = sheetNames[idx];

        int progressBase = 5 + static_cast<int>(90.0 * idx / nSheets);
        notify(progressBase, QStringLiteral("Processing sheet: ") + sheetName);

        if (!reader.selectSheet(sheetName)) {
            qWarning() << "[DataProcessor] Could not select sheet:" << sheetName
                       << "-" << reader.getLastError();
            continue;
        }

        SheetResult sheetResult = processSheet(reader, sheetName,
                                               hasManifest ? &manifest : nullptr);
        sheetResult.templateVersion = templateVersion;
        // W1 poka-yoke, WITHDRAWN as a load gate by W3b (2026-08-31):
        // SheetResult::strippedFormulaCells stays 0 on every routing fork.
        // App-template-lineage workbooks are byte-indistinguishable from
        // old-write-back wrecks (the bundled template is itself openpyxl-born
        // with zero caches and pre-filled formula chains), so any load-time
        // gate misfires on legitimate files - v2.10.7 shipped one and broke
        // the New File lineage. Wreck protection lives at the ROOT instead:
        // the surgical write-back preserves caches (tst_excelsurgery /
        // tst_excelwritepayload gates). See processSheet's standard-fork
        // comment for the full evidence. Transient - never serialized.

        if (sheetResult.fromInferredSchema) {
            m_lastUsedInference = true;
            // columnHeaders were already set by the inference lowering to the
            // sheet's real (13/8-wide) block-1 header texts; don't clobber them
            // with getColumnHeaders()'s fixed 12-wide positional read.
        } else if (!manifestRoutes(sheetName)) {
            // Capture column headers (populated after selectSheet).
            // Manifest-routed sheets skip this too: the key-based lowering
            // already recorded the schema's display names.
            sheetResult.columnHeaders = reader.getColumnHeaders();
        }

        result.sheets.append(sheetResult);
    }

    notify(100, QStringLiteral("Done. Processed ")
               + QString::number(result.sheets.size())
               + QStringLiteral(" sheet(s)."));

    return result;
}

// ===========================================================================
// processSheet — production single-sheet read (schema-driven)
//
// Extraction goes through model::SchemaDrivenReader (positional resolution, to
// mirror ExcelReader::extractRow byte-for-byte) + model::LegacyAdapter, then the
// UNCHANGED SheetProcessor chain. Header/metadata for each block is resolved by
// sniffing the same landmark cells ExcelReader::extractMetadata sniffs and
// applying the matching builtin header layout (Standard / Cart / Project).
// ===========================================================================

namespace {

// Read a header-field set for one block into a raw headers map, exactly like
// SchemaDrivenReader::parseSheet does (raw QVariant cells, block-relative
// 1-based row/col). Used to (re)build a block's headers from a chosen layout.
QMap<QString, QVariant> extractBlockHeaders(const QVector<QVector<QVariant>>& cells,
                                            int off,
                                            const QVector<model::HeaderFieldDef>& fields)
{
    auto cellAt = [&](int r, int c) -> QVariant {
        if (r < 0 || r >= cells.size() || c < 0 || c >= cells[r].size()) return {};
        return cells[r][c];
    };
    QMap<QString, QVariant> headers;
    for (const model::HeaderFieldDef& h : fields)
        headers.insert(h.key, cellAt(h.row - 1, off + h.col - 1));
    return headers;
}

} // namespace

SheetResult DataProcessor::processSheet(ExcelReader& reader, const QString& sheetName,
                                        const model::Manifest::ParseResult* manifest)
{
    logDebug(QStringLiteral("processSheet: ") + sheetName);

    SheetResult empty;
    empty.sheetName = sheetName;

    // ── SOP / instruction sheet — identical raw-table branch as legacy.
    //    Precedence over the resolver: manifests do not describe SOP sheets. ──
    if (sheetName.contains(QStringLiteral("SOP"), Qt::CaseInsensitive))
        return processSopSheet(reader, sheetName);

    const QVector<QVector<QVariant>> cells = reader.currentSheetCells();
    if (cells.isEmpty())
        return empty;

    const QString templateVersion = reader.detectTemplateVersion();

    // ── Routing ladder: manifest -> compiled standard -> header-driven
    //    inference, unified in model::SchemaResolver (the standardFits fork,
    //    the perRowRegime derivation, and the per-block Cart/Project landmark
    //    sniff moved there verbatim). `manifest` is the workbook's parsed
    //    _dve_schema (processFile reads it before the sheet loop); nullptr
    //    keeps the standard/inference ladder identical by construction. ──
    const model::SchemaResolver::Resolution res =
        model::SchemaResolver::resolve(manifest, sheetName, cells, templateVersion);

    if (res.source == model::SchemaResolver::Source::Manifest) {
        // Manifest sheets parse with the DECLARED schema + NameFirst column
        // resolution (reordered/renamed columns track by header text) and
        // lower through the generalized key-based lowering, which records
        // slot-ordered write provenance (columnKeys by resolved physical
        // slot, headerCells from schema.headerFields). The standard-path
        // provenance recording below is deliberately NOT reached - running
        // it too would clobber the slot order with schema order.
        const model::Sheet mMan = model::SchemaDrivenReader::parseSheet(
            cells, sheetName, res.schema, res.perRowRegime,
            res.columnResolution);
        if (mMan.samples.isEmpty())
            return empty;   // blank sheets stay non-errors
        // W3b: manifest sheets are OUR template's lineage by construction
        // (only the app's v3 template carries _dve_schema), so a cache-less
        // workbook here is the template's normal openpyxl-born state, not
        // destruction - strippedFormulaCells stays 0 (the default).
        return model::LegacyAdapter::lowerSchemaSheet(
            mMan, sheetName, templateVersion,
            /*fromInference=*/false, res.perRowRegime);
    }

    if (res.source == model::SchemaResolver::Source::Inference) {
        const model::Sheet mInf = model::SchemaDrivenReader::parseSheet(
            cells, sheetName, res.schema, /*perRowRegime=*/false,
            res.columnResolution);
        if (mInf.samples.isEmpty())
            return empty;   // blank sheets stay non-errors
        // W3b: no stripped-count gate here either - historical Cart-era
        // files the app edited for years are zero-cache AND inference-routed,
        // so the same indistinguishability argument applies (see the
        // standard-fork comment below). The lowering's parameter stays at
        // its dormant default.
        return model::LegacyAdapter::lowerInferredSheet(
            mInf, sheetName, templateVersion);
    }

    const bool perRowRegime = res.perRowRegime;

    // Data columns: positional resolution reproduces ExcelReader::extractRow's
    // fixed 12-wide-by-position read (name-first matching would shuffle data on
    // the historical Cart / 13-wide layouts). Header fields from the Standard
    // layout are extracted here too; cart/project blocks get theirs replaced
    // below.
    const model::TemplateSchema& schema = res.schema;
    model::Sheet mSheet = model::SchemaDrivenReader::parseSheet(
        cells, sheetName, schema, perRowRegime, res.columnResolution);
    if (mSheet.samples.isEmpty())
        return empty;

    // ── Per-block metadata layout: the resolver sniffed the SAME landmark
    //    cells ExcelReader::extractMetadata sniffs and chose each block's
    //    header layout (res.blockLayouts lines up 1:1 with the samples -
    //    both sides derive the block count from header-row width / blockCols).
    //    APPLY the choice here, where the parsed data lives: Cart / Project
    //    fully replace the block's headers; a Standard "old"-template block
    //    drops Heating Technology (extractMetadata's old branch never reads
    //    it). ──
    const model::TemplateSchema cartSchema =
        model::standardV1(perRowRegime, model::HeaderLayout::Cart);
    const model::TemplateSchema projectSchema =
        model::standardV1(perRowRegime, model::HeaderLayout::Project);

    // Resolved header layout per block - also consumed by the write-provenance
    // recording below (the header-cell map is sheet-level, so it needs to
    // know which layout the blocks actually used).
    const QVector<model::HeaderLayout>& blockLayouts = res.blockLayouts;

    for (int b = 0; b < mSheet.samples.size(); ++b) {
        const int off = b * schema.blockCols;

        model::Sample& sample = mSheet.samples[b];
        if (blockLayouts[b] == model::HeaderLayout::Cart) {
            sample.headers = extractBlockHeaders(cells, off, cartSchema.headerFields);
        } else if (blockLayouts[b] == model::HeaderLayout::Project) {
            sample.headers = extractBlockHeaders(cells, off, projectSchema.headerFields);
            // Assemble sampleID = "<project> <suffix>" (space, not dash), exactly
            // as extractMetadata does; suffix already carries its own hyphen.
            const QString project = sample.headers.value(QStringLiteral("project_name")).toString().trimmed();
            const QString suffix  = sample.headers.value(QStringLiteral("sample_suffix")).toString().trimmed();
            sample.headers.insert(QStringLiteral("sample_id"),
                                  suffix.isEmpty() ? project : project + QStringLiteral(" ") + suffix);
        } else if (templateVersion != QLatin1String("new")) {
            // Mirror extractMetadata's else-structure: only the "new" branch
            // reads Heating Technology. Every other standardized block (the
            // "old" template, and the theoretical "unknown") assigns none, so
            // drop the key here.
            sample.headers.remove(QStringLiteral("heating_technology"));
        }
    }

    // ── Lower each v3 Sample to the legacy raw shape, then run the UNCHANGED
    //    SheetProcessor chain. ──
    QVector<ExcelReader::SampleData> rawSamples;
    rawSamples.reserve(mSheet.samples.size());
    for (int i = 0; i < mSheet.samples.size(); ++i)
        rawSamples.append(model::LegacyAdapter::lowerSample(mSheet.samples[i], schema, i));

    std::unique_ptr<SheetProcessor> processor(createProcessor(sheetName));
    processor->setPerRowRegime(perRowRegime);
    // W3b (2026-08-31): the reader's stripped-cache count deliberately does
    // NOT gate behavior here. The bundled New File template is openpyxl-born
    // (zero formula caches by construction, =prev+K puff chains and =C5
    // before-weight chains pre-filled), so every app-template-lineage
    // workbook is BYTE-INDISTINGUISHABLE from a workbook the old openpyxl
    // write-back destroyed - same standard layout, same formula shapes, same
    // zero-cache state. Any load-time gate therefore misfires on one class
    // or the other (v2.10.7 shipped the gate and broke the New File lineage:
    // warning dialog + DB-save block + disabled repairs on the app's own
    // files). The repairs below ARE the correct load path for that lineage;
    // wreck creation itself is prevented at the root (the surgical write-back
    // preserves caches - tst_excelsurgery / tst_excelwritepayload gates), and
    // one Excel save re-caches any historical wreck. The gate machinery
    // (setStrippedFormulaCells / lowerSchemaSheet's parameter) stays dormant
    // and unit-pinned for a future provenance-stamped classifier.

    SheetResult result;
    try {
        result = processor->process(rawSamples, sheetName, templateVersion);
    } catch (const std::exception& ex) {
        setError(QString("Exception processing sheet '%1': %2")
                     .arg(sheetName, QString::fromLatin1(ex.what())));
        return empty;
    } catch (...) {
        setError(QString("Unknown exception processing sheet '%1'").arg(sheetName));
        return empty;
    }

    // ── Write provenance (Phase 2b): record where this sheet's cells came
    //    from so write-back derives addresses instead of assuming the 12-wide
    //    standardized layout (which silently corrupts Cart/Project header
    //    cells). Per-sample startColumn was copied from SampleData in
    //    SheetProcessors::buildSampleResult (the reader's own block origin).
    //    Data-column geometry comes from the standard schema; header VALUE
    //    cells come from the layout the per-block sniff above resolved. If
    //    blocks ever disagreed on layout, no single sheet-level header map
    //    would be right - leave headerCells empty so header edits are
    //    rejected instead of landing in wrong cells. Raw-table (SOP) and
    //    blank sheets return earlier and keep the no-provenance defaults. ──
    result.blockCols    = schema.blockCols;
    result.dataStartRow = schema.dataStartRow;
    for (const model::MetricDef& c : schema.columns)
        result.columnKeys.append(c.key);

    bool uniformLayout = true;
    for (int b = 1; b < blockLayouts.size(); ++b) {
        if (blockLayouts[b] != blockLayouts[0]) { uniformLayout = false; break; }
    }
    if (uniformLayout && !blockLayouts.isEmpty()) {
        const model::TemplateSchema* headerSchema = &schema;
        if (blockLayouts[0] == model::HeaderLayout::Cart)
            headerSchema = &cartSchema;
        else if (blockLayouts[0] == model::HeaderLayout::Project)
            headerSchema = &projectSchema;
        for (const model::HeaderFieldDef& h : headerSchema->headerFields)
            result.headerCells.insert(h.key, QPoint(h.col, h.row));
    }

    logDebug(QString("Sheet '%1' done: %2 sample(s), overallAvgTPM=%3")
                 .arg(sheetName)
                 .arg(result.samples.size())
                 .arg(result.overallAvgTPM, 0, 'f', 4));

    return result;
}

} // namespace DVE
