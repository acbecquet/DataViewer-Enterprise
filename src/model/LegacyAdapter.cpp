#include "LegacyAdapter.h"
#include "MetricRegistry.h"
#include "../pipeline/HeatingTech.h"
#include "../pipeline/SheetProcessors.h"
#include <QSet>

namespace DVE { namespace model {

ExcelReader::SampleData LegacyAdapter::lowerSample(const Sample& s,
                                                   const TemplateSchema& schema,
                                                   int blockIndex)
{
    ExcelReader::SampleData out;
    auto str = [&](const char* k){ return s.headers.value(QLatin1String(k)).toString().trimmed(); };
    auto num = [&](const char* k){ return ExcelReader::tolerantCellDouble(s.headers.value(QLatin1String(k))); };

    // Mirrors ExcelReader::extractMetadata's standardised ("new") layout branch
    // field-by-field: text fields go through toString().trimmed() (== getCellString),
    // numeric fields through tolerantCellDouble (== getCellDouble).
    out.metadata.testName          = str("test_name");
    out.metadata.date              = str("date");
    out.metadata.sampleID          = str("sample_id");
    out.metadata.heatingTechnology = str("heating_technology");
    out.metadata.media             = str("media");
    out.metadata.resistance        = num("resistance");
    out.metadata.puffingRegime     = str("puffing_regime");
    out.metadata.viscosity         = num("viscosity");
    out.metadata.tester            = str("tester");
    out.metadata.voltage           = num("voltage");
    out.metadata.initialOilMass    = num("initial_oil_mass");

    // Power: P = V^2 / (R + tech offset), guarded exactly as extractMetadata
    // guards it (voltage AND denominator must both be positive) - metadata.power
    // stays its default-constructed 0.0 otherwise.
    const double denom = out.metadata.resistance
        + DVE::heatingTechResistanceOffset(out.metadata.heatingTechnology);
    if (out.metadata.voltage > 0 && denom > 0)
        out.metadata.power = (out.metadata.voltage * out.metadata.voltage) / denom;

    out.startColumn = blockIndex * schema.blockCols;

    const int rows = s.rowCount();
    out.dataRows.resize(rows);
    for (int r = 0; r < rows; ++r) {
        QVector<QVariant> row(schema.blockCols);
        for (int c = 0; c < s.data.size() && c < schema.blockCols; ++c)
            row[c] = (r < s.data[c].values.size()) ? s.data[c].values[r] : QVariant();
        out.dataRows[r] = row;
    }
    return out;
}

// ===========================================================================
// lowerSchemaSheet — schema-described sheet (inference or manifest) ->
// SheetResult; lowerInferredSheet forwards here with the pre-2c defaults.
// ===========================================================================

namespace {

// Per-row coercions IDENTICAL to SheetProcessor::buildSampleResult: numeric
// fields via plain QVariant::toDouble (NOT unit-tolerant - that is only for the
// header band), string fields trimmed.
double rowDouble(const QVariant& v) { bool ok = false; const double d = v.toDouble(&ok); return ok ? d : 0.0; }
QString rowString(const QVariant& v) { return v.toString().trimmed(); }
bool cellPresent(const QVariant& v) { return v.isValid() && !v.toString().trimmed().isEmpty(); }

// Column keys that map to a fixed DataRow field OR are derived (recomputed) -
// everything NOT in this set is an open per-row metric that rides in
// DataRow::extra.
const QSet<QString>& fixedOrDerivedColumnKeys()
{
    static const QSet<QString> k = {
        QStringLiteral("puffs"), QStringLiteral("before_weight"), QStringLiteral("after_weight"),
        QStringLiteral("draw_pressure"), QStringLiteral("resistance"), QStringLiteral("puffing_regime"),
        QStringLiteral("smell"), QStringLiteral("clog"), QStringLiteral("notes"),
        QStringLiteral("tpm"), QStringLiteral("tpm_power_density"),
        QStringLiteral("variation_tpm"), QStringLiteral("oil_consumed")
    };
    return k;
}

// Header keys that map to a fixed SampleResult member - everything else rides
// in SampleResult::extra.
const QSet<QString>& fixedHeaderKeys()
{
    static const QSet<QString> k = {
        QStringLiteral("test_name"), QStringLiteral("date"), QStringLiteral("sample_id"),
        QStringLiteral("heating_technology"), QStringLiteral("media"), QStringLiteral("resistance"),
        QStringLiteral("power"), QStringLiteral("puffing_regime"), QStringLiteral("viscosity"),
        QStringLiteral("tester"), QStringLiteral("voltage"), QStringLiteral("initial_oil_mass")
    };
    return k;
}

// buildSampleResult's burn/clog/leak normaliser (strategy 2, notes scan).
QString normaliseBCL(const QString& raw)
{
    const QString u = raw.trimmed().toUpper();
    if (u == QLatin1String("Y") || u == QLatin1String("YES")) return QStringLiteral("Y");
    if (u == QLatin1String("N") || u == QLatin1String("NO"))  return QStringLiteral("N");
    return QString();
}

} // namespace

SheetResult LegacyAdapter::lowerSchemaSheet(const Sheet& sheet,
                                            const QString& sheetName,
                                            const QString& templateVersion,
                                            bool fromInference,
                                            bool perRowRegime,
                                            int strippedFormulaCells)
{
    SheetResult result;
    result.sheetName            = sheetName;
    result.templateVersion      = templateVersion;
    result.hasPerRowRegime      = perRowRegime;
    result.fromInferredSchema   = fromInference;
    result.strippedFormulaCells = strippedFormulaCells;

    // GenericSheetProcessor owns the metric recompute; every concrete processor
    // delegates to it, so the sheet-name-based factory choice is irrelevant here.
    GenericSheetProcessor proc;
    proc.setPerRowRegime(perRowRegime);

    for (int si = 0; si < sheet.samples.size(); ++si) {
        const Sample& m = sheet.samples[si];
        SampleResult sr;

        // Write provenance (Phase 2b): parseSheet reads block si at physical
        // origin si * blockCols - record it so write-back can derive addresses
        // on these (13/8-wide) layouts too.
        sr.startColumn = si * sheet.schema.blockCols;

        // ── Header band -> SampleResult members. Text via toString().trimmed()
        //    (== getCellString), numeric via tolerantCellDouble (== getCellDouble),
        //    mirroring ExcelReader::extractMetadata / LegacyAdapter::lowerSample. ──
        auto hstr = [&](const char* k){ return m.headers.value(QLatin1String(k)).toString().trimmed(); };
        auto hnum = [&](const char* k){ return ExcelReader::tolerantCellDouble(m.headers.value(QLatin1String(k))); };

        sr.sampleID          = hstr("sample_id");
        // sampleName fallback identical to buildSampleResult.
        sr.sampleName        = sr.sampleID.isEmpty()
                               ? QStringLiteral("Sample %1").arg(si + 1)
                               : sr.sampleID;
        sr.date              = hstr("date");
        sr.tester            = hstr("tester");
        sr.media             = hstr("media");
        sr.viscosity         = hnum("viscosity");
        // Resistance: a plain "Resistance:" label wins when present. Cart-era
        // sheets label the initial reading "Ri (Ohms)", which the registry
        // canonicalizes to resistance_initial (D11) - that reading IS the
        // sample resistance in the legacy model, so project it into the fixed
        // member (it also rides through SampleResult::extra under its own key).
        sr.resistance        = hnum("resistance");
        if (sr.resistance == 0.0)
            sr.resistance    = hnum("resistance_initial");
        sr.voltage           = hnum("voltage");
        sr.heatingTechnology = hstr("heating_technology");
        sr.puffingRegime     = hstr("puffing_regime");
        sr.initialOilMass    = hnum("initial_oil_mass");

        // Power: a present numeric header value wins (UserSim carries a computed
        // Power cell); else derive P = V^2 / (R + tech offset) exactly as
        // extractMetadata (guarded V>0 && denom>0); else stays 0.0 so
        // calculateMetrics yields clean-zero power density / normalized TPM
        // (S26 shape offers neither a power value nor a derivable one).
        double power = 0.0;
        const QVariant powerVar = m.headers.value(QStringLiteral("power"));
        if (cellPresent(powerVar)) {
            const double pv = ExcelReader::tolerantCellDouble(powerVar);
            if (pv > 0.0) power = pv;
        }
        if (power == 0.0) {
            const double denom = sr.resistance
                + DVE::heatingTechResistanceOffset(sr.heatingTechnology);
            if (sr.voltage > 0.0 && denom > 0.0)
                power = (sr.voltage * sr.voltage) / denom;
        }
        sr.power = power;

        // Unknown header keys -> SampleResult::extra (skip empties / known keys).
        for (auto it = m.headers.constBegin(); it != m.headers.constEnd(); ++it) {
            if (fixedHeaderKeys().contains(it.key())) continue;
            if (!cellPresent(it.value())) continue;
            sr.extra.insert(it.key(), it.value());
        }

        // ── Data rows: known metric keys -> fixed DataRow fields; unknown -> extra.
        //    Derived columns from the sheet are ignored (recomputed below). ──
        const int rows = m.rowCount();
        sr.rows.reserve(rows);
        for (int r = 0; r < rows; ++r) {
            DataRow dr;
            // PV1..PV5 -> draw_pressure_per_puff list (registry D5/8.2): buffer
            // per-puff aliases by element index; everything else rides through
            // under its own key as before.
            QVector<QVariant> perPuff(5);
            bool anyPerPuff = false;
            for (const MetricSeries& ser : m.data) {
                const QVariant v = (r < ser.values.size()) ? ser.values[r] : QVariant();
                const QString& key = ser.key;
                if      (key == QLatin1String("puffs"))          dr.puffs        = rowDouble(v);
                else if (key == QLatin1String("before_weight"))  dr.beforeWeight = rowDouble(v);
                else if (key == QLatin1String("after_weight"))   dr.afterWeight  = rowDouble(v);
                else if (key == QLatin1String("draw_pressure"))  dr.drawPressure = rowDouble(v);
                else if (key == QLatin1String("resistance"))     dr.resistance   = rowDouble(v);
                else if (key == QLatin1String("puffing_regime")) dr.puffingRegime = rowString(v);
                else if (key == QLatin1String("smell"))          dr.smell        = rowString(v);
                else if (key == QLatin1String("clog"))           dr.clog         = rowString(v);
                else if (key == QLatin1String("notes"))          dr.notes        = rowString(v);
                else if (fixedOrDerivedColumnKeys().contains(key)) { /* derived: recomputed */ }
                else if (cellPresent(v)) {
                    const MetricRegistry::PerPuffAlias pp = MetricRegistry::perPuffAlias(
                        SchemaDrivenReader::normalizeHeader(key));
                    if (pp.index > 0) {
                        // Coerce at assembly so the in-memory list already has
                        // the declared NumberList element type - a text cell
                        // ("N/A") must not differ before vs after a recovery
                        // round-trip through the {"a"} envelope's toDouble().
                        perPuff[pp.index - 1] = QVariant(rowDouble(v));
                        anyPerPuff = true;
                    } else {
                        dr.extra.insert(key, v);
                    }
                }
            }
            if (anyPerPuff) {
                // Trim unused TRAILING slots; interior gaps keep their position
                // as 0.0 so element i is always puff i+1.
                int last = perPuff.size();
                while (last > 0 && !perPuff[last - 1].isValid()) --last;
                QVariantList list;
                for (int i = 0; i < last; ++i)
                    list.append(perPuff[i].isValid() ? perPuff[i] : QVariant(0.0));
                dr.extra.insert(QStringLiteral("draw_pressure_per_puff"), list);
            }
            sr.rows.append(dr);
        }

        // ── Repair rules identical to buildSampleResult: openpyxl strips cached
        //    formula values, so fill-forward before-weight and extrapolate puffs. ──
        for (int i = 1; i < sr.rows.size(); ++i)
            if (sr.rows[i].beforeWeight == 0.0 && sr.rows[i - 1].afterWeight > 0.0)
                sr.rows[i].beforeWeight = sr.rows[i - 1].afterWeight;

        // W3b: NOT when the reader flagged this sheet's caches as destroyed -
        // extrapolating those zeros fabricates plausible puff chains (the
        // v2.10.6 smoke's +30 chains). Mirrors buildSampleResult's gate.
        if (strippedFormulaCells == 0
            && !sr.rows.isEmpty() && sr.rows[0].puffs > 0.0) {
            double increment = sr.rows[0].puffs;
            for (int i = 1; i < sr.rows.size(); ++i) {
                if (sr.rows[i].puffs > 0.0) {
                    const double diff = sr.rows[i].puffs - sr.rows[i - 1].puffs;
                    if (diff > 0.0) increment = diff;
                    break;
                }
            }
            for (int i = 1; i < sr.rows.size(); ++i)
                if (sr.rows[i].puffs == 0.0)
                    sr.rows[i].puffs = sr.rows[i - 1].puffs + increment;
        }

        // ── Burn / clog / leak: notes-keyword scan (buildSampleResult strategy 2;
        //    strategy 1's raw cols 8-10 Y/N flags don't exist on these layouts). ──
        QString burnVal, clogVal, leakVal;
        for (const DataRow& dr : sr.rows) {
            const QString n = dr.notes.toUpper();
            if (n.contains(QLatin1String("BURN"))) burnVal = n.contains(QLatin1String("NO BURN")) ? "N" : "Y";
            if (n.contains(QLatin1String("CLOG"))) clogVal = n.contains(QLatin1String("NO CLOG")) ? "N" : "Y";
            if (n.contains(QLatin1String("LEAK"))) leakVal = n.contains(QLatin1String("NO LEAK")) ? "N" : "Y";
        }
        sr.burnStatus = normaliseBCL(burnVal);
        sr.clogStatus = normaliseBCL(clogVal);
        sr.leakStatus = normaliseBCL(leakVal);

        proc.calculateMetrics(sr);
        result.samples.append(sr);
    }

    proc.computeSheetAggregates(result);

    // Column headers = the schema's displayNames in SCHEMA order (physical
    // order on identity resolutions, i.e. every inference sheet). For
    // UNMATCHED inferred columns that is the sheet's own row-4 text; for
    // registry-matched columns it is the CANONICAL display name (e.g.
    // "Failure" instead of "Failure? (if yes, when)") - deliberate under
    // naming-policy rule 6: display names are presentation hints, vocabulary
    // is normalized.
    for (const MetricDef& c : sheet.schema.columns)
        result.columnHeaders.append(c.displayName);

    // Write provenance (Phase 2b): sheet-level geometry + vocabulary from the
    // schema. HeaderFieldDef.row/col already point at the VALUE cell
    // (1-based, block-relative) - same QPoint(x=col, y=row) convention as the
    // standard path (see SheetResult::headerCells in ReportData.h).
    result.blockCols    = sheet.schema.blockCols;
    result.dataStartRow = sheet.schema.dataStartRow;

    // columnKeys in PHYSICAL slot order (Phase 2c): invert Sheet::columnSlots
    // (keysBySlot[slot] = key of the schema column resolved onto that slot) so
    // CellAddress::dataCell - which maps key -> columnKeys index -> physical
    // column - derives correct addresses on NameFirst sheets with reordered
    // columns. Inference schemas resolve identity, so their recorded keys are
    // byte-unchanged; a Sheet built without parseSheet (empty columnSlots) is
    // treated as identity too.
    const int nCols = int(sheet.schema.columns.size());
    QVector<int> slotOf = sheet.columnSlots;   // "slots" is a Qt keyword macro
    if (slotOf.size() != nCols) {
        slotOf.resize(nCols);
        for (int i = 0; i < nCols; ++i) slotOf[i] = i;
    }
    int slotSpan = nCols;
    for (int s : slotOf) slotSpan = qMax(slotSpan, s + 1);
    QVector<QString> keysBySlot(slotSpan);
    for (int i = 0; i < nCols; ++i) {
        const int s = slotOf[i];
        if (s >= 0 && keysBySlot[s].isEmpty())
            keysBySlot[s] = sheet.schema.columns[i].key;   // first claim wins
    }
    // Unclaimed slots (schema narrower than the block, or a name-match
    // collision dropped a claim - see the resolveColumns pass-2 note): fall
    // back to the schema's own positional key at that slot when that key is
    // not already claimed elsewhere; otherwise the slot stays keyless ("") so
    // CellAddress lookups reject instead of mis-addressing.
    QSet<QString> claimed;
    for (const QString& k : keysBySlot)
        if (!k.isEmpty()) claimed.insert(k);
    for (int s = 0; s < slotSpan; ++s)
        if (keysBySlot[s].isEmpty() && s < nCols
            && !claimed.contains(sheet.schema.columns[s].key))
            keysBySlot[s] = sheet.schema.columns[s].key;
    for (const QString& k : keysBySlot)
        result.columnKeys.append(k);

    for (const HeaderFieldDef& h : sheet.schema.headerFields)
        result.headerCells.insert(h.key, QPoint(h.col, h.row));

    return result;
}

SheetResult LegacyAdapter::lowerInferredSheet(const Sheet& sheet,
                                              const QString& sheetName,
                                              const QString& templateVersion,
                                              int strippedFormulaCells)
{
    return lowerSchemaSheet(sheet, sheetName, templateVersion,
                            /*fromInference=*/true, /*perRowRegime=*/false,
                            strippedFormulaCells);
}

}} // namespace DVE::model
