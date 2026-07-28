#include "StandardSchema.h"
#include "MetricRegistry.h"

namespace DVE { namespace model {

namespace {

// Copy a registry def and apply the standard layout's presentation flags
// (flags are layout policy, not vocabulary - same rules as before).
MetricDef col(const QString& key)
{
    const MetricDef* d = MetricRegistry::metric(key);
    Q_ASSERT(d);
    MetricDef m = *d;
    m.editable  = (m.role == Role::Qualitative);
    m.plottable = (m.key == QLatin1String("tpm"));
    return m;
}

HeaderFieldDef hf(const QString& key, const QString& displayName, ValueType type, int row, int col,
                  const QString& unit = QString(), const QString& calculator = QString(),
                  const QStringList& inputs = QStringList())
{
    HeaderFieldDef h;
    h.key         = key;
    h.displayName = displayName;
    h.type        = type;
    h.unit        = unit;
    h.row         = row;
    h.col         = col;
    h.calculator  = calculator;
    h.inputs      = inputs;
    return h;
}

AggregateDef agg(const QString& key, const QString& calculator, const QStringList& inputs)
{
    AggregateDef a;
    a.key        = key;
    a.calculator = calculator;
    a.inputs     = inputs;
    return a;
}

} // namespace

TemplateSchema standardV1(bool perRowRegime, HeaderLayout layout)
{
    TemplateSchema s;
    s.schemaId       = QStringLiteral("standard");
    s.version        = 1;
    s.headerRows      = 3;
    s.columnHeaderRow = 4;
    s.dataStartRow    = 5;
    s.blockCols       = 12;

    // ── Columns (physical column order == legacy DVE::Cols order; defs are
    //    registry copies - vocabulary lives in MetricRegistry, order here) ──
    s.columns.append(col(QStringLiteral("puffs")));
    s.columns.append(col(QStringLiteral("before_weight")));
    s.columns.append(col(QStringLiteral("after_weight")));
    s.columns.append(col(QStringLiteral("draw_pressure")));
    s.columns.append(col(perRowRegime ? QStringLiteral("puffing_regime")
                                      : QStringLiteral("resistance")));
    s.columns.append(col(QStringLiteral("smell")));
    s.columns.append(col(QStringLiteral("clog")));
    s.columns.append(col(QStringLiteral("notes")));
    s.columns.append(col(QStringLiteral("tpm")));
    s.columns.append(col(QStringLiteral("tpm_power_density")));
    s.columns.append(col(QStringLiteral("variation_tpm")));
    s.columns.append(col(QStringLiteral("oil_consumed")));

    // ── Header-band fields (block-relative, 1-based) ──
    // Each layout mirrors one branch of ExcelReader::extractMetadata cell-for-cell
    // (verified against src/ExcelReader.cpp). Fields absent from a layout lower to
    // "" / 0 in LegacyAdapter, exactly as extractMetadata leaves them unassigned.
    if (layout == HeaderLayout::Cart) {
        // isCartFormat branch: sample id / resistance / viscosity on row 2;
        // media / puff-regime / voltage on row 3. No test name / date / tester /
        // heating tech / initial oil mass.
        s.headerFields.append(hf(QStringLiteral("sample_id"), QStringLiteral("Sample ID"), ValueType::Text, 2, 2));
        s.headerFields.append(hf(QStringLiteral("resistance"), QStringLiteral("Resistance"), ValueType::Number, 2, 4, QStringLiteral("ohm")));
        s.headerFields.append(hf(QStringLiteral("viscosity"), QStringLiteral("Viscosity"), ValueType::Number, 2, 8, QStringLiteral("cP")));
        s.headerFields.append(hf(QStringLiteral("media"), QStringLiteral("Media"), ValueType::Text, 3, 2));
        s.headerFields.append(hf(QStringLiteral("puffing_regime"), QStringLiteral("Puffing Regime"), ValueType::Text, 3, 6));
        s.headerFields.append(hf(QStringLiteral("voltage"), QStringLiteral("Voltage"), ValueType::Number, 3, 8, QStringLiteral("V")));
    } else if (layout == HeaderLayout::Project) {
        // isProjectFormat branch: sample id is assembled from project_name +
        // sample_suffix (joined in processSheet); date + tester on row 1. No
        // test name / heating tech / initial oil mass.
        s.headerFields.append(hf(QStringLiteral("project_name"), QStringLiteral("Project"), ValueType::Text, 1, 7));
        s.headerFields.append(hf(QStringLiteral("sample_suffix"), QStringLiteral("Sample"), ValueType::Text, 1, 9));
        s.headerFields.append(hf(QStringLiteral("date"), QStringLiteral("Date"), ValueType::Text, 1, 3));
        s.headerFields.append(hf(QStringLiteral("tester"), QStringLiteral("Tester"), ValueType::Text, 1, 5));
        s.headerFields.append(hf(QStringLiteral("media"), QStringLiteral("Media"), ValueType::Text, 2, 2));
        s.headerFields.append(hf(QStringLiteral("resistance"), QStringLiteral("Resistance"), ValueType::Number, 2, 4, QStringLiteral("ohm")));
        s.headerFields.append(hf(QStringLiteral("puffing_regime"), QStringLiteral("Puffing Regime"), ValueType::Text, 2, 8));
        s.headerFields.append(hf(QStringLiteral("viscosity"), QStringLiteral("Viscosity"), ValueType::Number, 3, 2, QStringLiteral("cP")));
        s.headerFields.append(hf(QStringLiteral("voltage"), QStringLiteral("Voltage"), ValueType::Number, 3, 6, QStringLiteral("V")));
    } else {
        // Standardized layout (extractMetadata "new"/"old" branches - identical
        // cell positions; the old branch's missing Heating Technology is handled
        // by processSheet dropping that key when tv != "new").
        s.headerFields.append(hf(QStringLiteral("test_name"), QStringLiteral("Test Name"), ValueType::Text, 1, 1));
        s.headerFields.append(hf(QStringLiteral("date"), QStringLiteral("Date"), ValueType::Text, 1, 4));
        s.headerFields.append(hf(QStringLiteral("sample_id"), QStringLiteral("Sample ID"), ValueType::Text, 1, 6));
        s.headerFields.append(hf(QStringLiteral("heating_technology"), QStringLiteral("Heating Technology"), ValueType::Text, 1, 8));
        s.headerFields.append(hf(QStringLiteral("media"), QStringLiteral("Media"), ValueType::Text, 2, 2));
        s.headerFields.append(hf(QStringLiteral("resistance"), QStringLiteral("Resistance"), ValueType::Number, 2, 4, QStringLiteral("ohm")));
        s.headerFields.append(hf(QStringLiteral("power"), QStringLiteral("Power"), ValueType::Number, 2, 6, QStringLiteral("W"),
                                  QStringLiteral("power_v1"),
                                  {QStringLiteral("header:voltage"), QStringLiteral("header:resistance"), QStringLiteral("header:heating_technology")}));
        s.headerFields.append(hf(QStringLiteral("puffing_regime"), QStringLiteral("Puffing Regime"), ValueType::Text, 2, 8));
        s.headerFields.append(hf(QStringLiteral("viscosity"), QStringLiteral("Viscosity"), ValueType::Number, 3, 2, QStringLiteral("cP")));
        s.headerFields.append(hf(QStringLiteral("tester"), QStringLiteral("Tester"), ValueType::Text, 3, 4));
        s.headerFields.append(hf(QStringLiteral("voltage"), QStringLiteral("Voltage"), ValueType::Number, 3, 6, QStringLiteral("V")));
        s.headerFields.append(hf(QStringLiteral("initial_oil_mass"), QStringLiteral("Initial Oil Mass"), ValueType::Number, 3, 8, QStringLiteral("g")));
    }

    // ── Aggregates ──
    s.aggregates.append(agg(QStringLiteral("average_tpm"), QStringLiteral("mean"), {QStringLiteral("tpm")}));
    s.aggregates.append(agg(QStringLiteral("stddev_tpm"), QStringLiteral("stddev"), {QStringLiteral("tpm")}));
    s.aggregates.append(agg(QStringLiteral("avg_power_density"), QStringLiteral("mean_over_power"),
                             {QStringLiteral("tpm"), QStringLiteral("header:power")}));
    s.aggregates.append(agg(QStringLiteral("normalized_tpm"), QStringLiteral("mean_over_power"),
                             {QStringLiteral("tpm"), QStringLiteral("header:power")}));
    s.aggregates.append(agg(QStringLiteral("total_puffs"), QStringLiteral("last"), {QStringLiteral("puffs")}));
    s.aggregates.append(agg(QStringLiteral("total_oil_consumed"), QStringLiteral("last"), {QStringLiteral("oil_consumed")}));
    s.aggregates.append(agg(QStringLiteral("efficiency_percent"), QStringLiteral("efficiency_v1"),
                             {QStringLiteral("total_oil_consumed"), QStringLiteral("header:initial_oil_mass")}));
    s.aggregates.append(agg(QStringLiteral("burn_clog_leak"), QStringLiteral("status_scan_v1"),
                             {QStringLiteral("smell"), QStringLiteral("clog"), QStringLiteral("notes")}));

    return s;
}

}} // namespace DVE::model
