#include "StandardSchema.h"

namespace DVE { namespace model {

namespace {

MetricDef col(const QString& key, const QString& displayName, ValueType type, const QString& unit,
              Role role, const QString& calculator = QString(), const QStringList& inputs = QStringList(),
              bool plottable = false, bool editable = false, const QStringList& aliases = QStringList())
{
    MetricDef m;
    m.key           = key;
    m.displayName   = displayName;
    m.headerAliases = aliases;
    m.type          = type;
    m.unit          = unit;
    m.role          = role;
    m.calculator    = calculator;
    m.inputs        = inputs;
    m.plottable     = plottable;
    m.editable      = editable;
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

TemplateSchema standardV1(bool perRowRegime)
{
    TemplateSchema s;
    s.schemaId       = QStringLiteral("standard");
    s.version        = 1;
    s.headerRows      = 3;
    s.columnHeaderRow = 4;
    s.dataStartRow    = 5;
    s.blockCols       = 12;

    // ── Columns (physical column order == legacy DVE::Cols order) ──
    s.columns.append(col(QStringLiteral("puffs"), QStringLiteral("puffs"),
                          ValueType::Number, QStringLiteral("count"), Role::Measured));
    s.columns.append(col(QStringLiteral("before_weight"), QStringLiteral("Before Weight (g)"),
                          ValueType::Number, QStringLiteral("g"), Role::Measured));
    s.columns.append(col(QStringLiteral("after_weight"), QStringLiteral("After Weight (g)"),
                          ValueType::Number, QStringLiteral("g"), Role::Measured));
    s.columns.append(col(QStringLiteral("draw_pressure"), QStringLiteral("Draw Pressure"),
                          ValueType::Number, QStringLiteral("kPa"), Role::Measured,
                          QString(), QStringList(), false, false,
                          {QStringLiteral("Draw Pressure (kpa)")}));

    if (perRowRegime) {
        s.columns.append(col(QStringLiteral("puffing_regime"), QStringLiteral("Puffing Regime"),
                              ValueType::Text, QString(), Role::Qualitative,
                              QString(), QStringList(), false, true));
    } else {
        s.columns.append(col(QStringLiteral("resistance"), QStringLiteral("Resistance"),
                              ValueType::Number, QStringLiteral("ohm"), Role::Measured,
                              QString(), QStringList(), false, false,
                              {QStringLiteral("Resistance (Ω)")}));
    }

    s.columns.append(col(QStringLiteral("smell"), QStringLiteral("Smell"),
                          ValueType::Text, QString(), Role::Qualitative,
                          QString(), QStringList(), false, true,
                          {QStringLiteral("Smell (1-4)")}));
    s.columns.append(col(QStringLiteral("clog"), QStringLiteral("Clog"),
                          ValueType::Text, QString(), Role::Qualitative,
                          QString(), QStringList(), false, true,
                          {QStringLiteral("Clog (Y/N)")}));
    s.columns.append(col(QStringLiteral("notes"), QStringLiteral("Notes"),
                          ValueType::Text, QString(), Role::Qualitative,
                          QString(), QStringList(), false, true));
    s.columns.append(col(QStringLiteral("tpm"), QStringLiteral("TPM"),
                          ValueType::Number, QStringLiteral("mg/puff"), Role::Derived,
                          QStringLiteral("tpm_v1"),
                          {QStringLiteral("puffs"), QStringLiteral("before_weight"), QStringLiteral("after_weight")},
                          true, false,
                          {QStringLiteral("TPM (mg/puff)")}));
    s.columns.append(col(QStringLiteral("tpm_power_density"), QStringLiteral("TPM/PD"),
                          ValueType::Number, QStringLiteral("mg/(puff*W)"), Role::Derived,
                          QStringLiteral("power_density_v1"),
                          {QStringLiteral("tpm"), QStringLiteral("header:power")},
                          false, false,
                          {QStringLiteral("TPM Power Density (mg/(W*s))")}));
    s.columns.append(col(QStringLiteral("variation_tpm"), QStringLiteral("Variation"),
                          ValueType::Number, QStringLiteral("%"), Role::Derived,
                          QStringLiteral("variation_v1"),
                          {QStringLiteral("tpm")},
                          false, false,
                          {QStringLiteral("Variation in TPM (%)")}));
    s.columns.append(col(QStringLiteral("oil_consumed"), QStringLiteral("Oil Consumed"),
                          ValueType::Number, QStringLiteral("mg"), Role::Derived,
                          QStringLiteral("oil_consumed_v1"),
                          {QStringLiteral("before_weight"), QStringLiteral("after_weight")},
                          false, false,
                          {QStringLiteral("Oil Consumed (Cumulative, g)")}));

    // ── Header-band fields (block-relative, 1-based) ──
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
