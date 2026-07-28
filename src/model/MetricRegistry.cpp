#include "MetricRegistry.h"
#include "SchemaDrivenReader.h"
#include <QHash>

namespace DVE { namespace model {

namespace {

MetricDef m(const QString& key, const QString& displayName, ValueType type, const QString& unit,
            Role role, const QStringList& aliases = QStringList(),
            const QString& calculator = QString(), const QStringList& inputs = QStringList(),
            const QMap<QString, QString>& tags = {})
{
    MetricDef d;
    d.key = key; d.displayName = displayName; d.headerAliases = aliases;
    d.type = type; d.unit = unit; d.role = role;
    d.calculator = calculator; d.inputs = inputs; d.tags = tags;
    return d;
}

HeaderFieldDef h(const QString& key, const QString& displayName, ValueType type,
                 const QString& unit = QString(), const QString& calculator = QString(),
                 const QStringList& inputs = QStringList())
{
    HeaderFieldDef f;
    f.key = key; f.displayName = displayName; f.type = type; f.unit = unit;
    f.calculator = calculator; f.inputs = inputs;
    return f;   // row/col stay 0: positions are layout-specific, not registry data
}

QVector<MetricDef> buildMetrics()
{
    QVector<MetricDef> v;
    // -- The standard column set (registry section 2) --
    v << m("puffs", "puffs", ValueType::Number, "count", Role::Measured);
    v << m("before_weight", "Before Weight (g)", ValueType::Number, "g", Role::Measured,
           {"Before weight/g", "Before Weight/g", "Before weight (g)"});
    v << m("after_weight", "After Weight (g)", ValueType::Number, "g", Role::Measured,
           {"After weight/g", "After Weight/g", "After weight (g)"});
    v << m("draw_pressure", "Draw Pressure", ValueType::Number, "kPa", Role::Measured,
           {"Draw Pressure (kpa)"});
    v << m("resistance", "Resistance", ValueType::Number, "ohm", Role::Measured,
           {QString::fromUtf8("Resistance (\xCE\xA9)")});
    v << m("puffing_regime", "Puffing Regime", ValueType::Text, "", Role::Qualitative, {}, {}, {},
           {{"encoding", "legacy composite string; canonical form is the 4 split metrics (9.2)"}});
    v << m("smell", "Smell", ValueType::Text, "", Role::Qualitative,
           {"Smell (1-4)", "Smell (0-4)"}, {}, {},
           {{"scale", "0-4; 1-4-era sheets leave blank for no event (D9)"}});
    v << m("clog", "Clog", ValueType::Text, "", Role::Qualitative, {"Clog (Y/N)", "Clog?"});
    v << m("notes", "Notes", ValueType::Text, "", Role::Qualitative);
    v << m("tpm", "TPM", ValueType::Number, "mg/puff", Role::Derived, {"TPM (mg/puff)"},
           "tpm_v1", {"puffs", "before_weight", "after_weight"});
    v << m("tpm_power_density", "TPM Power Density", ValueType::Number, "mg/(W*s)", Role::Derived,
           {"TPM Power Density (mg/(W*s))", "TPM/PD"},
           "power_density_v1", {"tpm", "header:power", "puff_time"},
           {{"approximation", "no transient power; TCR shifts instantaneous power (9.1)"}});
    v << m("tpm_puff_density", "TPM Puff Density", ValueType::Number, "mg/((n s) puff*W)", Role::Derived,
           {"TPM Power Density (mg/puff/W)", "TPM Puff Density (mg/(puff*W))"},
           "puff_density_v1", {"tpm", "header:power"},
           {{"approximation", "no transient power; TCR shifts instantaneous power (9.1)"},
            {"n", "puff length taken from the puff_time metric"}});
    v << m("variation_tpm", "Variation in TPM", ValueType::Number, "%", Role::Derived,
           {"Variation in TPM (%)", "Variation"},
           "variation_v1", {"tpm"},
           {{"definition", "template rolling CV in percent is canonical (D1)"}});
    v << m("tpm_consistency", "TPM Consistency", ValueType::Number, "", Role::Derived, {},
           "consistency_v1", {"tpm"},
           {{"definition", "STDEV.P/AVERAGE over the session window, FRACTION (9.4)"}});
    v << m("rolling_avg_tpm", "Rolling Average TPM", ValueType::Number, "mg/puff", Role::Derived, {},
           "rolling_avg_v1", {"tpm"});
    v << m("oil_consumed", "Oil Consumed", ValueType::Number, "g", Role::Derived,
           {"Oil Consumed (Cumulative, g)"},
           "oil_consumed_v1", {"before_weight", "after_weight"});
    // -- Ratified additions (registry section 8.2) --
    v << m("puff_volume", "Puff Volume", ValueType::Number, "mL", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"});
    v << m("puff_time", "Puff Time", ValueType::Number, "s", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"});
    v << m("puff_rest_time", "Puff Rest Time", ValueType::Number, "s", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"});
    v << m("session_rest_time", "Session Rest Time", ValueType::Number, "s", Role::Derived, {},
           "regime_split_v1", {"puffing_regime"},
           {{"default", "0 when the regime string has only 3 parts"}});
    v << m("draw_pressure_per_puff", "Draw Pressure (per puff)", ValueType::NumberList, "kPa",
           Role::Measured, {}, {}, {},
           {{"source", "data-logger series; list length = puffs in the session"},
            {"lowering", "historical PV1-PV5 columns assemble into this (D5)"}});
    v << m("voltage", "Voltage", ValueType::Mixed, "V", Role::Measured, {}, {}, {},
           {{"grouping", "groupable like puffing regime; string values name curves (D1)"}});
    v << m("image", "Image", ValueType::Image, "", Role::Qualitative);
    // -- Open metrics observed in the wild (registry section 2) --
    v << m("chronology", "Chronology", ValueType::Text, "", Role::Qualitative);
    v << m("failure", "Failure", ValueType::Text, "", Role::Qualitative,
           {"Failure? (if yes, add detailed notes)", "Failure? (if yes, when)"});
    return v;
}

QVector<HeaderFieldDef> buildHeaderFields()
{
    QVector<HeaderFieldDef> v;
    v << h("test_name", "Test Name", ValueType::Text);
    v << h("date", "Date", ValueType::Text);
    v << h("tester", "Tester", ValueType::Text);
    v << h("sample_id", "Sample ID", ValueType::Text);
    v << h("project_name", "Project", ValueType::Text);
    v << h("sample_suffix", "Sample", ValueType::Text);
    v << h("distributor", "Distributor", ValueType::Text);
    v << h("resistance", "Resistance", ValueType::Number, "ohm");
    v << h("resistance_initial", "Resistance (Initial)", ValueType::Number, "ohm");
    v << h("resistance_final", "Resistance (Final)", ValueType::Number, "ohm");
    v << h("voltage", "Voltage", ValueType::Number, "V");
    v << h("power", "Power", ValueType::Number, "W", "power_v1",
           {"header:voltage", "header:resistance", "header:heating_technology"});
    v << h("heating_technology", "Heating Technology", ValueType::Text);
    v << h("media", "Media", ValueType::Text);
    v << h("viscosity", "Viscosity", ValueType::Number, "cP");
    v << h("initial_oil_mass", "Initial Oil Mass", ValueType::Number, "g");
    v << h("fill_volume", "Fill Volume", ValueType::Number, "mL");
    v << h("number_of_samples", "Number of Samples", ValueType::Number);
    v << h("puffing_regime", "Puffing Regime", ValueType::Text);
    v << h("did_burn", "Did this burn?", ValueType::Bool);
    v << h("did_clog", "Did this clog?", ValueType::Bool);
    v << h("did_leak", "Did this leak?", ValueType::Bool);
    // Cart-era design specs (registry 3.4; optional standard headers per D6)
    v << h("coil_material", "Coil Material", ValueType::Text);
    v << h("thermal_conductivity", "Thermal Conductivity", ValueType::Number);
    v << h("column_inner_diameter", "Column inner diameter", ValueType::Number, "mm");
    v << h("column_length", "Column length", ValueType::Number, "mm");
    v << h("coil_shape", "Coil shape", ValueType::Text);
    v << h("cotton_length", "Cotton length", ValueType::Number, "mm");
    return v;
}

// Label spellings per header key (registry section 3). displayName and key are
// implicit aliases; this table adds the punctuated / historical spellings.
QHash<QString, QStringList> buildHeaderAliasTable()
{
    QHash<QString, QStringList> t;
    // Only OBSERVED spellings are registered (review 2a: bare "Test"/"Ri"/"Rf"
    // were unobserved extrapolations that widen the inference label scan).
    t.insert("test_name", {"Test:"});
    t.insert("date", {"Date:"});
    t.insert("tester", {"Tester:"});
    t.insert("sample_id", {"Sample ID:", "Cart #"});
    t.insert("project_name", {"Project:"});
    t.insert("sample_suffix", {"Sample:"});
    t.insert("distributor", {"Distributor:"});
    t.insert("resistance", {"Resistance:", QString::fromUtf8("Resistance (\xCE\xA9):"), "Resistance (Ohms):"});
    t.insert("resistance_initial", {"Ri (Ohms)"});
    t.insert("resistance_final", {"Rf (Ohms)"});
    t.insert("voltage", {"Voltage:"});
    t.insert("power", {"Power:"});
    t.insert("heating_technology", {"Heating Technology:", "Heater Technology", "Heater Technology:"});
    t.insert("media", {"Media:"});
    t.insert("viscosity", {"Viscosity:"});
    t.insert("initial_oil_mass", {"Initial Oil Mass:"});
    t.insert("fill_volume", {"Fill Volume:"});
    t.insert("puffing_regime", {"Puffing Regime:", "Puff Regime"});
    t.insert("cotton_length", {"Cotton length (if applicable)"});
    return t;
}

} // namespace

const QVector<MetricDef>& MetricRegistry::allMetrics()
{
    static const QVector<MetricDef> v = buildMetrics();
    return v;
}

const QVector<HeaderFieldDef>& MetricRegistry::allHeaderFields()
{
    static const QVector<HeaderFieldDef> v = buildHeaderFields();
    return v;
}

const MetricDef* MetricRegistry::metric(const QString& key)
{
    for (const MetricDef& d : allMetrics())
        if (d.key == key) return &d;
    return nullptr;
}

const HeaderFieldDef* MetricRegistry::headerField(const QString& key)
{
    for (const HeaderFieldDef& d : allHeaderFields())
        if (d.key == key) return &d;
    return nullptr;
}

const MetricDef* MetricRegistry::metricByAlias(const QString& normalized)
{
    static const QHash<QString, int> index = [] {
        QHash<QString, int> ix;
        const QVector<MetricDef>& all = allMetrics();
        for (int i = 0; i < all.size(); ++i) {
            QStringList names{all[i].displayName, all[i].key};
            names += all[i].headerAliases;
            for (const QString& n : names) {
                const QString norm = SchemaDrivenReader::normalizeHeader(n);
                if (!norm.isEmpty()) ix.insert(norm, i);
            }
        }
        return ix;
    }();
    const auto it = index.constFind(normalized);
    return it == index.constEnd() ? nullptr : &allMetrics()[it.value()];
}

QStringList MetricRegistry::headerAliasesFor(const QString& key)
{
    static const QHash<QString, QStringList> table = buildHeaderAliasTable();
    return table.value(key);
}

const HeaderFieldDef* MetricRegistry::headerByLabel(const QString& normalized)
{
    static const QHash<QString, int> index = [] {
        QHash<QString, int> ix;
        const QVector<HeaderFieldDef>& all = allHeaderFields();
        for (int i = 0; i < all.size(); ++i) {
            QStringList names{all[i].displayName, all[i].key};
            names += headerAliasesFor(all[i].key);
            for (const QString& n : names) {
                const QString norm = SchemaDrivenReader::normalizeHeader(n);
                if (!norm.isEmpty()) ix.insert(norm, i);
            }
        }
        return ix;
    }();
    const auto it = index.constFind(normalized);
    return it == index.constEnd() ? nullptr : &allHeaderFields()[it.value()];
}

MetricRegistry::PerPuffAlias MetricRegistry::perPuffAlias(const QString& normalized)
{
    PerPuffAlias hit;
    if (normalized.size() == 3 && normalized.startsWith(QLatin1String("pv"))) {
        const int n = normalized.mid(2).toInt();
        if (n >= 1 && n <= 5) {
            hit.targetKey = QStringLiteral("draw_pressure_per_puff");
            hit.index = n;
        }
    }
    return hit;
}

}} // namespace DVE::model
