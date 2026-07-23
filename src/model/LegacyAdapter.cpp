#include "LegacyAdapter.h"
#include "../pipeline/HeatingTech.h"

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

}} // namespace DVE::model
