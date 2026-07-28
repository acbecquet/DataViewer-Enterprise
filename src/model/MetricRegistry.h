#pragma once
#include "MetricDef.h"
#include <QVector>

namespace DVE { namespace model {

// Compiled form of the RATIFIED vocabulary registry
// (docs/superpowers/specs/2026-07-27-tpm-v3-vocabulary-registry-draft.md).
// Single source of truth for canonical keys, display names, aliases, types,
// units, and tags. StandardSchema and SchemaInference copy defs OUT of here;
// layout positions and presentation flags stay layout-side.
// Naming policy (registry section 5): keys are forever; matching is by
// normalized alias (SchemaDrivenReader::normalizeHeader); metrics and header
// fields are separate namespaces.
class MetricRegistry {
public:
    static const QVector<MetricDef>& allMetrics();
    static const QVector<HeaderFieldDef>& allHeaderFields();

    static const MetricDef* metric(const QString& key);
    static const HeaderFieldDef* headerField(const QString& key);

    // Lookup by NORMALIZED alias text (caller normalizes). nullptr on miss.
    static const MetricDef* metricByAlias(const QString& normalized);
    static const HeaderFieldDef* headerByLabel(const QString& normalized);

    // Label spellings registered for a header key (normalized-alias sources).
    static QStringList headerAliasesFor(const QString& key);

    // PV1..PV5 -> element index of the draw_pressure_per_puff list metric
    // (registry 9.1/D5). index == 0 means "not a per-puff alias".
    struct PerPuffAlias { QString targetKey; int index = 0; };
    static PerPuffAlias perPuffAlias(const QString& normalized);
};

}} // namespace DVE::model
