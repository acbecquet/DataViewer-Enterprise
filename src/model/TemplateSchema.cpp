#include "TemplateSchema.h"

namespace DVE { namespace model {

const MetricDef* TemplateSchema::column(const QString& key) const
{
    for (const MetricDef& m : columns)
        if (m.key == key) return &m;
    return nullptr;
}

int TemplateSchema::columnPos(const QString& key) const
{
    for (int i = 0; i < columns.size(); ++i)
        if (columns[i].key == key) return i;
    return -1;
}

const HeaderFieldDef* TemplateSchema::headerField(const QString& key) const
{
    for (const HeaderFieldDef& h : headerFields)
        if (h.key == key) return &h;
    return nullptr;
}

}} // namespace DVE::model
