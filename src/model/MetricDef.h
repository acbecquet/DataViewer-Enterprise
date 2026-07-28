#pragma once
#include <QMap>
#include <QString>
#include <QStringList>

namespace DVE { namespace model {

enum class ValueType { Number, Text, Bool, Mixed, NumberList, Image };
enum class Role { Measured, Qualitative, Derived, Identity };

// One free-standing metric (a data column, in template terms).
struct MetricDef {
    QString     key;             // stable snake_case id, e.g. "before_weight"
    QString     displayName;     // canonical header text
    QStringList headerAliases;   // extra header spellings matched on read
    ValueType   type = ValueType::Number;
    QString     unit;
    Role        role = Role::Measured;
    QString     calculator;      // Derived only; evaluated from Phase 2
    QStringList inputs;          // metric keys or "header:<key>"
    bool        plottable = false;
    bool        editable  = false;
    int         precision = 2;
    // Open-ended key->value annotations (design spec section 18: "tag anything
    // on to the metrics"). Registry-authored today; manifest-authored in 2c.
    QMap<QString, QString> tags;
};

// One header-band field (applies to the whole sample).
struct HeaderFieldDef {
    QString     key;
    QString     displayName;
    ValueType   type = ValueType::Text;
    QString     unit;
    int         row = 0;         // 1-based, block-relative template row
    int         col = 0;         // 1-based, block-relative template col
    QString     calculator;      // derived header values (e.g. power)
    QStringList inputs;
};

struct AggregateDef {
    QString     key;
    QString     calculator;
    QStringList inputs;
};

}} // namespace DVE::model
