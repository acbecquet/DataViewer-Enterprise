#pragma once
#include <QJsonValue>
#include <QStringList>

namespace DVE { namespace testutil {

// Deep-compare two JSON values. Returns one line per difference in the form
// "/path[idx]/key: <a> != <b>"; empty list means identical. Numbers compare
// exactly (both sides of every v3 diff come from the same serializer, so
// float formatting is deterministic; the Task 3 determinism test proves it).
QStringList diffJson(const QJsonValue& a, const QJsonValue& b,
                     const QString& path = QString());

}} // namespace DVE::testutil
