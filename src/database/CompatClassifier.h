#pragma once

#include <QString>
#include <QStringList>
#include <QDate>

// v2.4.2 SP4-T1 (A3): CompatClassifier — pure-C++ triage classification for the
// Database Browser version/health filter. NO Qt object, NO DB, NO GUI: it maps a
// row's app_version (or, when unstamped, its creation date) to a coarse era
// bucket, and maps already-derived data-shape booleans to health flags. Mirrors
// the VersionLookup pattern (free functions in namespace DVE) and is unit-tested
// by tst_compatclassifier.
//
// Score-shape note: legacy-string-score detection is NOT done here — by the time
// JSON reaches a C++ struct the tolerant reader has coerced string scores to
// doubles, so the string-vs-number distinction is invisible. The caller passes
// `hasLegacyStringScores` in from the server query (jsonb_typeof, SP4-T2).

namespace DVE {

// Coarse era buckets. eraLabel is the canonical filter-bucket key (stable across
// approx/exact rows); `approx` is set when the era was inferred from the creation
// date because app_version was NULL/unstamped (the UI may append "(approx.)").
struct CompatClass {
    QString     eraLabel;        // "v2.0.x".."v2.4.x", or "pre-v2.4.2 (unknown)"
    bool        approx = false;  // era inferred from creation date (no app_version)
    QStringList health;          // empty == Healthy
    bool isHealthy() const { return health.isEmpty(); }
};

// The canonical "unstamped / pre-stamping" bucket label (used when there is no
// app_version and no usable creation date to infer an era).
QString unstampedEraLabel();

// Classify a row's era. `appVersion` is the stored value (e.g. "DataViewer/2.2.5",
// "2.2.5", "v2.2.5", or empty/NULL). When it parses to a vN.M version the bucket
// is "vN.M.x" (approx=false). When empty/unparseable, the era is inferred from
// `creation` against the release-date table (approx=true); if `creation` is also
// invalid (or predates v2.0.0), returns the unstamped bucket.
CompatClass classifyEra(const QString& appVersion, const QDate& creation);

// Health flags for a sensory / detailed-sensory session, from data shape.
//   hasLegacyStringScores (server-derived) -> "Legacy string scores"
//   isPlaceholder (default name / no real content) OR sampleCount==0 -> "Junk candidate"
// Empty result == Healthy.
QStringList sensoryHealth(bool hasLegacyStringScores, bool isPlaceholder, int sampleCount);

// Health flags for a TPM file, from data shape.
//   sampleCount==0 -> "No samples"
//   missingRegimes -> "Missing puff regimes"
// Empty result == Healthy.
QStringList fileHealth(int sampleCount, bool missingRegimes);

} // namespace DVE
