#pragma once

#include <QByteArray>
#include <QString>
#include <QRectF>

namespace DVE {

// Maximum dimension (px) for embedded images. Larger images are downscaled
// to keep PPTX / database size manageable. 1920 = full-HD width.
inline constexpr int kMaxImageDim = 1920;

// JPEG quality (0-100) for re-encoded images. 85 is a good visual/size tradeoff.
inline constexpr int kJpegQuality = 85;

// Load an image from disk, apply a normalised [0,1] crop rect, downscale if
// oversized, and re-encode as JPEG.  Returns raw JPEG bytes, or the original
// file bytes when no transformation is needed (avoids re-encoding).
//
// This is the single shared implementation used by both ReportGenerator and
// SensoryPanel.
QByteArray loadAndCropImage(const QString& path, const QRectF& crop);

// Compress an in-memory image blob: downscale if oversized, re-encode as
// JPEG.  Returns the original blob unchanged when it is already a small-
// enough JPEG.
QByteArray compressImageBlob(const QByteArray& raw);

} // namespace DVE
