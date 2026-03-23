#include "ImageUtils.h"

#include <QFile>
#include <QBuffer>
#include <QImage>
#include <QImageReader>

namespace DVE {

// ── Helpers ─────────────────────────────────────────────────────────────────

static bool isJpegHeader(const QByteArray& data)
{
    return data.size() >= 3
        && static_cast<unsigned char>(data[0]) == 0xFF
        && static_cast<unsigned char>(data[1]) == 0xD8
        && static_cast<unsigned char>(data[2]) == 0xFF;
}

static QByteArray encodeAsJpeg(const QImage& img)
{
    QByteArray result;
    QBuffer buf(&result);
    buf.open(QIODevice::WriteOnly);
    img.save(&buf, "JPEG", kJpegQuality);
    return result;
}

// ── loadAndCropImage ────────────────────────────────────────────────────────

QByteArray loadAndCropImage(const QString& path, const QRectF& crop)
{
    QFile f(path);
    if (!f.open(QIODevice::ReadOnly)) return {};
    QByteArray raw = f.readAll();

    const bool needsCrop = crop.isValid()
                        && crop != QRectF(0.0, 0.0, 1.0, 1.0);

    // Fast path: no crop needed -- check dimensions without a full decode.
    if (!needsCrop) {
        QBuffer probeBuf(const_cast<QByteArray*>(&raw));
        probeBuf.open(QIODevice::ReadOnly);
        QImageReader reader(&probeBuf);
        const QSize imgSize = reader.size();

        if (imgSize.isValid()
            && imgSize.width()  <= kMaxImageDim
            && imgSize.height() <= kMaxImageDim) {
            // Image fits within limits. If it's already JPEG, skip decode entirely.
            if (isJpegHeader(raw)) return raw;
            // Non-JPEG but within dimension limits -- re-encode to JPEG for
            // consistent format in the output PPTX, but no scale/crop needed.
        }
    }

    // Full decode required (crop, scale, or format conversion).
    QImage img;
    img.loadFromData(raw);
    if (img.isNull()) return raw;

    bool modified = false;

    // Apply crop.
    if (needsCrop) {
        const QRect srcRect(qRound(crop.x()      * img.width()),
                            qRound(crop.y()      * img.height()),
                            qRound(crop.width()  * img.width()),
                            qRound(crop.height() * img.height()));
        img = img.copy(srcRect);
        modified = true;
    }

    // Downscale if oversized.
    if (img.width() > kMaxImageDim || img.height() > kMaxImageDim) {
        img = img.scaled(kMaxImageDim, kMaxImageDim,
                         Qt::KeepAspectRatio, Qt::SmoothTransformation);
        modified = true;
    }

    // If nothing was modified and the source is already JPEG, return raw bytes
    // to avoid a lossy re-encode.
    if (!modified && isJpegHeader(raw))
        return raw;

    return encodeAsJpeg(img);
}

// ── compressImageBlob ───────────────────────────────────────────────────────

QByteArray compressImageBlob(const QByteArray& raw)
{
    if (raw.isEmpty()) return raw;

    // Check dimensions without a full decode using QImageReader.
    // This avoids the expensive QImage::loadFromData() for images that are
    // already small enough JPEGs.
    {
        QBuffer probeBuf(const_cast<QByteArray*>(&raw));
        probeBuf.open(QIODevice::ReadOnly);
        QImageReader reader(&probeBuf);
        const QSize sz = reader.size();

        if (sz.isValid()
            && sz.width()  <= kMaxImageDim
            && sz.height() <= kMaxImageDim
            && isJpegHeader(raw)) {
            return raw;  // Already a small-enough JPEG -- skip decode entirely.
        }
    }

    QImage img;
    img.loadFromData(raw);
    if (img.isNull()) return raw;

    if (img.width() > kMaxImageDim || img.height() > kMaxImageDim) {
        img = img.scaled(kMaxImageDim, kMaxImageDim,
                         Qt::KeepAspectRatio, Qt::SmoothTransformation);
    }

    return encodeAsJpeg(img);
}

} // namespace DVE
