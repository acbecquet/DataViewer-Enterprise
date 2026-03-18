#pragma once
#include <QString>
#include <QByteArray>
#include <QMap>
#include <QVector>

// ---------------------------------------------------------------------------
// ZipWriter – minimal ZIP 2.0 writer backed by zlib raw deflate.
// Supports stored (no compression) and deflated entries.
// Compatible with Qt 6 / MinGW / MSVC builds that ship zlib.
// ---------------------------------------------------------------------------
class ZipWriter {
public:
    ZipWriter();
    ~ZipWriter();

    // Add a file entry to the archive.
    //   compress = true  → DEFLATE (method 8)
    //   compress = false → STORE  (method 0)
    void addFile(const QString& name, const QByteArray& data, bool compress = true);

    // Serialise to disk.  Returns false on error; inspect lastError().
    bool save(const QString& filePath);

    // Serialise to an in-memory buffer.
    QByteArray toByteArray() const;

    QString lastError() const { return m_lastError; }

private:
    struct Entry {
        QString    name;
        QByteArray data;            // original (uncompressed) bytes
        bool       compressed;      // true → store with method 8
        QByteArray compressedData;  // deflate output (empty if stored)
    };

    QVector<Entry> m_entries;
    mutable QString m_lastError;

    // Raw DEFLATE (windowBits = -15) using zlib.
    QByteArray deflate(const QByteArray& input) const;

    // Standard CRC-32 (polynomial 0xEDB88320).
    static quint32 crc32(const QByteArray& data);

    // Binary layout helpers.
    QByteArray buildLocalFileHeader(const Entry& e,
                                    quint32 crc,
                                    quint32 compSize,
                                    quint32 uncompSize) const;

    QByteArray buildCentralDirEntry(const Entry& e,
                                    quint32 offset,
                                    quint32 crc,
                                    quint32 compSize,
                                    quint32 uncompSize) const;

    QByteArray buildEndOfCentralDir(int     numEntries,
                                    quint32 centralDirSize,
                                    quint32 centralDirOffset) const;
};
