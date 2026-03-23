#include "ZipWriter.h"

#include <QFile>
#include <QDateTime>
#include <zlib.h>

// ---------------------------------------------------------------------------
// Internal little-endian write helpers
// ---------------------------------------------------------------------------
static void writeLE16(QByteArray& buf, quint16 v)
{
    buf.append(static_cast<char>(v & 0xFF));
    buf.append(static_cast<char>((v >> 8) & 0xFF));
}

static void writeLE32(QByteArray& buf, quint32 v)
{
    buf.append(static_cast<char>(v & 0xFF));
    buf.append(static_cast<char>((v >> 8)  & 0xFF));
    buf.append(static_cast<char>((v >> 16) & 0xFF));
    buf.append(static_cast<char>((v >> 24) & 0xFF));
}

// ---------------------------------------------------------------------------
// ZIP method / version constants
// ---------------------------------------------------------------------------
static constexpr quint32 SIG_LOCAL          = 0x04034b50u;
static constexpr quint32 SIG_CENTRAL        = 0x02014b50u;
static constexpr quint32 SIG_END_OF_CENTRAL = 0x06054b50u;

static constexpr quint16 VERSION_NEEDED     = 20;   // 2.0 – needed for deflate
static constexpr quint16 VERSION_MADE_BY    = 20;
static constexpr quint16 METHOD_STORE       = 0;
static constexpr quint16 METHOD_DEFLATE     = 8;

// MS-DOS epoch date/time used for reproducible output (1980-01-01 00:00:00).
static constexpr quint16 DOS_DATE           = 0x0021; // 1980-01-01
static constexpr quint16 DOS_TIME           = 0x0000;

// ---------------------------------------------------------------------------
ZipWriter::ZipWriter()  = default;
ZipWriter::~ZipWriter() = default;

// ---------------------------------------------------------------------------
void ZipWriter::addFile(const QString& name, const QByteArray& data, bool compress)
{
    Entry e;
    e.name       = name;
    e.data       = data;
    e.compressed = compress;

    if (compress && !data.isEmpty()) {
        e.compressedData = deflate(data);
        // If deflated output is larger than raw, fall back to STORE.
        if (e.compressedData.isEmpty() ||
            e.compressedData.size() >= data.size()) {
            e.compressed     = false;
            e.compressedData.clear();
        }
    }

    m_entries.append(e);
}

// ---------------------------------------------------------------------------
QByteArray ZipWriter::deflate(const QByteArray& input) const
{
    if (input.isEmpty())
        return QByteArray();

    // windowBits = -15 → raw deflate (no zlib wrapper, no gzip wrapper).
    z_stream strm{};
    strm.zalloc   = Z_NULL;
    strm.zfree    = Z_NULL;
    strm.opaque   = Z_NULL;

    int ret = deflateInit2(&strm,
                           Z_DEFAULT_COMPRESSION,
                           Z_DEFLATED,
                           -15,          // raw deflate
                           8,            // default memory level
                           Z_DEFAULT_STRATEGY);
    if (ret != Z_OK) {
        m_lastError = QStringLiteral("deflateInit2 failed: %1").arg(ret);
        return QByteArray();
    }

    strm.avail_in  = static_cast<uInt>(input.size());
    strm.next_in   = reinterpret_cast<Bytef*>(const_cast<char*>(input.data()));

    QByteArray output;
    output.resize(static_cast<int>(deflateBound(&strm, strm.avail_in)));

    strm.avail_out = static_cast<uInt>(output.size());
    strm.next_out  = reinterpret_cast<Bytef*>(output.data());

    ret = ::deflate(&strm, Z_FINISH);
    deflateEnd(&strm);

    if (ret != Z_STREAM_END) {
        m_lastError = QStringLiteral("deflate failed: %1").arg(ret);
        return QByteArray();
    }

    output.resize(static_cast<int>(strm.total_out));
    return output;
}

// ---------------------------------------------------------------------------
quint32 ZipWriter::crc32(const QByteArray& data)
{
    // Build table on first call (ISO 3309 / ITU-T V.42 polynomial 0xEDB88320).
    static quint32 table[256] = {};
    static bool    tableReady = false;
    if (!tableReady) {
        for (int i = 0; i < 256; ++i) {
            quint32 c = static_cast<quint32>(i);
            for (int k = 0; k < 8; ++k)
                c = (c & 1) ? (0xEDB88320u ^ (c >> 1)) : (c >> 1);
            table[i] = c;
        }
        tableReady = true;
    }

    quint32 crc = 0xFFFFFFFFu;
    const auto* bytes = reinterpret_cast<const unsigned char*>(data.constData());
    for (int i = 0; i < data.size(); ++i)
        crc = table[(crc ^ bytes[i]) & 0xFF] ^ (crc >> 8);
    return crc ^ 0xFFFFFFFFu;
}

// ---------------------------------------------------------------------------
QByteArray ZipWriter::buildLocalFileHeader(const Entry& e,
                                           quint32 crc,
                                           quint32 compSize,
                                           quint32 uncompSize) const
{
    QByteArray name = e.name.toUtf8();
    quint16 method  = e.compressed ? METHOD_DEFLATE : METHOD_STORE;

    QByteArray hdr;
    writeLE32(hdr, SIG_LOCAL);
    writeLE16(hdr, VERSION_NEEDED);
    writeLE16(hdr, 0);              // general purpose flags
    writeLE16(hdr, method);
    writeLE16(hdr, DOS_TIME);
    writeLE16(hdr, DOS_DATE);
    writeLE32(hdr, crc);
    writeLE32(hdr, compSize);
    writeLE32(hdr, uncompSize);
    writeLE16(hdr, static_cast<quint16>(name.size()));
    writeLE16(hdr, 0);              // extra field length
    hdr.append(name);
    return hdr;
}

// ---------------------------------------------------------------------------
QByteArray ZipWriter::buildCentralDirEntry(const Entry& e,
                                           quint32 offset,
                                           quint32 crc,
                                           quint32 compSize,
                                           quint32 uncompSize) const
{
    QByteArray name = e.name.toUtf8();
    quint16 method  = e.compressed ? METHOD_DEFLATE : METHOD_STORE;

    QByteArray cd;
    writeLE32(cd, SIG_CENTRAL);
    writeLE16(cd, VERSION_MADE_BY);
    writeLE16(cd, VERSION_NEEDED);
    writeLE16(cd, 0);               // general purpose flags
    writeLE16(cd, method);
    writeLE16(cd, DOS_TIME);
    writeLE16(cd, DOS_DATE);
    writeLE32(cd, crc);
    writeLE32(cd, compSize);
    writeLE32(cd, uncompSize);
    writeLE16(cd, static_cast<quint16>(name.size()));
    writeLE16(cd, 0);               // extra field length
    writeLE16(cd, 0);               // file comment length
    writeLE16(cd, 0);               // disk number start
    writeLE16(cd, 0);               // internal attributes
    writeLE32(cd, 0);               // external attributes
    writeLE32(cd, offset);
    cd.append(name);
    return cd;
}

// ---------------------------------------------------------------------------
QByteArray ZipWriter::buildEndOfCentralDir(int     numEntries,
                                           quint32 centralDirSize,
                                           quint32 centralDirOffset) const
{
    QByteArray eocd;
    writeLE32(eocd, SIG_END_OF_CENTRAL);
    writeLE16(eocd, 0);                                         // disk number
    writeLE16(eocd, 0);                                         // start disk
    writeLE16(eocd, static_cast<quint16>(numEntries));          // entries on disk
    writeLE16(eocd, static_cast<quint16>(numEntries));          // total entries
    writeLE32(eocd, centralDirSize);
    writeLE32(eocd, centralDirOffset);
    writeLE16(eocd, 0);                                         // comment length
    return eocd;
}

// ---------------------------------------------------------------------------
QByteArray ZipWriter::toByteArray() const
{
    // Pre-compute total size to avoid repeated QByteArray reallocations
    qint64 estimatedSize = 0;
    for (const Entry& e : m_entries) {
        const QByteArray& payload = e.compressed ? e.compressedData : e.data;
        // Local file header (~30 + name) + payload + central dir entry (~46 + name)
        estimatedSize += 30 + e.name.size() + payload.size() + 46 + e.name.size();
    }
    estimatedSize += 22;  // End of central directory

    QByteArray zip;
    zip.reserve(static_cast<int>(qMin(estimatedSize, qint64(INT_MAX))));

    // Offsets into `zip` at which each local-file-header starts.
    QVector<quint32>  offsets;
    QVector<quint32>  crcs;
    QVector<quint32>  compSizes;
    QVector<quint32>  uncompSizes;
    offsets.reserve(m_entries.size());
    crcs.reserve(m_entries.size());
    compSizes.reserve(m_entries.size());
    uncompSizes.reserve(m_entries.size());

    for (const Entry& e : m_entries) {
        const QByteArray& payload    = e.compressed ? e.compressedData : e.data;
        quint32           crc        = crc32(e.data);
        quint32           compSize   = static_cast<quint32>(payload.size());
        quint32           uncompSize = static_cast<quint32>(e.data.size());

        offsets.append(static_cast<quint32>(zip.size()));
        crcs.append(crc);
        compSizes.append(compSize);
        uncompSizes.append(uncompSize);

        zip.append(buildLocalFileHeader(e, crc, compSize, uncompSize));
        zip.append(payload);
    }

    // Central directory
    quint32 cdOffset = static_cast<quint32>(zip.size());
    QByteArray cd;
    for (int i = 0; i < m_entries.size(); ++i)
        cd.append(buildCentralDirEntry(m_entries[i], offsets[i],
                                       crcs[i], compSizes[i], uncompSizes[i]));

    zip.append(cd);
    zip.append(buildEndOfCentralDir(m_entries.size(),
                                    static_cast<quint32>(cd.size()),
                                    cdOffset));
    return zip;
}

// ---------------------------------------------------------------------------
bool ZipWriter::save(const QString& filePath)
{
    QFile f(filePath);
    if (!f.open(QIODevice::WriteOnly | QIODevice::Truncate)) {
        m_lastError = QStringLiteral("Cannot open for writing: %1").arg(filePath);
        return false;
    }
    const QByteArray data = toByteArray();
    if (f.write(data) != data.size()) {
        m_lastError = QStringLiteral("Write error: %1").arg(f.errorString());
        return false;
    }
    return true;
}
