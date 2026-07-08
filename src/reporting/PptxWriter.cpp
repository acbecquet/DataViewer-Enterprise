#include "PptxWriter.h"
#include "../utils/XmlBuilder.h"
#include "../utils/ZipWriter.h"

#include <QFile>
#include <QDir>
#include <QFileInfo>
#include <QImageReader>
#include <QBuffer>
#include <QImage>
#include <QTextStream>

namespace DVE {

// ============================================================================
// Namespace-level XML constants
// ============================================================================

// DrawingML / PresentationML namespaces (used repeatedly).
static const QString NS_PML  = QStringLiteral("http://schemas.openxmlformats.org/presentationml/2006/main");
static const QString NS_DML  = QStringLiteral("http://schemas.openxmlformats.org/drawingml/2006/main");
static const QString NS_REL  = QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
static const QString NS_PKG  = QStringLiteral("http://schemas.openxmlformats.org/package/2006/relationships");
static const QString NS_CT   = QStringLiteral("http://schemas.openxmlformats.org/package/2006/content-types");

// Relationship type base.
static const QString RT_BASE = QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships/");
static const QString RT_PKG  = QStringLiteral("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
static const QString RT_SLIDE = RT_BASE + QStringLiteral("slide");
static const QString RT_SM   = RT_BASE + QStringLiteral("slideMaster");
static const QString RT_SL   = RT_BASE + QStringLiteral("slideLayout");
static const QString RT_THEME= RT_BASE + QStringLiteral("theme");
static const QString RT_IMG  = RT_BASE + QStringLiteral("image");

// Slide canvas dimensions (EMU).
static const long long SLIDE_CX = 12192000LL;
static const long long SLIDE_CY = 6858000LL;

// ============================================================================
// Construction / destruction
// ============================================================================
PptxWriter::PptxWriter()  = default;
PptxWriter::~PptxWriter() = default;

void PptxWriter::setResourcePath(const QString& resourceDir)
{
    m_resourceDir = resourceDir;
}

// ============================================================================
// EMU / colour helpers
// ============================================================================
long long PptxWriter::toEmu(double inches)
{
    return static_cast<long long>(inches * 914400.0 + 0.5);
}

QString PptxWriter::emu(double inches)
{
    return QString::number(toEmu(inches));
}

QString PptxWriter::colorHex(int color)
{
    return QStringLiteral("%1").arg(color & 0xFFFFFF, 6, 16, QLatin1Char('0')).toUpper();
}

// ============================================================================
// Resource loading
// ============================================================================
QByteArray PptxWriter::loadResourceImage(const QString& filename) const
{
    if (m_resourceCache.contains(filename))
        return m_resourceCache[filename];

    // Try direct path first, then images/ subdirectory
    QStringList candidates = {
        QDir(m_resourceDir).filePath(filename),
        QDir(m_resourceDir).filePath("images/" + filename)
    };

    for (const QString& path : candidates) {
        QFile f(path);
        if (f.open(QIODevice::ReadOnly)) {
            QByteArray data = f.readAll();
            qDebug() << "PptxWriter: loaded resource" << filename
                     << "from" << path << "(" << data.size() << "bytes)";
            m_resourceCache[filename] = data;
            return data;
        }
    }

    qWarning() << "PptxWriter: FAILED to load resource" << filename
               << "| resourceDir:" << m_resourceDir
               << "| candidates:" << candidates;
    m_lastError = QStringLiteral("Cannot open resource: %1 (searched %2)")
        .arg(filename, m_resourceDir);
    return QByteArray();
}

// ============================================================================
// Media registration
// ============================================================================

// Detect whether raw bytes are JPEG (starts with FF D8 FF).
static bool isJpeg(const QByteArray& data)
{
    return data.size() >= 3
        && static_cast<unsigned char>(data[0]) == 0xFF
        && static_cast<unsigned char>(data[1]) == 0xD8
        && static_cast<unsigned char>(data[2]) == 0xFF;
}

QString PptxWriter::addMedia(const QByteArray&               data,
                              const QString&                  ext,
                              QVector<QPair<QString, QByteArray>>& media)
{
    ++m_mediaCounter;

    // Determine actual extension: honour the caller's hint, but override when
    // the raw bytes prove the format is different (e.g. JPEG data labelled "png").
    QString actualExt = ext;
    if (ext == QStringLiteral("png") && isJpeg(data))
        actualExt = QStringLiteral("jpg");

    const QString zipPath = QStringLiteral("ppt/media/image%1.%2")
                            .arg(m_mediaCounter).arg(actualExt);
    media.append({zipPath, data});
    return QStringLiteral("rId%1").arg(media.size());
}

// ============================================================================
// Public slide-building API
// ============================================================================

void PptxWriter::addCoverSlide(const QString& title, const QString& dateStr)
{
    // Defer to the 4-arg overload with 0/0 = legacy fonts.
    addCoverSlide(title, dateStr, /*titleFontPt=*/0, /*dateFontPt=*/0);
}

void PptxWriter::addCoverSlide(const QString& title, const QString& dateStr,
                                int titleFontPt, int dateFontPt,
                                const QRectF& titleRect,
                                const QRectF& subtitleRect)
{
    Slide slide;

    QByteArray bgData   = loadResourceImage(QStringLiteral("Cover_Page_Logo.jpg"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full_white.png"));

    // rId assignment: bg → rId1, logo → rId2.
    QString bgRid   = addMedia(bgData,   QStringLiteral("jpg"),  slide.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"),  slide.media);

    slide.xml = buildCoverSlideXml(title, dateStr, bgRid, logoRid,
                                    titleFontPt, dateFontPt,
                                    titleRect, subtitleRect);
    m_slides.append(slide);
}

// ---------------------------------------------------------------------------
// 4-arg wrapper: synthesises a ContentSlideLayout from the legacy positions
// embedded in `table` / `plots[0]`, then delegates to the 5-arg overload.
void PptxWriter::addContentSlide(const QString&          sheetTitle,
                                  const SlideTable&       table,
                                  const QVector<SlideImage>& plots,
                                  const QString&          extraShapesXml)
{
    ContentSlideLayout dl;
    dl.title = QRectF();   // null → buildContentSlideXml keeps the hardcoded
                           // title position for 4-arg (non-layout) callers.
    dl.table = QRectF(table.x, table.y, table.w, table.h);
    if (!plots.isEmpty()) {
        const SlideImage& p = plots.first();
        dl.radar = QRectF(p.x, p.y, p.w, p.h);
    }
    addContentSlide(sheetTitle, table, plots, dl, extraShapesXml);
}

// 5-arg overload: contains the actual slide-building logic. Mutates copies
// of `tableIn`/`plotsIn` with values from `layout` before emitting XML.
void PptxWriter::addContentSlide(const QString&          sheetTitle,
                                  const SlideTable&       tableIn,
                                  const QVector<SlideImage>& plotsIn,
                                  const ContentSlideLayout& layout,
                                  const QString&          extraShapesXml)
{
    // Apply layout overrides (a null QRectF means "use the position already
    // baked into the input table/plot").
    SlideTable table = tableIn;
    if (!layout.table.isNull()) {
        table.x = layout.table.x();
        table.y = layout.table.y();
        table.w = layout.table.width();
        table.h = layout.table.height();
    }
    // Font override: layout.tableFontPt > 0 wins over whatever the caller
    // baked into SlideTable.fontPt. 0 falls through to whatever the table
    // already had (which itself defaults to 0 = legacy hardcoded 9 pt).
    if (layout.tableFontPt > 0) table.fontPt = layout.tableFontPt;

    QVector<SlideImage> plots = plotsIn;
    if (!plots.isEmpty() && !layout.radar.isNull()) {
        plots[0].x = layout.radar.x();
        plots[0].y = layout.radar.y();
        plots[0].w = layout.radar.width();
        plots[0].h = layout.radar.height();
    }

    Slide slide;

    QByteArray bgData   = loadResourceImage(QStringLiteral("ccell_background.png"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full.png"));

    QString bgRid   = addMedia(bgData,   QStringLiteral("png"), slide.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"), slide.media);

    // Register each plot image; build map of plot-index → rId.
    QMap<QString, QString> plotRids;
    for (int i = 0; i < plots.size(); ++i) {
        const SlideImage& si = plots[i];
        if (!si.pngData.isEmpty()) {
            QString rId = addMedia(si.pngData, QStringLiteral("png"), slide.media);
            plotRids[QStringLiteral("plot%1").arg(i)] = rId;
        }
    }

    slide.xml = buildContentSlideXml(sheetTitle, table, plots,
                                     bgRid, logoRid, plotRids,
                                     extraShapesXml, layout.titleFontPt,
                                     layout.title);
    m_slides.append(slide);
}

// ---------------------------------------------------------------------------
static QString slideHeader(const QString& bgRid);
static QString slideFooter();

void PptxWriter::addDualTableSlide(const QString&          sheetTitle,
                                    const SlideTable&       topTable,
                                    const SlideTable&       bottomTable,
                                    const QVector<SlideImage>& plots)
{
    Slide slide;

    QByteArray bgData   = loadResourceImage(QStringLiteral("ccell_background.png"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full.png"));

    QString bgRid   = addMedia(bgData,   QStringLiteral("png"), slide.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"), slide.media);

    // Register each plot image
    QMap<QString, QString> plotRids;
    for (int i = 0; i < plots.size(); ++i) {
        const SlideImage& si = plots[i];
        if (!si.pngData.isEmpty()) {
            QString rId = addMedia(si.pngData, QStringLiteral("png"), slide.media);
            plotRids[QStringLiteral("plot%1").arg(i)] = rId;
        }
    }

    // Build shapes XML directly (same pattern as buildContentSlideXml)
    QString shapes;
    int id = 2;

    // Title — dynamic height for wrapping
    int dtTitleLines = qMax(1, (sheetTitle.length() + 37) / 38);
    double dtTitleH = dtTitleLines * 0.50;
    shapes += makeTextBox(id++, 0.4, 0.1, 11.0, dtTitleH, sheetTitle,
                          QStringLiteral("Montserrat"), ReportLayout::kPptxContentTitlePt * 100, true,
                          QStringLiteral("1F497D"), QStringLiteral("l"));

    // Logo top-right
    shapes += makePic(id++, 11.9, 0.1, 1.22, 0.4, logoRid);

    // Top table (raw data)
    if (!topTable.headers.isEmpty())
        shapes += buildTableXml(topTable, id++);

    // Bottom table (averages/stats)
    if (!bottomTable.headers.isEmpty())
        shapes += buildTableXml(bottomTable, id++);

    // Plot images
    for (int i = 0; i < plots.size(); ++i) {
        const QString key = QStringLiteral("plot%1").arg(i);
        if (!plotRids.contains(key)) continue;
        const SlideImage& si = plots[i];
        shapes += makePic(id++, si.x, si.y,
                          (si.w > 0) ? si.w : 2.7,
                          (si.h > 0) ? si.h : 2.0,
                          plotRids[key]);
    }

    slide.xml = slideHeader(bgRid) + shapes + slideFooter();
    m_slides.append(slide);
}

// ---------------------------------------------------------------------------
void PptxWriter::addImageSlide(const QString&          sheetTitle,
                                const QVector<QByteArray>& images)
{
    Slide slide;

    QByteArray bgData   = loadResourceImage(QStringLiteral("ccell_background.png"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full.png"));

    QString bgRid   = addMedia(bgData,   QStringLiteral("png"), slide.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"), slide.media);

    QMap<int, QString> imageRids;
    int count = qMin(images.size(), 6);  // max 6 per slide (2×3 grid)
    for (int i = 0; i < count; ++i) {
        if (!images[i].isEmpty()) {
            QString rId = addMedia(images[i], QStringLiteral("png"), slide.media);
            imageRids[i] = rId;
        }
    }

    slide.xml = buildImageSlideXml(sheetTitle, count, bgRid, logoRid, imageRids);
    m_slides.append(slide);
}

// ---------------------------------------------------------------------------
void PptxWriter::addImageSlide(const QString&          sheetTitle,
                                const QVector<SlideImage>& images)
{
    Slide slide;

    QByteArray bgData   = loadResourceImage(QStringLiteral("ccell_background.png"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full.png"));

    QString bgRid   = addMedia(bgData,   QStringLiteral("png"), slide.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"), slide.media);

    QVector<QPair<QString, QRectF>> imageRids;
    for (const SlideImage& img : images) {
        if (!img.pngData.isEmpty()) {
            QString rId = addMedia(img.pngData, QStringLiteral("png"), slide.media);
            imageRids.append({rId, QRectF(img.x, img.y, img.w, img.h)});
        }
    }

    slide.xml = buildCustomImageSlideXml(sheetTitle, bgRid, logoRid, imageRids);
    m_slides.append(slide);
}

// ---------------------------------------------------------------------------
void PptxWriter::addConclusionsSlide()
{
    Slide s;

    QByteArray bgData   = loadResourceImage(QStringLiteral("ccell_background.png"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full_white.png"));

    QString bgRid   = addMedia(bgData,   QStringLiteral("png"), s.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"), s.media);

    // Title is rendered by buildContentSlideXml; only the body box is added here.
    QString shapes;
    shapes += makeTextBox(101, 0.80, 1.40, 11.7, 5.50,
                          QStringLiteral(""),
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    s.xml = buildContentSlideXml(QStringLiteral("Conclusions"),
                                 SlideTable{}, QVector<SlideImage>{},
                                 bgRid, logoRid, {}, shapes);
    m_slides.append(std::move(s));
}

// ---------------------------------------------------------------------------
void PptxWriter::addTestProtocolSlide(const SlideTable& sopTable)
{
    Slide s;

    QByteArray bgData   = loadResourceImage(QStringLiteral("ccell_background.png"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full_white.png"));

    QString bgRid   = addMedia(bgData,   QStringLiteral("png"), s.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"), s.media);

    s.xml = buildContentSlideXml(QStringLiteral("Test Protocol"),
                                 sopTable, QVector<SlideImage>{},
                                 bgRid, logoRid, {}, QString());
    m_slides.append(std::move(s));
}

// ---------------------------------------------------------------------------
void PptxWriter::addTestOverviewSlide(const QString& description,
                                      const QStringList& testNames)
{
    Slide s;

    QByteArray bgData   = loadResourceImage(QStringLiteral("ccell_background.png"));
    QByteArray logoData = loadResourceImage(QStringLiteral("ccell_logo_full_white.png"));

    QString bgRid   = addMedia(bgData,   QStringLiteral("png"), s.media);
    QString logoRid = addMedia(logoData, QStringLiteral("png"), s.media);

    // Title is rendered by buildContentSlideXml; description/bullets/summary follow.
    QString shapes;
    shapes += makeTextBox(101, 0.80, 1.30, 11.7, 0.55,
                          description,
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    QStringList bulletLines;
    for (const QString& name : testNames)
        bulletLines << QStringLiteral("• ") + name;
    shapes += makeTextBox(102, 1.00, 2.10, 11.5, 4.30,
                          bulletLines.join("\n"),
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    shapes += makeTextBox(103, 0.80, 6.50, 11.7, 0.55,
                          QStringLiteral(""),
                          QStringLiteral("Segoe UI"), 1800, false,
                          QStringLiteral("000000"), QStringLiteral("l"));

    s.xml = buildContentSlideXml(QStringLiteral("Test Overview"),
                                 SlideTable{}, QVector<SlideImage>{},
                                 bgRid, logoRid, {}, shapes);
    m_slides.append(std::move(s));
}

// ---------------------------------------------------------------------------
void PptxWriter::addSectionDividerSlide(const QString& filename)
{
    addSectionDividerSlide(filename, /*titleFontPt=*/0);
}

void PptxWriter::addSectionDividerSlide(const QString& filename, int titleFontPt)
{
    Slide s;
    // Match the cover slide: full-bleed Cover_Page_Logo.jpg with white CCELL logo overlay.
    QString bgRid   = addMedia(loadResourceImage(QStringLiteral("Cover_Page_Logo.jpg")), QStringLiteral("jpg"), s.media);
    QString logoRid = addMedia(loadResourceImage(QStringLiteral("ccell_logo_full_white.png")), QStringLiteral("png"), s.media);

    s.xml = buildCoverSlideXml(filename, /*date=*/QString(), bgRid, logoRid,
                                titleFontPt, /*dateFontPt=*/0);
    m_slides.append(std::move(s));
}

// ============================================================================
// save()
// ============================================================================
bool PptxWriter::save(const QString& filePath)
{
    ZipWriter zip;

    // ── Static skeleton parts ──────────────────────────────────────────────
    zip.addFile(QStringLiteral("[Content_Types].xml"),
                buildContentTypesXml().toUtf8());

    zip.addFile(QStringLiteral("_rels/.rels"),
                buildRootRelsXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/presentation.xml"),
                buildPresentationXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/_rels/presentation.xml.rels"),
                buildPresRelsXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/presProps.xml"),
                buildPresPropsXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/viewProps.xml"),
                buildViewPropsXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/tableStyles.xml"),
                buildTableStylesXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/theme/theme1.xml"),
                buildThemeXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/slideMasters/slideMaster1.xml"),
                buildSlideMasterXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/slideMasters/_rels/slideMaster1.xml.rels"),
                buildSlideMasterRelsXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/slideLayouts/slideLayout1.xml"),
                buildSlideLayoutXml().toUtf8());

    zip.addFile(QStringLiteral("ppt/slideLayouts/_rels/slideLayout1.xml.rels"),
                buildSlideLayoutRelsXml().toUtf8());

    // ── Per-slide content ──────────────────────────────────────────────────
    for (int i = 0; i < m_slides.size(); ++i) {
        const Slide& s = m_slides[i];
        const int n    = i + 1;

        zip.addFile(QStringLiteral("ppt/slides/slide%1.xml").arg(n),
                    s.xml.toUtf8());

        zip.addFile(QStringLiteral("ppt/slides/_rels/slide%1.xml.rels").arg(n),
                    buildSlideRelsXml(s.media).toUtf8());

        for (const auto& [path, data] : s.media) {
            zip.addFile(path, data, /*compress=*/true);
        }
    }

    if (!zip.save(filePath)) {
        m_lastError = zip.lastError();
        return false;
    }
    return true;
}

// ============================================================================
// Presentation XML
// ============================================================================
QString PptxWriter::buildPresentationXml() const
{
    QString sldIdLst;
    // Slide IDs start at 256; relationship IDs start at rId2
    // (rId1 is the slide master).
    for (int i = 0; i < m_slides.size(); ++i) {
        sldIdLst += QStringLiteral(
            R"(<p:sldId id="%1" r:id="rId%2"/>)")
            .arg(256 + i).arg(2 + i);
    }

    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<p:presentation)"
        R"( xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main")"
        R"( xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main")"
        R"( xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships")"
        R"( saveSubsetFonts="1">)"
        R"(<p:sldMasterIdLst>)"
        R"(<p:sldMasterId id="2147483648" r:id="rId1"/>)"
        R"(</p:sldMasterIdLst>)"
        R"(<p:sldIdLst>%1</p:sldIdLst>)"
        R"(<p:sldSz cx="12192000" cy="6858000"/>)"
        R"(<p:notesSz cx="6858000" cy="9144000"/>)"
        R"(</p:presentation>)").arg(sldIdLst);
}

// ============================================================================
// [Content_Types].xml
// ============================================================================
QString PptxWriter::buildContentTypesXml() const
{
    QString overrides;

    overrides += QStringLiteral(
        R"(<Override PartName="/ppt/presentation.xml")"
        R"( ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>)");

    overrides += QStringLiteral(
        R"(<Override PartName="/ppt/presProps.xml")"
        R"( ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>)");

    overrides += QStringLiteral(
        R"(<Override PartName="/ppt/viewProps.xml")"
        R"( ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>)");

    overrides += QStringLiteral(
        R"(<Override PartName="/ppt/tableStyles.xml")"
        R"( ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>)");

    overrides += QStringLiteral(
        R"(<Override PartName="/ppt/theme/theme1.xml")"
        R"( ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>)");

    overrides += QStringLiteral(
        R"(<Override PartName="/ppt/slideMasters/slideMaster1.xml")"
        R"( ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>)");

    overrides += QStringLiteral(
        R"(<Override PartName="/ppt/slideLayouts/slideLayout1.xml")"
        R"( ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>)");

    for (int i = 0; i < m_slides.size(); ++i) {
        overrides += QStringLiteral(
            R"(<Override PartName="/ppt/slides/slide%1.xml")"
            R"( ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>)")
            .arg(i + 1);
    }

    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">)"
        R"(<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>)"
        R"(<Default Extension="xml"  ContentType="application/xml"/>)"
        R"(<Default Extension="png"  ContentType="image/png"/>)"
        R"(<Default Extension="jpg"  ContentType="image/jpeg"/>)"
        R"(<Default Extension="jpeg" ContentType="image/jpeg"/>)"
        R"(%1)"
        R"(</Types>)").arg(overrides);
}

// ============================================================================
// _rels/.rels
// ============================================================================
QString PptxWriter::buildRootRelsXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
        R"(<Relationship Id="rId1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")"
        R"( Target="ppt/presentation.xml"/>)"
        R"(</Relationships>)");
}

// ============================================================================
// ppt/_rels/presentation.xml.rels
// ============================================================================
QString PptxWriter::buildPresRelsXml() const
{
    QString rels;

    // rId1 = slide master
    rels += QStringLiteral(
        R"(<Relationship Id="rId1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")"
        R"( Target="slideMasters/slideMaster1.xml"/>)");

    // rId2..rIdN = slides
    for (int i = 0; i < m_slides.size(); ++i) {
        rels += QStringLiteral(
            R"(<Relationship Id="rId%1")"
            R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")"
            R"( Target="slides/slide%2.xml"/>)")
            .arg(2 + i).arg(i + 1);
    }

    // theme, presProps, viewProps, tableStyles
    int nextId = 2 + m_slides.size();

    rels += QStringLiteral(
        R"(<Relationship Id="rId%1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")"
        R"( Target="theme/theme1.xml"/>)").arg(nextId++);

    rels += QStringLiteral(
        R"(<Relationship Id="rId%1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps")"
        R"( Target="presProps.xml"/>)").arg(nextId++);

    rels += QStringLiteral(
        R"(<Relationship Id="rId%1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps")"
        R"( Target="viewProps.xml"/>)").arg(nextId++);

    rels += QStringLiteral(
        R"(<Relationship Id="rId%1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles")"
        R"( Target="tableStyles.xml"/>)").arg(nextId);

    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
        R"(%1)"
        R"(</Relationships>)").arg(rels);
}

// ============================================================================
// ppt/slideMasters/_rels/slideMaster1.xml.rels
// ============================================================================
QString PptxWriter::buildSlideMasterRelsXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
        R"(<Relationship Id="rId1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")"
        R"( Target="../theme/theme1.xml"/>)"
        R"(<Relationship Id="rId2")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")"
        R"( Target="../slideLayouts/slideLayout1.xml"/>)"
        R"(</Relationships>)");
}

// ============================================================================
// ppt/slideLayouts/_rels/slideLayout1.xml.rels
// ============================================================================
QString PptxWriter::buildSlideLayoutRelsXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
        R"(<Relationship Id="rId1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")"
        R"( Target="../slideMasters/slideMaster1.xml"/>)"
        R"(</Relationships>)");
}

// ============================================================================
// Per-slide .rels
// ============================================================================
QString PptxWriter::buildSlideRelsXml(
        const QVector<QPair<QString, QByteArray>>& media) const
{
    QString rels;

    // rId1..rIdN = media (images)
    for (int i = 0; i < media.size(); ++i) {
        // Convert absolute zip path to relative path from ppt/slides/
        // e.g. "ppt/media/image3.png" → "../media/image3.png"
        const QString& zipPath = media[i].first;
        const QString relTarget = QStringLiteral("../") +
                                  zipPath.section(QLatin1Char('/'), 1);  // drop "ppt/"

        rels += QStringLiteral(
            R"(<Relationship Id="rId%1")"
            R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")"
            R"( Target="%2"/>)")
            .arg(i + 1).arg(relTarget);
    }

    // Every slide must reference a slideLayout — without this PowerPoint
    // shows a "repair" dialog on open.
    int layoutRid = media.size() + 1;
    rels += QStringLiteral(
        R"(<Relationship Id="rId%1")"
        R"( Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout")"
        R"( Target="../slideLayouts/slideLayout1.xml"/>)")
        .arg(layoutRid);

    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)"
        R"(%1)"
        R"(</Relationships>)").arg(rels);
}

// ============================================================================
// Slide Master XML (minimal blank master)
// ============================================================================
QString PptxWriter::buildSlideMasterXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<p:sldMaster)"
        R"( xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main")"
        R"( xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main")"
        R"( xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">)"
        R"(<p:cSld>)"
        R"(<p:spTree>)"
        R"(<p:nvGrpSpPr>)"
        R"(<p:cNvPr id="1" name=""/>)"
        R"(<p:cNvGrpSpPr/>)"
        R"(<p:nvPr/>)"
        R"(</p:nvGrpSpPr>)"
        R"(<p:grpSpPr>)"
        R"(<a:xfrm>)"
        R"(<a:off x="0" y="0"/>)"
        R"(<a:ext cx="0" cy="0"/>)"
        R"(<a:chOff x="0" y="0"/>)"
        R"(<a:chExt cx="0" cy="0"/>)"
        R"(</a:xfrm>)"
        R"(</p:grpSpPr>)"
        R"(</p:spTree>)"
        R"(</p:cSld>)"
        R"(<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2")"
        R"( accent1="accent1" accent2="accent2" accent3="accent3")"
        R"( accent4="accent4" accent5="accent5" accent6="accent6")"
        R"( hlink="hlink" folHlink="folHlink"/>)"
        R"(<p:sldLayoutIdLst>)"
        R"(<p:sldLayoutId id="2147483649" r:id="rId2"/>)"
        R"(</p:sldLayoutIdLst>)"
        R"(<p:txStyles>)"
        R"(<p:titleStyle><a:lstStyle/></p:titleStyle>)"
        R"(<p:bodyStyle><a:lstStyle/></p:bodyStyle>)"
        R"(<p:otherStyle><a:lstStyle/></p:otherStyle>)"
        R"(</p:txStyles>)"
        R"(</p:sldMaster>)");
}

// ============================================================================
// Slide Layout XML (blank layout)
// ============================================================================
QString PptxWriter::buildSlideLayoutXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<p:sldLayout)"
        R"( xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main")"
        R"( xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main")"
        R"( xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships")"
        R"( type="blank" preserve="1">)"
        R"(<p:cSld name="Blank">)"
        R"(<p:spTree>)"
        R"(<p:nvGrpSpPr>)"
        R"(<p:cNvPr id="1" name=""/>)"
        R"(<p:cNvGrpSpPr/>)"
        R"(<p:nvPr/>)"
        R"(</p:nvGrpSpPr>)"
        R"(<p:grpSpPr>)"
        R"(<a:xfrm>)"
        R"(<a:off x="0" y="0"/>)"
        R"(<a:ext cx="0" cy="0"/>)"
        R"(<a:chOff x="0" y="0"/>)"
        R"(<a:chExt cx="0" cy="0"/>)"
        R"(</a:xfrm>)"
        R"(</p:grpSpPr>)"
        R"(</p:spTree>)"
        R"(</p:cSld>)"
        R"(<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>)"
        R"(</p:sldLayout>)");
}

// ============================================================================
// Theme XML (minimal Office theme)
// ============================================================================
QString PptxWriter::buildThemeXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">)"
        R"(<a:themeElements>)"
        R"(<a:clrScheme name="Office">)"
        R"(<a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1>)"
        R"(<a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1>)"
        R"(<a:dk2><a:srgbClr val="44546A"/></a:dk2>)"
        R"(<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>)"
        R"(<a:accent1><a:srgbClr val="4472C4"/></a:accent1>)"
        R"(<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>)"
        R"(<a:accent3><a:srgbClr val="A9D18E"/></a:accent3>)"
        R"(<a:accent4><a:srgbClr val="FFC000"/></a:accent4>)"
        R"(<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>)"
        R"(<a:accent6><a:srgbClr val="70AD47"/></a:accent6>)"
        R"(<a:hlink><a:srgbClr val="0563C1"/></a:hlink>)"
        R"(<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>)"
        R"(</a:clrScheme>)"
        R"(<a:fontScheme name="Office">)"
        R"(<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>)"
        R"(<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>)"
        R"(</a:fontScheme>)"
        R"(<a:fmtScheme name="Office">)"
        R"(<a:fillStyleLst>)"
        R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
        R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
        R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
        R"(</a:fillStyleLst>)"
        R"(<a:lnStyleLst>)"
        R"(<a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>)"
        R"(<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>)"
        R"(<a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>)"
        R"(</a:lnStyleLst>)"
        R"(<a:effectStyleLst>)"
        R"(<a:effectStyle><a:effectLst/></a:effectStyle>)"
        R"(<a:effectStyle><a:effectLst/></a:effectStyle>)"
        R"(<a:effectStyle><a:effectLst/></a:effectStyle>)"
        R"(</a:effectStyleLst>)"
        R"(<a:bgFillStyleLst>)"
        R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
        R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
        R"(<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>)"
        R"(</a:bgFillStyleLst>)"
        R"(</a:fmtScheme>)"
        R"(</a:themeElements>)"
        R"(</a:theme>)");
}

// ============================================================================
// presProps / viewProps / tableStyles (minimal stubs)
// ============================================================================
QString PptxWriter::buildPresPropsXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<p:presProps)"
        R"( xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main")"
        R"( xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">)"
        R"(<p:extLst/>)"
        R"(</p:presProps>)");
}

QString PptxWriter::buildViewPropsXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<p:viewPr)"
        R"( xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main")"
        R"( xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">)"
        R"(<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>)"
        R"(<p:slideViewPr>)"
        R"(<p:cSldViewPr>)"
        R"(<p:cViewPr varScale="1"><p:scale><a:sx n="64" d="100"/><a:sy n="64" d="100"/></p:scale>)"
        R"(<p:origin x="-1608" y="-114"/></p:cViewPr>)"
        R"(<p:guideLst/>)"
        R"(</p:cSldViewPr>)"
        R"(</p:slideViewPr>)"
        R"(<p:lastView>sldView</p:lastView>)"
        R"(</p:viewPr>)");
}

QString PptxWriter::buildTableStylesXml() const
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<p:tblStyleLst)"
        R"( xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main")"
        R"( def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>)");
}

// ============================================================================
// Shape XML helpers
// ============================================================================

// ---------------------------------------------------------------------------
// Text box
// ---------------------------------------------------------------------------
QString PptxWriter::makeTextBox(int    id,
                                 double xIn, double yIn,
                                 double wIn, double hIn,
                                 const QString& text,
                                 const QString& font,
                                 int    sizePt100,
                                 bool   bold,
                                 const QString& colorHex,
                                 const QString& alignHoriz) const
{
    const QString xStr = emu(xIn);
    const QString yStr = emu(yIn);
    const QString wStr = emu(wIn);
    const QString hStr = emu(hIn);

    const QString safeText = XmlBuilder::escapeXml(text);

    const QString boldStr = bold ? QStringLiteral("1") : QStringLiteral("0");

    return QStringLiteral(
        R"(<p:sp>)"
        R"(<p:nvSpPr>)"
        R"(<p:cNvPr id="%1" name="TextBox %1"/>)"
        R"(<p:cNvSpPr txBox="1"/>)"
        R"(<p:nvPr/>)"
        R"(</p:nvSpPr>)"
        R"(<p:spPr>)"
        R"(<a:xfrm><a:off x="%2" y="%3"/><a:ext cx="%4" cy="%5"/></a:xfrm>)"
        R"(<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>)"
        R"(<a:noFill/>)"
        R"(</p:spPr>)"
        R"(<p:txBody>)"
        R"(<a:bodyPr wrap="square" rtlCol="0"/>)"
        R"(<a:lstStyle/>)"
        R"(<a:p>)"
        R"(<a:pPr algn="%9"/>)"
        R"(<a:r>)"
        R"(<a:rPr lang="en-US" sz="%6" b="%7" dirty="0">)"
        R"(<a:solidFill><a:srgbClr val="%8"/></a:solidFill>)"
        R"(<a:latin typeface="%10"/>)"
        R"(</a:rPr>)"
        R"(<a:t>%11</a:t>)"
        R"(</a:r>)"
        R"(</a:p>)"
        R"(</p:txBody>)"
        R"(</p:sp>)")
        .arg(id)
        .arg(xStr, yStr, wStr, hStr)
        .arg(sizePt100)
        .arg(boldStr)
        .arg(colorHex)
        .arg(alignHoriz)
        .arg(font)
        .arg(safeText);
}

// ---------------------------------------------------------------------------
// Image (p:pic)
// ---------------------------------------------------------------------------
QString PptxWriter::makePic(int    id,
                             double xIn, double yIn,
                             double wIn, double hIn,
                             const QString& rId) const
{
    return QStringLiteral(
        R"(<p:pic>)"
        R"(<p:nvPicPr>)"
        R"(<p:cNvPr id="%1" name="Image %1"/>)"
        R"(<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>)"
        R"(<p:nvPr/>)"
        R"(</p:nvPicPr>)"
        R"(<p:blipFill>)"
        R"(<a:blip r:embed="%2"/>)"
        R"(<a:stretch><a:fillRect/></a:stretch>)"
        R"(</p:blipFill>)"
        R"(<p:spPr>)"
        R"(<a:xfrm><a:off x="%3" y="%4"/><a:ext cx="%5" cy="%6"/></a:xfrm>)"
        R"(<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>)"
        R"(</p:spPr>)"
        R"(</p:pic>)")
        .arg(id).arg(rId)
        .arg(emu(xIn), emu(yIn), emu(wIn), emu(hIn));
}

// ============================================================================
// Table XML
// ============================================================================

QString PptxWriter::makeTableCell(const QString& text,
                                   bool           isHeader,
                                   const QString& bgColorHex,
                                   const QString& textColorHex,
                                   int            customSizePt100) const
{
    const QString boldStr = isHeader ? QStringLiteral("1") : QStringLiteral("0");
    // Header cells 10pt (1000), data cells 9pt (900); override with customSizePt100.
    const QString szStr = (customSizePt100 > 0)
                          ? QString::number(customSizePt100)
                          : (isHeader ? QStringLiteral("1000") : QStringLiteral("900"));

    // Split on newlines → one <a:p> per line
    QStringList lines = text.split(QLatin1Char('\n'));
    QString paras;
    for (const QString& line : lines) {
        const QString safe = XmlBuilder::escapeXml(line);
        paras += QStringLiteral(
            R"(<a:p>)"
            R"(<a:pPr algn="ctr"/>)"
            R"(<a:r>)"
            R"(<a:rPr lang="en-US" sz="%1" b="%2" dirty="0">)"
            R"(<a:solidFill><a:srgbClr val="%3"/></a:solidFill>)"
            R"(</a:rPr>)"
            R"(<a:t>%4</a:t>)"
            R"(</a:r>)"
            R"(</a:p>)").arg(szStr, boldStr, textColorHex, safe);
    }

    return QStringLiteral(
        R"(<a:tc>)"
        R"(<a:txBody>)"
        R"(<a:bodyPr wrap="square"/>)"
        R"(<a:lstStyle/>)"
        R"(%1)"
        R"(</a:txBody>)"
        R"(<a:tcPr anchor="ctr">)"
        R"(<a:solidFill><a:srgbClr val="%2"/></a:solidFill>)"
        R"(<a:lnL w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnL>)"
        R"(<a:lnR w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnR>)"
        R"(<a:lnT w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnT>)"
        R"(<a:lnB w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:lnB>)"
        R"(</a:tcPr>)"
        R"(</a:tc>)")
        .arg(paras, bgColorHex);
}

// ---------------------------------------------------------------------------
QString PptxWriter::buildTableXml(const SlideTable& table, int shapeId) const
{
    const int nCols = table.headers.size();
    if (nCols == 0) return QString();

    const QString hdrBg   = colorHex(table.headerBgColor);
    const QString altBg   = colorHex(table.altRowColor);
    const QString whiteBg = QStringLiteral("FFFFFF");

    // Column widths: use colWidthFractions if provided, else equal split.
    const long long totalEmu = toEmu(table.w);
    // Height computed from actual rows so the table always fits content.
    const long long kHeaderRowH = 274320LL;   // 0.30 in
    const long long kDataRowH   = 201168LL;   // 0.22 in
    const long long tableH = kHeaderRowH + (long long)table.rows.size() * kDataRowH;

    QVector<long long> colEmus;
    if (table.colWidthFractions.size() == nCols) {
        long long used = 0;
        for (int c = 0; c < nCols; ++c) {
            long long w = static_cast<long long>(totalEmu * table.colWidthFractions[c]);
            colEmus.append(w);
            used += w;
        }
        // Distribute any rounding remainder to last column
        colEmus.last() += (totalEmu - used);
    } else {
        const long long colEmu = (nCols > 0) ? totalEmu / nCols : totalEmu;
        for (int c = 0; c < nCols; ++c)
            colEmus.append(colEmu);
    }

    // Grid columns
    QString gridCols;
    gridCols.reserve(nCols * 30);
    for (long long cw : colEmus)
        gridCols += QStringLiteral(R"(<a:gridCol w="%1"/>)").arg(cw);

    // Per-cell font sizes. 0 in SlideTable::fontPt means "use legacy default"
    // (10 pt header / 9 pt body via makeTableCell's 0 sentinel). When set,
    // we apply the same value to both header and body cells; the Notes column
    // keeps its 8 pt override for compactness even under a body-font override.
    const int bodyFontPt100   = (table.fontPt > 0 ? table.fontPt : 0) * 100;
    const int headerFontPt100 = bodyFontPt100;  // headers track body by default

    // Header row
    QString headerRow;
    headerRow.reserve(nCols * 400 + 40);
    headerRow += QStringLiteral(R"(<a:tr h="%1">)").arg(kHeaderRowH);
    for (const QString& hdr : table.headers)
        headerRow += makeTableCell(hdr, true, hdrBg, QStringLiteral("FFFFFF"),
                                    headerFontPt100);
    headerRow += QStringLiteral("</a:tr>");

    // Data rows — pre-allocate to avoid repeated reallocation
    QString dataRows;
    dataRows.reserve(table.rows.size() * nCols * 400);
    for (int r = 0; r < table.rows.size(); ++r) {
        const QString& rowBg = (r % 2 == 0) ? whiteBg : altBg;
        dataRows += QStringLiteral(R"(<a:tr h="%1">)").arg(kDataRowH);
        const QStringList& row = table.rows[r];
        for (int c = 0; c < nCols; ++c) {
            const QString cellText = (c < row.size()) ? row[c] : QString();
            // Notes column: 8pt font (800 hundredths of a point) preserved
            // even when the rest of the table uses a custom size.
            const bool isNotes = (c < table.headers.size() &&
                                  table.headers[c].contains(
                                      QStringLiteral("Note"), Qt::CaseInsensitive));
            dataRows += makeTableCell(cellText, false, rowBg,
                                      QStringLiteral("000000"),
                                      isNotes ? 800 : bodyFontPt100);
        }
        dataRows += QStringLiteral("</a:tr>");
    }

    // Wrap in graphicFrame
    return QStringLiteral(
        R"(<p:graphicFrame>)"
        R"(<p:nvGraphicFramePr>)"
        R"(<p:cNvPr id="%8" name="Table %8"/>)"
        R"(<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>)"
        R"(<p:nvPr/>)"
        R"(</p:nvGraphicFramePr>)"
        R"(<p:xfrm><a:off x="%1" y="%2"/><a:ext cx="%3" cy="%4"/></p:xfrm>)"
        R"(<a:graphic>)"
        R"(<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">)"
        R"(<a:tbl>)"
        R"(<a:tblPr firstRow="1" bandRow="1">)"
        R"(<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>)"
        R"(</a:tblPr>)"
        R"(<a:tblGrid>%5</a:tblGrid>)"
        R"(%6%7)"
        R"(</a:tbl>)"
        R"(</a:graphicData>)"
        R"(</a:graphic>)"
        R"(</p:graphicFrame>)")
        .arg(emu(table.x), emu(table.y),
             QString::number(totalEmu), QString::number(tableH),
             gridCols, headerRow, dataRows, QString::number(shapeId));
}

// ============================================================================
// Slide XML builders
// ============================================================================

// Common slide wrapper (background + spTree open/close).
static QString slideHeader(const QString& bgRid)
{
    return QStringLiteral(
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        R"(<p:sld)"
        R"( xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main")"
        R"( xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main")"
        R"( xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">)"
        R"(<p:cSld>)"
        R"(<p:bg><p:bgPr>)"
        R"(<a:blipFill><a:blip r:embed="%1"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>)"
        R"(<a:effectLst/>)"
        R"(</p:bgPr></p:bg>)"
        R"(<p:spTree>)"
        R"(<p:nvGrpSpPr>)"
        R"(<p:cNvPr id="1" name=""/>)"
        R"(<p:cNvGrpSpPr/>)"
        R"(<p:nvPr/>)"
        R"(</p:nvGrpSpPr>)"
        R"(<p:grpSpPr>)"
        R"(<a:xfrm>)"
        R"(<a:off x="0" y="0"/>)"
        R"(<a:ext cx="0" cy="0"/>)"
        R"(<a:chOff x="0" y="0"/>)"
        R"(<a:chExt cx="0" cy="0"/>)"
        R"(</a:xfrm>)"
        R"(</p:grpSpPr>)").arg(bgRid);
}

static QString slideFooter()
{
    return QStringLiteral(
        R"(</p:spTree>)"
        R"(</p:cSld>)"
        R"(<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>)"
        R"(</p:sld>)");
}

// ---------------------------------------------------------------------------
// Cover slide
// ---------------------------------------------------------------------------
QString PptxWriter::buildCoverSlideXml(const QString& title,
                                        const QString& date,
                                        const QString& bgRid,
                                        const QString& logoRid,
                                        int titleFontPt,
                                        int dateFontPt,
                                        const QRectF& titleRect,
                                        const QRectF& subtitleRect) const
{
    QString shapes;
    int id = 2;

    // Title: Montserrat 46pt bold white, centred (or caller's override).
    // makeTextBox takes hundredths-of-a-point.
    const int titleSz100 = (titleFontPt > 0 ? titleFontPt : ReportLayout::kPptxCoverTitlePt) * 100;
    // Bug 1: honor a non-null layout rect; else legacy hardcoded box.
    double tX = 0.5, tY = 2.0, tW = 12.3, tH = 1.5;
    if (!titleRect.isNull()) {
        tX = titleRect.x(); tY = titleRect.y();
        tW = titleRect.width(); tH = titleRect.height();
    }
    shapes += makeTextBox(id++,
                          tX, tY, tW, tH,
                          title,
                          QStringLiteral("Montserrat"),
                          titleSz100, true, QStringLiteral("FFFFFF"),
                          QStringLiteral("ctr"));

    // Date: Montserrat 24pt white, centred (omitted when date is empty).
    if (!date.isEmpty()) {
        const int dateSz100 = (dateFontPt > 0 ? dateFontPt : 24) * 100;
        double dX = 0.5, dY = 4.6, dW = 12.3, dH = 0.7;
        if (!subtitleRect.isNull()) {
            dX = subtitleRect.x(); dY = subtitleRect.y();
            dW = subtitleRect.width(); dH = subtitleRect.height();
        }
        shapes += makeTextBox(id++,
                              dX, dY, dW, dH,
                              date,
                              QStringLiteral("Montserrat"),
                              dateSz100, false, QStringLiteral("FFFFFF"),
                              QStringLiteral("ctr"));
    }

    // White logo top-right — fixed aspect ratio 1.22" × 0.4"
    shapes += makePic(id++,
                      11.9, 0.1, 1.22, 0.4,
                      logoRid);

    return slideHeader(bgRid) + shapes + slideFooter();
}

// ---------------------------------------------------------------------------
// Content slide
// ---------------------------------------------------------------------------
QString PptxWriter::buildContentSlideXml(const QString&          title,
                                          const SlideTable&       table,
                                          const QVector<SlideImage>& plots,
                                          const QString&          bgRid,
                                          const QString&          logoRid,
                                          const QMap<QString, QString>& plotRids,
                                          const QString&          extraShapesXml,
                                          int                     titleFontPt,
                                          const QRectF&           titleRect) const
{
    QString shapes;
    int id = 2;

    // Sheet title: Montserrat 32pt by default, #1F497D. The 0 sentinel comes
    // through addContentSlide's 5-arg overload when the caller's layout has
    // titleFontPt == 0 (i.e. user hasn't customized).
    const int titleSz100 = (titleFontPt > 0 ? titleFontPt : ReportLayout::kPptxContentTitlePt) * 100;
    // Estimate height: ~38 chars/line at 32pt in 11", 0.50" per line. The
    // wrap-width assumption stays valid for reasonable user overrides (~20-40
    // pt); extreme values may overflow but the text box auto-expands at render.
    int titleLines = qMax(1, (title.length() + 37) / 38);
    double titleH = titleLines * 0.50;
    // Bug 1: a non-null layout rect (user-moved or canvas default) overrides
    // the legacy hardcoded title box, so the pptx matches the Preview canvas.
    double titleX = 0.4, titleY = 0.1, titleW = 11.0;
    if (!titleRect.isNull()) {
        titleX = titleRect.x();
        titleY = titleRect.y();
        titleW = titleRect.width();
        titleH = titleRect.height();
    }
    shapes += makeTextBox(id++,
                          titleX, titleY, titleW, titleH,
                          title,
                          QStringLiteral("Montserrat"),
                          titleSz100, true, QStringLiteral("1F497D"),
                          QStringLiteral("l"));

    // Logo top-right: ccell_logo_full.png — fixed aspect ratio 1.22" × 0.4"
    shapes += makePic(id++,
                      11.9, 0.1, 1.22, 0.4,
                      logoRid);

    // Data table
    if (!table.headers.isEmpty())
        shapes += buildTableXml(table, id++);

    // Plot images — use position/size from each SlideImage directly
    for (int i = 0; i < plots.size(); ++i) {
        const QString key = QStringLiteral("plot%1").arg(i);
        if (!plotRids.contains(key)) continue;

        const SlideImage& si = plots[i];
        shapes += makePic(id++, si.x, si.y,
                          (si.w > 0) ? si.w : 2.7,
                          (si.h > 0) ? si.h : 2.0,
                          plotRids[key]);
    }

    if (!extraShapesXml.isEmpty())
        shapes += extraShapesXml;

    return slideHeader(bgRid) + shapes + slideFooter();
}

// ---------------------------------------------------------------------------
// Image grid slide
// ---------------------------------------------------------------------------
QString PptxWriter::buildImageSlideXml(const QString& title,
                                        int imageCount,
                                        const QString& bgRid,
                                        const QString& logoRid,
                                        const QMap<int, QString>& imageRids) const
{
    QString shapes;
    int id = 2;

    // Title
    shapes += makeTextBox(id++,
                          0.4, 0.1, 11.0, 0.5,
                          title,
                          QStringLiteral("Montserrat"),
                          3200, true, QStringLiteral("1F497D"),
                          QStringLiteral("l"));

    // Logo
    shapes += makePic(id++,
                      11.9, 0.1, 1.22, 0.4,
                      logoRid);

    // 2-column × 3-row grid
    // Cell size: 4.2" × 3.0"; start: x=0.3", y=0.7"; gap: 0.15"
    const double cellW   = 4.2;
    const double cellH   = 3.0;
    const double startX  = 0.3;
    const double startY  = 0.7;
    const double gapX    = 0.15;
    const double gapY    = 0.1;

    int shown = 0;
    for (int i = 0; i < imageCount && i < 6; ++i) {
        if (!imageRids.contains(i)) continue;

        const int col  = shown % 2;
        const int row  = shown / 2;
        const double x = startX + col * (cellW + gapX);
        const double y = startY + row * (cellH + gapY);

        shapes += makePic(id++, x, y, cellW, cellH, imageRids[i]);
        ++shown;
    }

    return slideHeader(bgRid) + shapes + slideFooter();
}

// ---------------------------------------------------------------------------
QString PptxWriter::buildCustomImageSlideXml(
        const QString& title,
        const QString& bgRid,
        const QString& logoRid,
        const QVector<QPair<QString, QRectF>>& imageRids) const
{
    QString shapes;
    int id = 2;

    // Title
    shapes += makeTextBox(id++,
                          0.4, 0.1, 11.0, 0.5,
                          title,
                          QStringLiteral("Montserrat"),
                          3200, true, QStringLiteral("1F497D"),
                          QStringLiteral("l"));

    // Logo
    shapes += makePic(id++, 11.9, 0.1, 1.22, 0.4, logoRid);

    // Images at custom positions
    for (const auto& pair : imageRids) {
        const QRectF& r = pair.second;
        shapes += makePic(id++, r.x(), r.y(), r.width(), r.height(), pair.first);
    }

    return slideHeader(bgRid) + shapes + slideFooter();
}

} // namespace DVE
