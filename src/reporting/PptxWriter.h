#pragma once
#include "../pipeline/ReportData.h"
#include "../utils/ZipWriter.h"
#include <QString>
#include <QStringList>
#include <QVector>
#include <QByteArray>
#include <QMap>
#include <QPair>
#include <QRectF>

namespace DVE {

// ---------------------------------------------------------------------------
// SlideImage – a positioned image on a content slide (inches).
// ---------------------------------------------------------------------------
struct SlideImage {
    QByteArray pngData;
    double x = 0.0;
    double y = 0.0;
    double w = 2.7;
    double h = 2.0;
};

// ---------------------------------------------------------------------------
// SlideTable – table data plus layout hints.
// ---------------------------------------------------------------------------
struct SlideTable {
    QStringList          headers;
    QVector<QStringList> rows;
    double x = 0.5;
    double y = 0.64;
    double w = 12.3;   // total width (inches); column widths split evenly
    double h = 4.5;    // max height hint (not enforced in XML)
    int headerBgColor = 0x1F4E79;  // dark blue
    int altRowColor   = 0xE8F4FD;  // light blue

    // Optional: fractional width per column (values should sum to 1.0).
    // Empty = equal widths.
    QVector<double> colWidthFractions;
};

// ---------------------------------------------------------------------------
// PptxWriter – builds a PowerPoint (.pptx) file matching the Python output.
//
// Typical usage:
//   PptxWriter w;
//   w.setResourcePath("/path/to/resources");
//   w.addCoverSlide("My Report", "2024-01-15");
//   w.addContentSlide("Sheet 1", table, plots);
//   w.save("/output/report.pptx");
// ---------------------------------------------------------------------------
class PptxWriter {
public:
    PptxWriter();
    ~PptxWriter();

    // Directory containing Cover_Page_Logo.jpg, ccell_background.png,
    // ccell_logo_full.png, ccell_logo_full_white.png.
    void setResourcePath(const QString& resourceDir);

    // Add a full-background cover slide.
    void addCoverSlide(const QString& title, const QString& dateStr);

    // Add a slide with a data table and optional plot images below.
    void addContentSlide(const QString& sheetTitle,
                         const SlideTable& table,
                         const QVector<SlideImage>& plots,
                         const QString& extraShapesXml = QString());

    // Add a slide with two tables (stacked vertically on the left) and plot images.
    void addDualTableSlide(const QString& sheetTitle,
                           const SlideTable& topTable,
                           const SlideTable& bottomTable,
                           const QVector<SlideImage>& plots);

    // Add a slide with a grid of embedded photos (up to 6 per slide, 2×3).
    void addImageSlide(const QString& sheetTitle,
                       const QVector<QByteArray>& images);

    // Add a slide with custom-positioned images (positions in inches).
    void addImageSlide(const QString& sheetTitle,
                       const QVector<SlideImage>& images);

    // Add a conclusions slide with a title and one large empty textbox.
    void addConclusionsSlide();

    // Add a Test Protocol slide with title and the provided table.
    void addTestProtocolSlide(const SlideTable& sopTable);

    // Add a Test Overview slide with title, description, bullet list of test names,
    // and an empty trailing textbox for manual performance summary.
    void addTestOverviewSlide(const QString& description,
                              const QStringList& testNames);

    // Add a cover-style section divider slide showing the filename centred on
    // both axes (no date).
    void addSectionDividerSlide(const QString& filename);

    // Serialise to a .pptx file.  Returns false on error.
    bool    save(const QString& filePath);
    QString lastError() const { return m_lastError; }

private:
    // -----------------------------------------------------------------------
    // Internal slide record: the slide XML + any media files referenced.
    // -----------------------------------------------------------------------
    struct Slide {
        QString xml;
        // zip-path → raw bytes for all images embedded in this slide.
        QVector<QPair<QString, QByteArray>> media;
    };

    QVector<Slide> m_slides;
    QString        m_resourceDir;
    mutable QString m_lastError;
    int            m_mediaCounter = 0;   // global counter → unique media names

    // Cached resource images (loaded lazily).
    mutable QMap<QString, QByteArray> m_resourceCache;

    // -----------------------------------------------------------------------
    // Static OOXML skeleton generators
    // -----------------------------------------------------------------------
    QString buildPresentationXml()    const;
    QString buildContentTypesXml()    const;
    QString buildRootRelsXml()        const;
    QString buildPresRelsXml()        const;
    QString buildSlideMasterXml()     const;
    QString buildSlideMasterRelsXml() const;
    QString buildSlideLayoutXml()     const;
    QString buildSlideLayoutRelsXml() const;
    QString buildThemeXml()           const;
    QString buildPresPropsXml()       const;
    QString buildViewPropsXml()       const;
    QString buildTableStylesXml()     const;

    // Per-slide relationship file.
    // media is the list of (zipPath, data) pairs added for that slide.
    QString buildSlideRelsXml(const QVector<QPair<QString, QByteArray>>& media) const;

    // -----------------------------------------------------------------------
    // Slide XML builders
    // -----------------------------------------------------------------------
    QString buildCoverSlideXml(const QString& title,
                               const QString& date,
                               const QString& bgRid,
                               const QString& logoRid) const;

    QString buildContentSlideXml(const QString& title,
                                 const SlideTable& table,
                                 const QVector<SlideImage>& plots,
                                 const QString& bgRid,
                                 const QString& logoRid,
                                 const QMap<QString, QString>& plotRids,
                                 const QString& extraShapesXml = QString()) const;

    QString buildImageSlideXml(const QString& title,
                               int imageCount,
                               const QString& bgRid,
                               const QString& logoRid,
                               const QMap<int, QString>& imageRids) const;

    QString buildCustomImageSlideXml(const QString& title,
                                     const QString& bgRid,
                                     const QString& logoRid,
                                     const QVector<QPair<QString, QRectF>>& imageRids) const;

    // -----------------------------------------------------------------------
    // Shape XML helpers
    // -----------------------------------------------------------------------
    // Text box (auto-sizing body).
    QString makeTextBox(int id,
                        double xIn, double yIn,
                        double wIn, double hIn,
                        const QString& text,
                        const QString& font,
                        int    sizePt100,      // hundredths of a point  (3200 = 32pt)
                        bool   bold,
                        const QString& colorHex,
                        const QString& alignHoriz = QStringLiteral("l")) const;

    // Inline image (p:pic).
    QString makePic(int id,
                    double xIn, double yIn,
                    double wIn, double hIn,
                    const QString& rId) const;

    // DrawingML table (p:graphicFrame).
    QString buildTableXml(const SlideTable& table, int shapeId) const;

    // Single table cell XML.
    // customSizePt100: override font size in hundredths of a pt (0 = auto).
    QString makeTableCell(const QString& text,
                          bool  isHeader,
                          const QString& bgColorHex,
                          const QString& textColorHex,
                          int   customSizePt100 = 0) const;

    // -----------------------------------------------------------------------
    // Utilities
    // -----------------------------------------------------------------------
    // inches → EMU string.
    static QString emu(double inches);
    static long long toEmu(double inches);

    // Integer colour → "RRGGBB" hex string (no leading #).
    static QString colorHex(int color);

    // Load a resource image by filename; returns PNG/JPEG bytes.
    QByteArray loadResourceImage(const QString& filename) const;

    // Register a media blob in this slide's media list; return its rId string.
    // ext should be "png" or "jpg".
    QString addMedia(const QByteArray& data,
                     const QString&    ext,
                     QVector<QPair<QString, QByteArray>>& media);
};

} // namespace DVE
