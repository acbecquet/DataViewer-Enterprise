#pragma once
#include <QJsonObject>
#include <QRectF>
#include <QString>
#include <QHash>
#include <QVector>
#include <QStringList>

namespace DVE {

struct PropertiesBox {
    QRectF rect;
    QString text;
    int    fontPt = 0;        // 0 = use renderer's hardcoded default
};

struct ContentSlideLayout {
    QRectF title;
    int    titleFontPt = 0;   // 0 = renderer default (canvas 18, PPTX 32)
    QRectF table;
    int    tableFontPt = 0;   // 0 = renderer default (canvas 11, PPTX 11)
    QRectF radar;
    PropertiesBox propertiesBox;
};

struct ImageSlideLayout {
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
};

struct TableSort {
    QString column;            // empty = insertion order
    Qt::SortOrder order = Qt::DescendingOrder;
};

struct ReportLayout {
    static constexpr int kCurrentVersion = 1;
    int version = kCurrentVersion;
    QString modeId = QStringLiteral("sensory");

    TableSort tableSort;
    QHash<QString, ContentSlideLayout> contentSlides;   // key: "content_<sessionId>"
    QHash<QString, ImageSlideLayout>   imageSlides;     // key: "image_<sessionId>"
    QHash<QString, QRectF>             dividerTitles;   // key: "divider_<sessionId>"
    QHash<QString, int>                dividerTitleFontPts; // parallel; 0 = default
    ContentSlideLayout cumulative;
    QRectF coverTitle;
    int    coverTitleFontPt = 0;     // 0 = renderer default
    QRectF coverSubtitle;
    int    coverSubtitleFontPt = 0;
    QStringList zOrder;                                  // top-to-bottom ids per slide

    // Renderer-side default font sizes — consumed when a layout field is 0
    // ("not set"). Two distinct sets because the canvas is a small preview
    // (60 px/in) while the PPTX is full-scale: matching sizes for both would
    // make canvas text impractically large or PPTX text too small. Legacy
    // flow (empty ReportLayout from SensoryPanel::generateCombinedPptx) hits
    // the PPTX defaults, preserving the v1.0.x output exactly.
    static constexpr int kCanvasCoverTitlePt        = 28;
    static constexpr int kCanvasCoverSubtitlePt     = 16;
    static constexpr int kCanvasDividerTitlePt      = 32;
    static constexpr int kCanvasContentTitlePt      = 18;
    static constexpr int kCanvasPropertiesBoxPt     = 12;
    static constexpr int kCanvasTablePt             = 11;

    static constexpr int kPptxCoverTitlePt          = 44;   // DV-24: -2 pt
    static constexpr int kPptxCoverSubtitlePt       = 24;
    static constexpr int kPptxDividerTitlePt        = 44;   // DV-24: -2 pt (uses cover-slide xml)
    static constexpr int kPptxContentTitlePt        = 30;   // DV-24: -2 pt
    static constexpr int kPptxPropertiesBoxPt       = 16;
    static constexpr int kPptxTablePt               = 11;   // user-requested bump from legacy 9

    QJsonObject toJson() const;
    static ReportLayout fromJson(const QJsonObject& obj, bool* ok = nullptr);
    bool isEmpty() const;       // true if all rects are default-constructed
};

// Helper rect <-> JSON conversion. Exposed so other modules
// (e.g. SensoryReportSource cumulative-layout serialization) can reuse them.
QJsonArray rectToJsonArray(const QRectF&);
QRectF     rectFromJsonArray(const QJsonArray&);

} // namespace DVE
