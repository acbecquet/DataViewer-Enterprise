#include "ReportLayout.h"
#include <QJsonArray>

namespace DVE {

QJsonArray rectToJsonArray(const QRectF& r) {
    return QJsonArray{ r.x(), r.y(), r.width(), r.height() };
}

QRectF rectFromJsonArray(const QJsonArray& a) {
    if (a.size() != 4) return {};
    return { a[0].toDouble(), a[1].toDouble(), a[2].toDouble(), a[3].toDouble() };
}

static QJsonObject contentToJson(const ContentSlideLayout& c) {
    return {
        { "title",         rectToJsonArray(c.title) },
        { "titleFontPt",   c.titleFontPt },
        { "table",         rectToJsonArray(c.table) },
        { "tableFontPt",   c.tableFontPt },
        { "radar",         rectToJsonArray(c.radar) },
        { "propertiesBox", QJsonObject{
            { "rect",   rectToJsonArray(c.propertiesBox.rect) },
            { "text",   c.propertiesBox.text },
            { "fontPt", c.propertiesBox.fontPt }
        } }
    };
}

static ContentSlideLayout contentFromJson(const QJsonObject& o) {
    // toInt(default) returns `default` when the value is missing or non-numeric,
    // giving us free backward-compatibility: old JSON without fontPt fields
    // loads with the struct's hardcoded defaults (which themselves match the
    // legacy v1.0.x canvas sizes).
    ContentSlideLayout c;
    c.title       = rectFromJsonArray(o.value("title").toArray());
    c.titleFontPt = o.value("titleFontPt").toInt(c.titleFontPt);
    c.table       = rectFromJsonArray(o.value("table").toArray());
    c.tableFontPt = o.value("tableFontPt").toInt(c.tableFontPt);
    c.radar       = rectFromJsonArray(o.value("radar").toArray());
    const QJsonObject pb = o.value("propertiesBox").toObject();
    c.propertiesBox.rect   = rectFromJsonArray(pb.value("rect").toArray());
    c.propertiesBox.text   = pb.value("text").toString();
    c.propertiesBox.fontPt = pb.value("fontPt").toInt(c.propertiesBox.fontPt);
    return c;
}

QJsonObject ReportLayout::toJson() const {
    QJsonObject slides;
    slides["cover"] = QJsonObject{
        { "title",            rectToJsonArray(coverTitle) },
        { "titleFontPt",      coverTitleFontPt },
        { "subtitle",         rectToJsonArray(coverSubtitle) },
        { "subtitleFontPt",   coverSubtitleFontPt }
    };
    for (auto it = dividerTitles.cbegin(); it != dividerTitles.cend(); ++it) {
        // Sentinel 0 = "use renderer default" — keep that in the JSON so
        // round-trip is symmetric and the renderer can decide.
        const int fontPt = dividerTitleFontPts.value(it.key(), 0);
        slides[it.key()] = QJsonObject{
            { "title",       rectToJsonArray(it.value()) },
            { "titleFontPt", fontPt }
        };
    }
    for (auto it = contentSlides.cbegin(); it != contentSlides.cend(); ++it)
        slides[it.key()] = contentToJson(it.value());
    for (auto it = imageSlides.cbegin(); it != imageSlides.cend(); ++it) {
        QJsonArray layouts, crops;
        for (const QRectF& r : it.value().imageLayouts) layouts.append(rectToJsonArray(r));
        for (const QRectF& r : it.value().imageCrops)   crops.append(rectToJsonArray(r));
        slides[it.key()] = QJsonObject{
            { "imageLayouts", layouts },
            { "imageCrops",   crops }
        };
    }
    slides["cumulative"] = contentToJson(cumulative);

    QJsonArray zOrderArr;
    for (const QString& s : zOrder) zOrderArr.append(s);

    return {
        { "version", version },
        { "mode",    modeId },
        { "tableSort", QJsonObject{
            { "column", tableSort.column },
            { "order",  tableSort.order == Qt::AscendingOrder ? "asc" : "desc" }
        } },
        { "slides",  slides },
        { "zOrder",  zOrderArr }
    };
}

ReportLayout ReportLayout::fromJson(const QJsonObject& obj, bool* ok) {
    ReportLayout r;
    if (ok) *ok = false;
    if (!obj.contains("version") || !obj.contains("mode")) return r;
    r.version = obj.value("version").toInt(kCurrentVersion);
    r.modeId  = obj.value("mode").toString();
    if (r.modeId != "sensory") return r;            // v1 only handles sensory

    const QJsonObject ts = obj.value("tableSort").toObject();
    r.tableSort.column = ts.value("column").toString();
    r.tableSort.order  = ts.value("order").toString() == "asc"
                         ? Qt::AscendingOrder : Qt::DescendingOrder;

    const QJsonObject slides = obj.value("slides").toObject();
    for (auto it = slides.begin(); it != slides.end(); ++it) {
        const QString& key = it.key();
        const QJsonObject v = it.value().toObject();
        if (key == "cover") {
            r.coverTitle          = rectFromJsonArray(v.value("title").toArray());
            r.coverTitleFontPt    = v.value("titleFontPt").toInt(r.coverTitleFontPt);
            r.coverSubtitle       = rectFromJsonArray(v.value("subtitle").toArray());
            r.coverSubtitleFontPt = v.value("subtitleFontPt").toInt(r.coverSubtitleFontPt);
        } else if (key == "cumulative") {
            r.cumulative = contentFromJson(v);
        } else if (key.startsWith("content_")) {
            r.contentSlides[key] = contentFromJson(v);
        } else if (key.startsWith("image_")) {
            ImageSlideLayout img;
            for (const QJsonValue& jv : v.value("imageLayouts").toArray())
                img.imageLayouts.append(rectFromJsonArray(jv.toArray()));
            for (const QJsonValue& jv : v.value("imageCrops").toArray())
                img.imageCrops.append(rectFromJsonArray(jv.toArray()));
            r.imageSlides[key] = img;
        } else if (key.startsWith("divider_")) {
            r.dividerTitles[key] = rectFromJsonArray(v.value("title").toArray());
            // Missing field → 0 sentinel → renderer uses its hardcoded default.
            r.dividerTitleFontPts[key] = v.value("titleFontPt").toInt(0);
        }
    }

    for (const QJsonValue& jv : obj.value("zOrder").toArray())
        r.zOrder.append(jv.toString());

    if (ok) *ok = true;
    return r;
}

bool ReportLayout::isEmpty() const {
    return contentSlides.isEmpty() && imageSlides.isEmpty()
        && dividerTitles.isEmpty() && coverTitle.isNull();
}

} // namespace DVE
