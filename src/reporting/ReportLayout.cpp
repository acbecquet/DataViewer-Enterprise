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
        { "table",         rectToJsonArray(c.table) },
        { "radar",         rectToJsonArray(c.radar) },
        { "propertiesBox", QJsonObject{
            { "rect", rectToJsonArray(c.propertiesBox.rect) },
            { "text", c.propertiesBox.text }
        } }
    };
}

static ContentSlideLayout contentFromJson(const QJsonObject& o) {
    ContentSlideLayout c;
    c.title = rectFromJsonArray(o.value("title").toArray());
    c.table = rectFromJsonArray(o.value("table").toArray());
    c.radar = rectFromJsonArray(o.value("radar").toArray());
    const QJsonObject pb = o.value("propertiesBox").toObject();
    c.propertiesBox.rect = rectFromJsonArray(pb.value("rect").toArray());
    c.propertiesBox.text = pb.value("text").toString();
    return c;
}

QJsonObject ReportLayout::toJson() const {
    QJsonObject slides;
    slides["cover"] = QJsonObject{
        { "title",    rectToJsonArray(coverTitle) },
        { "subtitle", rectToJsonArray(coverSubtitle) }
    };
    for (auto it = dividerTitles.cbegin(); it != dividerTitles.cend(); ++it)
        slides[it.key()] = QJsonObject{ { "title", rectToJsonArray(it.value()) } };
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
            r.coverTitle    = rectFromJsonArray(v.value("title").toArray());
            r.coverSubtitle = rectFromJsonArray(v.value("subtitle").toArray());
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
