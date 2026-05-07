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
};

struct ContentSlideLayout {
    QRectF title;
    QRectF table;
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
    ContentSlideLayout cumulative;
    QRectF coverTitle;
    QRectF coverSubtitle;
    QStringList zOrder;                                  // top-to-bottom ids per slide

    QJsonObject toJson() const;
    static ReportLayout fromJson(const QJsonObject& obj, bool* ok = nullptr);
    bool isEmpty() const;       // true if all rects are default-constructed
};

// Helper rect <-> JSON conversion. Exposed so other modules
// (e.g. SensoryReportSource cumulative-layout serialization) can reuse them.
QJsonArray rectToJsonArray(const QRectF&);
QRectF     rectFromJsonArray(const QJsonArray&);

} // namespace DVE
