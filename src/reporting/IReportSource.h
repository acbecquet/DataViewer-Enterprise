#pragma once
#include <QString>
#include <QStringList>
#include <QVector>
#include <QSet>
#include <QImage>
#include <QRectF>
#include "ReportLayout.h"

namespace DVE {

enum class SlideKind { Cover, Divider, Content, Image, Cumulative };

struct SampleRef {
    QString slideKey;     // "content_<id>" -- the slide this sample belongs to
    QString sampleId;     // unique within slideKey, used in excludedSamples set
    QString displayName;
    QString sessionLabel; // shown as group header in checkbox panel
};

struct ReportSlideSpec {
    SlideKind kind;
    QString   slideKey;       // "cover" | "divider_<id>" | "content_<id>" | etc.
    QString   title;          // current title text (editable)
    // Resolved geometry + content for the canvas, after layout overrides
    // applied. Field-by-kind meaningfulness:
    //   Cover, Divider, Cumulative: title + layout
    //   Content:                    title, table*, radarPixmap, propertiesText, layout
    //   Image:                      imagePaths, imageLayouts, imageCrops
    QImage    radarPixmap;
    QStringList tableHeaders;
    QVector<QStringList> tableRows;
    QString   propertiesText;
    QStringList imagePaths;
    QVector<QRectF> imageLayouts;
    QVector<QRectF> imageCrops;
    ContentSlideLayout layout;     // resolved geometry
};

class IReportSource {
public:
    virtual ~IReportSource() = default;

    // Polymorphic interface — concrete adapters own non-trivial state
    // (sessions, DB handles). Disable copy/move at the base to prevent
    // accidental slicing. Adapters can re-enable if they need to.
    IReportSource() = default;
    IReportSource(const IReportSource&) = delete;
    IReportSource& operator=(const IReportSource&) = delete;
    IReportSource(IReportSource&&) = delete;
    IReportSource& operator=(IReportSource&&) = delete;

    // Identity
    virtual QString modeId() const = 0;          // "sensory" (only mode in v1)
    virtual QString sourceLabel() const = 0;     // shown in dialog title bar

    // Slides
    virtual int slideCount() const = 0;
    virtual SlideKind slideKind(int idx) const = 0;
    virtual ReportSlideSpec buildSlide(int idx, const ReportLayout&,
                                       const QSet<QString>& excludedSamples) const = 0;

    // Sample enumeration (for the checkbox panel)
    virtual QVector<SampleRef> allSamples() const = 0;  // {sessionId, sampleId, displayName}

    // Persistence
    virtual ReportLayout loadLayout() const = 0;
    virtual void saveLayout(const ReportLayout&) = 0;

    // Final PPTX write
    [[nodiscard]] virtual bool writePptx(const QString& outPath, const ReportLayout&,
                                          const QSet<QString>& excludedSamples,
                                          QString* errorOut) = 0;
};

} // namespace DVE
