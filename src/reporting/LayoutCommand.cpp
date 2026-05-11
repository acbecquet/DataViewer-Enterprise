#include "LayoutCommand.h"

namespace DVE {

// --- RectCommand ----------------------------------------------------------
RectCommand::RectCommand(const QString& slideKey, const QString& elementId,
                         const QRectF& oldRect, const QRectF& newRect)
    : m_slideKey(slideKey), m_elementId(elementId),
      m_old(oldRect), m_new(newRect) {}

void RectCommand::apply(ReportLayout& l) { setRect(l, m_new); }
void RectCommand::undo (ReportLayout& l) { setRect(l, m_old); }

QString RectCommand::description() const {
    return QStringLiteral("Move/resize %1 on %2").arg(m_elementId, m_slideKey);
}

void RectCommand::setRect(ReportLayout& l, const QRectF& r) const {
    if (m_slideKey == QLatin1String("cover")) {
        if      (m_elementId == QLatin1String("title"))    l.coverTitle    = r;
        else if (m_elementId == QLatin1String("subtitle")) l.coverSubtitle = r;
    } else if (m_slideKey.startsWith(QLatin1String("divider_"))) {
        // Divider title rect lives in ReportLayout::dividerTitles
        // (QHash<QString,QRectF>) keyed by slide id ("divider_<sessionId>").
        if (m_elementId == QLatin1String("title"))
            l.dividerTitles[m_slideKey] = r;
    } else if (m_slideKey == QLatin1String("cumulative")) {
        if      (m_elementId == QLatin1String("title")) l.cumulative.title = r;
        else if (m_elementId == QLatin1String("table")) l.cumulative.table = r;
        else if (m_elementId == QLatin1String("radar")) l.cumulative.radar = r;
        else if (m_elementId == QLatin1String("propertiesBox"))
            l.cumulative.propertiesBox.rect = r;
    } else {
        // content_<sessionId>
        ContentSlideLayout cs = l.contentSlides.value(m_slideKey);
        if      (m_elementId == QLatin1String("title")) cs.title = r;
        else if (m_elementId == QLatin1String("table")) cs.table = r;
        else if (m_elementId == QLatin1String("radar")) cs.radar = r;
        else if (m_elementId == QLatin1String("propertiesBox"))
            cs.propertiesBox.rect = r;
        l.contentSlides[m_slideKey] = cs;
    }
}

// --- FontSizeCommand ------------------------------------------------------
FontSizeCommand::FontSizeCommand(const QString& slideKey, const QString& elementId,
                                 int oldPt, int newPt)
    : m_slideKey(slideKey), m_elementId(elementId), m_old(oldPt), m_new(newPt) {}

void FontSizeCommand::apply(ReportLayout& l) { setFontPt(l, m_new); }
void FontSizeCommand::undo (ReportLayout& l) { setFontPt(l, m_old); }

QString FontSizeCommand::description() const {
    return QStringLiteral("Font %1pt -> %2pt on %3").arg(m_old).arg(m_new).arg(m_elementId);
}

void FontSizeCommand::setFontPt(ReportLayout& l, int pt) const {
    if (m_slideKey == QLatin1String("cover")) {
        if      (m_elementId == QLatin1String("title"))    l.coverTitleFontPt    = pt;
        else if (m_elementId == QLatin1String("subtitle")) l.coverSubtitleFontPt = pt;
    } else if (m_slideKey.startsWith(QLatin1String("divider_"))) {
        if (m_elementId == QLatin1String("title"))
            l.dividerTitleFontPts[m_slideKey] = pt;
    } else if (m_slideKey == QLatin1String("cumulative")) {
        if      (m_elementId == QLatin1String("title"))         l.cumulative.titleFontPt          = pt;
        else if (m_elementId == QLatin1String("table"))         l.cumulative.tableFontPt          = pt;
        else if (m_elementId == QLatin1String("propertiesBox")) l.cumulative.propertiesBox.fontPt = pt;
    } else {
        ContentSlideLayout cs = l.contentSlides.value(m_slideKey);
        if      (m_elementId == QLatin1String("title"))         cs.titleFontPt          = pt;
        else if (m_elementId == QLatin1String("table"))         cs.tableFontPt          = pt;
        else if (m_elementId == QLatin1String("propertiesBox")) cs.propertiesBox.fontPt = pt;
        l.contentSlides[m_slideKey] = cs;
    }
}

// --- SortCommand ----------------------------------------------------------
SortCommand::SortCommand(const QString& oldC, Qt::SortOrder oldO,
                         const QString& newC, Qt::SortOrder newO)
    : m_oldCol(oldC), m_newCol(newC), m_oldOrd(oldO), m_newOrd(newO) {}

void SortCommand::apply(ReportLayout& l) {
    l.tableSort.column = m_newCol;
    l.tableSort.order  = m_newOrd;
}
void SortCommand::undo(ReportLayout& l) {
    l.tableSort.column = m_oldCol;
    l.tableSort.order  = m_oldOrd;
}

} // namespace DVE
